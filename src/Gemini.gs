/**
 * ============================================================================
 * Gemini.gs — Receipt data extraction via the Gemini API
 * ----------------------------------------------------------------------------
 * Created:  2026-10-30
 * Author:   Allen McIntosh (with Claude)
 *
 * Sends the receipt image/PDF inline to Gemini with a strict JSON response
 * schema, so the reply is machine-readable fields rather than free text.
 * The category list comes from the Categories tab and is passed as an enum,
 * so Gemini can only answer with one of your categories.
 * ============================================================================
 */

/**
 * Error type that tells the processor whether to leave the file in the Inbox
 * and retry on the next run (true) or route it to Needs Attention (false).
 */
class ExtractionError extends Error {
  /**
   * @param {string} message
   * @param {boolean} retryable - leave file in Inbox and try again next run
   * @param {boolean=} fatal    - stop processing the rest of the Inbox this run
   */
  constructor(message, retryable, fatal) {
    super(message);
    this.name = 'ExtractionError';
    this.retryable = !!retryable;
    this.fatal = !!fatal;
  }
}

/**
 * Extracts structured fields from a receipt file.
 * @param {Blob} blob - receipt image or PDF
 * @param {{name: string, description: string}[]} categories
 * @return {Object} parsed receipt fields (see buildReceiptSchema_)
 * @throws {ExtractionError}
 */
function extractReceiptData_(blob, categories) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_API_KEY);
  if (!apiKey) throw new ExtractionError('Gemini API key is not set.', true, true);

  const model = getSetting_('Gemini Model') || CONFIG.DEFAULT_GEMINI_MODEL;
  const url = CONFIG.GEMINI_ENDPOINT + encodeURIComponent(model) + ':generateContent';

  const payload = {
    contents: [{
      role: 'user',
      parts: [
        { text: buildPrompt_(categories) },
        { inlineData: { mimeType: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) } }
      ]
    }],
    generationConfig: {
      temperature: 0,
      responseMimeType: 'application/json',
      responseSchema: buildReceiptSchema_(categories.map(c => c.name))
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },   // header, so the key never appears in a URL/log
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let lastError = '';
  for (let attempt = 1; attempt <= CONFIG.GEMINI_MAX_ATTEMPTS; attempt++) {
    const resp = UrlFetchApp.fetch(url, options);
    const code = resp.getResponseCode();
    const body = resp.getContentText();

    if (code === 200) return parseGeminiResponse_(body);

    lastError = `HTTP ${code}: ${body.substring(0, 300)}`;
    // 429 (rate limit) and 5xx are transient: back off and retry
    if (code === 429 || code >= 500) {
      Utilities.sleep(Math.pow(2, attempt) * 1000);
      continue;
    }
    // Bad key / bad model name / no permission are setup problems, not file
    // problems: keep the file in the Inbox and stop this run (fatal).
    const configProblem = code === 401 || code === 403 || code === 404 ||
      /API_KEY_INVALID|API key not valid|not found for API version|is not supported/i.test(body);
    if (configProblem) {
      throw new ExtractionError('Gemini setup problem (check API key / Gemini Model setting). ' + lastError, true, true);
    }
    // Anything else (e.g., 400 on a corrupt file) is specific to this file
    throw new ExtractionError('Gemini request failed. ' + lastError, false);
  }
  throw new ExtractionError('Gemini unavailable after retries. ' + lastError, true, true);
}

/**
 * Builds the instruction text sent alongside the receipt.
 * @param {{name: string, description: string}[]} categories
 * @return {string}
 */
function buildPrompt_(categories) {
  const catList = categories.map(c => `- ${c.name}: ${c.description || ''}`).join('\n');
  return [
    'You are extracting data from a purchase receipt, invoice, or order confirmation for an expense tracker.',
    'Return only what is printed on the document. Never guess a number; use null if a value is missing or unreadable.',
    '',
    'Rules:',
    '- receipt_date: the transaction/purchase date in YYYY-MM-DD format.',
    '- total: the final amount charged, including tax, tips, and fees.',
    '- subtotal and tax: as printed; null if not shown.',
    '- payment_last4: last 4 digits of the card if printed, else null. Never return more than 4 digits.',
    '- vendor: the merchant or store name as a customer would recognize it (e.g., "Shell", "Amazon", "Best Buy").',
    '- For fuel purchases, fill gallons, price_per_gallon, and fuel_grade if printed.',
    '- items_summary: one short line listing what was bought (e.g., "2x Cat6 patch cable 10ft, USB-C hub").',
    '- is_receipt: false if the document is not a receipt/invoice at all.',
    '- confidence: "low" if the image is blurry, cut off, faded, or totals are ambiguous.',
    '- warnings: brief notes on anything a human should double-check, or an empty string.',
    '',
    'Choose the single best category from this list:',
    catList
  ].join('\n');
}

/**
 * JSON schema for the structured response (OpenAPI subset used by Gemini).
 * @param {string[]} categoryNames
 * @return {Object}
 */
function buildReceiptSchema_(categoryNames) {
  const num = { type: 'NUMBER', nullable: true };
  const str = { type: 'STRING', nullable: true };
  return {
    type: 'OBJECT',
    properties: {
      is_receipt: { type: 'BOOLEAN' },
      vendor: str,
      receipt_date: { type: 'STRING', nullable: true, description: 'YYYY-MM-DD' },
      subtotal: num,
      tax: num,
      total: num,
      currency: str,
      payment_last4: str,
      category: { type: 'STRING', enum: categoryNames },
      gallons: num,
      price_per_gallon: num,
      fuel_grade: str,
      items_summary: str,
      confidence: { type: 'STRING', enum: ['high', 'medium', 'low'] },
      warnings: str
    },
    required: ['is_receipt', 'vendor', 'receipt_date', 'total', 'category', 'confidence'],
    propertyOrdering: [
      'is_receipt', 'vendor', 'receipt_date', 'subtotal', 'tax', 'total', 'currency',
      'payment_last4', 'category', 'gallons', 'price_per_gallon', 'fuel_grade',
      'items_summary', 'confidence', 'warnings'
    ]
  };
}

/**
 * Pulls the JSON object out of a Gemini generateContent response.
 * @param {string} body - raw HTTP response text
 * @return {Object}
 * @throws {ExtractionError}
 */
function parseGeminiResponse_(body) {
  let json;
  try {
    json = JSON.parse(body);
  } catch (e) {
    throw new ExtractionError('Gemini returned non-JSON response.', true);
  }

  const cand = json.candidates && json.candidates[0];
  const parts = cand && cand.content && cand.content.parts;
  if (!parts || !parts.length) {
    const reason = (cand && cand.finishReason) ||
                   (json.promptFeedback && json.promptFeedback.blockReason) || 'unknown';
    throw new ExtractionError('Gemini returned no content (reason: ' + reason + ').', false);
  }

  // Join text parts and strip any accidental ```json fences
  const text = parts.map(p => p.text || '').join('').replace(/^\s*```(?:json)?|```\s*$/g, '').trim();
  try {
    return JSON.parse(text);
  } catch (e) {
    throw new ExtractionError('Could not parse extracted fields: ' + text.substring(0, 200), false);
  }
}
