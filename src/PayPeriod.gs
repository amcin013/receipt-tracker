/**
 * ============================================================================
 * PayPeriod.gs — Biweekly pay period math for Work reimbursements
 * ----------------------------------------------------------------------------
 * Created:  2026-11-02
 * Author:   Allen McIntosh (with Claude)
 *
 * Pay periods are fixed-length blocks (default 14 days) counted forward and
 * backward from one known period start on the Settings tab:
 *
 *   Pay Period Anchor Start   2026-10-18   (Sun Oct 18 - Sat Oct 31, 2026)
 *   Pay Period Length (days)  14
 *
 * A period's label matches how payroll names it, e.g. "OCT 18- OCT 31".
 * The reimbursement email subject leads with that label.
 *
 * All date math uses calendar-day arithmetic (new Date(y, m, d + n)) rather
 * than adding milliseconds, so daylight-saving changes can't shift a day.
 * ============================================================================
 */

const PAY_PERIOD_DEFAULT_ANCHOR = '2026-10-18';
const PAY_PERIOD_DEFAULT_LENGTH = 14;

/**
 * Returns the pay period containing a date.
 * @param {Date} date
 * @return {{start: Date, end: Date, label: string, key: string}}
 */
function payPeriodFor_(date) {
  const anchor = payPeriodAnchor_();
  const len = payPeriodLength_();
  const day = dayNumber_(date) - dayNumber_(anchor);
  const index = Math.floor(day / len);          // negative before the anchor
  const start = addDays_(anchor, index * len);
  const end = addDays_(start, len - 1);
  return { start: start, end: end, label: payPeriodLabel_(start, end), key: dayNumber_(start) + '' };
}

/**
 * Returns the most recent pay period that has fully ended before today.
 * On Mon Nov 16, 2026 that is Nov 1 - Nov 14.
 * @param {Date=} today - defaults to now
 * @return {{start: Date, end: Date, label: string, key: string}}
 */
function lastCompletedPayPeriod_(today) {
  const current = payPeriodFor_(today || new Date());
  return payPeriodFor_(addDays_(current.start, -1));
}

/**
 * Builds the payroll-style label, e.g. "OCT 18- OCT 31".
 * @param {Date} start
 * @param {Date} end
 * @return {string}
 */
function payPeriodLabel_(start, end) {
  const tz = Session.getScriptTimeZone();
  const f = d => Utilities.formatDate(d, tz, 'MMM d').toUpperCase();
  return `${f(start)}- ${f(end)}`;
}

/**
 * True if a date falls inside a pay period (inclusive).
 * @param {Date} date
 * @param {{start: Date, end: Date}} period
 * @return {boolean}
 */
function inPayPeriod_(date, period) {
  const n = dayNumber_(date);
  return n >= dayNumber_(period.start) && n <= dayNumber_(period.end);
}

/* ============================================================================
 * Settings + date helpers
 * ========================================================================== */

/**
 * Reads the anchor start date from Settings. Accepts a real date cell or
 * YYYY-MM-DD text; falls back to the default if blank or unreadable.
 * @return {Date}
 */
function payPeriodAnchor_() {
  const raw = getSettingRaw_('Pay Period Anchor Start');
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return new Date(raw.getFullYear(), raw.getMonth(), raw.getDate());
  }
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(raw || '').trim()) ||
            /^(\d{4})-(\d{2})-(\d{2})/.exec(PAY_PERIOD_DEFAULT_ANCHOR);
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/**
 * Reads the period length in days from Settings (default 14).
 * @return {number}
 */
function payPeriodLength_() {
  const n = Math.round(Number(getSetting_('Pay Period Length (days)')));
  return n > 0 ? n : PAY_PERIOD_DEFAULT_LENGTH;
}

/**
 * Whole-day number for a date's calendar day (ignores time and DST).
 * @param {Date} d
 * @return {number}
 */
function dayNumber_(d) {
  return Math.round(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000);
}

/**
 * Returns a new local-midnight date n calendar days after d.
 * @param {Date} d
 * @param {number} n
 * @return {Date}
 */
function addDays_(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}
