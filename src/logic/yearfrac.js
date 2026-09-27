/**
 * Year fraction for lease terms.
 * A date that falls on the same month and day N years later is exactly N years
 * (actual/actual). Partial years use the actual days in the year that contains them.
 * Actual/365 is not used: leap days would turn the specified 10-year and 5-year
 * anniversary cases into values that no longer display as 10.00 and 5.00.
 */

function utcDay(date) {
  return Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
}

export function asDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === 'number' && Number.isFinite(value) && value > 20000 && value < 80000) {
    var utc = new Date(Math.round((value - 25569) * 86400 * 1000));
    return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
  }
  return null;
}

function addYears(date, years) {
  var next = new Date(date.getFullYear() + years, date.getMonth(), date.getDate());
  if (next.getMonth() !== date.getMonth()) {
    next = new Date(date.getFullYear() + years, date.getMonth() + 1, 0);
  }
  return next;
}

export function yearFrac(start, end) {
  var s = asDate(start);
  var e = asDate(end);
  if (!s || !e) return null;
  if (utcDay(e) <= utcDay(s)) return 0;
  var years = e.getFullYear() - s.getFullYear();
  var anniversary = addYears(s, years);
  if (utcDay(anniversary) > utcDay(e)) {
    years -= 1;
    anniversary = addYears(s, years);
  }
  var next = addYears(anniversary, 1);
  var remaining = (utcDay(e) - utcDay(anniversary)) / 86400000;
  var span = (utcDay(next) - utcDay(anniversary)) / 86400000;
  if (!span) return years;
  return years + remaining / span;
}
