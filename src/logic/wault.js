import { asDate, yearFrac } from './yearfrac.js';

function positiveRent(value, minimum) {
  var floor = minimum == null ? 0 : minimum;
  var number = null;
  if (typeof value === 'number' && Number.isFinite(value)) number = value;
  else if (typeof value === 'string' && value.trim()) {
    var parsed = Number(value.replace(/,/g, '').trim());
    if (Number.isFinite(parsed)) number = parsed;
  }
  if (number === null || !(number > floor)) return null;
  return number;
}

/**
 * Future effective break on or before lease expiry, otherwise lease expiry.
 * A break on or before the valuation date falls back to expiry.
 */
export function resolveBreakEnd(valuation, effectiveBreak, leaseExpiry) {
  var val = asDate(valuation);
  var brk = asDate(effectiveBreak);
  var exp = asDate(leaseExpiry);
  if (val && brk && utcAfter(brk, val)) {
    if (!exp || !utcAfter(brk, exp)) return brk;
  }
  return exp;
}

function utcAfter(a, b) {
  return Date.UTC(a.getFullYear(), a.getMonth(), a.getDate()) > Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
}

function weighted(records, yearsFor) {
  var numerator = 0;
  var denominator = 0;
  var used = 0;
  for (var i = 0; i < records.length; i += 1) {
    var record = records[i];
    var rent = positiveRent(record.rent, record.rentGreaterThan == null ? 0 : record.rentGreaterThan);
    if (rent === null) continue;
    var years = yearsFor(record);
    if (years === null) continue;
    numerator += years * rent;
    denominator += rent;
    used += 1;
  }
  if (!used || denominator === 0) return null;
  return numerator / denominator;
}

export function waultToExpiry(records) {
  return weighted(records, function (record) {
    if (record.requireValidExpiry !== false && !asDate(record.expiry)) return null;
    var years = yearFrac(record.valuation, record.expiry);
    if (years === null) return null;
    return Math.max(0, years);
  });
}

export function waultToBreak(records) {
  return weighted(records, function (record) {
    var end = resolveBreakEnd(record.valuation, record.breakDate, record.expiry);
    if (!asDate(end)) return null;
    var years = yearFrac(record.valuation, end);
    if (years === null) return null;
    return Math.max(0, years);
  });
}
