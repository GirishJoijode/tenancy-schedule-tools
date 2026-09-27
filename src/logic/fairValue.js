import { asDate } from './yearfrac.js';

var MONTH_NAMES = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];

export var MULTIPLE_VALUATION_DATES_WARNING =
  'Multiple valuation dates detected. Fair Value heading cannot be resolved to a single portfolio date.';

export function formatFairValueHeading(value) {
  var date = asDate(value);
  if (!date) return 'Fair Value';
  return 'Fair Value ' + MONTH_NAMES[date.getMonth()] + ' ' + date.getFullYear();
}

export function resolvePortfolioValuation(dates) {
  var unique = [];
  var resolved = [];
  for (var i = 0; i < (dates || []).length; i += 1) {
    var date = asDate(dates[i]);
    if (!date) continue;
    var key = date.getFullYear() + '-' + date.getMonth() + '-' + date.getDate();
    if (unique.indexOf(key) !== -1) continue;
    unique.push(key);
    resolved.push(date);
  }
  if (resolved.length === 0) {
    return { date: null, consistent: true, heading: 'Fair Value' };
  }
  if (resolved.length > 1) {
    return { date: null, consistent: false, heading: 'Fair Value' };
  }
  return {
    date: resolved[0],
    consistent: true,
    heading: formatFairValueHeading(resolved[0]),
  };
}
