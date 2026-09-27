import { normaliseHeader } from './text.js';

var MONTHS = {
  jan: 1,
  feb: 2,
  mar: 3,
  apr: 4,
  may: 5,
  jun: 6,
  jul: 7,
  aug: 8,
  sep: 9,
  oct: 10,
  nov: 11,
  dec: 12,
};

var MONTH_LABELS = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function makePeriod(year, month) {
  return {
    year: year,
    month: month,
    label: MONTH_LABELS[month] + '-' + year,
  };
}

function periodFromDate(year, month) {
  if (!Number.isFinite(year) || !Number.isFinite(month)) return null;
  if (month < 1 || month > 12 || year < 1900 || year > 2500) return null;
  return makePeriod(year, month);
}

export function parsePeriod(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return periodFromDate(value.getFullYear(), value.getMonth() + 1);
  }

  if (typeof value === 'number' && Number.isFinite(value) && value > 20000 && value < 80000) {
    var serial = new Date(Math.round((value - 25569) * 86400 * 1000));
    return periodFromDate(serial.getUTCFullYear(), serial.getUTCMonth() + 1);
  }

  if (value === null || value === undefined) return null;
  var text = String(value).replace(/\u00a0/g, ' ').trim();
  var match = text.match(/^([A-Za-z]{3})-(\d{4})$/);
  if (!match) return null;
  var month = MONTHS[match[1].toLowerCase()];
  if (!month) return null;
  return periodFromDate(Number(match[2]), month);
}

export function isPeriodField(value) {
  return parsePeriod(value) !== null;
}

export function periodLabelFromHeader(value) {
  var period = parsePeriod(value);
  return period ? period.label : null;
}

export function getPeriodColumns(sourceColumns, section) {
  var wanted = normaliseHeader(section);
  var found = [];
  var columns = sourceColumns || [];
  for (var i = 0; i < columns.length; i += 1) {
    var column = columns[i];
    if (normaliseHeader(column.section) !== wanted) continue;
    var period = parsePeriod(column.field);
    if (!period) continue;
    found.push({
      field: period.label,
      year: period.year,
      month: period.month,
      columnIndex: column.columnIndex,
      column: column,
    });
  }
  return found;
}

export function getPeriodFields(options) {
  var request = options || {};
  return getPeriodColumns(request.columns, request.section);
}

export function sortPeriodsChronologically(periods) {
  return (periods || []).slice().sort(function (a, b) {
    if (a.year !== b.year) return a.year - b.year;
    return a.month - b.month;
  });
}

export function selectPeriods(periods, selection) {
  var list = periods || [];
  if (!list.length) return [];
  if (selection === 'firstPeriod') return [list[0]];
  if (selection === 'lastPeriod') return [list[list.length - 1]];
  if (selection === 'earliestPeriod') return [sortPeriodsChronologically(list)[0]];
  if (selection === 'latestPeriod') return [sortPeriodsChronologically(list)[list.length - 1]];
  return list.slice();
}

function numericPeriodValue(raw) {
  if (raw === null || raw === undefined || raw === '') return null;
  if (typeof raw === 'number' && Number.isFinite(raw)) return raw;
  if (typeof raw === 'string' && raw.trim()) {
    var parsed = Number(raw.replace(/,/g, '').trim());
    if (Number.isFinite(parsed)) return parsed;
  }
  return null;
}

export function periodValues(parsedRow, periods) {
  var numbers = [];
  for (var i = 0; i < periods.length; i += 1) {
    var column = periods[i].column;
    var raw = parsedRow && parsedRow.values ? parsedRow.values[column.key] : null;
    var number = numericPeriodValue(raw);
    if (number !== null) numbers.push(number);
  }
  return numbers;
}

export function aggregatePeriodValues(numbers, aggregation) {
  if (!numbers.length) return null;
  if (aggregation === 'max') return Math.max.apply(null, numbers);
  if (aggregation === 'min') return Math.min.apply(null, numbers);
  if (aggregation === 'sum') {
    return numbers.reduce(function (total, value) {
      return total + value;
    }, 0);
  }
  if (aggregation === 'average') {
    return (
      numbers.reduce(function (total, value) {
        return total + value;
      }, 0) / numbers.length
    );
  }
  return numbers.length === 1 ? numbers[0] : null;
}

export function readPeriodGroupValue(parsedRow, columns, ref) {
  if (!ref) return null;
  var selected = selectPeriods(getPeriodColumns(columns, ref.section), ref.selection || 'allPeriods');
  return aggregatePeriodValues(periodValues(parsedRow, selected), ref.aggregation || null);
}

export function describePeriodGroup(columns, section) {
  var periods = getPeriodColumns(columns, section);
  var lines = [
    section,
    periods.length + ' period column' + (periods.length === 1 ? '' : 's') + ' detected:',
  ];
  for (var i = 0; i < periods.length; i += 1) lines.push(periods[i].field);
  return lines.join('\n');
}
