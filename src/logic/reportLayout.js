import { DEFAULT_SCHEDULE_COLUMNS } from '../templates/defaultSchedule.js';
import { comparisonColumnActive, comparisonHeading, directionHighlight } from './comparison.js';

var sessionOptionalIds = null;
var sessionComparison = {};

function isActiveOptional(column) {
  return !!column && column.role === 'optional' && column.rendered !== false && column.feature !== 'comparison';
}

export function defaultOptionalIds() {
  var ids = [];
  for (var i = 0; i < DEFAULT_SCHEDULE_COLUMNS.length; i += 1) {
    if (isActiveOptional(DEFAULT_SCHEDULE_COLUMNS[i])) ids.push(DEFAULT_SCHEDULE_COLUMNS[i].id);
  }
  return ids;
}

export function getReportLayout() {
  if (!sessionOptionalIds) sessionOptionalIds = defaultOptionalIds();
  return sessionOptionalIds.slice();
}

export function setReportLayout(ids, comparison) {
  var allowed = {};
  var defaults = defaultOptionalIds();
  for (var i = 0; i < defaults.length; i += 1) allowed[defaults[i]] = true;
  var next = [];
  var incoming = ids || [];
  for (var j = 0; j < incoming.length; j += 1) {
    if (allowed[incoming[j]] && next.indexOf(incoming[j]) === -1) next.push(incoming[j]);
  }
  sessionOptionalIds = next;
  if (comparison) {
    sessionComparison = {};
    Object.keys(comparison).forEach(function (id) {
      if (allowed[id]) sessionComparison[id] = !!comparison[id];
    });
  }
  return getReportLayout();
}

export function resetReportLayout() {
  sessionOptionalIds = defaultOptionalIds();
  sessionComparison = {};
  return getReportLayout();
}

export function isComparisonEnabled(column) {
  if (!column || !column.comparison || !column.comparison.supported) return false;
  if (Object.prototype.hasOwnProperty.call(sessionComparison, column.id)) return !!sessionComparison[column.id];
  return !!column.comparison.defaultEnabled;
}

export function getComparisonSelection() {
  var selection = {};
  var columns = listOptionalColumns();
  for (var i = 0; i < columns.length; i += 1) {
    if (!columns[i].comparison || !columns[i].comparison.supported) continue;
    selection[columns[i].id] = isComparisonEnabled(columns[i]);
  }
  return selection;
}

function outputEntry(column, label, extras) {
  var extra = extras || {};
  return {
    id: extra.id || column.id,
    label: label,
    width: column.width,
    format: extra.format != null ? extra.format : column.format,
    align: column.align,
    hiddenInExcel: !!column.hiddenInExcel,
    role: column.role,
    parentId: extra.parentId || column.id,
    comparisonRole: extra.comparisonRole || '',
    directionHighlight: !!extra.directionHighlight,
  };
}

function appendOptional(selected, column, labels, options) {
  var baseLabel = column.label;
  var active = comparisonColumnActive(column, isComparisonEnabled(column), options || {});
  if (!active) {
    selected.push(outputEntry(column, labels[column.id] || baseLabel));
    return;
  }
  var currentPeriod = options.currentPeriod || '';
  var previousPeriod = options.previousDatesDiffer ? '' : (options.previousPeriod || '');
  selected.push(outputEntry(column, comparisonHeading(baseLabel, currentPeriod, false), {
    comparisonRole: 'current',
    parentId: column.id,
    directionHighlight: directionHighlight(column.format),
  }));
  selected.push(outputEntry(column, comparisonHeading(baseLabel, previousPeriod, !!options.previousDatesDiffer || !previousPeriod), {
    id: column.id + '__previous',
    comparisonRole: 'previous',
    parentId: column.id,
  }));
  if (column.comparison.showPercentChange) {
    selected.push(outputEntry(column, 'Change %', {
      id: column.id + '__change',
      format: '0.00%',
      comparisonRole: 'change',
      parentId: column.id,
    }));
  }
}

export function buildOutputColumns(headingById, options) {
  var labels = headingById || {};
  var byId = {};
  var systems = [];
  var required = [];
  for (var i = 0; i < DEFAULT_SCHEDULE_COLUMNS.length; i += 1) {
    var column = DEFAULT_SCHEDULE_COLUMNS[i];
    byId[column.id] = column;
    if (column.rendered === false || column.feature === 'comparison') continue;
    if (column.role === 'system') systems.push(outputEntry(column, labels[column.id] || column.label));
    else if (column.role === 'required') required.push(outputEntry(column, labels[column.id] || column.label));
  }
  var selected = [];
  var order = getReportLayout();
  for (var j = 0; j < order.length; j += 1) {
    var chosen = byId[order[j]];
    if (!isActiveOptional(chosen)) continue;
    appendOptional(selected, chosen, labels, options);
  }
  return systems.concat(required, selected);
}

export function listOptionalColumns() {
  var columns = [];
  for (var i = 0; i < DEFAULT_SCHEDULE_COLUMNS.length; i += 1) {
    if (isActiveOptional(DEFAULT_SCHEDULE_COLUMNS[i])) columns.push(DEFAULT_SCHEDULE_COLUMNS[i]);
  }
  return columns;
}

export function listFixedColumns(role) {
  var columns = [];
  for (var i = 0; i < DEFAULT_SCHEDULE_COLUMNS.length; i += 1) {
    if (DEFAULT_SCHEDULE_COLUMNS[i].role === role && DEFAULT_SCHEDULE_COLUMNS[i].rendered !== false) {
      columns.push(DEFAULT_SCHEDULE_COLUMNS[i]);
    }
  }
  return columns;
}
