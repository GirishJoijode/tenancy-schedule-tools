import { DEFAULT_SCHEDULE_COLUMNS } from '../templates/defaultSchedule.js';

export function normaliseComparisonPart(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(/\u00a0/g, ' ').trim().toLowerCase();
}

export function comparisonKey(propertyReference, unitName) {
  var propertyKey = normaliseComparisonPart(propertyReference);
  var unitKey = normaliseComparisonPart(unitName);
  if (!propertyKey || !unitKey) return '';
  return propertyKey + '||' + unitKey;
}

export function comparableColumns() {
  return DEFAULT_SCHEDULE_COLUMNS.filter(function (column) {
    return column.comparison && column.comparison.supported && column.rendered !== false && column.feature !== 'comparison';
  });
}

function ruleHasValue(rule) {
  return !!(rule && rule.method !== 'blank' && rule.method !== 'label');
}

export function comparisonColumnActive(column, enabled, sources) {
  if (!enabled || !column || !column.comparison || !column.comparison.supported) return false;
  var files = sources || {};
  if (column.comparison.propertyOnly) return !!files.previousProperty;
  var propertySide = ruleHasValue(column.logic && column.logic.property) && !!files.previousProperty;
  var unitSide = ruleHasValue(column.logic && column.logic.unit) && !!files.previousProperty && !!files.previousUnit;
  return propertySide || unitSide;
}

export function changeRatio(current, previous) {
  var left = numeric(current);
  var right = numeric(previous);
  if (left === null || right === null || right === 0) return null;
  return (left / right) - 1;
}

export function comparisonTone(current, previous) {
  var left = numeric(current);
  var right = numeric(previous);
  if (left === null || right === null || right === 0) return '';
  if (left === right) return '';
  return left > right ? 'up' : 'down';
}

export function changeTone(ratio) {
  var value = numeric(ratio);
  if (value === null || value === 0) return '';
  return value > 0 ? 'up' : 'down';
}

export function roundedChangePercentagePoints(ratio) {
  var value = numeric(ratio);
  if (value === null) return null;
  var scaled = Number((value * 10000).toFixed(8));
  var sign = scaled < 0 ? -1 : 1;
  var rounded = sign * Math.round(Math.abs(scaled));
  if (rounded === 0) return 0;
  return rounded / 100;
}

export function fairValueChangeTone(ratio) {
  var rounded = roundedChangePercentagePoints(ratio);
  if (rounded === null || rounded === 0) return '';
  return rounded > 0 ? 'up' : 'down';
}

function numeric(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.trim() && value.trim() !== '-') {
    var parsed = Number(value.replace(/,/g, '').trim());
    if (Number.isFinite(parsed)) return parsed;
  }
  return null;
}

export function indexPreviousPortfolio(portfolio) {
  var byProperty = new Map();
  var byUnit = new Map();
  var duplicatePropertyRefs = [];
  var properties = (portfolio && portfolio.properties) || [];
  for (var i = 0; i < properties.length; i += 1) {
    var property = properties[i];
    var ref = normaliseComparisonPart(property.standardValues && property.standardValues.property_reference);
    if (!ref) continue;
    if (byProperty.has(ref)) {
      if (duplicatePropertyRefs.indexOf(ref) === -1) duplicatePropertyRefs.push(ref);
      continue;
    }
    byProperty.set(ref, property);
  }
  duplicatePropertyRefs.forEach(function (ref) { byProperty.delete(ref); });
  byProperty.forEach(function (property) {
    var ref = property.standardValues && property.standardValues.property_reference;
    var units = property.units || [];
    for (var u = 0; u < units.length; u += 1) {
      var key = comparisonKey(ref, units[u].standardValues && units[u].standardValues.unit);
      if (!key || byUnit.has(key)) continue;
      byUnit.set(key, units[u]);
    }
  });
  return {
    byProperty: byProperty,
    byUnit: byUnit,
    duplicatePropertyRefs: duplicatePropertyRefs,
  };
}

export function attachPreviousMatches(currentProperties, index) {
  var lookup = index || { byProperty: new Map(), byUnit: new Map() };
  for (var i = 0; i < currentProperties.length; i += 1) {
    var property = currentProperties[i];
    var ref = property.standardValues && property.standardValues.property_reference;
    property.previousProperty = lookup.byProperty.get(normaliseComparisonPart(ref)) || null;
    var units = property.units || [];
    for (var u = 0; u < units.length; u += 1) {
      var key = comparisonKey(ref, units[u].standardValues && units[u].standardValues.unit);
      units[u].previousUnit = key ? lookup.byUnit.get(key) || null : null;
    }
  }
}

function previousValue(record, fieldId) {
  if (!record || !record.standardValues) return null;
  var value = record.standardValues[fieldId];
  if (value === undefined || value === null || value === '') return null;
  return value;
}

export function withRowComparison(kind, values, previousRecord) {
  var next = Object.assign({}, values || {});
  var fields = comparableColumns();
  for (var i = 0; i < fields.length; i += 1) {
    var field = fields[i];
    var blankUnit = kind === 'unit' && field.comparison.propertyOnly;
    var previous = blankUnit ? null : previousValue(previousRecord, field.id);
    next[field.id + '__previous'] = previous;
    if (field.comparison.showPercentChange) {
      next[field.id + '__change'] = blankUnit ? null : changeRatio(next[field.id], previous);
    }
  }
  return next;
}

export function overlaySummaryComparison(currentValues, previousValues) {
  var next = Object.assign({}, currentValues || {});
  var fields = comparableColumns();
  for (var i = 0; i < fields.length; i += 1) {
    var field = fields[i];
    var previous = previousValues ? previousValues[field.id] : null;
    if (previous === undefined || previous === '') previous = null;
    next[field.id + '__previous'] = previous;
    if (field.comparison.showPercentChange) next[field.id + '__change'] = changeRatio(next[field.id], previous);
  }
  return next;
}

export function directionHighlight(format) {
  if (!format) return false;
  if (String(format).indexOf('yy') !== -1) return false;
  return /[#0]/.test(String(format));
}

export function comparisonHeading(baseLabel, period, genericPrevious) {
  if (genericPrevious) return baseLabel + '\nPrevious Valuation';
  if (!period) return baseLabel;
  return baseLabel + '\n' + period;
}
