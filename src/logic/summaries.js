import { resolveBreakEnd, waultToBreak, waultToExpiry } from './wault.js';

function numeric(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.trim()) {
    var parsed = Number(value.replace(/,/g, '').trim());
    if (Number.isFinite(parsed)) return parsed;
  }
  return null;
}

function sumProperties(properties, columnId) {
  var total = 0;
  var any = false;
  for (var i = 0; i < properties.length; i += 1) {
    var value = numeric(properties[i].standardValues && properties[i].standardValues[columnId]);
    if (value === null) continue;
    total += value;
    any = true;
  }
  return any ? total : null;
}

export function sumMatchedUnitValues(properties, columnId) {
  var total = 0;
  var any = false;
  for (var i = 0; i < (properties || []).length; i += 1) {
    var units = properties[i].units || [];
    for (var u = 0; u < units.length; u += 1) {
      var value = numeric(units[u].standardValues && units[u].standardValues[columnId]);
      if (value === null) continue;
      total += value;
      any = true;
    }
  }
  return any ? total : null;
}

function recordsFor(properties, rule) {
  var records = [];
  var minimum = rule.eligibility && rule.eligibility.rentGreaterThan != null ? rule.eligibility.rentGreaterThan : 0;
  for (var p = 0; p < properties.length; p += 1) {
    var property = properties[p];
    var valuation = property.standardValues ? property.standardValues[rule.start.field] : null;
    var units = property.units || [];
    for (var u = 0; u < units.length; u += 1) {
      var values = units[u].standardValues || {};
      records.push({
        valuation: valuation,
        expiry: rule.end ? values[rule.end.id] : values[rule.fallbackEnd.id],
        breakDate: rule.preferredEnd ? values[rule.preferredEnd.id] : null,
        rent: values[rule.weight.id],
        rentGreaterThan: minimum,
        requireValidExpiry: rule.measure === 'unexpiredLeaseTerm',
      });
    }
  }
  return records;
}

function divideValues(numerator, denominator) {
  if (numerator === null || denominator === null || denominator === 0) return null;
  return numerator / denominator;
}

export function evaluateRule(rule, properties, columnId, sectorName) {
  if (!rule || rule.method === 'blank' || (rule.method === 'summary' && rule.summary === 'blank')) return null;
  if (rule.method === 'label') {
    if (rule.label === 'grandTotal') return 'GRAND TOTAL';
    return (sectorName || 'Sector') + ' TOTAL';
  }
  if (rule.method === 'fixed') return rule.value;
  if (rule.method === 'summary' && rule.summary === 'sum') {
    var sumId = (rule.source && rule.source.id) || columnId;
    if (rule.scope === 'matchedUnits') return sumMatchedUnitValues(properties, sumId);
    return sumProperties(properties, columnId);
  }
  if (rule.method === 'summary' && rule.summary === 'recalculate') {
    return divideValues(
      sumProperties(properties, rule.numerator.id),
      sumProperties(properties, rule.denominator.id)
    );
  }
  if (rule.method === 'derived' && rule.calculation === 'divide') {
    return divideValues(
      sumProperties(properties, rule.numerator.id),
      sumProperties(properties, rule.denominator.id)
    );
  }
  if (rule.method === 'derived' || (rule.method === 'summary' && rule.summary === 'weightedAverage')) {
    var derived = rule.method === 'derived' ? rule : rule.calculation;
    var records = recordsFor(properties, derived);
    if (derived.measure === 'unexpiredLeaseTermToBreak') return waultToBreak(records);
    return waultToExpiry(records);
  }
  return null;
}

export function applyPropertyDerived(property, columns) {
  var values = property.standardValues;
  for (var i = 0; i < columns.length; i += 1) {
    var column = columns[i];
    var rule = column.logic && column.logic.property;
    if (!rule || rule.method !== 'derived') continue;
    if (rule.calculation === 'divide') {
      values[column.id] = divideValues(numeric(values[rule.numerator.id]), numeric(values[rule.denominator.id]));
      continue;
    }
    values[column.id] = evaluateRule(rule, [property], column.id, null);
  }
}

export function applyPropertyContext(property, columns) {
  var values = property.standardValues;
  for (var i = 0; i < columns.length; i += 1) {
    var column = columns[i];
    var rule = column.logic && column.logic.property;
    if (!rule || rule.method !== 'summary' || rule.summary !== 'sum' || rule.scope !== 'matchedUnits') continue;
    values[column.id] = sumMatchedUnitValues([property], (rule.source && rule.source.id) || column.id);
  }
  applyPropertyDerived(property, columns);
}

export function summaryValues(properties, columns, context, sectorName) {
  var values = {};
  for (var i = 0; i < columns.length; i += 1) {
    var column = columns[i];
    var rule = column.logic ? column.logic[context] : null;
    values[column.id] = evaluateRule(rule, properties, column.id, sectorName);
  }
  return values;
}

export { resolveBreakEnd };
