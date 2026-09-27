import { findSourceColumn } from '../schema/sourceLookup.js';

function push(issues, column, context, message) {
  issues.push({
    columnId: column.id,
    label: column.label,
    context: context,
    message: message,
  });
}

function checkRef(issues, column, context, rule) {
  var ref = rule.ref;
  if (!ref) {
    push(issues, column, context, 'Source reference is missing.');
    return;
  }
  if (ref.type === 'periodGroup') {
    if (!ref.section) push(issues, column, context, 'Period group section is missing.');
    if (ref.selection === 'allPeriods' && !ref.aggregation) {
      push(issues, column, context, 'All Period Columns requires an aggregation.');
    }
    return;
  }
  if (!ref.section || !ref.field) push(issues, column, context, 'Source field needs a section and field.');
}

function checkDerived(issues, column, context, rule) {
  if (rule.calculation === 'divide') {
    if (!rule.numerator || !rule.numerator.id) push(issues, column, context, 'Derived calculation is missing a numerator.');
    if (!rule.denominator || !rule.denominator.id) push(issues, column, context, 'Derived calculation is missing a denominator.');
    return;
  }
  if (!rule.measure) push(issues, column, context, 'Derived calculation measure is missing.');
  if (!rule.start || !rule.start.field) push(issues, column, context, 'WAULT start date is missing.');
  if (!rule.weight || !rule.weight.id) push(issues, column, context, 'WAULT weight is missing.');
  if (rule.measure === 'unexpiredLeaseTerm' && (!rule.end || !rule.end.id)) {
    push(issues, column, context, 'WAULT to expiry needs a Lease Expiry input.');
  }
  if (rule.measure === 'unexpiredLeaseTermToBreak') {
    if (!rule.preferredEnd || !rule.preferredEnd.id) push(issues, column, context, 'WAULT to break needs an Effective Break input.');
    if (!rule.fallbackEnd || !rule.fallbackEnd.id) push(issues, column, context, 'WAULT to break needs a Lease Expiry fallback.');
  }
}

function visitSummary(issues, columnsById, column, context, rule, stack) {
  if (!rule || rule.method !== 'summary' || rule.summary !== 'recalculate') return;
  if (stack.indexOf(column.id) !== -1) {
    push(issues, column, context, 'Circular summary dependency: ' + stack.concat(column.id).join(' → '));
    return;
  }
  var next = stack.concat(column.id);
  [rule.numerator, rule.denominator].forEach(function (input) {
    if (!input || !input.id) {
      push(issues, column, context, 'Recalculated summary is missing an input.');
      return;
    }
    var other = columnsById[input.id];
    if (!other) {
      push(issues, column, context, 'Summary input is not a template field: ' + input.id);
      return;
    }
    var otherRule = other.logic && other.logic[context];
    visitSummary(issues, columnsById, other, context, otherRule, next);
  });
}

export function validateTemplateLogic(columns, contexts) {
  var issues = [];
  var columnsById = {};
  for (var i = 0; i < columns.length; i += 1) columnsById[columns[i].id] = columns[i];
  var seen = {};

  for (var c = 0; c < columns.length; c += 1) {
    var column = columns[c];
    if (seen[column.id]) push(issues, column, 'property', 'Duplicate output field id.');
    seen[column.id] = true;
    if (!column.logic) {
      push(issues, column, 'property', 'Field has no logic definition.');
      continue;
    }
    for (var k = 0; k < contexts.length; k += 1) {
      var context = contexts[k];
      var rule = column.logic[context];
      if (!rule) {
        push(issues, column, context, 'Missing logic for this context.');
        continue;
      }
      if (rule.method === 'source' || rule.method === 'aggregatedSource') checkRef(issues, column, context, rule);
      if (rule.method === 'derived') checkDerived(issues, column, context, rule);
      if (rule.method === 'summary') visitSummary(issues, columnsById, column, context, rule, []);
      if (rule.method === 'fixed' && (rule.value === undefined || rule.value === '')) {
        push(issues, column, context, 'Fixed value is empty.');
      }
    }
  }
  return issues;
}

export function validateLoadedSources(columns, contexts, propertyColumns, unitColumns) {
  var issues = [];
  if (!propertyColumns && !unitColumns) return issues;
  for (var c = 0; c < columns.length; c += 1) {
    var column = columns[c];
    for (var k = 0; k < contexts.length; k += 1) {
      var context = contexts[k];
      var rule = column.logic && column.logic[context];
      if (!rule || (rule.method !== 'source' && rule.method !== 'aggregatedSource')) continue;
      var ref = rule.ref;
      if (!ref) continue;
      var dataset = ref.dataset || (context === 'unit' ? 'unit' : 'property');
      var sourceColumns = dataset === 'unit' ? unitColumns : propertyColumns;
      if (!sourceColumns) continue;
      if (ref.type === 'periodGroup') {
        var wanted = String(ref.section || '').toLowerCase();
        var found = sourceColumns.some(function (sourceColumn) {
          return String(sourceColumn.section || '').toLowerCase() === wanted;
        });
        if (!found) {
          issues.push({
            columnId: column.id,
            label: column.label,
            context: context,
            message: 'Loaded ' + dataset + ' workbook has no source group: ' + ref.section,
          });
        }
        continue;
      }
      if (!findSourceColumn(sourceColumns, dataset, ref)) {
        issues.push({
          columnId: column.id,
          label: column.label,
          context: context,
          message: 'Loaded ' + dataset + ' workbook is missing ' + ref.section + ' > ' + ref.field,
        });
      }
    }
  }
  return issues;
}
