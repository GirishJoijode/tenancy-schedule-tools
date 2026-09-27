var METHOD_LABELS = {
  source: 'Source Field',
  derived: 'Derived Calculation',
  aggregatedSource: 'Aggregated Source',
  summary: 'Summary Aggregation',
  label: 'Summary Aggregation',
  fixed: 'Fixed Value',
  blank: 'Blank',
};

var DATASET_LABELS = {
  property: 'Property',
  unit: 'Unit',
};

var PERIOD_PHRASE = {
  allPeriods: 'all periods',
  earliestPeriod: 'earliest period',
  latestPeriod: 'latest period',
  firstPeriod: 'first period column',
  lastPeriod: 'last period column',
};

function datasetOf(rule, context) {
  var ref = rule && rule.ref;
  if (ref && ref.dataset) return ref.dataset;
  return context === 'unit' ? 'unit' : 'property';
}

function datasetLabel(dataset) {
  return DATASET_LABELS[dataset] || dataset || 'Property';
}

export function fieldLabel(columns, id) {
  var list = columns || [];
  for (var i = 0; i < list.length; i += 1) {
    if (list[i].id === id) return list[i].label;
  }
  return id || '';
}

function aggregationToken(ref) {
  if (!ref) return 'VALUE';
  if (ref.aggregation === 'max') return 'MAX';
  if (ref.aggregation === 'min') return 'MIN';
  if (ref.aggregation === 'sum') return 'SUM';
  if (ref.aggregation === 'average') return 'AVERAGE';
  if (ref.selection === 'earliestPeriod') return 'EARLIEST';
  if (ref.selection === 'latestPeriod') return 'LATEST';
  if (ref.selection === 'firstPeriod') return 'FIRST';
  if (ref.selection === 'lastPeriod') return 'LAST';
  return 'VALUE';
}

export function sourcePath(rule, context) {
  var ref = (rule && rule.ref) || {};
  var dataset = datasetLabel(datasetOf(rule, context));
  if (ref.type === 'periodGroup') return dataset + ' › ' + (ref.section || 'Period group');
  if (!ref.section && !ref.field) return 'Source not set';
  return dataset + ' › ' + (ref.section || '') + ' › ' + (ref.field || '');
}

export function aggregatedPreview(rule, context) {
  var ref = (rule && rule.ref) || {};
  var phrase = PERIOD_PHRASE[ref.selection] || 'all periods';
  return aggregationToken(ref) + '(' + sourcePath(rule, context) + ' › ' + phrase + ')';
}

function operationSymbol(operation) {
  if (operation === 'add') return '+';
  if (operation === 'subtract') return '−';
  if (operation === 'multiply') return '×';
  return '÷';
}

export function scopeLine(rule) {
  if (!rule) return 'Matched Units';
  if (rule.scopeNote === 'sector') return 'All qualifying matched Units in Sector';
  if (rule.scopeNote === 'portfolio') return 'All qualifying matched Units in Portfolio';
  return 'Matched Units';
}

export function methodLabel(rule) {
  if (!rule || rule.method === 'blank' || (rule.method === 'summary' && rule.summary === 'blank')) return 'Blank';
  return METHOD_LABELS[rule.method] || 'Blank';
}

export function describeRule(rule, context, columns, columnId) {
  if (!rule || rule.method === 'blank' || (rule.method === 'summary' && rule.summary === 'blank')) return 'Blank';
  if (rule.method === 'label') {
    return rule.label === 'grandTotal' ? 'LABEL\nGRAND TOTAL' : 'LABEL\n<Sector name> TOTAL';
  }
  if (rule.method === 'fixed') return 'Fixed value\n' + (rule.value == null ? '' : rule.value);
  if (rule.method === 'source') {
    var source = sourcePath(rule, context);
    if (columnId === 'break_date' && context === 'unit') return source + '\nFormat: dd/mm/yyyy';
    return source;
  }
  if (rule.method === 'aggregatedSource') {
    var preview = aggregatedPreview(rule, context);
    if (rule.ref && rule.ref.scope === 'currentUnitRow') {
      return preview + '\nScope: Current Unit Source Row\nUses the maximum occupied area across all available period columns for this Unit only.';
    }
    return preview;
  }
  if (rule.method === 'summary' && rule.summary === 'sum') {
    if (rule.scope === 'matchedUnits') {
      if (rule.scopeNote === 'sector') return 'SUM Unit Areas in Sector';
      if (rule.scopeNote === 'portfolio') return 'SUM all Unit Areas in Portfolio';
      return 'SUM matched Unit Areas';
    }
    return 'SUM(Property › ' + fieldLabel(columns, columnId) + ')';
  }
  if (rule.method === 'summary' && rule.summary === 'recalculate') {
    if (columnId === 'cv_psf' && context === 'sectorTotal') return 'Sector Fair Value ÷ Sector Area';
    if (columnId === 'cv_psf' && context === 'grandTotal') return 'Grand Fair Value ÷ Grand Area';
    return fieldLabel(columns, rule.numerator && rule.numerator.id) + ' ' + operationSymbol(rule.operation) + ' ' + fieldLabel(columns, rule.denominator && rule.denominator.id);
  }
  if (rule.method === 'derived' && rule.calculation === 'divide') {
    if (rule.scopeNote === 'sector') return 'Sector Fair Value ÷ Sector Area';
    if (rule.scopeNote === 'portfolio') return 'Grand Fair Value ÷ Grand Area';
    return fieldLabel(columns, rule.numerator && rule.numerator.id) + ' ÷ ' + fieldLabel(columns, rule.denominator && rule.denominator.id);
  }
  if (rule.method === 'derived' && rule.measure === 'unexpiredLeaseTermToBreak') {
    return [
      'WEIGHTED AVG',
      'Remaining Term to Break',
      'Weight: Contracted Rent',
      'Scope: ' + scopeLine(rule),
      'Format: 0.00',
    ].join('\n');
  }
  if (rule.method === 'derived') {
    return [
      'WEIGHTED AVG',
      'Remaining Term to Expiry',
      'Weight: Contracted Rent',
      'Scope: ' + scopeLine(rule),
      'Format: 0.00',
    ].join('\n');
  }
  return 'Blank';
}

export function contextLabel(context) {
  if (context === 'property') return 'Property';
  if (context === 'unit') return 'Unit';
  if (context === 'sectorTotal') return 'Sector Total';
  if (context === 'grandTotal') return 'Grand Total';
  return context;
}

export function isSummaryRule(rule) {
  if (!rule) return false;
  return rule.method === 'summary' || rule.method === 'label';
}

export function isCalculatedRule(rule) {
  return !!(rule && rule.method === 'derived');
}
