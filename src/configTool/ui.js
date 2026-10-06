import { APP_STATE, PROPERTY_SUMMARY_STATE } from '../state.js';
import { escapeHtml } from '../ui/common.js';
import { getLogicTemplate, listLogicTemplates } from '../logic/registry.js';
import { getDraftTemplate, resetDraft, updateDraftRule } from '../logic/draftState.js';
import {
  aggregatedPreview,
  contextLabel,
  describeRule,
  isCalculatedRule,
  isSummaryRule,
  methodLabel,
  scopeLine,
  sourcePath,
} from '../logic/describe.js';
import { validateLoadedSources, validateTemplateLogic } from '../logic/validateLogic.js';
import { buildDraftSpecification } from '../logic/exportDraft.js';
import { waultBreakRule, waultExpiryRule } from '../logic/fieldRules.js';
import { MULTIPLE_VALUATION_DATES_WARNING, resolvePortfolioValuation } from '../logic/fairValue.js';
import { readSourceValue } from '../schema/sourceLookup.js';
import { PROPERTY_SOURCE_FIELDS } from '../schema/sourceFields.js';

var viewState = {
  templateId: 'tenancy_schedule',
  search: '',
  filter: 'all',
  openKey: '',
  restoreSearch: false,
};

function scopeNote(context) {
  if (context === 'sectorTotal') return 'sector';
  if (context === 'grandTotal') return 'portfolio';
  return 'property';
}

function productionColumn(columnId) {
  var production = getLogicTemplate(viewState.templateId);
  for (var i = 0; i < production.columns.length; i += 1) {
    if (production.columns[i].id === columnId) return production.columns[i];
  }
  return null;
}

function rulesEqual(left, right) {
  return JSON.stringify(left || null) === JSON.stringify(right || null);
}

function isModified(column, context) {
  var production = productionColumn(column.id);
  var before = production && production.logic ? production.logic[context] : null;
  var after = column.logic ? column.logic[context] : null;
  return !rulesEqual(before, after);
}

function changeCount(columns, contexts) {
  var count = 0;
  for (var i = 0; i < columns.length; i += 1) {
    for (var c = 0; c < contexts.length; c += 1) {
      if (isModified(columns[i], contexts[c])) count += 1;
    }
  }
  return count;
}

function issuesFor(issues, columnId, context) {
  return issues.filter(function (issue) {
    return issue.columnId === columnId && issue.context === context;
  });
}

function statusFor(schemaHits, runtimeHits) {
  if (schemaHits.length) return { kind: 'error', text: '✕ Error' };
  if (runtimeHits.length) return { kind: 'warning', text: '⚠ Warning' };
  return { kind: 'valid', text: '✓ Valid' };
}

function counts(columns, contexts) {
  var result = { fields: columns.length, source: 0, calculated: 0, aggregated: 0, summary: 0 };
  for (var i = 0; i < columns.length; i += 1) {
    for (var c = 0; c < contexts.length; c += 1) {
      var rule = columns[i].logic && columns[i].logic[contexts[c]];
      if (!rule) continue;
      if (rule.method === 'source') result.source += 1;
      if (rule.method === 'derived') result.calculated += 1;
      if (rule.method === 'aggregatedSource') result.aggregated += 1;
      if (isSummaryRule(rule)) result.summary += 1;
    }
  }
  return result;
}

function matchesFilter(column, context, rule, issues) {
  if (viewState.filter === 'modified') return isModified(column, context);
  if (viewState.filter === 'issues') return issuesFor(issues, column.id, context).length > 0;
  if (viewState.filter === 'calculated') return isCalculatedRule(rule);
  if (viewState.filter === 'summary') return isSummaryRule(rule);
  return true;
}

function matchesSearch(column, context, rule) {
  var query = viewState.search.trim().toLowerCase();
  if (!query) return true;
  var haystack = [
    column.label,
    contextLabel(context),
    methodLabel(rule),
    describeRule(rule, context, getDraftTemplate(viewState.templateId).columns, column.id),
  ].join(' ').toLowerCase();
  return haystack.indexOf(query) !== -1;
}

function isComparisonField(column) {
  return !!(column && (column.rendered === false || column.feature === 'comparison'));
}

function resolvedFairValue() {
  var propertyData = viewState.templateId === 'property_summary' ? PROPERTY_SUMMARY_STATE.property : APP_STATE.property;
  if (!propertyData || !propertyData.rows) return { heading: 'Fair Value', consistent: true };
  return resolvePortfolioValuation(
    propertyData.rows.map(function (row) {
      return readSourceValue(row, propertyData.columns, 'property', PROPERTY_SOURCE_FIELDS.traditionalValuationDate);
    })
  );
}

function displayLabel(column) {
  if (column.id === 'fair_value') return resolvedFairValue().heading;
  return column.label;
}

function fieldOptions(selectedId) {
  var columns = getDraftTemplate(viewState.templateId).columns;
  return columns.map(function (column) {
    return '<option value="' + escapeHtml(column.id) + '"' + (column.id === selectedId ? ' selected' : '') + '>' + escapeHtml(column.label) + '</option>';
  }).join('');
}

function formula(numerator, denominator) {
  return (
    '<div class="config-formula"><div>' + escapeHtml(numerator) + '</div><div class="config-formula-line"></div><div>' +
    escapeHtml(denominator) + '</div></div>'
  );
}

function diffBlock(column, context) {
  if (!isModified(column, context)) return '';
  var production = productionColumn(column.id);
  var before = production && production.logic ? production.logic[context] : null;
  var after = column.logic ? column.logic[context] : null;
  var columns = getDraftTemplate(viewState.templateId).columns;
  return (
    '<div class="config-diff"><div><h3>Production</h3><p>' + escapeHtml(describeRule(before, context, columns, column.id)) + '</p></div><div><h3>Draft</h3><p>' +
    escapeHtml(describeRule(after, context, columns, column.id)) + '</p></div></div>'
  );
}

function validationLine(hits) {
  if (!hits.length) return '<p class="config-check">✓ Valid</p>';
  return '<p class="config-check config-check-' + (hits[0].runtime ? 'warning' : 'error') + '">' + escapeHtml((hits[0].runtime ? '⚠ Warning' : '✕ Error') + ' — ' + hits[0].message) + '</p>';
}

function sourceEditor(column, context, rule, hits) {
  var ref = rule.ref || { type: 'field', dataset: 'property', section: '', field: '', occurrence: 1 };
  var dataset = ref.dataset || (context === 'unit' ? 'unit' : 'property');
  return (
    '<div class="config-edit-grid">' +
    '<label>Dataset<select data-edit="dataset"><option value="property"' + (dataset !== 'unit' ? ' selected' : '') + '>Property</option><option value="unit"' + (dataset === 'unit' ? ' selected' : '') + '>Unit</option></select></label>' +
    '<label>Section<input data-edit="section" value="' + escapeHtml(ref.section || '') + '"></label>' +
    '<label>Field<input data-edit="field" value="' + escapeHtml(ref.field || '') + '"></label>' +
    '</div>' +
    '<p class="config-kicker">Resolved source</p><p class="config-preview" data-role="resolved">' + escapeHtml(sourcePath(rule, context)) + '</p>' +
    (column.id === 'cap_ex' ? '<p class="config-note">Uses Capital Costs directly from the Property Level source.</p>' : '') +
    (column.id === 'break_date' && context === 'unit' ? '<p class="config-kicker">Format</p><p class="config-preview">dd/mm/yyyy</p><p class="config-note">Unit rows show the Effective Break date. Property, Sector Total and Grand Total rows in this column show WAULT to Break.</p>' : '') +
    validationLine(hits)
  );
}

function aggregatedEditor(rule, context, hits) {
  var ref = rule.ref || { dataset: 'unit', section: '', selection: 'allPeriods', aggregation: 'max' };
  function option(name, value, current) {
    return '<option value="' + name + '"' + (current === name ? ' selected' : '') + '>' + value + '</option>';
  }
  return (
    '<div class="config-edit-grid">' +
    '<label>Aggregation<select data-edit="aggregation">' +
    option('max', 'Maximum', ref.aggregation) + option('min', 'Minimum', ref.aggregation) + option('sum', 'Sum', ref.aggregation) + option('average', 'Average', ref.aggregation) +
    '</select></label>' +
    '<label>Dataset<select data-edit="dataset"><option value="property"' + (ref.dataset !== 'unit' ? ' selected' : '') + '>Property</option><option value="unit"' + (ref.dataset === 'unit' ? ' selected' : '') + '>Unit</option></select></label>' +
    '<label>Section<input data-edit="section" value="' + escapeHtml(ref.section || '') + '"></label>' +
    '<label>Field scope<select data-edit="selection">' +
    option('allPeriods', 'All Period Columns', ref.selection) + option('earliestPeriod', 'Earliest Period', ref.selection) + option('latestPeriod', 'Latest Period', ref.selection) + option('firstPeriod', 'First Period Column', ref.selection) + option('lastPeriod', 'Last Period Column', ref.selection) +
    '</select></label></div>' +
    '<p class="config-kicker">Preview</p><p class="config-preview" data-role="resolved">' + escapeHtml(aggregatedPreview(rule, context)) + '</p>' +
    (ref.section === 'Capital Cost'
      ? '<p class="config-preview">Aggregation: Sum</p><p class="config-preview">Sign Convention: Cost / Outflow</p><p class="config-preview">Output: Negative absolute value</p><p class="config-note">Sums all Capital Cost period values for the current Unit and reports the result as a negative cost.</p>'
      : (ref.scope === 'currentUnitRow' || ref.section === 'Occupied Area (SF)'
        ? '<p class="config-preview">Scope: Current Unit Source Row</p><p class="config-note">Uses the maximum occupied area across all available period columns for this Unit only.</p>'
        : '')) +
    validationLine(hits)
  );
}

function summaryEditor(column, context, rule, hits) {
  var columns = getDraftTemplate(viewState.templateId).columns;
  var kind = 'sum';
  if (rule.method === 'label') kind = 'label';
  else if (rule.method === 'derived') kind = 'weightedAverage';
  else if (rule.summary === 'recalculate') kind = 'recalculate';
  else if (rule.summary === 'blank' || rule.method === 'blank') kind = 'blank';
  var html = '<div class="config-edit-grid"><label>Summary method<select data-edit="summaryKind">';
  [['sum', 'Sum'], ['recalculate', 'Recalculate'], ['weightedAverage', 'Weighted Average'], ['label', 'Label'], ['blank', 'Blank']].forEach(function (item) {
    html += '<option value="' + item[0] + '"' + (kind === item[0] ? ' selected' : '') + '>' + item[1] + '</option>';
  });
  html += '</select></label>';
  if (kind === 'sum') {
    if (rule.scope === 'matchedUnits') {
      html += '<label>Source<select disabled><option>Matched Unit Areas</option></select></label></div>';
      html += '<p class="config-preview">' + escapeHtml(describeRule(rule, context, columns, column.id)) + '</p>';
      html += '<p class="config-note">Each matched Unit is counted once. Property source area is not used.</p>';
    } else {
      html += '<label>Source<select disabled><option>Property › ' + escapeHtml(column.id === 'cap_ex' ? 'General › Capital Costs' : column.label) + '</option></select></label></div>';
      html += '<p class="config-preview">' + escapeHtml(column.id === 'cap_ex' ? describeRule(rule, context, columns, column.id) : ('SUM(Property › ' + column.label + ')')) + '</p><p class="config-note">Adds the property-level value once per property. Unit rows are not included.</p>';
    }
  } else if (kind === 'recalculate') {
    html += '<label>Operation<select data-edit="operation">';
    [['divide', 'Divide'], ['multiply', 'Multiply'], ['add', 'Add'], ['subtract', 'Subtract']].forEach(function (item) {
      html += '<option value="' + item[0] + '"' + ((rule.operation || 'divide') === item[0] ? ' selected' : '') + '>' + item[1] + '</option>';
    });
    html += '</select></label>';
    html += '<label>Left operand<select data-edit="numerator">' + fieldOptions(rule.numerator && rule.numerator.id) + '</select></label>';
    html += '<label>Right operand<select data-edit="denominator">' + fieldOptions(rule.denominator && rule.denominator.id) + '</select></label></div>';
    html += '<p class="config-kicker">Equation</p><p class="config-preview">' + escapeHtml(describeRule(rule, context, columns, column.id)) + '</p>';
    if ((rule.operation || 'divide') !== 'divide') {
      html += '<p class="config-note">Production generation currently evaluates divide. Other operations are stored in the draft specification.</p>';
    }
  } else if (kind === 'weightedAverage') {
    html += '</div>' + waultEditor(column, context, rule.method === 'derived' ? rule : waultExpiryRule(scopeNote(context)), hits, true);
    return html;
  } else if (kind === 'label') {
    html += '</div><p class="config-preview">' + (context === 'grandTotal' ? 'GRAND TOTAL' : '&lt;Sector name&gt; TOTAL') + '</p>';
  } else {
    html += '</div><p class="config-preview">Blank</p><p class="config-note">No summary is calculated for this field.</p>';
  }
  html += validationLine(hits);
  return html;
}

function waultEditor(column, context, rule, hits, embedded) {
  var measure = rule.measure === 'unexpiredLeaseTermToBreak' ? 'unexpiredLeaseTermToBreak' : 'unexpiredLeaseTerm';
  var isBreak = measure === 'unexpiredLeaseTermToBreak';
  var html = '';
  if (!embedded) {
    html += '<div class="config-edit-grid"><label>Value method<select data-edit="method"><option value="derived" selected>Derived Calculation</option><option value="source">Source Field</option><option value="aggregatedSource">Aggregated Source</option><option value="summary">Summary Aggregation</option><option value="fixed">Fixed Value</option><option value="blank">Blank</option></select></label></div>';
  }
  html += '<div class="config-edit-grid">';
  html += '<label>Calculation type<select disabled><option>Weighted Average</option></select></label>';
  html += '<label>Measure<select data-edit="measure"><option value="unexpiredLeaseTerm"' + (!isBreak ? ' selected' : '') + '>Remaining Lease Term to Expiry</option><option value="unexpiredLeaseTermToBreak"' + (isBreak ? ' selected' : '') + '>Remaining Lease Term to Break</option></select></label>';
  html += '<label>Weight<select data-edit="weight">' + fieldOptions(rule.weight && rule.weight.id) + '</select></label>';
  html += '</div><dl class="config-facts">';
  html += '<div><dt>Start date</dt><dd>Property › Valuation</dd></div>';
  if (isBreak) {
    html += '<div><dt>Preferred end</dt><dd>Unit › Effective Break</dd></div>';
    html += '<div><dt>Fallback end</dt><dd>Unit › Lease Expiry</dd></div>';
    html += '<div><dt>Use Effective Break when</dt><dd>Valid future break on or before lease expiry</dd></div>';
  } else {
    html += '<div><dt>End date</dt><dd>Unit › Lease Expiry</dd></div>';
  }
  html += '<div><dt>Scope</dt><dd>' + escapeHtml(scopeLine(rule)) + '</dd></div>';
  html += '<div><dt>Format</dt><dd>0.00</dd></div></dl>';
  html += '<p class="config-kicker">Eligibility</p><div class="config-checks">';
  html += '<label><input type="checkbox" data-edit="validEnd"' + (!rule.eligibility || rule.eligibility.validEndDate !== false ? ' checked' : '') + '> ' + (isBreak ? 'Usable end date' : 'Valid expiry') + '</label>';
  html += '<label><input type="checkbox" data-edit="positiveRent"' + (!rule.eligibility || rule.eligibility.rentGreaterThan !== false ? ' checked' : '') + '> Contracted Rent &gt; 0</label></div>';
  html += '<p class="config-kicker">Formula preview</p>';
  html += formula(
    isBreak ? 'Σ (Unexpired Break Years × Contracted Rent)' : 'Σ (Unexpired Expiry Years × Contracted Rent)',
    'Σ (Contracted Rent)'
  );
  html += '<p class="config-note">' + (isBreak
    ? 'Uses the next valid future Effective Break where available; otherwise uses Lease Expiry, then weights remaining terms by Contracted Rent.'
    : "Calculates each matched unit's remaining lease term from the Property valuation date to Lease Expiry and weights it by Contracted Rent.") + '</p>';
  if (context === 'sectorTotal' || context === 'grandTotal') {
    html += '<p class="config-note">Calculated from the underlying units in this ' + (context === 'sectorTotal' ? 'sector' : 'portfolio') + '. Property WAULT values are not averaged.</p>';
  }
  html += validationLine(hits);
  return html;
}

function methodSelect(rule) {
  var method = !rule || rule.method === 'blank' ? 'blank' : rule.method === 'label' ? 'summary' : rule.method;
  var options = [
    ['source', 'Source Field'],
    ['derived', 'Derived Calculation'],
    ['aggregatedSource', rule && rule.sign === 'cost' ? 'Aggregated Source / Derived Aggregation' : 'Aggregated Source'],
    ['summary', 'Summary Aggregation'],
    ['fixed', 'Fixed Value'],
    ['blank', 'Blank'],
  ];
  return '<label>Value method<select data-edit="method">' + options.map(function (item) {
    return '<option value="' + item[0] + '"' + (method === item[0] ? ' selected' : '') + '>' + item[1] + '</option>';
  }).join('') + '</select></label>';
}

function divideEditor(column, context, rule, hits) {
  var columns = getDraftTemplate(viewState.templateId).columns;
  return (
    '<div class="config-edit-grid">' + methodSelect(rule) +
    '<label>Operation<select disabled><option>Divide</option></select></label>' +
    '<label>Numerator<select data-edit="numerator">' + fieldOptions(rule.numerator && rule.numerator.id) + '</select></label>' +
    '<label>Denominator<select data-edit="denominator">' + fieldOptions(rule.denominator && rule.denominator.id) + '</select></label>' +
    '</div><p class="config-kicker">Equation</p><p class="config-preview">' + escapeHtml(describeRule(rule, context, columns, column.id)) + '</p>' +
    '<p class="config-note">Blank when the denominator is blank or zero. Property, Sector and Grand Total all divide Fair Value by Area.</p>' +
    validationLine(hits)
  );
}

function editorBody(column, context, rule, hits) {
  var current = rule || { method: 'blank' };
  if (current.method === 'derived' && current.calculation === 'divide') return divideEditor(column, context, current, hits);
  if (current.method === 'derived') return waultEditor(column, context, current, hits, false);
  if (current.method === 'source') {
    return '<div class="config-edit-grid">' + methodSelect(current) + '</div>' + sourceEditor(column, context, current, hits);
  }
  if (current.method === 'aggregatedSource') {
    return '<div class="config-edit-grid">' + methodSelect(current) + '</div>' + aggregatedEditor(current, context, hits);
  }
  if (current.method === 'summary' || current.method === 'label' || (current.method === 'blank' && (context === 'sectorTotal' || context === 'grandTotal'))) {
    var summaryRule = current.method === 'blank' && (context === 'sectorTotal' || context === 'grandTotal')
      ? { method: 'summary', summary: 'blank' }
      : current;
    return '<div class="config-edit-grid">' + methodSelect(summaryRule.method === 'label' ? { method: 'summary' } : summaryRule) + '</div>' + summaryEditor(column, context, summaryRule, hits);
  }
  if (current.method === 'fixed') {
    return '<div class="config-edit-grid">' + methodSelect(current) + '<label>Value<input data-edit="fixed" value="' + escapeHtml(current.value == null ? '' : current.value) + '"></label></div>' + validationLine(hits);
  }
  return '<div class="config-edit-grid">' + methodSelect(current) + '</div><p class="config-preview">Blank</p>' + validationLine(hits);
}

function canonicalRule(method, context) {
  if (method === 'source') {
    return {
      method: 'source',
      ref: { type: 'field', dataset: context === 'unit' ? 'unit' : 'property', section: 'General', field: '', occurrence: 1 },
    };
  }
  if (method === 'aggregatedSource') {
    return {
      method: 'aggregatedSource',
      ref: { type: 'periodGroup', dataset: context === 'unit' ? 'unit' : 'property', section: '', selection: 'allPeriods', aggregation: 'max' },
    };
  }
  if (method === 'summary') return { method: 'summary', summary: 'sum', scope: 'properties' };
  if (method === 'fixed') return { method: 'fixed', value: '' };
  if (method === 'derived') return waultExpiryRule(scopeNote(context));
  if (method === 'label') return { method: 'label', label: context === 'grandTotal' ? 'grandTotal' : 'sectorTotal' };
  return { method: 'blank' };
}

function readWault(form, context, measure) {
  var next = measure === 'unexpiredLeaseTermToBreak' ? waultBreakRule(scopeNote(context)) : waultExpiryRule(scopeNote(context));
  var weight = form.querySelector('[data-edit="weight"]');
  if (weight) next.weight = { kind: 'templateField', id: weight.value, dataset: 'unit' };
  var validEnd = form.querySelector('[data-edit="validEnd"]');
  var positiveRent = form.querySelector('[data-edit="positiveRent"]');
  if (validEnd) next.eligibility.validEndDate = validEnd.checked;
  if (positiveRent) next.eligibility.rentGreaterThan = positiveRent.checked ? 0 : false;
  return next;
}

function readRule(form, column, context) {
  var current = (column.logic && column.logic[context]) || { method: 'blank' };
  var methodInput = form.querySelector('[data-edit="method"]');
  var summaryInput = form.querySelector('[data-edit="summaryKind"]');
  var method = methodInput ? methodInput.value : current.method;
  var summaryKind = summaryInput ? summaryInput.value : '';
  if (method === 'summary' && summaryKind === 'label') method = 'label';
  if (method === 'summary' && summaryKind === 'blank') method = 'blank';
  if (method === 'derived' || (method === 'summary' && summaryKind === 'weightedAverage')) {
    if (form.querySelector('[data-edit="numerator"]') && !form.querySelector('[data-edit="measure"]')) {
      return {
        method: 'derived',
        calculation: 'divide',
        numerator: { kind: 'templateField', id: form.querySelector('[data-edit="numerator"]').value },
        denominator: { kind: 'templateField', id: form.querySelector('[data-edit="denominator"]').value },
        scopeNote: scopeNote(context),
      };
    }
    var measureInput = form.querySelector('[data-edit="measure"]');
    if (!measureInput) return waultExpiryRule(scopeNote(context));
    return readWault(form, context, measureInput.value);
  }
  if (method === 'blank' || method === 'label') return canonicalRule(method, context);
  if (method === 'source') {
    if (!form.querySelector('[data-edit="section"]')) return canonicalRule('source', context);
    return {
      method: 'source',
      ref: {
        type: 'field',
        dataset: form.querySelector('[data-edit="dataset"]').value,
        section: form.querySelector('[data-edit="section"]').value.trim(),
        field: form.querySelector('[data-edit="field"]').value.trim(),
        occurrence: 1,
      },
    };
  }
  if (method === 'aggregatedSource') {
    if (!form.querySelector('[data-edit="selection"]')) return canonicalRule('aggregatedSource', context);
    return {
      method: 'aggregatedSource',
      ref: {
        type: 'periodGroup',
        dataset: form.querySelector('[data-edit="dataset"]').value,
        section: form.querySelector('[data-edit="section"]').value.trim(),
        selection: form.querySelector('[data-edit="selection"]').value,
        aggregation: form.querySelector('[data-edit="aggregation"]').value,
      },
    };
  }
  if (method === 'fixed') {
    var fixed = form.querySelector('[data-edit="fixed"]');
    if (!fixed) return canonicalRule('fixed', context);
    return { method: 'fixed', value: fixed.value };
  }
  if (summaryKind === 'sum') return { method: 'summary', summary: 'sum', scope: 'properties' };
  if (!form.querySelector('[data-edit="numerator"]')) {
    return {
      method: 'summary',
      summary: 'recalculate',
      operation: 'divide',
      numerator: { kind: 'templateField', id: 'contracted_rent_pa' },
      denominator: { kind: 'templateField', id: 'area_sqft' },
    };
  }
  return {
    method: 'summary',
    summary: 'recalculate',
    operation: form.querySelector('[data-edit="operation"]').value,
    numerator: { kind: 'templateField', id: form.querySelector('[data-edit="numerator"]').value },
    denominator: { kind: 'templateField', id: form.querySelector('[data-edit="denominator"]').value },
  };
}

function renderConfiguration() {
  var root = document.getElementById('config-root');
  if (!root) return;
  var templates = listLogicTemplates();
  var production = getLogicTemplate(viewState.templateId);
  var draft = getDraftTemplate(viewState.templateId);
  var schemaIssues = validateTemplateLogic(draft.columns, draft.contexts);
  var productionIssues = validateTemplateLogic(production.columns, production.contexts);
  var propertyWorkbook = viewState.templateId === 'property_summary' ? PROPERTY_SUMMARY_STATE.property : APP_STATE.property;
  var unitWorkbook = viewState.templateId === 'property_summary' ? null : APP_STATE.unit;
  var runtimeIssues = validateLoadedSources(
    draft.columns,
    draft.contexts,
    propertyWorkbook && propertyWorkbook.columns,
    unitWorkbook && unitWorkbook.columns
  ).map(function (issue) {
    return Object.assign({ runtime: true }, issue);
  });
  var issues = schemaIssues.concat(runtimeIssues);
  var tally = counts(draft.columns.filter(function (column) { return !isComparisonField(column); }), draft.contexts);
  var changes = changeCount(draft.columns, draft.contexts);
  var columns = draft.columns.filter(function (column) { return !isComparisonField(column) && column.role !== 'system'; });
  var systemColumns = draft.columns.filter(function (column) { return column.role === 'system'; });
  var disabledColumns = draft.columns.filter(isComparisonField);
  var body = '';
  var shown = 0;
  var group = 0;
  var fairValue = resolvedFairValue();

  for (var i = 0; i < columns.length; i += 1) {
    var column = columns[i];
    var contexts = draft.contexts.filter(function (context) {
      var rule = column.logic ? column.logic[context] : null;
      return matchesFilter(column, context, rule, issues) && matchesSearch(column, context, rule);
    });
    if (!contexts.length) continue;
    var openHere = contexts.some(function (context) { return column.id + '|' + context === viewState.openKey; }) ? 1 : 0;
    var span = contexts.length + openHere;
    var stripe = group % 2 === 1 ? ' config-group-alt' : '';
    group += 1;
    for (var c = 0; c < contexts.length; c += 1) {
      var context = contexts[c];
      var rule = column.logic ? column.logic[context] : null;
      var key = column.id + '|' + context;
      var rowIssues = issuesFor(issues, column.id, context);
      var schemaHits = rowIssues.filter(function (issue) { return !issue.runtime; });
      var runtimeHits = rowIssues.filter(function (issue) { return issue.runtime; });
      var status = statusFor(schemaHits, runtimeHits);
      var modified = isModified(column, context);
      body += '<tr class="' + stripe.trim() + '" data-row-key="' + escapeHtml(key) + '">';
      if (c === 0) {
        body += '<td class="config-field" rowspan="' + span + '">' + escapeHtml(displayLabel(column));
        if (column.format) body += '<span class="config-format">' + escapeHtml(column.format) + '</span>';
        if (column.comparison && column.comparison.supported) {
          body += '<span class="config-format">Comparison supported'
            + (column.comparison.defaultEnabled ? ' · Default on' : '')
            + (column.comparison.showPercentChange ? ' · Change %' : '')
            + (column.comparison.propertyOnly ? ' · Property only' : '')
            + '</span>';
        }
        body += '</td>';
      }
      body += '<td>' + escapeHtml(contextLabel(context)) + '</td>';
      body += '<td>' + escapeHtml(methodLabel(rule)) + '</td>';
      body += '<td class="config-definition" data-definition-for="' + escapeHtml(key) + '">' + escapeHtml(describeRule(rule, context, columns, column.id)) + '</td>';
      body += '<td class="config-status config-status-' + status.kind + '">' + status.text + (modified ? '<span class="config-modified">Modified</span>' : '') + '</td>';
      body += '<td><button type="button" class="config-action" data-action="toggle" data-key="' + escapeHtml(key) + '">' + (viewState.openKey === key ? 'Close' : 'Edit') + '</button></td></tr>';
      if (viewState.openKey === key) {
        body += '<tr class="config-editor-row' + stripe + '" data-editor-for="' + escapeHtml(key) + '"><td colspan="5"><div class="config-inline-editor" data-column="' + escapeHtml(column.id) + '" data-context="' + escapeHtml(context) + '">';
        body += '<p class="config-editor-title">' + escapeHtml(displayLabel(column)) + ' · ' + escapeHtml(contextLabel(context)) + '</p>';
        body += editorBody(column, context, rule, rowIssues);
        body += diffBlock(column, context);
        body += '</div></td></tr>';
      }
      shown += 1;
    }
  }

  var issueList = '';
  if (issues.length) {
    issueList = '<div class="config-issues"><h3>Validation</h3><ul>' + issues.map(function (issue) {
      var key = issue.columnId + '|' + issue.context;
      return '<li><button type="button" data-action="jump" data-key="' + escapeHtml(key) + '">' + escapeHtml(issue.label + ' · ' + contextLabel(issue.context) + ' — ' + issue.message) + '</button></li>';
    }).join('') + '</ul></div>';
  }

  var systemBody = '';
  if (systemColumns.length) {
    systemBody = '<section class="config-system"><h3>System Fields</h3><p>Always included and hidden in Excel. These references are not part of the report layout.</p><table class="config-table"><thead><tr><th>Output Field</th><th>Rule</th><th>Placement</th></tr></thead><tbody>' +
      systemColumns.map(function (column) {
        var rule = column.id === 'asset_unit_reference' ? 'Unit only. Address + "_" + Unit' : 'Unit only. Unit › General › Lease ID';
        return '<tr><td>' + escapeHtml(column.label) + '</td><td>' + escapeHtml(rule) + '</td><td>Always included · Hidden in Excel</td></tr>';
      }).join('') + '</tbody></table></section>';
  }

  var disabledBody = '';
  if (disabledColumns.length) {
    disabledBody = '<section class="config-disabled"><h3>Comparison fields — currently disabled</h3><p>These historical / comparison columns remain in the template definition but are not rendered on the generated schedule.</p><table class="config-table config-table-disabled"><thead><tr><th>Output Field</th><th>Status</th></tr></thead><tbody>' +
      disabledColumns.map(function (column) {
        return '<tr><td>' + escapeHtml(column.label) + '</td><td>Not rendered</td></tr>';
      }).join('') +
      '</tbody></table></section>';
  }

  root.innerHTML =
    '<div class="config-workspace">' +
    '<div class="config-controls">' +
    '<label class="config-template">Template<select id="config-template">' + templates.map(function (template) {
      return '<option value="' + escapeHtml(template.id) + '"' + (template.id === viewState.templateId ? ' selected' : '') + '>' + escapeHtml(template.label) + '</option>';
    }).join('') + '</select></label>' +
    '<span class="config-production">' + (productionIssues.length ? '✕ Production logic has errors' : '✓ Production logic valid') + '</span>' +
    '<span class="config-changes">' + (changes ? changes + ' draft change' + (changes === 1 ? '' : 's') : 'No draft changes') + '</span>' +
    '<button type="button" id="config-reset" class="btn-secondary">Reset Draft</button>' +
    '<button type="button" id="config-download" class="generate-btn">Download Draft Specification</button>' +
    '</div>' +
    '<div class="config-banner"><strong>Draft Logic Editor</strong><p>Changes made here do not alter production generation. Download the draft specification to apply changes in source code.</p></div>' +
    (fairValue.consistent ? '' : '<p class="config-date-warning">' + escapeHtml(MULTIPLE_VALUATION_DATES_WARNING) + '</p>') +
    '<div class="config-toolbar"><input id="config-search" type="search" placeholder="Search fields..." value="' + escapeHtml(viewState.search) + '">' +
    '<div class="config-filters" role="group" aria-label="Logic filters">' +
    [['all', 'All'], ['modified', 'Modified'], ['issues', 'Issues'], ['calculated', 'Calculated'], ['summary', 'Summary Rules']].map(function (item) {
      return '<button type="button" data-filter="' + item[0] + '"' + (viewState.filter === item[0] ? ' aria-pressed="true"' : ' aria-pressed="false"') + '>' + item[1] + '</button>';
    }).join('') + '</div>' +
    '<span class="config-metrics">' + tally.fields + ' Fields · ' + tally.source + ' Source · ' + tally.calculated + ' Calculated · ' + tally.aggregated + ' Aggregated · ' + tally.summary + ' Summary Rules · ' + issues.length + ' Errors</span></div>' +
    '<p class="config-legend">' + draft.contexts.map(function (context) {
      var text = context === 'property' ? 'Value shown on Property summary rows' : context === 'unit' ? 'Value shown on tenancy/unit rows' : context === 'sectorTotal' ? 'Value shown on <Sector> TOTAL' : 'Value shown on portfolio GRAND TOTAL';
      return '<strong>' + escapeHtml(contextLabel(context)) + '</strong> ' + escapeHtml(text);
    }).join(' ') + '</p>' +
    issueList +
    systemBody +
    '<div class="config-table-wrap"><table class="config-table"><thead><tr><th class="config-field">Output Field</th><th>Context</th><th>Value Method</th><th>Logic Definition</th><th>Status</th><th>Action</th></tr></thead><tbody>' +
    (body || '<tr><td colspan="6" class="config-empty">No fields match the current search or filter.</td></tr>') +
    '</tbody></table></div>' + disabledBody +
    '<p class="config-footnote">' + shown + ' visible rules. ' +
    (propertyWorkbook || unitWorkbook ? 'Loaded workbooks are used for schema checks only.' : 'Schema validation only until a workbook is loaded in Standardiser.') + '</p></div>';

  if (viewState.restoreEdit) {
    var editField = root.querySelector('.config-inline-editor [data-edit="' + viewState.restoreEdit.edit + '"]');
    if (editField) {
      editField.focus();
      if (editField.setSelectionRange && viewState.restoreEdit.caret != null) {
        editField.setSelectionRange(viewState.restoreEdit.caret, viewState.restoreEdit.caret);
      }
    }
    viewState.restoreEdit = null;
  }
  if (viewState.restoreSearch) {
    var search = document.getElementById('config-search');
    if (search) {
      search.focus();
      var end = search.value.length;
      search.setSelectionRange(end, end);
    }
    viewState.restoreSearch = false;
  }
  var editor = root.querySelector('[data-editor-for="' + viewState.openKey + '"]');
  if (editor && viewState.scrollEditor) {
    editor.scrollIntoView({ block: 'nearest' });
    viewState.scrollEditor = false;
  }
}

function commitEditor(form) {
  var draft = getDraftTemplate(viewState.templateId);
  var columnId = form.getAttribute('data-column');
  var context = form.getAttribute('data-context');
  var column = null;
  for (var i = 0; i < draft.columns.length; i += 1) {
    if (draft.columns[i].id === columnId) column = draft.columns[i];
  }
  if (!column) return;
  updateDraftRule(viewState.templateId, columnId, context, readRule(form, column, context));
}

function bindConfiguration(root) {
  root.addEventListener('click', function (event) {
    var filter = event.target.closest('[data-filter]');
    if (filter) {
      viewState.filter = filter.getAttribute('data-filter');
      renderConfiguration();
      return;
    }
    var toggle = event.target.closest('[data-action="toggle"]');
    if (toggle) {
      var key = toggle.getAttribute('data-key');
      viewState.openKey = viewState.openKey === key ? '' : key;
      viewState.scrollEditor = true;
      renderConfiguration();
      return;
    }
    var jump = event.target.closest('[data-action="jump"]');
    if (jump) {
      viewState.openKey = jump.getAttribute('data-key');
      viewState.filter = 'all';
      viewState.search = '';
      viewState.scrollEditor = true;
      renderConfiguration();
      return;
    }
    if (event.target.id === 'config-reset') {
      resetDraft(viewState.templateId);
      viewState.openKey = '';
      renderConfiguration();
      return;
    }
    if (event.target.id === 'config-download') {
      var spec = buildDraftSpecification(viewState.templateId, getDraftTemplate(viewState.templateId));
      var blob = new Blob([JSON.stringify(spec.document, null, 2)], { type: 'application/json' });
      var url = URL.createObjectURL(blob);
      var link = document.createElement('a');
      link.href = url;
      link.download = spec.filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
    }
  });
  root.addEventListener('change', function (event) {
    if (event.target.id === 'config-template') {
      viewState.templateId = event.target.value;
      viewState.openKey = '';
      viewState.search = '';
      viewState.filter = 'all';
      renderConfiguration();
      return;
    }
    var form = event.target.closest('.config-inline-editor');
    if (!form) return;
    commitEditor(form);
    renderConfiguration();
  });
  root.addEventListener('input', function (event) {
    if (event.target.id === 'config-search') {
      viewState.search = event.target.value;
      viewState.restoreSearch = true;
      renderConfiguration();
      return;
    }
    if (event.target.tagName === 'SELECT') return;
    var form = event.target.closest('.config-inline-editor');
    if (!form) return;
    var editName = event.target.getAttribute('data-edit');
    var caret = event.target.selectionStart;
    commitEditor(form);
    viewState.restoreEdit = { edit: editName, caret: caret };
    renderConfiguration();
  });
}

var bound = false;
var lastFocus = null;

function isConfigOpen() {
  var modal = document.getElementById('config-modal');
  return !!(modal && !modal.classList.contains('hidden'));
}

export function closeConfiguration() {
  var modal = document.getElementById('config-modal');
  var openBtn = document.getElementById('config-open');
  if (modal) {
    modal.classList.add('hidden');
    modal.setAttribute('aria-hidden', 'true');
  }
  document.body.classList.remove('config-open');
  if (lastFocus && lastFocus.focus) lastFocus.focus();
  else if (openBtn) openBtn.focus();
}

export function openConfiguration() {
  lastFocus = document.activeElement;
  renderConfiguration();
  var modal = document.getElementById('config-modal');
  var closeBtn = document.getElementById('config-close');
  if (modal) {
    modal.classList.remove('hidden');
    modal.setAttribute('aria-hidden', 'false');
  }
  document.body.classList.add('config-open');
  if (closeBtn) closeBtn.focus();
}

export function setupConfigurationTool() {
  var root = document.getElementById('config-root');
  if (root && !bound) {
    bindConfiguration(root);
    bound = true;
  }
  var openBtn = document.getElementById('config-open');
  var closeBtn = document.getElementById('config-close');
  var modal = document.getElementById('config-modal');
  if (openBtn && !openBtn.dataset.bound) {
    openBtn.dataset.bound = '1';
    openBtn.addEventListener('click', openConfiguration);
  }
  if (closeBtn && !closeBtn.dataset.bound) {
    closeBtn.dataset.bound = '1';
    closeBtn.addEventListener('click', closeConfiguration);
  }
  if (modal && !modal.dataset.bound) {
    modal.dataset.bound = '1';
    modal.addEventListener('click', function (event) {
      if (event.target && event.target.getAttribute('data-config-close') === 'true') closeConfiguration();
    });
  }
  document.addEventListener('keydown', function (event) {
    if (event.key === 'Escape' && isConfigOpen()) {
      closeConfiguration();
      event.preventDefault();
    }
  });
  renderConfiguration();
}

export function refreshConfiguration() {
  renderConfiguration();
}

