import { getLogicTemplate } from './registry.js';
import { cloneDefinition } from './draftState.js';
import { validateTemplateLogic } from './validateLogic.js';
import { DEFAULT_SCHEDULE_STRUCTURE } from './scheduleStructure.js';
import { breakdownSpecification } from './portfolioBreakdown.js';

var APPLICATION_VERSION = '1.0.0';

function logicIndex(columns, contexts) {
  var waultToExpiry = [];
  var waultToBreak = [];
  var summaryAggregation = [];
  for (var i = 0; i < columns.length; i += 1) {
    var column = columns[i];
    var expiryContexts = {};
    var breakContexts = {};
    var expiryCount = 0;
    var breakCount = 0;
    for (var c = 0; c < contexts.length; c += 1) {
      var context = contexts[c];
      var rule = column.logic && column.logic[context];
      if (!rule) continue;
      if (rule.method === 'derived' && rule.measure === 'unexpiredLeaseTermToBreak') {
        breakContexts[context] = cloneDefinition(rule);
        breakCount += 1;
      } else if (rule.method === 'derived' && rule.measure === 'unexpiredLeaseTerm') {
        expiryContexts[context] = cloneDefinition(rule);
        expiryCount += 1;
      }
      if (rule.method === 'summary' || rule.method === 'label') {
        summaryAggregation.push({
          columnId: column.id,
          label: column.label,
          context: context,
          rule: cloneDefinition(rule),
        });
      }
    }
    if (expiryCount) {
      waultToExpiry.push({ columnId: column.id, label: column.label, contexts: expiryContexts });
    }
    if (breakCount) {
      waultToBreak.push({ columnId: column.id, label: column.label, contexts: breakContexts });
    }
  }
  var area = columns.filter(function (column) { return column.id === 'area_sqft'; }).map(function (column) {
    return { columnId: column.id, label: column.label, logic: cloneDefinition(column.logic) };
  });
  var capitalValue = columns.filter(function (column) { return column.id === 'cv_psf'; }).map(function (column) {
    return { columnId: column.id, label: column.label, logic: cloneDefinition(column.logic) };
  });
  var currentFairValue = columns.filter(function (column) { return column.id === 'fair_value'; }).map(function (column) {
    return {
      columnId: column.id,
      label: column.label,
      heading: 'Fair Value <Month YYYY> from Property › General › Traditional Valuation Date when all valid Property dates agree; otherwise Fair Value',
      logic: cloneDefinition(column.logic),
    };
  });
  var disabledComparison = columns.filter(function (column) {
    return column.rendered === false || column.feature === 'comparison';
  }).map(function (column) {
    return {
      columnId: column.id,
      label: column.label,
      rendered: false,
      feature: column.feature || 'comparison',
      logic: cloneDefinition(column.logic),
    };
  });
  return {
    contexts: contexts.slice(),
    waultToExpiry: waultToExpiry,
    waultToBreak: waultToBreak,
    summaryAggregation: summaryAggregation,
    area: area,
    capitalValue: capitalValue,
    currentFairValue: currentFairValue,
    disabledComparison: disabledComparison,
  };
}

function logicChanges(productionColumns, draftColumns, contexts) {
  var changes = [];
  var draftById = {};
  for (var i = 0; i < draftColumns.length; i += 1) draftById[draftColumns[i].id] = draftColumns[i];
  for (var p = 0; p < productionColumns.length; p += 1) {
    var column = productionColumns[p];
    var draft = draftById[column.id];
    for (var c = 0; c < contexts.length; c += 1) {
      var context = contexts[c];
      var before = column.logic ? column.logic[context] : null;
      var after = draft && draft.logic ? draft.logic[context] : null;
      if (JSON.stringify(before) === JSON.stringify(after)) continue;
      changes.push({
        columnId: column.id,
        label: column.label,
        context: context,
        production: cloneDefinition(before),
        draft: cloneDefinition(after),
      });
    }
  }
  return changes;
}

export function buildDraftSpecification(templateId, draft) {
  var production = getLogicTemplate(templateId);
  var contexts = production.contexts;
  var issues = validateTemplateLogic(draft.columns, contexts);
  var stamp = new Date();
  var y = stamp.getFullYear();
  var m = String(stamp.getMonth() + 1).padStart(2, '0');
  var d = String(stamp.getDate()).padStart(2, '0');
  return {
    filename: 'Tenancy_Schedule_Tools_Logic_Draft_' + y + m + d + '.json',
    document: {
      schemaVersion: 1,
      applicationVersion: APPLICATION_VERSION,
      exportedAt: stamp.toISOString(),
      template: { id: production.id, label: production.label },
      productionDefinition: {
        columns: cloneDefinition(production.columns),
        structure: cloneDefinition(production.structure || null),
        waultToExpiry: 'rent-weighted unexpired term from Property Valuation to Unit Lease Expiry',
        waultToBreak: 'rent-weighted unexpired term from Property Valuation to future Effective Break, otherwise Lease Expiry',
      },
      draftDefinition: {
        columns: cloneDefinition(draft.columns),
        structure: cloneDefinition(draft.structure || production.structure || null),
      },
      changes: logicChanges(production.columns, draft.columns, contexts),
      validationSummary: {
        issueCount: issues.length,
        issues: issues,
      },
      logicIndex: logicIndex(draft.columns, contexts),
      scheduleStructure: production.id === 'tenancy_schedule' ? cloneDefinition(DEFAULT_SCHEDULE_STRUCTURE) : null,
      portfolioBreakdown: production.id === 'tenancy_schedule' ? breakdownSpecification() : null,
      developerNotes:
        'Draft only. Production generation reads the source-controlled template. Apply accepted changes in source code. This file contains logic metadata and no workbook business data.',
    },
  };
}
