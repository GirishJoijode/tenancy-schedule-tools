import { normaliseHeader } from '../schema/text.js';
import { findSourceColumn, readSourceValue } from '../schema/sourceLookup.js';
import { buildPortfolio } from '../model/portfolio.js';
import { UNIT_SOURCE_FIELDS } from '../schema/sourceFields.js';
import { buildPortfolioGroups, getBreakdownMode, getCustomConfig } from '../logic/portfolioBreakdown.js';
import { DEFAULT_SCHEDULE_COLUMNS } from '../templates/defaultSchedule.js';
import { buildOutputColumns } from '../logic/reportLayout.js';
import { DEFAULT_SCHEDULE_STRUCTURE } from '../logic/scheduleStructure.js';
import { applyPropertyContext, calculateGroupTotals, summaryValues } from '../logic/summaries.js';
import {
  MULTIPLE_PREVIOUS_VALUATION_DATES_WARNING,
  MULTIPLE_VALUATION_DATES_WARNING,
  resolvePortfolioValuation,
} from '../logic/fairValue.js';
import {
  attachPreviousMatches,
  indexPreviousPortfolio,
  overlaySummaryComparison,
  withRowComparison,
} from '../logic/comparison.js';

export function findDiagnosticFieldValue(parsedRow, sourceColumns, fieldNames) {
  var wanted = fieldNames.map(function (field) {
    return normaliseHeader(field);
  });
  var matches = sourceColumns.filter(function (column) {
    return wanted.indexOf(normaliseHeader(column.field)) !== -1;
  });
  if (matches.length === 0) return null;

  matches.sort(function (a, b) {
    var aGeneral = normaliseHeader(a.section) === 'general' ? 0 : 1;
    var bGeneral = normaliseHeader(b.section) === 'general' ? 0 : 1;
    if (aGeneral !== bGeneral) return aGeneral - bGeneral;
    if (a.occurrence !== b.occurrence) return a.occurrence - b.occurrence;
    return a.columnIndex - b.columnIndex;
  });

  var key = matches[0].key;
  var raw = parsedRow && parsedRow.values && parsedRow.values[key];
  if (raw === undefined || raw === null || raw === '') return null;
  return raw;
}

function referenceText(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(/\u00a0/g, ' ').trim();
}

export function assetUnitReference(address, unitName) {
  var left = referenceText(address);
  var right = referenceText(unitName);
  if (left && right) return left + '_' + right;
  return left || right || null;
}

export function getActiveOutputColumns(headingById, options) {
  return buildOutputColumns(headingById, options);
}

function matchedPreviousProperties(properties) {
  var matched = [];
  for (var i = 0; i < properties.length; i += 1) {
    if (properties[i].previousProperty) matched.push(properties[i].previousProperty);
  }
  return matched;
}

export function transformToSchedule(propertyData, unitData, options) {
  var errors = [];
  var warnings = [];
  var columns = getActiveOutputColumns();

  var emptySummary = {
    propertiesProcessed: 0,
    propertiesMatched: 0,
    propertiesWithoutMatchedUnits: 0,
    propertiesEmptyMatchValue: 0,
    unitRecordsProcessed: unitData && unitData.rows ? unitData.rows.length : 0,
    matchedUnitRecords: 0,
    unmatchedUnitRecords: 0,
    unitsEmptyMatchValue: 0,
    duplicatePropertyKeys: 0,
  };

  if (!propertyData) errors.push('Property Level workbook is not loaded.');
  if (!unitData) errors.push('Unit / TS Level workbook is not loaded.');
  if (errors.length) {
    return {
      ok: false,
      errors: errors,
      warnings: warnings,
      summary: emptySummary,
      columns: columns,
      scheduleRows: [],
      unmatchedRows: [],
    };
  }

  var portfolio = buildPortfolio(propertyData, unitData);

  if (portfolio.propertiesEmptyMatchValue > 0) {
    warnings.push(
      portfolio.propertiesEmptyMatchValue +
        ' Property row' +
        (portfolio.propertiesEmptyMatchValue === 1 ? ' has' : 's have') +
        ' no Property Name. Those rows are not included.'
    );
  }

  if (portfolio.duplicateDetails.length > 0) {
    warnings.push(
      portfolio.duplicateDetails.length +
        ' duplicate Property Name value(s) found. Matching Unit records for those names are treated as unmatched: ' +
        portfolio.duplicateDetails.slice(0, 15).join('; ') +
        (portfolio.duplicateDetails.length > 15
          ? ' \u2026and ' + (portfolio.duplicateDetails.length - 15) + ' more.'
          : '')
    );
  }

  if (portfolio.unitsEmptyMatchValue > 0) {
    warnings.push(
      portfolio.unitsEmptyMatchValue +
        ' Unit row' +
        (portfolio.unitsEmptyMatchValue === 1 ? ' has' : 's have') +
        ' no Property Name.'
    );
  }

  var previousProperty = options && options.previousProperty;
  var previousUnit = options && options.previousUnit;
  var previousValuation = { period: '', consistent: true };
  if (previousProperty) {
    var previousPortfolio = buildPortfolio(previousProperty, previousUnit || { rows: [], columns: [] });
    for (var pp = 0; pp < previousPortfolio.properties.length; pp += 1) {
      applyPropertyContext(previousPortfolio.properties[pp], DEFAULT_SCHEDULE_COLUMNS);
    }
    var previousIndex = indexPreviousPortfolio(previousPortfolio);
    if (previousIndex.duplicatePropertyRefs.length) {
      warnings.push('Duplicate previous Property Reference values were ignored for comparison.');
    }
    previousValuation = resolvePortfolioValuation(previousPortfolio.properties.map(function (property) {
      return property.standardValues && property.standardValues.valuation;
    }));
    if (!previousValuation.consistent) warnings.push(MULTIPLE_PREVIOUS_VALUATION_DATES_WARNING);
    attachPreviousMatches(portfolio.properties, previousIndex);
  }

  var requested = options && options.breakdown;
  var grouped = buildPortfolioGroups(portfolio.properties, {
    mode: requested && requested.mode ? requested.mode : getBreakdownMode(),
    propertyData: propertyData,
    customConfig: requested && requested.customConfig ? requested.customConfig : getCustomConfig(),
  });
  var scheduleRows = [];
  var allProperties = [];

  function remember(properties) {
    for (var prep = 0; prep < properties.length; prep += 1) {
      applyPropertyContext(properties[prep], DEFAULT_SCHEDULE_COLUMNS);
      allProperties.push(properties[prep]);
    }
  }

  function appendProperties(properties, groupLabel) {
    for (var i = 0; i < properties.length; i += 1) {
      var property = properties[i];
      scheduleRows.push({
        kind: 'property',
        group: groupLabel,
        sector: groupLabel,
        values: withRowComparison('property', property.standardValues, property.previousProperty),
        sourceRowIndex: property.sourceRow.rowIndex,
        matchKey: property.matchKey,
      });
      for (var u = 0; u < property.units.length; u += 1) {
        var unit = property.units[u];
        unit.standardValues.asset_unit_reference = assetUnitReference(
          property.standardValues && property.standardValues.address,
          unit.standardValues && unit.standardValues.unit
        );
        scheduleRows.push({
          kind: 'unit',
          group: groupLabel,
          sector: groupLabel,
          values: withRowComparison('unit', unit.standardValues, unit.previousUnit),
          sourceRowIndex: unit.sourceRow.rowIndex,
          matchKey: unit.matchKey,
        });
      }
    }
  }

  function appendGrandTotal() {
    scheduleRows.push({
      kind: 'grandTotal',
      values: overlaySummaryComparison(
        summaryValues(allProperties, DEFAULT_SCHEDULE_COLUMNS, 'grandTotal', null),
        (function () {
          var matched = matchedPreviousProperties(allProperties);
          return matched.length ? summaryValues(matched, DEFAULT_SCHEDULE_COLUMNS, 'grandTotal', null) : null;
        })()
      ),
    });
  }

  if (!grouped.effective) {
    remember(grouped.properties);
    for (var plain = 0; plain < DEFAULT_SCHEDULE_STRUCTURE.ungrouped.length; plain += 1) {
      var plainToken = DEFAULT_SCHEDULE_STRUCTURE.ungrouped[plain];
      if (plainToken === 'properties') appendProperties(grouped.properties, '');
      else if (plainToken === 'separator') scheduleRows.push({ kind: 'separator', values: {} });
      else if (plainToken === 'grandTotal') appendGrandTotal();
    }
  } else {
    for (var s = 0; s < grouped.groups.length; s += 1) {
      var group = grouped.groups[s];
      remember(group.properties);
      for (var step = 0; step < DEFAULT_SCHEDULE_STRUCTURE.perGroup.length; step += 1) {
        var token = DEFAULT_SCHEDULE_STRUCTURE.perGroup[step];
        if (token === 'groupHeader') {
          scheduleRows.push({ kind: 'group', group: group.label, sector: group.label, values: {} });
        } else if (token === 'properties') {
          appendProperties(group.properties, group.label);
        } else if (token === 'groupTotal') {
          scheduleRows.push({
            kind: 'groupTotal',
            group: group.label,
            sector: group.label,
            values: overlaySummaryComparison(
              calculateGroupTotals(group.properties, DEFAULT_SCHEDULE_COLUMNS, group.label),
              (function (members, label) {
                var matched = matchedPreviousProperties(members);
                return matched.length ? calculateGroupTotals(matched, DEFAULT_SCHEDULE_COLUMNS, label) : null;
              })(group.properties, group.label)
            ),
          });
        } else if (token === 'separator') {
          scheduleRows.push({ kind: 'separator', values: {} });
        }
      }
    }
    for (var close = 0; close < DEFAULT_SCHEDULE_STRUCTURE.afterGroups.length; close += 1) {
      if (DEFAULT_SCHEDULE_STRUCTURE.afterGroups[close] === 'grandTotal') appendGrandTotal();
    }
  }

  var unitNameCol = portfolio.unitNameColumn;
  var unmatchedRows = portfolio.orphanUnits.map(function (row) {
    var rowMatch = unitNameCol ? row.values[unitNameCol.key] : null;
    return {
      originalExcelRow: row.rowIndex + 1,
      matchValue: rowMatch === undefined || rowMatch === '' ? null : rowMatch,
      propertyName: findDiagnosticFieldValue(row, unitData.columns, [UNIT_SOURCE_FIELDS.propertyName.field]),
      tenantName: findDiagnosticFieldValue(row, unitData.columns, [UNIT_SOURCE_FIELDS.tenantName.field]),
      suite: findDiagnosticFieldValue(row, unitData.columns, [UNIT_SOURCE_FIELDS.suite.field]),
    };
  });

  if (unmatchedRows.length > 0) {
    warnings.push(
      unmatchedRows.length +
        ' Unit / TS record(s) did not match any Property' +
        (portfolio.ambiguousUnitCount > 0
          ? ' (including ' + portfolio.ambiguousUnitCount + ' ambiguous due to duplicate property names)'
          : '') +
        ' and will be included on the "Unmatched Records" sheet.'
    );
  }

  var valuation = resolvePortfolioValuation(
    allProperties.map(function (property) {
      return property.standardValues && property.standardValues.valuation;
    })
  );
  if (!valuation.consistent) warnings.push(MULTIPLE_VALUATION_DATES_WARNING);
  columns = getActiveOutputColumns({ fair_value: valuation.heading }, {
    previousProperty: !!previousProperty,
    previousUnit: !!previousUnit,
    currentPeriod: valuation.period || '',
    previousPeriod: previousValuation.period || '',
    previousDatesDiffer: !!(previousProperty && !previousValuation.consistent),
  });

  return {
    ok: true,
    errors: [],
    warnings: warnings,
    valuation: valuation,
    portfolio: portfolio,
    summary: {
      propertiesProcessed: portfolio.properties.length,
      propertiesMatched: portfolio.propertiesMatched,
      propertiesWithoutMatchedUnits: portfolio.properties.length - portfolio.propertiesMatched,
      propertiesEmptyMatchValue: portfolio.propertiesEmptyMatchValue,
      unitRecordsProcessed: unitData.rows.length,
      matchedUnitRecords: portfolio.matchedUnitCount,
      unmatchedUnitRecords: unmatchedRows.length,
      unitsEmptyMatchValue: portfolio.unitsEmptyMatchValue,
      duplicatePropertyKeys: portfolio.duplicateDetails.length,
    },
    columns: columns,
    scheduleRows: scheduleRows,
    unmatchedRows: unmatchedRows,
  };
}

export function analyseMatching(propertyData, unitData) {
  if (!propertyData || !unitData) return null;
  var model = transformToSchedule(propertyData, unitData);
  return {
    ok: model.ok,
    errors: model.errors,
    warnings: model.warnings,
    summary: model.summary,
  };
}

export function findExactSourceColumn(columns, section, field) {
  return findSourceColumn(columns, columns && columns[0] ? columns[0].fileType : 'unit', {
    section: section,
    field: field,
    occurrence: 1,
  });
}
