import { normaliseHeader } from '../schema/text.js';
import { findSourceColumn, readSourceValue } from '../schema/sourceLookup.js';
import { buildPortfolio } from '../model/portfolio.js';
import { UNIT_SOURCE_FIELDS } from '../schema/sourceFields.js';
import {
  DEFAULT_SCHEDULE_COLUMNS,
  getRenderedScheduleColumns,
  PROPERTY_SECTOR,
  UNCLASSIFIED_SECTOR,
} from '../templates/defaultSchedule.js';
import { DEFAULT_SCHEDULE_STRUCTURE } from '../logic/scheduleStructure.js';
import { applyPropertyContext, summaryValues } from '../logic/summaries.js';
import { MULTIPLE_VALUATION_DATES_WARNING, resolvePortfolioValuation } from '../logic/fairValue.js';

function sectorNameFor(property, propertyData) {
  var raw = readSourceValue(property.sourceRow, propertyData.columns, 'property', PROPERTY_SECTOR);
  if (raw === null || raw === undefined) return UNCLASSIFIED_SECTOR;
  var text = String(raw).replace(/\u00a0/g, ' ').trim();
  return text || UNCLASSIFIED_SECTOR;
}

function groupPropertiesBySector(properties, propertyData) {
  var order = [];
  var groups = new Map();
  for (var i = 0; i < properties.length; i += 1) {
    var property = properties[i];
    var sector = sectorNameFor(property, propertyData);
    if (!groups.has(sector)) {
      groups.set(sector, []);
      order.push(sector);
    }
    groups.get(sector).push(property);
  }
  return { order: order, groups: groups };
}

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

export function getActiveOutputColumns(headingById) {
  return getRenderedScheduleColumns(headingById);
}

export function transformToSchedule(propertyData, unitData) {
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

  var grouped = groupPropertiesBySector(portfolio.properties, propertyData);
  var scheduleRows = [];
  var allProperties = [];
  for (var s = 0; s < grouped.order.length; s += 1) {
    var sector = grouped.order[s];
    var sectorProperties = grouped.groups.get(sector);
    for (var prep = 0; prep < sectorProperties.length; prep += 1) {
      applyPropertyContext(sectorProperties[prep], DEFAULT_SCHEDULE_COLUMNS);
      allProperties.push(sectorProperties[prep]);
    }
    for (var step = 0; step < DEFAULT_SCHEDULE_STRUCTURE.perSector.length; step += 1) {
      var token = DEFAULT_SCHEDULE_STRUCTURE.perSector[step];
      if (token === 'sectorHeader') {
        scheduleRows.push({ kind: 'sector', sector: sector, values: {} });
      } else if (token === 'properties') {
        for (var i = 0; i < sectorProperties.length; i += 1) {
          var property = sectorProperties[i];
          scheduleRows.push({
            kind: 'property',
            sector: sector,
            values: property.standardValues,
            sourceRowIndex: property.sourceRow.rowIndex,
            matchKey: property.matchKey,
          });
          for (var u = 0; u < property.units.length; u += 1) {
            var unit = property.units[u];
            scheduleRows.push({
              kind: 'unit',
              sector: sector,
              values: unit.standardValues,
              sourceRowIndex: unit.sourceRow.rowIndex,
              matchKey: unit.matchKey,
            });
          }
        }
      } else if (token === 'sectorTotal') {
        scheduleRows.push({
          kind: 'sectorTotal',
          sector: sector,
          values: summaryValues(sectorProperties, DEFAULT_SCHEDULE_COLUMNS, 'sectorTotal', sector),
        });
      } else if (token === 'separator') {
        scheduleRows.push({ kind: 'separator', values: {} });
      }
    }
  }
  for (var close = 0; close < DEFAULT_SCHEDULE_STRUCTURE.afterSectors.length; close += 1) {
    if (DEFAULT_SCHEDULE_STRUCTURE.afterSectors[close] === 'grandTotal') {
      scheduleRows.push({
        kind: 'grandTotal',
        values: summaryValues(allProperties, DEFAULT_SCHEDULE_COLUMNS, 'grandTotal', null),
      });
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
  columns = getActiveOutputColumns({ fair_value: valuation.heading });

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
