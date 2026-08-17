import { STANDARD_TEMPLATE_COLUMNS } from '../config/constants.js';
import {
  getMappingSummary,
  resolveSourceKey,
} from '../config/mapping.js';
import { normaliseHeader, normaliseJoinValue } from '../excel/parser.js';

export function resolveMappedValue(mapping, parsedRow, sourceColumns) {
  if (!mapping || mapping.type !== 'source' || !mapping.sourceKey) {
    return null;
  }

  var col = resolveSourceKey(sourceColumns, mapping.sourceKey);
  if (!col) return null;

  var raw = parsedRow && parsedRow.values && parsedRow.values[col.key];
  if (raw === undefined || raw === null || raw === '') return null;
  return raw;
}

export function mapRowValues(kind, mappingConfig, parsedRow, parsedWorkbook) {
  var values = {};
  var side = kind === 'property' ? 'propertyMapping' : 'unitMapping';
  var columns = parsedWorkbook.columns;

  for (var i = 0; i < STANDARD_TEMPLATE_COLUMNS.length; i += 1) {
    var templateCol = STANDARD_TEMPLATE_COLUMNS[i];
    var entry = null;
    for (var j = 0; j < mappingConfig.columnMappings.length; j += 1) {
      if (mappingConfig.columnMappings[j].targetColumnId === templateCol.id) {
        entry = mappingConfig.columnMappings[j];
        break;
      }
    }
    var mapping = (entry && entry[side]) || { type: 'blank' };
    values[templateCol.id] = resolveMappedValue(mapping, parsedRow, columns);
  }

  return values;
}

export function findDiagnosticFieldValue(parsedRow, sourceColumns, fieldNames) {
  var wanted = fieldNames.map(function (f) {
    return normaliseHeader(f);
  });
  var matches = sourceColumns.filter(function (c) {
    return wanted.indexOf(normaliseHeader(c.field)) !== -1;
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

export function rowHasMeaningfulData(row) {
  if (!row || !row.values) return false;
  var keys = Object.keys(row.values);
  for (var i = 0; i < keys.length; i += 1) {
    var value = row.values[keys[i]];
    if (value === null || value === undefined) continue;
    if (typeof value === 'string' && value.trim() === '') continue;
    return true;
  }
  return false;
}

export function getActiveOutputColumns(mappingConfig) {
  var columns = [];
  if (!mappingConfig || !mappingConfig.columnMappings) return columns;

  for (var i = 0; i < STANDARD_TEMPLATE_COLUMNS.length; i += 1) {
    var templateCol = STANDARD_TEMPLATE_COLUMNS[i];
    var entry = null;
    for (var j = 0; j < mappingConfig.columnMappings.length; j += 1) {
      if (mappingConfig.columnMappings[j].targetColumnId === templateCol.id) {
        entry = mappingConfig.columnMappings[j];
        break;
      }
    }

    if (entry && entry.ignore === true) continue;

    columns.push({ id: templateCol.id, label: templateCol.label });
  }

  return columns;
}

export function validateForTransform(propertyData, unitData, mappingConfig) {
  var errors = [];
  var warnings = [];

  if (!propertyData) errors.push('Property Level workbook is not loaded.');
  if (!unitData) errors.push('Unit / TS Level workbook is not loaded.');
  if (!mappingConfig) errors.push('Mapping configuration is not available.');

  if (errors.length) {
    return { errors: errors, warnings: warnings, propertyMatchCol: null, unitMatchCol: null };
  }

  // Incomplete / unresolved mappings are warnings only - never block generation.
  var summary = getMappingSummary(mappingConfig, propertyData, unitData);
  if (summary.missingMappings > 0) {
    warnings.push(
      summary.missingMappings +
        ' configured source mapping(s) could not be resolved against the uploaded files. Those fields will be left blank.'
    );
  }

  var propertyMatchCol = null;
  var unitMatchCol = null;

  if (!mappingConfig.propertyMatch || mappingConfig.propertyMatch.type !== 'source') {
    warnings.push(
      'Property Level Match Field is not configured. Property summary rows will still be written, but Unit rows cannot be matched.'
    );
  } else {
    propertyMatchCol = resolveSourceKey(
      propertyData.columns,
      mappingConfig.propertyMatch.sourceKey
    );
    if (!propertyMatchCol) {
      warnings.push(
        'Property Level Match Field source could not be resolved. Property summary rows will still be written, but Unit rows cannot be matched.'
      );
    }
  }

  if (!mappingConfig.unitMatch || mappingConfig.unitMatch.type !== 'source') {
    warnings.push(
      'Unit / TS Match Field is not configured. Unit rows cannot be matched underneath properties.'
    );
  } else {
    unitMatchCol = resolveSourceKey(unitData.columns, mappingConfig.unitMatch.sourceKey);
    if (!unitMatchCol) {
      warnings.push(
        'Unit / TS Match Field source could not be resolved. Unit rows cannot be matched underneath properties.'
      );
    }
  }

  return {
    errors: errors,
    warnings: warnings,
    propertyMatchCol: propertyMatchCol,
    unitMatchCol: unitMatchCol,
  };
}

export function transformToSchedule(propertyData, unitData, mappingConfig) {
  var validation = validateForTransform(propertyData, unitData, mappingConfig);
  var errors = validation.errors;
  var warnings = validation.warnings.slice();
  var propertyMatchCol = validation.propertyMatchCol;
  var unitMatchCol = validation.unitMatchCol;

  var columns = getActiveOutputColumns(mappingConfig);

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

  // Valid property data rows only (ignore trailing blanks / empty non-data rows).
  var propertyRows = [];
  for (var pri = 0; pri < propertyData.rows.length; pri += 1) {
    if (rowHasMeaningfulData(propertyData.rows[pri])) {
      propertyRows.push(propertyData.rows[pri]);
    }
  }

  var keyIndex = new Map();
  var propertiesEmptyMatchValue = 0;

  if (propertyMatchCol) {
    for (var pki = 0; pki < propertyRows.length; pki += 1) {
      var prow = propertyRows[pki];
      var rawKey = prow.values[propertyMatchCol.key];
      var normKey = normaliseJoinValue(rawKey);
      if (!normKey) {
        propertiesEmptyMatchValue += 1;
        continue;
      }
      var keyEntry = keyIndex.get(normKey);
      if (!keyEntry) {
        keyEntry = { count: 0, displayValues: new Set(), rowNumbers: [] };
        keyIndex.set(normKey, keyEntry);
      }
      keyEntry.count += 1;
      keyEntry.rowNumbers.push(prow.rowIndex + 1);
      if (rawKey !== null && rawKey !== undefined && String(rawKey).trim() !== '') {
        keyEntry.displayValues.add(String(rawKey).trim());
      }
    }
  }

  var duplicateKeys = new Set();
  var duplicateDetails = [];
  keyIndex.forEach(function (v, k) {
    if (v.count > 1) {
      duplicateKeys.add(k);
      var label = v.displayValues.values().next().value || k;
      duplicateDetails.push('"' + label + '" (rows ' + v.rowNumbers.join(', ') + ')');
    }
  });

  if (propertiesEmptyMatchValue > 0) {
    warnings.push(
      propertiesEmptyMatchValue +
        ' Property row' +
        (propertiesEmptyMatchValue === 1 ? ' has' : 's have') +
        ' no match value. Their Property Summary rows will still be written with no Unit rows attached.'
    );
  }

  if (duplicateDetails.length > 0) {
    warnings.push(
      duplicateDetails.length +
        ' duplicate Property match key(s) found. Matching Unit records for those keys are treated as unmatched/ambiguous: ' +
        duplicateDetails.slice(0, 15).join('; ') +
        (duplicateDetails.length > 15
          ? ' \u2026and ' + (duplicateDetails.length - 15) + ' more.'
          : '')
    );
  }

  var unitsByKey = new Map();
  var unmatchedSourceRows = [];
  var unitsEmptyMatchValue = 0;
  var ambiguousUnitCount = 0;

  for (var ui = 0; ui < unitData.rows.length; ui += 1) {
    var unitRow = unitData.rows[ui];

    if (!unitMatchCol || !propertyMatchCol) {
      // Match fields unavailable/unresolved - cannot attach any units.
      unmatchedSourceRows.push(unitRow);
      continue;
    }

    var rawMatch = unitRow.values[unitMatchCol.key];
    var unitKey = normaliseJoinValue(rawMatch);

    if (!unitKey) {
      unitsEmptyMatchValue += 1;
      unmatchedSourceRows.push(unitRow);
      continue;
    }

    // Ambiguous when multiple properties share this key - do not attach to any.
    if (duplicateKeys.has(unitKey)) {
      ambiguousUnitCount += 1;
      unmatchedSourceRows.push(unitRow);
      continue;
    }

    var list = unitsByKey.get(unitKey);
    if (!list) {
      list = [];
      unitsByKey.set(unitKey, list);
    }
    list.push(unitRow);
  }

  if (unitsEmptyMatchValue > 0) {
    warnings.push(
      unitsEmptyMatchValue +
        ' Unit row' +
        (unitsEmptyMatchValue === 1 ? ' has' : 's have') +
        ' no match value.'
    );
  }

  var scheduleRows = [];
  var matchedUnitSet = new Set();
  var propertiesMatched = 0;

  for (var pi = 0; pi < propertyRows.length; pi += 1) {
    var propertyRow = propertyRows[pi];
    var propKey = '';
    if (propertyMatchCol) {
      propKey = normaliseJoinValue(propertyRow.values[propertyMatchCol.key]) || '';
    }

    scheduleRows.push({
      kind: 'property',
      values: mapRowValues('property', mappingConfig, propertyRow, propertyData),
      sourceRowIndex: propertyRow.rowIndex,
      matchKey: propKey,
    });

    var canAttach =
      Boolean(propKey) && !duplicateKeys.has(propKey) && unitsByKey.has(propKey);
    var units = canAttach ? unitsByKey.get(propKey) : [];

    if (units.length > 0) propertiesMatched += 1;

    for (var uj = 0; uj < units.length; uj += 1) {
      var matchedUnitRow = units[uj];
      matchedUnitSet.add(matchedUnitRow);
      scheduleRows.push({
        kind: 'unit',
        values: mapRowValues('unit', mappingConfig, matchedUnitRow, unitData),
        sourceRowIndex: matchedUnitRow.rowIndex,
        matchKey: propKey,
      });
    }
  }

  for (var uk = 0; uk < unitData.rows.length; uk += 1) {
    var checkRow = unitData.rows[uk];
    if (!matchedUnitSet.has(checkRow) && unmatchedSourceRows.indexOf(checkRow) === -1) {
      unmatchedSourceRows.push(checkRow);
    }
  }

  unmatchedSourceRows.sort(function (a, b) {
    return a.rowIndex - b.rowIndex;
  });

  var unmatchedRows = unmatchedSourceRows.map(function (row) {
    var rowMatch = unitMatchCol ? row.values[unitMatchCol.key] : null;
    return {
      originalExcelRow: row.rowIndex + 1,
      matchValue: rowMatch === undefined || rowMatch === '' ? null : rowMatch,
      propertyName: findDiagnosticFieldValue(row, unitData.columns, ['Property Name']),
      tenantName: findDiagnosticFieldValue(row, unitData.columns, ['Tenant Name']),
      suite: findDiagnosticFieldValue(row, unitData.columns, ['Suite']),
    };
  });

  if (unmatchedRows.length > 0) {
    warnings.push(
      unmatchedRows.length +
        ' Unit / TS record(s) did not match any Property' +
        (ambiguousUnitCount > 0
          ? ' (including ' + ambiguousUnitCount + ' ambiguous due to duplicate property keys)'
          : '') +
        ' and will be included on the "Unmatched Records" sheet.'
    );
  }

  var propertiesWithoutMatchedUnits = propertyRows.length - propertiesMatched;

  return {
    ok: true,
    errors: [],
    warnings: warnings,
    summary: {
      propertiesProcessed: propertyRows.length,
      propertiesMatched: propertiesMatched,
      propertiesWithoutMatchedUnits: propertiesWithoutMatchedUnits,
      propertiesEmptyMatchValue: propertiesEmptyMatchValue,
      unitRecordsProcessed: unitData.rows.length,
      matchedUnitRecords: matchedUnitSet.size,
      unmatchedUnitRecords: unmatchedRows.length,
      unitsEmptyMatchValue: unitsEmptyMatchValue,
      duplicatePropertyKeys: duplicateDetails.length,
    },
    columns: columns,
    scheduleRows: scheduleRows,
    unmatchedRows: unmatchedRows,
  };
}

/**
 * Lightweight matching analysis for the pre-generation Matching Summary UI.
 * Does not mutate configuration. Uses the same rules as transformToSchedule.
 */
export function analyseMatching(propertyData, unitData, mappingConfig) {
  if (!propertyData || !unitData || !mappingConfig) return null;
  var model = transformToSchedule(propertyData, unitData, mappingConfig);
  return {
    ok: model.ok,
    errors: model.errors,
    warnings: model.warnings,
    summary: model.summary,
  };
}
