import { normaliseJoinValue } from '../schema/text.js';
import { findSourceColumn, readSourceValue } from '../schema/sourceLookup.js';
import { DEFAULT_SCHEDULE_COLUMNS, PROPERTY_UNIT_JOIN } from '../templates/defaultSchedule.js';

function outputRule(column, kind) {
  if (!column.logic) return null;
  return kind === 'property' ? column.logic.property : column.logic.unit;
}

export function standardValuesFor(kind, parsedRow, parsedWorkbook) {
  var fileType = kind === 'property' ? 'property' : 'unit';
  var values = {};
  for (var i = 0; i < DEFAULT_SCHEDULE_COLUMNS.length; i += 1) {
    var col = DEFAULT_SCHEDULE_COLUMNS[i];
    var rule = outputRule(col, kind);
    if (rule && (rule.method === 'source' || rule.method === 'aggregatedSource')) {
      values[col.id] = readSourceValue(parsedRow, parsedWorkbook.columns, fileType, rule.ref);
    } else if (rule && rule.method === 'fixed') {
      values[col.id] = rule.value;
    } else {
      values[col.id] = null;
    }
  }
  return values;
}

function rawJoinValue(parsedRow, column) {
  if (!column || !parsedRow || !parsedRow.values) return null;
  var raw = parsedRow.values[column.key];
  if (raw === undefined || raw === '') return null;
  return raw;
}

export function buildPortfolio(propertyData, unitData) {
  var propertyRows = (propertyData && propertyData.rows) || [];
  var unitRows = (unitData && unitData.rows) || [];
  var propertyNameCol = propertyData
    ? findSourceColumn(propertyData.columns, 'property', PROPERTY_UNIT_JOIN)
    : null;
  var unitNameCol = unitData
    ? findSourceColumn(unitData.columns, 'unit', PROPERTY_UNIT_JOIN)
    : null;

  var keyIndex = new Map();
  var propertiesEmptyMatchValue = 0;

  for (var pki = 0; pki < propertyRows.length; pki += 1) {
    var prow = propertyRows[pki];
    var rawKey = rawJoinValue(prow, propertyNameCol);
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

  var duplicateKeys = new Set();
  var duplicateDetails = [];
  keyIndex.forEach(function (value, key) {
    if (value.count > 1) {
      duplicateKeys.add(key);
      var label = value.displayValues.values().next().value || key;
      duplicateDetails.push('"' + label + '" (rows ' + value.rowNumbers.join(', ') + ')');
    }
  });

  var unitsByKey = new Map();
  var orphanUnits = [];
  var unitsEmptyMatchValue = 0;
  var ambiguousUnitCount = 0;

  for (var ui = 0; ui < unitRows.length; ui += 1) {
    var unitRow = unitRows[ui];
    var rawMatch = rawJoinValue(unitRow, unitNameCol);
    var unitKey = normaliseJoinValue(rawMatch);
    if (!unitKey) {
      unitsEmptyMatchValue += 1;
      orphanUnits.push(unitRow);
      continue;
    }
    if (duplicateKeys.has(unitKey)) {
      ambiguousUnitCount += 1;
      orphanUnits.push(unitRow);
      continue;
    }
    var list = unitsByKey.get(unitKey);
    if (!list) {
      list = [];
      unitsByKey.set(unitKey, list);
    }
    list.push(unitRow);
  }

  var properties = [];
  var matchedUnitSet = new Set();
  var propertiesMatched = 0;

  for (var pi = 0; pi < propertyRows.length; pi += 1) {
    var propertyRow = propertyRows[pi];
    var propKey = normaliseJoinValue(rawJoinValue(propertyRow, propertyNameCol)) || '';
    var canAttach = Boolean(propKey) && !duplicateKeys.has(propKey) && unitsByKey.has(propKey);
    var attached = canAttach ? unitsByKey.get(propKey) : [];
    if (attached.length > 0) propertiesMatched += 1;

    var units = [];
    for (var uj = 0; uj < attached.length; uj += 1) {
      var matchedUnitRow = attached[uj];
      matchedUnitSet.add(matchedUnitRow);
      units.push({
        sourceRow: matchedUnitRow,
        matchKey: propKey,
        sourceValues: matchedUnitRow.values,
        standardValues: standardValuesFor('unit', matchedUnitRow, unitData),
      });
    }

    properties.push({
      sourceRow: propertyRow,
      matchKey: propKey,
      sourceValues: propertyRow.values,
      standardValues: standardValuesFor('property', propertyRow, propertyData),
      units: units,
    });
  }

  for (var uk = 0; uk < unitRows.length; uk += 1) {
    var checkRow = unitRows[uk];
    if (!matchedUnitSet.has(checkRow) && orphanUnits.indexOf(checkRow) === -1) {
      orphanUnits.push(checkRow);
    }
  }

  orphanUnits.sort(function (a, b) {
    return a.rowIndex - b.rowIndex;
  });

  return {
    properties: properties,
    orphanUnits: orphanUnits,
    propertyNameColumn: propertyNameCol,
    unitNameColumn: unitNameCol,
    duplicateKeys: duplicateKeys,
    duplicateDetails: duplicateDetails,
    propertiesEmptyMatchValue: propertiesEmptyMatchValue,
    unitsEmptyMatchValue: unitsEmptyMatchValue,
    ambiguousUnitCount: ambiguousUnitCount,
    propertiesMatched: propertiesMatched,
    matchedUnitCount: matchedUnitSet.size,
  };
}
