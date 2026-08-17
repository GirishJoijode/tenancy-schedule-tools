import {
  COMPARE_IDENTITY,
  COMPARE_NUMERIC_EPS,
  DATE_COMPARE_IDS,
  STANDARD_TEMPLATE_COLUMNS,
} from '../config/constants.js';
import { findExactSourceColumn } from '../config/mapping.js';
import { normaliseJoinValue } from '../excel/parser.js';
import { resolveMappedValue } from '../standardiser/transform.js';

export function getUnitMappedCompareColumns(mappingConfig) {
  var cols = [];
  if (!mappingConfig || !mappingConfig.columnMappings) return cols;
  for (var i = 0; i < STANDARD_TEMPLATE_COLUMNS.length; i += 1) {
    var templateCol = STANDARD_TEMPLATE_COLUMNS[i];
    var entry = null;
    for (var j = 0; j < mappingConfig.columnMappings.length; j += 1) {
      if (mappingConfig.columnMappings[j].targetColumnId === templateCol.id) {
        entry = mappingConfig.columnMappings[j];
        break;
      }
    }
    if (entry && entry.unitMapping && entry.unitMapping.type === 'source') {
      cols.push({
        id: templateCol.id,
        label: templateCol.label,
        unitMapping: entry.unitMapping,
      });
    }
  }
  return cols;
}

export function isBlankCompareValue(value) {
  if (value === null || value === undefined) return true;
  if (typeof value === 'string' && value.trim() === '') return true;
  return false;
}

function excelSerialToUtcDate(serial) {
  // Excel epoch 1899-12-30
  var ms = Math.round((Number(serial) - 25569) * 86400 * 1000);
  var d = new Date(ms);
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
}

export function toCompareDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
  }
  if (typeof value === 'number' && Number.isFinite(value) && value > 20000 && value < 80000) {
    return excelSerialToUtcDate(value);
  }
  if (typeof value === 'string' && value.trim()) {
    var parsed = Date.parse(value);
    if (!Number.isNaN(parsed)) {
      var d = new Date(parsed);
      return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    }
  }
  return null;
}

export function compareValuesEqual(a, b, columnId) {
  var aBlank = isBlankCompareValue(a);
  var bBlank = isBlankCompareValue(b);
  if (aBlank && bBlank) return true;
  if (aBlank || bBlank) return false;

  if (DATE_COMPARE_IDS[columnId]) {
    var da = toCompareDate(a);
    var db = toCompareDate(b);
    if (da && db) return da.getTime() === db.getTime();
  }

  var aDate = a instanceof Date ? toCompareDate(a) : null;
  var bDate = b instanceof Date ? toCompareDate(b) : null;
  if (aDate && bDate) return aDate.getTime() === bDate.getTime();

  var aNum = typeof a === 'number' && Number.isFinite(a);
  var bNum = typeof b === 'number' && Number.isFinite(b);
  if (aNum && bNum) {
    var diff = Math.abs(a - b);
    var scale = Math.max(1, Math.abs(a), Math.abs(b));
    return diff <= COMPARE_NUMERIC_EPS * scale;
  }

  // Text vs number: allow numeric text equivalence for display-only differences
  if (aNum || bNum) {
    var na = aNum ? a : Number(String(a).replace(/[£$,\s]/g, '').replace(/,/g, ''));
    var nb = bNum ? b : Number(String(b).replace(/[£$,\s]/g, '').replace(/,/g, ''));
    if (Number.isFinite(na) && Number.isFinite(nb)) {
      var d2 = Math.abs(na - nb);
      var s2 = Math.max(1, Math.abs(na), Math.abs(nb));
      return d2 <= COMPARE_NUMERIC_EPS * s2;
    }
  }

  var sa = normaliseJoinValue(a);
  var sb = normaliseJoinValue(b);
  return sa === sb;
}

export function formatCompareDisplay(value, columnId) {
  if (isBlankCompareValue(value)) return '';
  if (DATE_COMPARE_IDS[columnId] || value instanceof Date) {
    var d = toCompareDate(value);
    if (d) {
      var dd = String(d.getUTCDate()).padStart(2, '0');
      var mm = String(d.getUTCMonth() + 1).padStart(2, '0');
      var yyyy = d.getUTCFullYear();
      return dd + '/' + mm + '/' + yyyy;
    }
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    if (columnId === 'cap_rate' || columnId.indexOf('pct_') === 0 || columnId === 'niy' || columnId === 'ney' || columnId === 'ry') {
      return (value * 100).toFixed(2) + '%';
    }
    if (Math.abs(value) >= 1000 && Math.floor(value) === value) {
      return value.toLocaleString('en-GB');
    }
    return String(value);
  }
  return String(value);
}

function buildUnitCompareRecord(parsedRow, parsedWorkbook, compareColumns, propertyCol, unitCol) {
  var propertyName = propertyCol ? parsedRow.values[propertyCol.key] : null;
  var unitName = unitCol ? parsedRow.values[unitCol.key] : null;
  var values = {};
  for (var i = 0; i < compareColumns.length; i += 1) {
    var col = compareColumns[i];
    values[col.id] = resolveMappedValue(col.unitMapping, parsedRow, parsedWorkbook.columns);
  }
  var propKey = normaliseJoinValue(propertyName);
  var unitKey = normaliseJoinValue(unitName);
  var incompleteIdentity = !propKey || !unitKey;
  // Incomplete identities get a unique key so they never silently match another row.
  var key = incompleteIdentity
    ? '__incomplete__|' + parsedRow.rowIndex + '|' + propKey + '||' + unitKey
    : propKey + '||' + unitKey;
  return {
    sourceRowIndex: parsedRow.rowIndex,
    propertyName: propertyName,
    unitName: unitName,
    key: key,
    incompleteIdentity: incompleteIdentity,
    values: values,
    tenantName: values.tenant_name != null ? values.tenant_name : null,
  };
}

function indexCompareRecords(records) {
  var map = new Map();
  var duplicates = new Set();
  for (var i = 0; i < records.length; i += 1) {
    var rec = records[i];
    if (rec.incompleteIdentity) {
      map.set(rec.key, [rec]);
      duplicates.add(rec.key);
      continue;
    }
    var list = map.get(rec.key);
    if (!list) {
      map.set(rec.key, [rec]);
    } else {
      list.push(rec);
      duplicates.add(rec.key);
    }
  }
  return { map: map, duplicates: duplicates };
}

export function runTenancyCompare(currentParsed, previousParsed, mappingConfig) {
  var errors = [];
  var warnings = [];
  if (!currentParsed) errors.push('Current TS workbook is not loaded.');
  if (!previousParsed) errors.push('Previous TS workbook is not loaded.');
  if (!mappingConfig) errors.push('Mapping configuration is not available.');
  if (errors.length) {
    return { ok: false, errors: errors, warnings: warnings, pairs: [], summary: null, columns: [] };
  }

  var propertyColCurrent = findExactSourceColumn(
    currentParsed.columns,
    COMPARE_IDENTITY.property.section,
    COMPARE_IDENTITY.property.field
  );
  var unitColCurrent = findExactSourceColumn(
    currentParsed.columns,
    COMPARE_IDENTITY.unit.section,
    COMPARE_IDENTITY.unit.field
  );
  var propertyColPrevious = findExactSourceColumn(
    previousParsed.columns,
    COMPARE_IDENTITY.property.section,
    COMPARE_IDENTITY.property.field
  );
  var unitColPrevious = findExactSourceColumn(
    previousParsed.columns,
    COMPARE_IDENTITY.unit.section,
    COMPARE_IDENTITY.unit.field
  );

  if (!propertyColCurrent) errors.push('Property Name field could not be identified in Current TS.');
  if (!unitColCurrent) errors.push('Suite field could not be identified in Current TS.');
  if (!propertyColPrevious) errors.push('Property Name field could not be identified in Previous TS.');
  if (!unitColPrevious) errors.push('Suite field could not be identified in Previous TS.');
  if (errors.length) {
    return { ok: false, errors: errors, warnings: warnings, pairs: [], summary: null, columns: [] };
  }

  var compareColumns = getUnitMappedCompareColumns(mappingConfig);
  if (compareColumns.length === 0) {
    errors.push('No Unit / TS mappings are configured in Advanced Settings, so there are no fields to compare.');
    return { ok: false, errors: errors, warnings: warnings, pairs: [], summary: null, columns: [] };
  }

  var currentRecords = currentParsed.rows.map(function (row) {
    return buildUnitCompareRecord(row, currentParsed, compareColumns, propertyColCurrent, unitColCurrent);
  });
  var previousRecords = previousParsed.rows.map(function (row) {
    return buildUnitCompareRecord(row, previousParsed, compareColumns, propertyColPrevious, unitColPrevious);
  });

  var currentIndex = indexCompareRecords(currentRecords);
  var previousIndex = indexCompareRecords(previousRecords);

  if (currentIndex.duplicates.size > 0) {
    warnings.push(
      currentIndex.duplicates.size +
        ' ambiguous Property Name + Suite key(s) found in Current TS.'
    );
  }
  if (previousIndex.duplicates.size > 0) {
    warnings.push(
      previousIndex.duplicates.size +
        ' ambiguous Property Name + Suite key(s) found in Previous TS.'
    );
  }

  var pairs = [];
  var seenPrev = new Set();
  var propertyOrder = [];
  var propertySeen = {};

  function noteProperty(name) {
    var key = normaliseJoinValue(name) || '(blank property)';
    if (!propertySeen[key]) {
      propertySeen[key] = true;
      propertyOrder.push(name == null || name === '' ? '(Blank Property)' : String(name));
    }
    return key;
  }

  // Walk Current order
  for (var ci = 0; ci < currentRecords.length; ci += 1) {
    var cur = currentRecords[ci];
    noteProperty(cur.propertyName);
    var curList = currentIndex.map.get(cur.key) || [cur];
    var prevList = previousIndex.map.get(cur.key) || [];
    var ambiguous =
      currentIndex.duplicates.has(cur.key) || previousIndex.duplicates.has(cur.key);

    // Only emit once per key from the first current occurrence
    if (curList[0] !== cur) continue;

    if (ambiguous) {
      pairs.push({
        status: 'ambiguous',
        key: cur.key,
        propertyName: cur.propertyName,
        unitName: cur.unitName,
        propertyKey: normaliseJoinValue(cur.propertyName) || '(blank property)',
        changedFields: [],
        changedLabels: [],
        current: cur,
        previous: prevList[0] || null,
      });
      if (prevList[0]) seenPrev.add(cur.key);
      continue;
    }

    if (!prevList.length) {
      pairs.push({
        status: 'new',
        key: cur.key,
        propertyName: cur.propertyName,
        unitName: cur.unitName,
        propertyKey: normaliseJoinValue(cur.propertyName) || '(blank property)',
        changedFields: [],
        changedLabels: ['New unit'],
        current: cur,
        previous: null,
      });
      continue;
    }

    var prev = prevList[0];
    seenPrev.add(cur.key);
    var changedFields = [];
    var changedLabels = [];
    for (var k = 0; k < compareColumns.length; k += 1) {
      var col = compareColumns[k];
      var cv = cur.values[col.id];
      var pv = prev.values[col.id];
      if (!compareValuesEqual(cv, pv, col.id)) {
        changedFields.push(col.id);
        changedLabels.push(col.label);
      }
    }
    pairs.push({
      status: changedFields.length ? 'changed' : 'unchanged',
      key: cur.key,
      propertyName: cur.propertyName,
      unitName: cur.unitName,
      propertyKey: normaliseJoinValue(cur.propertyName) || '(blank property)',
      changedFields: changedFields,
      changedLabels: changedLabels,
      current: cur,
      previous: prev,
    });
  }

  // Removed / leftover previous keys
  for (var pi = 0; pi < previousRecords.length; pi += 1) {
    var prevRec = previousRecords[pi];
    var prevList2 = previousIndex.map.get(prevRec.key) || [prevRec];
    if (prevList2[0] !== prevRec) continue;
    if (seenPrev.has(prevRec.key)) continue;
    if (currentIndex.map.has(prevRec.key)) continue;

    noteProperty(prevRec.propertyName);
    var amb =
      previousIndex.duplicates.has(prevRec.key) || currentIndex.duplicates.has(prevRec.key);
    pairs.push({
      status: amb ? 'ambiguous' : 'removed',
      key: prevRec.key,
      propertyName: prevRec.propertyName,
      unitName: prevRec.unitName,
      propertyKey: normaliseJoinValue(prevRec.propertyName) || '(blank property)',
      changedFields: [],
      changedLabels: amb ? [] : ['Removed unit'],
      current: null,
      previous: prevRec,
    });
  }

  // Sort: property by first-seen order, then current order / removed after
  var propertyRank = {};
  for (var pr = 0; pr < propertyOrder.length; pr += 1) {
    propertyRank[normaliseJoinValue(propertyOrder[pr]) || '(blank property)'] = pr;
  }
  pairs.sort(function (a, b) {
    var ra = propertyRank[a.propertyKey] != null ? propertyRank[a.propertyKey] : 99999;
    var rb = propertyRank[b.propertyKey] != null ? propertyRank[b.propertyKey] : 99999;
    if (ra !== rb) return ra - rb;
    var ua = normaliseJoinValue(a.unitName);
    var ub = normaliseJoinValue(b.unitName);
    if (ua < ub) return -1;
    if (ua > ub) return 1;
    return 0;
  });

  var summary = {
    currentUnits: currentRecords.length,
    previousUnits: previousRecords.length,
    matched: 0,
    changed: 0,
    unchanged: 0,
    newUnits: 0,
    removedUnits: 0,
    ambiguous: 0,
  };
  for (var s = 0; s < pairs.length; s += 1) {
    var st = pairs[s].status;
    if (st === 'changed') {
      summary.changed += 1;
      summary.matched += 1;
    } else if (st === 'unchanged') {
      summary.unchanged += 1;
      summary.matched += 1;
    } else if (st === 'new') summary.newUnits += 1;
    else if (st === 'removed') summary.removedUnits += 1;
    else if (st === 'ambiguous') summary.ambiguous += 1;
  }

  return {
    ok: true,
    errors: [],
    warnings: warnings,
    pairs: pairs,
    summary: summary,
    columns: compareColumns,
    propertyOrder: propertyOrder,
  };
}
