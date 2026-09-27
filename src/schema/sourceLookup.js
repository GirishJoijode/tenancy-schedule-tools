import { readPeriodGroupValue } from './periods.js';
import { buildSourceColumnKey, normaliseJoinValue } from './text.js';

export function findSourceColumn(columns, fileType, ref) {
  if (!columns || !ref) return null;
  var key = buildSourceColumnKey(fileType, ref.section, ref.field, ref.occurrence || 1);
  for (var i = 0; i < columns.length; i += 1) {
    if (columns[i].key === key) return columns[i];
  }
  return null;
}

export function readSourceValue(parsedRow, columns, fileType, ref) {
  if (!ref) return null;
  if (ref.type === 'periodGroup') return readPeriodGroupValue(parsedRow, columns, ref);
  var col = findSourceColumn(columns, fileType, ref);
  if (!col) return null;
  var raw = parsedRow && parsedRow.values && parsedRow.values[col.key];
  if (raw === undefined || raw === null || raw === '') return null;
  return raw;
}

export function readMatchKey(parsedRow, columns, fileType, ref) {
  var col = findSourceColumn(columns, fileType, ref);
  if (!col || !parsedRow || !parsedRow.values) return '';
  return normaliseJoinValue(parsedRow.values[col.key]) || '';
}

export function isBlankSourceValue(value) {
  if (value === null || value === undefined) return true;
  if (typeof value === 'string' && value.trim() === '') return true;
  return false;
}
