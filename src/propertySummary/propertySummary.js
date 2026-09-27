import { readSourceValue } from '../schema/sourceLookup.js';
import { PROPERTY_ONLY_COLUMNS, PROPERTY_ONLY_OUTPUT_FILENAME } from '../templates/propertyOnly.js';
import { buildPropertySummaryWorkbook } from './propertySummaryExport.js';
import { downloadWorkbook } from '../excel/writer.js';

export function getActivePropertySummaryColumns() {
  return PROPERTY_ONLY_COLUMNS.map(function (column) {
    return {
      id: column.id,
      label: column.label,
      property: column.property,
    };
  });
}

export function transformToPropertySummary(propertyData) {
  var errors = [];
  var warnings = [];

  if (!propertyData) errors.push('Property Level workbook is not loaded.');
  if (errors.length) {
    return { ok: false, errors: errors, warnings: warnings, rows: [], columns: [], summary: null };
  }

  var columns = getActivePropertySummaryColumns();
  var rows = [];
  var sourceRows = propertyData.rows || [];

  for (var i = 0; i < sourceRows.length; i += 1) {
    var parsedRow = sourceRows[i];
    var values = {};
    for (var c = 0; c < columns.length; c += 1) {
      var col = columns[c];
      values[col.id] = readSourceValue(parsedRow, propertyData.columns, 'property', col.property);
    }
    rows.push({
      sourceRowIndex: parsedRow.rowIndex,
      values: values,
    });
  }

  return {
    ok: true,
    errors: [],
    warnings: warnings,
    rows: rows,
    columns: columns,
    summary: {
      propertiesProcessed: rows.length,
      columnsWritten: columns.length,
      sourceRows: sourceRows.length,
    },
  };
}

export function generatePropertySummary(propertyData) {
  var model = transformToPropertySummary(propertyData);
  if (!model.ok) {
    return {
      ok: false,
      errors: model.errors,
      warnings: model.warnings,
      summary: model.summary,
    };
  }

  try {
    var workbook = buildPropertySummaryWorkbook(model);
    downloadWorkbook(workbook, PROPERTY_ONLY_OUTPUT_FILENAME);
    return {
      ok: true,
      errors: [],
      warnings: model.warnings,
      summary: model.summary,
      filename: PROPERTY_ONLY_OUTPUT_FILENAME,
    };
  } catch (err) {
    console.error('Property Summary workbook generation failed:', err);
    return {
      ok: false,
      errors: ['The Property Only workbook could not be created. Please try again.'],
      warnings: model.warnings,
      summary: model.summary,
    };
  }
}
