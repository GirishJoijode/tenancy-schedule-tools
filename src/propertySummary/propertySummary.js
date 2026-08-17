import {
  PROPERTY_SUMMARY_COLUMNS,
  PROPERTY_SUMMARY_OUTPUT_FILENAME,
} from '../config/propertySummaryColumns.js';
import { resolveMappedValue, rowHasMeaningfulData } from '../standardiser/transform.js';
import { buildPropertySummaryWorkbook } from './propertySummaryExport.js';
import { downloadWorkbook } from '../excel/writer.js';

export function getActivePropertySummaryColumns(psConfig) {
  var columns = [];
  var mappings = (psConfig && psConfig.columnMappings) || [];

  for (var i = 0; i < PROPERTY_SUMMARY_COLUMNS.length; i += 1) {
    var templateCol = PROPERTY_SUMMARY_COLUMNS[i];
    var entry = null;
    for (var j = 0; j < mappings.length; j += 1) {
      if (mappings[j].targetColumnId === templateCol.id) {
        entry = mappings[j];
        break;
      }
    }
    if (entry && entry.ignore === true) continue;
    columns.push({
      id: templateCol.id,
      label: templateCol.label,
      propertyMapping: (entry && entry.propertyMapping) || { type: 'blank' },
    });
  }

  return columns;
}

export function transformToPropertySummary(propertyData, psConfig) {
  var errors = [];
  var warnings = [];

  if (!propertyData) errors.push('Property Level workbook is not loaded.');
  if (!psConfig) errors.push('Property Only mapping configuration is not available.');
  if (errors.length) {
    return { ok: false, errors: errors, warnings: warnings, rows: [], columns: [], summary: null };
  }

  var columns = getActivePropertySummaryColumns(psConfig);
  if (columns.length === 0) {
    warnings.push('All Property Only fields are marked Ignore, so the workbook contains no data columns.');
  }

  var rows = [];
  var skippedEmpty = 0;
  var sourceRows = propertyData.rows || [];

  for (var i = 0; i < sourceRows.length; i += 1) {
    var parsedRow = sourceRows[i];
    if (!rowHasMeaningfulData(parsedRow)) {
      skippedEmpty += 1;
      continue;
    }

    var values = {};
    for (var c = 0; c < columns.length; c += 1) {
      var col = columns[c];
      values[col.id] = resolveMappedValue(col.propertyMapping, parsedRow, propertyData.columns);
    }
    rows.push({
      sourceRowIndex: parsedRow.rowIndex,
      values: values,
    });
  }

  if (skippedEmpty > 0) {
    warnings.push(
      skippedEmpty +
        ' empty Property row' +
        (skippedEmpty === 1 ? ' was' : 's were') +
        ' omitted.'
    );
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

export function generatePropertySummary(propertyData, psConfig) {
  var model = transformToPropertySummary(propertyData, psConfig);
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
    downloadWorkbook(workbook, PROPERTY_SUMMARY_OUTPUT_FILENAME);
    return {
      ok: true,
      errors: [],
      warnings: model.warnings,
      summary: model.summary,
      filename: PROPERTY_SUMMARY_OUTPUT_FILENAME,
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
