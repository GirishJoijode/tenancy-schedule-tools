import { OUTPUT_FILENAME } from '../config/constants.js';
import { buildWorkbookFromModel, downloadWorkbook } from '../excel/writer.js';
import { transformToSchedule } from './transform.js';

export function generateStandardisedSchedule(propertyData, unitData) {
  var model = transformToSchedule(propertyData, unitData);

  if (!model.ok) {
    return {
      ok: false,
      errors: model.errors,
      warnings: model.warnings,
      summary: model.summary,
    };
  }

  try {
    var workbook = buildWorkbookFromModel(model);
    downloadWorkbook(workbook, OUTPUT_FILENAME, { freezeHeaderRows: 2 });

    return {
      ok: true,
      errors: [],
      warnings: model.warnings,
      summary: model.summary,
      filename: OUTPUT_FILENAME,
    };
  } catch (err) {
    console.error('Workbook generation failed:', err);
    return {
      ok: false,
      errors: ['The Excel workbook could not be created. Please try again.'],
      warnings: model.warnings,
      summary: model.summary,
    };
  }
}
