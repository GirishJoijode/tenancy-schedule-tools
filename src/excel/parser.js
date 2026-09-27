import * as XLSXNamespace from 'xlsx-js-style';

var XLSX = XLSXNamespace.utils ? XLSXNamespace : XLSXNamespace.default;
import { GROUP_HEADER_OFFSET, HEADER_SEARCH_ROWS } from '../config/constants.js';
import { HEADER_MARKERS } from '../schema/sourceFields.js';
import {
  buildSourceColumnKey,
  normaliseHeader,
  normaliseJoinValue,
  normaliseMatchingKey,
} from '../schema/text.js';
import { periodLabelFromHeader } from '../schema/periods.js';
import { assertRecognisedSource, omitBlankPropertyNameRows } from '../schema/validateSource.js';

export { buildSourceColumnKey, normaliseHeader, normaliseJoinValue, normaliseMatchingKey };

export function columnIndexToLetter(index) {
  var n = index + 1;
  var letter = '';
  while (n > 0) {
    var rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

export function readWorkbook(arrayBuffer) {
  return XLSX.read(arrayBuffer, {
    type: 'array',
    cellDates: true,
    dense: false,
  });
}

export function getFirstWorksheet(workbook) {
  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error('The workbook does not contain any worksheets.');
  }
  var sheetName = workbook.SheetNames[0];
  var sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    throw new Error('Worksheet "' + sheetName + '" could not be read.');
  }
  return { sheetName: sheetName, sheet: sheet };
}

export function sheetToArrayOfArrays(sheet) {
  var ref = sheet['!ref'];
  if (!ref) return [];

  var range = XLSX.utils.decode_range(ref);
  var rows = [];

  for (var r = 0; r <= range.e.r; r += 1) {
    var row = [];
    for (var c = 0; c <= range.e.c; c += 1) {
      var addr = XLSX.utils.encode_cell({ r: r, c: c });
      var cell = sheet[addr];
      if (cell === undefined || cell === null) {
        row.push(null);
      } else if (cell.v !== undefined) {
        row.push(cell.v);
      } else if (cell.w !== undefined) {
        row.push(cell.w);
      } else {
        row.push(null);
      }
    }
    rows.push(row);
  }

  return rows;
}

export function cellDisplayText(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value);
}

export function detectHeaderRow(rows, marker) {
  var target = normaliseHeader(marker);
  if (!target) {
    throw new Error('Header marker is empty.');
  }

  var limit = Math.min(rows.length, HEADER_SEARCH_ROWS);

  for (var r = 0; r < limit; r += 1) {
    var row = rows[r] || [];
    for (var c = 0; c < row.length; c += 1) {
      if (normaliseHeader(cellDisplayText(row[c])) === target) {
        return r;
      }
    }
  }

  throw new Error(
    'Could not find header marker "' + marker + '" within the first ' + HEADER_SEARCH_ROWS + ' rows.'
  );
}

export function resolveMergedGroupHeadings(sheet, rows, groupRowIndex, columnCount) {
  var sections = Array.from({ length: columnCount }, function () {
    return '';
  });
  var groupRow = rows[groupRowIndex] || [];

  for (var c = 0; c < columnCount; c += 1) {
    var text = cellDisplayText(groupRow[c]).trim();
    if (text) sections[c] = text;
  }

  var merges = sheet['!merges'] || [];
  for (var i = 0; i < merges.length; i += 1) {
    var merge = merges[i];
    if (groupRowIndex < merge.s.r || groupRowIndex > merge.e.r) continue;

    var topLeft = rows[merge.s.r] && rows[merge.s.r][merge.s.c];
    var label = cellDisplayText(topLeft).trim();
    if (!label) continue;

    for (var mc = merge.s.c; mc <= merge.e.c; mc += 1) {
      if (mc >= 0 && mc < columnCount) {
        sections[mc] = label;
      }
    }
  }

  return sections;
}

export function generateSourceColumns(fileType, fieldHeaders, sectionHeaders) {
  var occurrenceCounters = new Map();
  var columns = [];
  var columnCount = Math.max(fieldHeaders.length, sectionHeaders.length);

  for (var c = 0; c < columnCount; c += 1) {
    var field = (fieldHeaders[c] || '').trim();
    var section = (sectionHeaders[c] || '').trim();

    if (!field) continue;

    var occKey = normaliseMatchingKey(section) + '||' + normaliseMatchingKey(field);
    var occurrence = (occurrenceCounters.get(occKey) || 0) + 1;
    occurrenceCounters.set(occKey, occurrence);

    var columnLetter = columnIndexToLetter(c);
    var key = buildSourceColumnKey(fileType, section, field, occurrence);

    var sectionPart = section || '(No Section)';
    var occSuffix = occurrence > 1 ? ' (#' + occurrence + ')' : '';
    var displayLabel = sectionPart + ' > ' + field + occSuffix + ' [' + columnLetter + ']';

    columns.push({
      fileType: fileType,
      columnIndex: c,
      columnLetter: columnLetter,
      section: section,
      field: field,
      occurrence: occurrence,
      key: key,
      displayLabel: displayLabel,
    });
  }

  return columns;
}

export function extractDataRows(rows, headerRowIndex, columns) {
  var parsed = [];
  var dataRowCount = 0;

  for (var r = headerRowIndex + 1; r < rows.length; r += 1) {
    var sheetRow = rows[r] || [];
    var values = {};
    var hasAnyValue = false;

    for (var i = 0; i < columns.length; i += 1) {
      var col = columns[i];
      var raw = sheetRow[col.columnIndex];
      var value = raw === undefined ? null : raw;
      values[col.key] = value;
      if (value !== null && value !== '' && !(typeof value === 'string' && value.trim() === '')) {
        hasAnyValue = true;
      }
    }

    if (!hasAnyValue) continue;

    parsed.push({
      rowIndex: r,
      values: values,
    });
    dataRowCount += 1;
  }

  return { rows: parsed, dataRowCount: dataRowCount };
}

export function validateXlsxFile(file) {
  if (!file) {
    throw new Error('No file selected.');
  }
  var name = file.name || '';
  var lower = name.toLowerCase();
  if (!lower.endsWith('.xlsx')) {
    throw new Error('Please select an Excel .xlsx file.');
  }
}

export function readFileAsArrayBuffer(file) {
  if (typeof file.arrayBuffer === 'function') {
    return file.arrayBuffer().catch(function () {
      return readFileAsArrayBufferFallback(file);
    });
  }
  return readFileAsArrayBufferFallback(file);
}

export function readFileAsArrayBufferFallback(file) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload = function () {
      resolve(reader.result);
    };
    reader.onerror = function () {
      reject(new Error('Could not read the selected file.'));
    };
    reader.readAsArrayBuffer(file);
  });
}

export function parseSourceFile(file, fileType) {
  try {
    validateXlsxFile(file);
  } catch (err) {
    return Promise.reject(err);
  }

  var marker = HEADER_MARKERS[fileType];
  if (!marker) {
    return Promise.reject(new Error('Unsupported file type: ' + fileType));
  }

  return readFileAsArrayBuffer(file)
    .then(function (arrayBuffer) {
      var workbook;
      try {
        workbook = readWorkbook(arrayBuffer);
      } catch (err) {
        throw new Error('The file could not be opened as a valid Excel workbook.');
      }

      var wsInfo = getFirstWorksheet(workbook);
      var sheetName = wsInfo.sheetName;
      var sheet = wsInfo.sheet;
      var matrix = sheetToArrayOfArrays(sheet);

      if (matrix.length === 0) {
        throw new Error('The worksheet appears to be empty.');
      }

      var headerRowIndex = detectHeaderRow(matrix, marker);
      var groupRowIndex = headerRowIndex - GROUP_HEADER_OFFSET;

      var fieldRow = matrix[headerRowIndex] || [];
      var columnCount = fieldRow.length;
      var ref = sheet['!ref'];
      if (ref) {
        var range = XLSX.utils.decode_range(ref);
        columnCount = Math.max(columnCount, range.e.c + 1);
      }

      var fieldHeaders = Array.from({ length: columnCount }, function (_, c) {
        var periodLabel = periodLabelFromHeader(fieldRow[c]);
        if (periodLabel) return periodLabel;
        return cellDisplayText(fieldRow[c]).trim();
      });

      var sectionHeaders;
      if (groupRowIndex >= 0) {
        sectionHeaders = resolveMergedGroupHeadings(sheet, matrix, groupRowIndex, columnCount);
      } else {
        sectionHeaders = Array.from({ length: columnCount }, function () {
          return '';
        });
      }

      var columns = generateSourceColumns(fileType, fieldHeaders, sectionHeaders);

      if (columns.length === 0) {
        throw new Error('No source columns were detected on the header row.');
      }

      assertRecognisedSource(fileType, columns);

      var extracted = omitBlankPropertyNameRows(
        extractDataRows(matrix, headerRowIndex, columns),
        columns,
        fileType
      );

      return {
        fileName: file.name,
        fileType: fileType,
        sheetName: sheetName,
        headerRowIndex: headerRowIndex,
        groupRowIndex: groupRowIndex >= 0 ? groupRowIndex : null,
        columns: columns,
        rows: extracted.rows,
        dataRowCount: extracted.dataRowCount,
      };
    })
    .catch(function (err) {
      var message = err && typeof err.message === 'string' ? err.message : '';
      if (
        message.indexOf('Source format not recognised') === 0 ||
        message.indexOf('Could not find header marker') === 0 ||
        message.indexOf('The file could not be opened') === 0 ||
        message.indexOf('The worksheet') === 0 ||
        message.indexOf('The workbook') === 0 ||
        message.indexOf('No source columns') === 0 ||
        message.indexOf('Worksheet ') === 0 ||
        message.indexOf('Header marker') === 0 ||
        message.indexOf('Unsupported file type') === 0 ||
        message.indexOf('Could not read the selected file') === 0 ||
        message.indexOf('Please upload') === 0
      ) {
        throw err;
      }
      throw new Error('The file could not be read as a recognised Property or Unit workbook.');
    });
}
