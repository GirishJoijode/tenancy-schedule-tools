import * as XLSXNamespace from 'xlsx-js-style';

var XLSX = XLSXNamespace.utils ? XLSXNamespace : XLSXNamespace.default;
import {
  PROPERTY_ONLY_COL_WIDTHS,
  PROPERTY_ONLY_NUM_FMT,
  PROPERTY_ONLY_SHEET_NAME,
  PROPERTY_ONLY_VALUATION_IDS,
  PROPERTY_ONLY_VALUATION_TITLE,
  PROPERTY_ONLY_WRAP_IDS,
} from '../templates/propertyOnly.js';

const HEADER_FILL = { patternType: 'solid', fgColor: { rgb: 'FFD800' } };
const HEADER_FONT = { bold: true, color: { rgb: '000000' }, name: 'Calibri', sz: 11 };
const TITLE_FONT = { bold: true, color: { rgb: '000000' }, name: 'Calibri', sz: 12 };
const DATA_FONT = { name: 'Calibri', sz: 10, color: { rgb: '000000' } };
const BLACK = { rgb: '000000' };

const THIN_BORDER = {
  top: { style: 'thin', color: BLACK },
  bottom: { style: 'thin', color: BLACK },
  left: { style: 'thin', color: BLACK },
  right: { style: 'thin', color: BLACK },
};

const MEDIUM_BORDER = {
  top: { style: 'medium', color: BLACK },
  bottom: { style: 'medium', color: BLACK },
  left: { style: 'medium', color: BLACK },
  right: { style: 'medium', color: BLACK },
};

function valueToCell(value, columnId, style) {
  var numFmt = PROPERTY_ONLY_NUM_FMT[columnId] || null;
  var cellStyle = Object.assign({}, style);
  if (numFmt) cellStyle.numFmt = numFmt;

  if (value === null || value === undefined || value === '') {
    return { t: 's', v: '', s: cellStyle };
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return {
      t: 'd',
      v: value,
      z: 'dd/mm/yyyy',
      s: Object.assign({}, cellStyle, { numFmt: 'dd/mm/yyyy' }),
    };
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    var cell = { t: 'n', v: value, s: cellStyle };
    if (numFmt) cell.z = numFmt;
    return cell;
  }

  if (typeof value === 'boolean') {
    return { t: 'b', v: value, s: cellStyle };
  }

  return { t: 's', v: String(value), s: cellStyle };
}

function findValuationSpan(columns) {
  var first = -1;
  var last = -1;
  for (var i = 0; i < columns.length; i += 1) {
    if (PROPERTY_ONLY_VALUATION_IDS.indexOf(columns[i].id) !== -1) {
      if (first === -1) first = i;
      last = i;
    }
  }
  if (first === -1) return null;
  return { start: first, end: last };
}

/**
 * Builds a property-only workbook. One data row per property.
 * Title/sub-header architecture is here so a valuation date can be added later.
 * Grouping and totals are intentionally omitted in this version.
 */
export function buildPropertySummaryWorkbook(model) {
  var columns = model.columns || [];
  var rows = model.rows || [];
  var titleRow = 0;
  var headerRow = 1;
  var firstDataRow = 2;

  var ws = {};
  var colCount = Math.max(columns.length, 1);
  var lastRow = firstDataRow + Math.max(rows.length, 1) - 1;
  if (rows.length === 0) lastRow = headerRow;

  var titleStyle = {
    fill: HEADER_FILL,
    font: TITLE_FONT,
    border: MEDIUM_BORDER,
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  };
  var headerStyle = {
    fill: HEADER_FILL,
    font: HEADER_FONT,
    border: MEDIUM_BORDER,
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  };

  var span = columns.length ? findValuationSpan(columns) : null;

  for (var c = 0; c < colCount; c += 1) {
    var titleAddr = XLSX.utils.encode_cell({ r: titleRow, c: c });
    var inSpan = span && c >= span.start && c <= span.end;
    ws[titleAddr] = {
      t: 's',
      v: inSpan && c === span.start ? PROPERTY_ONLY_VALUATION_TITLE : '',
      s: inSpan ? titleStyle : {
        fill: HEADER_FILL,
        font: TITLE_FONT,
        border: THIN_BORDER,
        alignment: { horizontal: 'center', vertical: 'center' },
      },
    };

    var headerAddr = XLSX.utils.encode_cell({ r: headerRow, c: c });
    ws[headerAddr] = {
      t: 's',
      v: columns[c] ? columns[c].label : '',
      s: headerStyle,
    };
  }

  for (var r = 0; r < rows.length; r += 1) {
    var excelRow = firstDataRow + r;
    var values = rows[r].values || {};
    for (var d = 0; d < columns.length; d += 1) {
      var col = columns[d];
      var wrap = Boolean(PROPERTY_ONLY_WRAP_IDS[col.id]);
      var dataStyle = {
        font: DATA_FONT,
        border: THIN_BORDER,
        alignment: {
          vertical: 'center',
          horizontal: PROPERTY_ONLY_NUM_FMT[col.id] ? 'right' : 'left',
          wrapText: wrap,
        },
      };
      var addr = XLSX.utils.encode_cell({ r: excelRow, c: d });
      ws[addr] = valueToCell(values[col.id], col.id, dataStyle);
    }
  }

  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: lastRow, c: colCount - 1 },
  });

  ws['!merges'] = [];
  if (span && span.end > span.start) {
    ws['!merges'].push({
      s: { r: titleRow, c: span.start },
      e: { r: titleRow, c: span.end },
    });
  }

  ws['!cols'] = columns.map(function (col) {
    return { wch: PROPERTY_ONLY_COL_WIDTHS[col.id] || 14 };
  });
  if (!columns.length) ws['!cols'] = [{ wch: 20 }];

  ws['!rows'] = [{ hpt: 22 }, { hpt: 30 }];
  ws['!views'] = [{ state: 'frozen', ySplit: 2, topLeftCell: 'A3', activeCell: 'A3' }];

  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, PROPERTY_ONLY_SHEET_NAME);
  return wb;
}
