import * as XLSXNamespace from 'xlsx-js-style';

var XLSX = XLSXNamespace.utils ? XLSXNamespace : XLSXNamespace.default;
import { DATE_COMPARE_IDS } from '../config/constants.js';
import { DEFAULT_SCHEDULE_COLUMNS } from '../templates/defaultSchedule.js';

const HEADER_FILL = { patternType: 'solid', fgColor: { rgb: '002060' } };
const HEADER_FONT = { bold: true, color: { rgb: 'FFFFFF' }, name: 'Calibri', sz: 11 };
const UNIT_FONT = { bold: false, name: 'Calibri', sz: 11, color: { rgb: '000000' } };

const BORDER_COLOR = { rgb: 'BFBFBF' };
const HEADER_BORDER_COLOR = { rgb: '001538' };

const THIN_BORDER = {
  top: { style: 'thin', color: BORDER_COLOR },
  bottom: { style: 'thin', color: BORDER_COLOR },
  left: { style: 'thin', color: BORDER_COLOR },
  right: { style: 'thin', color: BORDER_COLOR },
};

const HEADER_BORDER = {
  top: { style: 'thin', color: HEADER_BORDER_COLOR },
  bottom: { style: 'thin', color: HEADER_BORDER_COLOR },
  left: { style: 'thin', color: HEADER_BORDER_COLOR },
  right: { style: 'thin', color: HEADER_BORDER_COLOR },
};

function excelNumberFormat(format) {
  if (!format) return null;
  return String(format).replace(/£/g, '\u00a3');
}

var COLUMN_NUM_FMT = {};
for (var formatIndex = 0; formatIndex < DEFAULT_SCHEDULE_COLUMNS.length; formatIndex += 1) {
  var formattedColumn = DEFAULT_SCHEDULE_COLUMNS[formatIndex];
  var excelFormat = excelNumberFormat(formattedColumn.format);
  if (excelFormat) COLUMN_NUM_FMT[formattedColumn.id] = excelFormat;
}

const COLUMN_WIDTHS = {
  property_reference: 18,
  valuation: 11,
  town: 16,
  address: 46,
  tenure_uk: 10,
  area_sqft: 12,
  unit: 16,
  tenant_name: 36,
  start_date: 12,
  review_date: 12,
  expiry_date: 12,
  break_date: 12,
  mwault_year: 10,
  contracted_rent_pa: 15,
  contracted_rent_psf: 13,
  deemed_rent_psf: 13,
  market_rent_pa: 15,
  market_rent_psf: 13,
  current_void: 12,
  rent_free_month: 11,
  expiry_void_month: 11,
  rent_free_expiry_month: 12,
  cap_rate: 10,
  niy: 9,
  ney: 9,
  ry: 9,
  cv_psf: 11,
  cap_ex: 13,
  fair_value: 16,
  fair_value_march_2026: 16,
  pct_change_26_april_24: 14,
  pct_change_x_ytd: 12,
  comments: 42,
};

const LEFT_COLUMN_IDS = new Set([
  'property_reference',
  'address',
  'tenant_name',
]);

export function getNumberFormat(columnId, value) {
  if (columnId === 'break_date') {
    if (typeof value === 'number' && Number.isFinite(value)) return '0.00';
    if (value instanceof Date && !Number.isNaN(value.getTime())) return 'dd/mm/yyyy';
    return null;
  }

  if (columnId === 'current_void') {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return 'dd/mm/yyyy';
    if (typeof value === 'number' && Number.isFinite(value)) return '0';
    return null;
  }

  if (COLUMN_NUM_FMT[columnId]) return COLUMN_NUM_FMT[columnId];

  if (value instanceof Date && !Number.isNaN(value.getTime())) return 'dd/mm/yyyy';

  return null;
}

export function valueToCell(value, baseStyle, columnId) {
  baseStyle = baseStyle || {};
  columnId = columnId || '';
  var numFmt = getNumberFormat(columnId, value);
  var style = numFmt ? Object.assign({}, baseStyle, { numFmt: numFmt }) : Object.assign({}, baseStyle);

  if (
    columnId === 'valuation' &&
    typeof value === 'number' &&
    Number.isFinite(value) &&
    value > 20000 &&
    value < 80000
  ) {
    var serialMs = Math.round((value - 25569) * 86400 * 1000);
    var serialDate = new Date(serialMs);
    value = new Date(serialDate.getUTCFullYear(), serialDate.getUTCMonth(), serialDate.getUTCDate());
  }

  if (value === null || value === undefined || value === '') {
    return { t: 's', v: '', s: style };
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return {
      t: 'd',
      v: value,
      z: numFmt || 'dd/mm/yyyy',
      s: Object.assign({}, style, { numFmt: numFmt || 'dd/mm/yyyy' }),
    };
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    var cell = {
      t: 'n',
      v: value,
      s: style,
    };
    if (numFmt) cell.z = numFmt;
    return cell;
  }

  if (typeof value === 'boolean') {
    return { t: 'b', v: value, s: style };
  }

  return {
    t: 's',
    v: String(value),
    s: style,
  };
}

function scheduleHorizontalAlign(columnId) {
  if (LEFT_COLUMN_IDS.has(columnId)) return 'left';
  return 'center';
}

export function buildDataCellStyle(columnId, isProperty) {
  return {
    font: isProperty ? SCHEDULE_PROPERTY_FONT : SCHEDULE_UNIT_FONT,
    fill: isProperty ? SCHEDULE_PROPERTY_FILL : SCHEDULE_WHITE_FILL,
    border: isProperty ? SCHEDULE_PROPERTY_BORDER : SCHEDULE_UNIT_BORDER,
    alignment: {
      vertical: 'center',
      horizontal: scheduleHorizontalAlign(columnId),
      wrapText: columnId === 'comments',
    },
  };
}

var SCHEDULE_BLACK = { rgb: '000000' };
var SCHEDULE_HEADER_FONT = { bold: true, color: { rgb: 'FFFFFF' }, name: 'Arial', sz: 8 };
var SCHEDULE_PROPERTY_FONT = { bold: true, name: 'Arial', sz: 8, color: { rgb: '000000' } };
var SCHEDULE_UNIT_FONT = { bold: false, name: 'Arial', sz: 8, color: { rgb: '000000' } };

var SCHEDULE_HEADER_BORDER = {
  top: { style: 'thin', color: SCHEDULE_BLACK },
  bottom: { style: 'medium', color: SCHEDULE_BLACK },
  left: { style: 'thin', color: SCHEDULE_BLACK },
  right: { style: 'thin', color: SCHEDULE_BLACK },
};

var SCHEDULE_SECTOR_EDGE = { style: 'medium', color: SCHEDULE_BLACK };

function sectorBandBorder(columnIndex, columnCount) {
  var border = {
    top: SCHEDULE_SECTOR_EDGE,
    bottom: SCHEDULE_SECTOR_EDGE,
  };
  if (columnIndex === 0) border.left = SCHEDULE_SECTOR_EDGE;
  if (columnIndex === columnCount - 1) border.right = SCHEDULE_SECTOR_EDGE;
  return border;
}

function summaryBandStyle(columnIndex, columnCount, horizontal, indent) {
  return {
    fill: HEADER_FILL,
    font: SCHEDULE_HEADER_FONT,
    border: sectorBandBorder(columnIndex, columnCount),
    alignment: {
      vertical: 'center',
      horizontal: horizontal,
      indent: indent || 0,
      wrapText: false,
    },
  };
}

var SCHEDULE_PROPERTY_BORDER = {
  top: { style: 'medium', color: SCHEDULE_BLACK },
  bottom: { style: 'thin', color: SCHEDULE_BLACK },
  left: { style: 'thin', color: SCHEDULE_BLACK },
  right: { style: 'thin', color: SCHEDULE_BLACK },
};

var SCHEDULE_UNIT_BORDER = {
  top: { style: 'thin', color: SCHEDULE_BLACK },
  bottom: { style: 'thin', color: SCHEDULE_BLACK },
  left: { style: 'thin', color: SCHEDULE_BLACK },
  right: { style: 'thin', color: SCHEDULE_BLACK },
};

var SCHEDULE_PROPERTY_FILL = { patternType: 'solid', fgColor: { rgb: 'F2F0ED' } };
var SCHEDULE_WHITE_FILL = { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } };
var SCHEDULE_HEADER_HEIGHT = 40;
var SCHEDULE_SEPARATOR_HEIGHT = 8;
var SCHEDULE_SECTOR_HEIGHT = 18;
var SCHEDULE_DATA_HEIGHT = 24;

export function buildTenancyScheduleSheet(model) {
  var columns = model.columns;
  var colCount = columns.length;
  var scheduleRows = model.scheduleRows || [];
  var rowCount = 2 + scheduleRows.length;
  var ws = {};

  if (colCount === 0) {
    ws.A1 = {
      t: 's',
      v: 'No output columns are defined.',
      s: {
        font: SCHEDULE_HEADER_FONT,
        fill: HEADER_FILL,
        border: SCHEDULE_HEADER_BORDER,
      },
    };
    ws['!ref'] = 'A1';
    ws['!cols'] = [{ wch: 80 }];
    return ws;
  }

  for (var c = 0; c < colCount; c += 1) {
    var headerAddr = XLSX.utils.encode_cell({ r: 0, c: c });
    ws[headerAddr] = {
      t: 's',
      v: columns[c].label,
      s: {
        fill: HEADER_FILL,
        font: SCHEDULE_HEADER_FONT,
        border: SCHEDULE_HEADER_BORDER,
        alignment: {
          vertical: 'center',
          horizontal: 'center',
          wrapText: true,
        },
      },
    };
  }

  for (var blankCol = 0; blankCol < colCount; blankCol += 1) {
    ws[XLSX.utils.encode_cell({ r: 1, c: blankCol })] = {
      t: 's',
      v: '',
      s: {
        font: SCHEDULE_UNIT_FONT,
        alignment: { vertical: 'center', horizontal: 'left', wrapText: false },
      },
    };
  }

  var merges = [];

  for (var r = 0; r < scheduleRows.length; r += 1) {
    var row = scheduleRows[r];
    var excelRow = r + 2;

    if (row.kind === 'sector') {
      merges.push({
        s: { r: excelRow, c: 0 },
        e: { r: excelRow, c: Math.max(colCount - 1, 0) },
      });
      for (var sc = 0; sc < colCount; sc += 1) {
        var sectorAddr = XLSX.utils.encode_cell({ r: excelRow, c: sc });
        ws[sectorAddr] = {
          t: 's',
          v: sc === 0 ? row.sector : '',
          s: summaryBandStyle(sc, colCount, 'left', sc === 0 ? 1 : 0),
        };
      }
      continue;
    }

    if (row.kind === 'separator') {
      for (var sepCol = 0; sepCol < colCount; sepCol += 1) {
        ws[XLSX.utils.encode_cell({ r: excelRow, c: sepCol })] = {
          t: 's',
          v: '',
          s: {
            font: SCHEDULE_UNIT_FONT,
            alignment: { vertical: 'center', horizontal: 'left', wrapText: false },
          },
        };
      }
      continue;
    }

    if (row.kind === 'sectorTotal' || row.kind === 'grandTotal') {
      for (var tc = 0; tc < colCount; tc += 1) {
        var totalId = columns[tc].id;
        var totalValue = row.values[totalId];
        var totalStyle = summaryBandStyle(tc, colCount, scheduleHorizontalAlign(totalId), 0);
        ws[XLSX.utils.encode_cell({ r: excelRow, c: tc })] = valueToCell(totalValue, totalStyle, totalId);
      }
      continue;
    }

    var isProperty = row.kind === 'property';
    for (var dc = 0; dc < colCount; dc += 1) {
      var colId = columns[dc].id;
      var dataAddr = XLSX.utils.encode_cell({ r: excelRow, c: dc });
      var cellStyle = buildDataCellStyle(colId, isProperty);
      ws[dataAddr] = valueToCell(row.values[colId], cellStyle, colId);
    }
  }

  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: Math.max(rowCount - 1, 0), c: Math.max(colCount - 1, 0) },
  });

  if (merges.length) ws['!merges'] = merges;

  ws['!freeze'] = {
    xSplit: 0,
    ySplit: 2,
    topLeftCell: 'A3',
    activePane: 'bottomLeft',
    state: 'frozen',
  };
  ws['!views'] = [{ state: 'frozen', ySplit: 2, topLeftCell: 'A3', activeCell: 'A3' }];

  if (colCount > 0 && rowCount > 2) {
    ws['!autofilter'] = {
      ref: XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: rowCount - 1, c: colCount - 1 },
      }),
    };
  }

  ws['!cols'] = columns.map(function (col) {
    return {
      wch:
        COLUMN_WIDTHS[col.id] != null
          ? COLUMN_WIDTHS[col.id]
          : col.width || Math.min(22, Math.max(10, String(col.label).length + 2)),
    };
  });

  var rowProps = [
    { hpt: SCHEDULE_HEADER_HEIGHT, customHeight: true },
    { hpt: SCHEDULE_SEPARATOR_HEIGHT, customHeight: true },
  ];
  for (var ri = 0; ri < scheduleRows.length; ri += 1) {
    rowProps.push({
      hpt:
        scheduleRows[ri].kind === 'sector' ||
        scheduleRows[ri].kind === 'sectorTotal' ||
        scheduleRows[ri].kind === 'grandTotal'
          ? SCHEDULE_SECTOR_HEIGHT
          : scheduleRows[ri].kind === 'separator'
            ? SCHEDULE_SEPARATOR_HEIGHT
            : SCHEDULE_DATA_HEIGHT,
      customHeight: true,
    });
  }
  ws['!rows'] = rowProps;

  return ws;
}

export function buildUnmatchedSheet(unmatchedRows) {
  var headers = ['Original Excel Row', 'Property Name', 'Tenant Name', 'Suite'];

  var ws = {};

  for (var c = 0; c < headers.length; c += 1) {
    var headerAddr = XLSX.utils.encode_cell({ r: 0, c: c });
    ws[headerAddr] = {
      t: 's',
      v: headers[c],
      s: {
        fill: HEADER_FILL,
        font: HEADER_FONT,
        border: HEADER_BORDER,
        alignment: { vertical: 'center', horizontal: 'center', wrapText: true },
      },
    };
  }

  for (var r = 0; r < unmatchedRows.length; r += 1) {
    var row = unmatchedRows[r];
    var values = [row.originalExcelRow, row.matchValue || row.propertyName, row.tenantName, row.suite];
    for (var dc = 0; dc < values.length; dc += 1) {
      var dataAddr = XLSX.utils.encode_cell({ r: r + 1, c: dc });
      var style = {
        font: UNIT_FONT,
        border: THIN_BORDER,
        alignment: { vertical: 'center', horizontal: dc === 0 ? 'right' : 'left' },
      };
      ws[dataAddr] = valueToCell(values[dc], style, dc === 0 ? 'area_sqft' : '');
    }
  }

  var lastRow = unmatchedRows.length;
  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: lastRow, c: headers.length - 1 },
  });
  ws['!views'] = [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];
  ws['!autofilter'] = {
    ref: XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: lastRow, c: headers.length - 1 },
    }),
  };
  ws['!cols'] = [{ wch: 18 }, { wch: 32 }, { wch: 28 }, { wch: 22 }];
  ws['!rows'] = [{ hpt: 28 }];

  return ws;
}

export function buildWorkbookFromModel(model) {
  var wb = XLSX.utils.book_new();

  var scheduleSheet = buildTenancyScheduleSheet(model);
  XLSX.utils.book_append_sheet(wb, scheduleSheet, 'Tenancy Schedule');

  if (model.unmatchedRows && model.unmatchedRows.length > 0) {
    var unmatchedSheet = buildUnmatchedSheet(model.unmatchedRows);
    XLSX.utils.book_append_sheet(wb, unmatchedSheet, 'Unmatched Records');
  }

  return wb;
}

var CRC32_TABLE = new Uint32Array(256);
for (var crcIndex = 0; crcIndex < 256; crcIndex += 1) {
  var crcValue = crcIndex;
  for (var crcBit = 0; crcBit < 8; crcBit += 1) {
    crcValue = crcValue & 1 ? 0xedb88320 ^ (crcValue >>> 1) : crcValue >>> 1;
  }
  CRC32_TABLE[crcIndex] = crcValue >>> 0;
}

function readU16(bytes, offset) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readU32(bytes, offset) {
  return (
    (bytes[offset] |
      (bytes[offset + 1] << 8) |
      (bytes[offset + 2] << 16) |
      (bytes[offset + 3] << 24)) >>>
    0
  );
}

function writeU16(bytes, offset, value) {
  bytes[offset] = value & 255;
  bytes[offset + 1] = (value >>> 8) & 255;
}

function writeU32(bytes, offset, value) {
  bytes[offset] = value & 255;
  bytes[offset + 1] = (value >>> 8) & 255;
  bytes[offset + 2] = (value >>> 16) & 255;
  bytes[offset + 3] = (value >>> 24) & 255;
}

function indexOfBytes(haystack, needle) {
  for (var i = 0; i <= haystack.length - needle.length; i += 1) {
    var matched = true;
    for (var j = 0; j < needle.length; j += 1) {
      if (haystack[i + j] !== needle[j]) {
        matched = false;
        break;
      }
    }
    if (matched) return i;
  }
  return -1;
}

function crc32(bytes) {
  var crc = 0xffffffff;
  for (var i = 0; i < bytes.length; i += 1) {
    crc = CRC32_TABLE[(crc ^ bytes[i]) & 255] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function findStoredZipEntry(bytes, targetName) {
  var offset = 0;
  var decoder = new TextDecoder();
  while (offset + 30 < bytes.length && readU32(bytes, offset) === 0x04034b50) {
    var method = readU16(bytes, offset + 8);
    var compressedSize = readU32(bytes, offset + 18);
    var nameLength = readU16(bytes, offset + 26);
    var extraLength = readU16(bytes, offset + 28);
    var name = decoder.decode(bytes.subarray(offset + 30, offset + 30 + nameLength));
    var dataStart = offset + 30 + nameLength + extraLength;
    if (name === targetName) {
      return {
        offset: offset,
        method: method,
        compressedSize: compressedSize,
        dataStart: dataStart,
      };
    }
    offset = dataStart + compressedSize;
  }
  return null;
}

function injectFrozenHeader(arrayBuffer, ySplit) {
  var bytes = arrayBuffer instanceof Uint8Array ? arrayBuffer : new Uint8Array(arrayBuffer);
  var entry = findStoredZipEntry(bytes, 'xl/worksheets/sheet1.xml');
  if (!entry || entry.method !== 0) return bytes;

  var marker = new TextEncoder().encode('<sheetViews><sheetView workbookViewId="0"/></sheetViews>');
  var topLeft = 'A' + (ySplit + 1);
  var replacement = new TextEncoder().encode(
    '<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="' +
      ySplit +
      '" topLeftCell="' +
      topLeft +
      '" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="' +
      topLeft +
      '" sqref="' +
      topLeft +
      '"/></sheetView></sheetViews>'
  );
  var sheet = bytes.subarray(entry.dataStart, entry.dataStart + entry.compressedSize);
  var markerAt = indexOfBytes(sheet, marker);
  if (markerAt < 0) return bytes;

  var delta = replacement.length - marker.length;
  var next = new Uint8Array(bytes.length + delta);
  var insertAt = entry.dataStart + markerAt;
  next.set(bytes.subarray(0, insertAt), 0);
  next.set(replacement, insertAt);
  next.set(bytes.subarray(insertAt + marker.length), insertAt + replacement.length);

  var newSize = entry.compressedSize + delta;
  var newSheet = next.subarray(entry.dataStart, entry.dataStart + newSize);
  writeU32(next, entry.offset + 14, crc32(newSheet));
  writeU32(next, entry.offset + 18, newSize);
  writeU32(next, entry.offset + 22, newSize);

  var eocd = -1;
  var scanFrom = Math.max(0, next.length - 22 - 65535);
  for (var scan = next.length - 22; scan >= scanFrom; scan -= 1) {
    if (readU32(next, scan) === 0x06054b50) {
      eocd = scan;
      break;
    }
  }
  if (eocd < 0) return bytes;

  var centralSize = readU32(next, eocd + 12);
  var centralOffset = readU32(next, eocd + 16);
  var movedCentral = centralOffset + delta;
  var cursor = movedCentral;
  var centralEnd = movedCentral + centralSize;
  while (cursor + 46 <= centralEnd && readU32(next, cursor) === 0x02014b50) {
    var centralNameLength = readU16(next, cursor + 28);
    var centralExtraLength = readU16(next, cursor + 30);
    var centralCommentLength = readU16(next, cursor + 32);
    var localOffset = readU32(next, cursor + 42);
    if (localOffset === entry.offset) {
      writeU32(next, cursor + 16, crc32(newSheet));
      writeU32(next, cursor + 20, newSize);
      writeU32(next, cursor + 24, newSize);
    } else if (localOffset > entry.offset) {
      writeU32(next, cursor + 42, localOffset + delta);
    }
    cursor += 46 + centralNameLength + centralExtraLength + centralCommentLength;
  }
  writeU32(next, eocd + 16, movedCentral);
  return next;
}

export function workbookToArrayBuffer(workbook, options) {
  var buffer = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array',
    cellDates: true,
  });
  if (options && options.freezeHeaderRows) {
    return injectFrozenHeader(buffer, options.freezeHeaderRows);
  }
  return buffer;
}

export function downloadWorkbook(workbook, filename, options) {
  var buffer = workbookToArrayBuffer(workbook, options);
  var blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.rel = 'noopener';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(function () {
    URL.revokeObjectURL(url);
  }, 2000);
}

/**
 * Builds the TS Comparison workbook (comparison rows + change summary sheet).
 * Kept here so XLSX stays confined to the excel/ modules.
 */
export function buildCompareWorkbook(result) {
  var cols = result.columns;
  var headers = ['Status', 'Property', 'Unit', 'Version', 'Changes'].concat(
    cols.map(function (c) {
      return c.label;
    })
  );

  var aoa = [headers];
  var metaRows = []; // track styling: {r, status, version, changedFields}

  for (var i = 0; i < result.pairs.length; i += 1) {
    var pair = result.pairs[i];
    var changesText =
      pair.status === 'changed'
        ? pair.changedLabels.join('; ')
        : pair.status === 'new'
          ? 'New unit'
          : pair.status === 'removed'
            ? 'Removed unit'
            : pair.status === 'ambiguous'
              ? 'Ambiguous match key'
              : '';

    var pushVersion = function (version, record) {
      var row = [
        version === 'current' ? pair.status.toUpperCase() : '',
        pair.propertyName,
        pair.unitName,
        version.toUpperCase(),
        version === 'current' ? changesText : '',
      ];
      for (var c = 0; c < cols.length; c += 1) {
        row.push(record ? record.values[cols[c].id] : null);
      }
      aoa.push(row);
      metaRows.push({
        r: aoa.length - 1,
        status: pair.status,
        version: version,
        changedFields: pair.changedFields || [],
      });
    };
    pushVersion('current', pair.current);
    pushVersion('previous', pair.previous);
  }

  var wb = XLSX.utils.book_new();
  var ws = XLSX.utils.aoa_to_sheet(aoa);

  // Style header + rows
  var range = XLSX.utils.decode_range(ws['!ref']);
  for (var c = range.s.c; c <= range.e.c; c += 1) {
    var addr = XLSX.utils.encode_cell({ r: 0, c: c });
    if (!ws[addr]) continue;
    ws[addr].s = {
      fill: { patternType: 'solid', fgColor: { rgb: '002060' } },
      font: { bold: true, color: { rgb: 'FFFFFF' }, name: 'Calibri', sz: 11 },
      alignment: { vertical: 'center', wrapText: true },
    };
  }

  for (var m = 0; m < metaRows.length; m += 1) {
    var meta = metaRows[m];
    var fill =
      meta.status === 'new'
        ? 'F3FAF6'
        : meta.status === 'removed'
          ? 'FFF7F6'
          : meta.version === 'current'
            ? 'F7F9FC'
            : 'FFFFFF';
    for (var cc = range.s.c; cc <= range.e.c; cc += 1) {
      var cellAddr = XLSX.utils.encode_cell({ r: meta.r, c: cc });
      if (!ws[cellAddr]) ws[cellAddr] = { t: 's', v: '' };
      var colId = cc >= 5 ? cols[cc - 5].id : null;
      var changed =
        meta.status === 'changed' && colId && meta.changedFields.indexOf(colId) !== -1;
      ws[cellAddr].s = {
        fill: {
          patternType: 'solid',
          fgColor: { rgb: changed ? 'FFF3BF' : fill },
        },
        font: { name: 'Calibri', sz: 10, bold: meta.version === 'current' },
      };
      if (colId && DATE_COMPARE_IDS[colId] && ws[cellAddr].v instanceof Date) {
        ws[cellAddr].t = 'd';
        ws[cellAddr].z = 'dd/mm/yyyy';
        ws[cellAddr].s.numFmt = 'dd/mm/yyyy';
      } else if (colId && typeof ws[cellAddr].v === 'number') {
        var fmt = getNumberFormat(colId, ws[cellAddr].v);
        if (fmt) {
          ws[cellAddr].z = fmt;
          ws[cellAddr].s.numFmt = fmt;
        }
      }
    }
  }

  ws['!views'] = [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];
  ws['!autofilter'] = { ref: ws['!ref'] };
  ws['!cols'] = headers.map(function (h) {
    return { wch: Math.min(28, Math.max(10, String(h).length + 2)) };
  });
  XLSX.utils.book_append_sheet(wb, ws, 'TS Comparison');

  // Change Summary sheet
  var summaryAoa = [['Status', 'Property', 'Unit', 'Changed Fields']];
  for (var si = 0; si < result.pairs.length; si += 1) {
    var sp = result.pairs[si];
    summaryAoa.push([
      sp.status.toUpperCase(),
      sp.propertyName,
      sp.unitName,
      sp.changedLabels.join('; '),
    ]);
  }
  var ws2 = XLSX.utils.aoa_to_sheet(summaryAoa);
  for (var sc = 0; sc < 4; sc += 1) {
    var saddr = XLSX.utils.encode_cell({ r: 0, c: sc });
    if (ws2[saddr]) {
      ws2[saddr].s = {
        fill: { patternType: 'solid', fgColor: { rgb: '002060' } },
        font: { bold: true, color: { rgb: 'FFFFFF' } },
      };
    }
  }
  ws2['!views'] = [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];
  ws2['!autofilter'] = { ref: ws2['!ref'] };
  XLSX.utils.book_append_sheet(wb, ws2, 'Change Summary');

  return wb;
}
