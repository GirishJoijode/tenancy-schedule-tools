import * as XLSX from 'xlsx-js-style';
import { DATE_COMPARE_IDS } from '../config/constants.js';

const HEADER_FILL = { patternType: 'solid', fgColor: { rgb: '002060' } };
const HEADER_FONT = { bold: true, color: { rgb: 'FFFFFF' }, name: 'Calibri', sz: 11 };
const PROPERTY_FILL = { patternType: 'solid', fgColor: { rgb: 'D9D9D9' } };
const PROPERTY_FONT = { bold: true, name: 'Calibri', sz: 11, color: { rgb: '000000' } };
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

const PROPERTY_BORDER = {
  top: { style: 'medium', color: { rgb: '595959' } },
  bottom: { style: 'thin', color: BORDER_COLOR },
  left: { style: 'thin', color: BORDER_COLOR },
  right: { style: 'thin', color: BORDER_COLOR },
};

const COLUMN_NUM_FMT = {
  area_sqft: '#,##0',
  contracted_rent_pa: '\u00a3#,##0',
  contracted_rent_psf: '\u00a30.00',
  deemed_rent_psf: '\u00a30.00',
  market_rent_pa: '\u00a3#,##0',
  market_rent_psf: '\u00a30.00',
  rent_free_month: '0',
  expiry_void_month: '0',
  rent_free_expiry_month: '0',
  cap_rate: '0.00%',
  niy: '0.00%',
  ney: '0.00%',
  ry: '0.00%',
  cv_psf: '\u00a30',
  cap_ex: '\u00a3#,##0;[Red]-\u00a3#,##0',
  fair_value_april_2024: '\u00a3#,##0',
  fair_value_march_2026: '\u00a3#,##0',
  pct_change_26_april_24: '0.00%',
  pct_change_x_ytd: '0.00%',
  start_date: 'dd/mm/yyyy',
  review_date: 'dd/mm/yyyy',
  expiry_date: 'dd/mm/yyyy',
  break_date: 'dd/mm/yyyy',
};

const COLUMN_WIDTHS = {
  property_reference: 14,
  town: 14,
  address: 38,
  tenure_uk: 8,
  area_sqft: 12,
  unit: 30,
  tenant_name: 30,
  start_date: 12,
  review_date: 12,
  expiry_date: 12,
  break_date: 12,
  mwault_year: 10,
  contracted_rent_pa: 14,
  contracted_rent_psf: 12,
  deemed_rent_psf: 12,
  market_rent_pa: 14,
  market_rent_psf: 12,
  current_void: 11,
  rent_free_month: 10,
  expiry_void_month: 10,
  rent_free_expiry_month: 11,
  cap_rate: 9,
  niy: 8,
  ney: 8,
  ry: 8,
  cv_psf: 10,
  cap_ex: 12,
  fair_value_april_2024: 14,
  fair_value_march_2026: 14,
  pct_change_26_april_24: 12,
  pct_change_x_ytd: 11,
  comments: 36,
};

const DATE_COLUMN_IDS = new Set(['start_date', 'review_date', 'expiry_date', 'break_date']);
const WRAP_COLUMN_IDS = new Set(['comments', 'address', 'unit', 'tenant_name']);

export function getNumberFormat(columnId, value) {
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

export function buildDataCellStyle(columnId, isProperty) {
  var wrap = WRAP_COLUMN_IDS.has(columnId);
  var horizontal =
    DATE_COLUMN_IDS.has(columnId) || COLUMN_NUM_FMT[columnId] || columnId === 'current_void'
      ? 'right'
      : 'left';

  return {
    font: isProperty ? PROPERTY_FONT : UNIT_FONT,
    fill: isProperty ? PROPERTY_FILL : undefined,
    border: isProperty ? PROPERTY_BORDER : THIN_BORDER,
    alignment: {
      vertical: 'center',
      horizontal: horizontal,
      wrapText: wrap,
    },
  };
}

export function buildTenancyScheduleSheet(model) {
  var columns = model.columns;
  var colCount = columns.length;
  var rowCount = 1 + model.scheduleRows.length;
  var ws = {};

  if (colCount === 0) {
    ws.A1 = {
      t: 's',
      v: 'No output columns configured (all Standardised Template mappings are blank).',
      s: {
        font: HEADER_FONT,
        fill: HEADER_FILL,
        border: HEADER_BORDER,
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
        font: HEADER_FONT,
        border: HEADER_BORDER,
        alignment: {
          vertical: 'center',
          horizontal: 'center',
          wrapText: true,
        },
      },
    };
  }

  for (var r = 0; r < model.scheduleRows.length; r += 1) {
    var row = model.scheduleRows[r];
    var isProperty = row.kind === 'property';

    for (var dc = 0; dc < colCount; dc += 1) {
      var colId = columns[dc].id;
      var dataAddr = XLSX.utils.encode_cell({ r: r + 1, c: dc });
      var style = buildDataCellStyle(colId, isProperty);
      if (!isProperty) delete style.fill;
      ws[dataAddr] = valueToCell(row.values[colId], style, colId);
    }
  }

  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: Math.max(rowCount - 1, 0), c: Math.max(colCount - 1, 0) },
  });

  ws['!freeze'] = {
    xSplit: 0,
    ySplit: 1,
    topLeftCell: 'A2',
    activePane: 'bottomLeft',
    state: 'frozen',
  };
  ws['!views'] = [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];

  if (colCount > 0) {
    ws['!autofilter'] = {
      ref: XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: Math.max(rowCount - 1, 0), c: colCount - 1 },
      }),
    };
  }

  ws['!cols'] = columns.map(function (col) {
    return {
      wch:
        COLUMN_WIDTHS[col.id] != null
          ? COLUMN_WIDTHS[col.id]
          : Math.min(22, Math.max(10, String(col.label).length + 2)),
    };
  });

  var rowProps = [{ hpt: 36 }];
  for (var ri = 0; ri < model.scheduleRows.length; ri += 1) {
    rowProps.push({ hpt: 18 });
  }
  ws['!rows'] = rowProps;

  return ws;
}

export function buildUnmatchedSheet(unmatchedRows) {
  var headers = [
    'Original Excel Row',
    'Configured Unit Match Value',
    'Property Name',
    'Tenant Name',
    'Suite',
  ];

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
    var values = [
      row.originalExcelRow,
      row.matchValue,
      row.propertyName,
      row.tenantName,
      row.suite,
    ];
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
  ws['!cols'] = [{ wch: 18 }, { wch: 32 }, { wch: 28 }, { wch: 28 }, { wch: 22 }];
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

export function workbookToArrayBuffer(workbook) {
  return XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array',
    cellDates: true,
  });
}

export function downloadWorkbook(workbook, filename) {
  var buffer = workbookToArrayBuffer(workbook);
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
