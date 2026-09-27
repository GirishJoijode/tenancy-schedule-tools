import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createServer } from 'vite';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx-js-style');
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const failures = [];

function fail(message) {
  failures.push(message);
  console.error('FAIL:', message);
}

function assert(condition, message) {
  if (!condition) fail(message);
}

const vite = await createServer({
  root: root,
  server: { middlewareMode: true },
  appType: 'custom',
  logLevel: 'error',
});

const { getRequiredPropertyFields } = await vite.ssrLoadModule('/src/schema/propertySourceSchema.js');
const { getRequiredUnitFields } = await vite.ssrLoadModule('/src/schema/unitSourceSchema.js');
const { PROPERTY_SOURCE_FIELDS, UNIT_SOURCE_FIELDS } = await vite.ssrLoadModule('/src/schema/sourceFields.js');
const { readSourceValue } = await vite.ssrLoadModule('/src/schema/sourceLookup.js');
const { parseSourceFile } = await vite.ssrLoadModule('/src/excel/parser.js');
const { transformToSchedule } = await vite.ssrLoadModule('/src/standardiser/transform.js');
const { transformToPropertySummary } = await vite.ssrLoadModule('/src/propertySummary/propertySummary.js');
const { runTenancyCompare } = await vite.ssrLoadModule('/src/compare/engine.js');
const { buildWorkbookFromModel, workbookToArrayBuffer } = await vite.ssrLoadModule('/src/excel/writer.js');
const { buildPropertySummaryWorkbook } = await vite.ssrLoadModule('/src/propertySummary/propertySummaryExport.js');

function asFile(bytes, name) {
  const view = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const copy = view.buffer.slice(view.byteOffset, view.byteOffset + view.byteLength);
  return {
    name: name,
    arrayBuffer: function () {
      return Promise.resolve(copy);
    },
  };
}

function makeWorkbook(refs, dataRows) {
  const aoa = [refs.map(function (ref) { return ref.section; }), refs.map(function (ref) { return ref.field; })];
  dataRows.forEach(function (row) { aoa.push(row); });
  const sheet = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, 'Data');
  return XLSX.write(book, { bookType: 'xlsx', type: 'array', cellDates: true });
}

function valuesRow(refs, map) {
  return refs.map(function (ref) {
    const key = ref.section + '|' + ref.field;
    return Object.prototype.hasOwnProperty.call(map, key) ? map[key] : null;
  });
}

function fieldKey(ref) {
  return ref.section + '|' + ref.field;
}

async function parseBuffer(bytes, name, fileType) {
  return parseSourceFile(asFile(bytes, name), fileType);
}

function walk(dir, hit) {
  readdirSync(dir).forEach(function (name) {
    if (name === 'node_modules' || name === 'dist' || name === '_migration') return;
    const full = resolve(dir, name);
    if (statSync(full).isDirectory()) {
      walk(full, hit);
      return;
    }
    if (!/\.(js|mjs|css|html)$/.test(name)) return;
    hit(full, readFileSync(full, 'utf8'));
  });
}

try {
  const propertyRefs = getRequiredPropertyFields().concat([
    { section: 'Occupied Area', field: 'Aug-2027' },
  ]);
  const unitRefs = [UNIT_SOURCE_FIELDS.leaseId].concat(getRequiredUnitFields(), [
    { section: 'Occupied Area (SF)', field: 'Aug-2027' },
  ]);
  const valuationDate = new Date(2026, 8, 30);

  function propertyMap(name, sector, extra) {
    const map = Object.assign({
      'General|Property Name': name,
      'General|Sector': sector,
      'General|External ID': name ? name.replace(/\s+/g, '-').toUpperCase() : '',
      'General|Traditional Valuation Date': name ? valuationDate : null,
      'General|City': name === 'Asset 6' ? null : 'Leeds',
    }, extra || {});
    return valuesRow(propertyRefs, map);
  }

  const propertyBytes = makeWorkbook(propertyRefs, [
    propertyMap('Asset 1', 'Healthcare'),
    propertyMap('Asset 2', 'Healthcare'),
    propertyMap('Asset 3', 'Offices'),
    propertyMap('Asset 4', 'Healthcare'),
    propertyMap('', 'Total', { 'General|External ID': 'TOTAL' }),
    propertyMap('Asset 5', '   '),
    propertyMap('Asset 6', 'Retail Warehouse'),
  ]);

  function unitMap(name, suite, tenant, extra) {
    const map = Object.assign({
      'General|Property Name': name,
      'General|Suite': suite,
      'General|Tenant Name': tenant,
      'General|Lease Start': new Date(2024, 0, 1),
      'General|Lease Expiry': new Date(2029, 0, 1),
    }, extra || {});
    return valuesRow(unitRefs, map);
  }

  const unitBytes = makeWorkbook(unitRefs, [
    unitMap('Asset 1', '1', 'Tenant A'),
    unitMap('Asset 1', '2', 'Tenant B'),
    unitMap('Asset 1', '3', 'Tenant C'),
    unitMap('Asset 3', 'G', 'Tenant D'),
    unitMap('Orphan House', '9', 'Nobody'),
    unitMap('', 'Total', ''),
  ]);

  const propertyParsed = await parseBuffer(propertyBytes, 'Property.xlsx', 'property');
  const unitParsed = await parseBuffer(unitBytes, 'TS.xlsx', 'unit');
  const names = propertyParsed.rows.map(function (row) {
    const col = propertyParsed.columns.find(function (column) {
      return column.field === 'Property Name';
    });
    return row.values[col.key];
  });
  assert(names.join('|') === 'Asset 1|Asset 2|Asset 3|Asset 4|Asset 5|Asset 6', 'property total row was not excluded: ' + names.join('|'));
  assert(names.indexOf('Asset 6') !== -1, 'property with a blank city was removed');
  assert(unitParsed.dataRowCount === 5, 'unit total row was not excluded');

  const model = transformToSchedule(propertyParsed, unitParsed);
  assert(model.ok, 'default transform failed');
  const outline = model.scheduleRows.map(function (row) {
    if (row.kind === 'sector') return 'sector:' + row.sector;
    if (row.kind === 'sectorTotal') return 'sectorTotal:' + row.values.property_reference;
    if (row.kind === 'separator') return 'separator';
    if (row.kind === 'grandTotal') return 'grandTotal:' + row.values.property_reference;
    return row.kind + ':' + (row.values.property_reference || row.values.unit || '');
  });
  const expectedOutline = [
    'sector:Healthcare', 'property:ASSET-1', 'unit:1', 'unit:2', 'unit:3',
    'property:ASSET-2', 'property:ASSET-4', 'sectorTotal:Healthcare TOTAL', 'separator',
    'sector:Offices', 'property:ASSET-3', 'unit:G', 'sectorTotal:Offices TOTAL', 'separator',
    'sector:Unclassified', 'property:ASSET-5', 'sectorTotal:Unclassified TOTAL', 'separator',
    'sector:Retail Warehouse', 'property:ASSET-6', 'sectorTotal:Retail Warehouse TOTAL', 'separator',
    'grandTotal:GRAND TOTAL',
  ];
  assert(JSON.stringify(outline) === JSON.stringify(expectedOutline), 'sector outline\n' + outline.join(' | '));
  const sectorTotals = model.scheduleRows.filter(function (row) { return row.kind === 'sectorTotal'; });
  const sectors = model.scheduleRows.filter(function (row) { return row.kind === 'sector'; });
  assert(sectorTotals.length === sectors.length, 'each sector needs one total');
  assert(model.scheduleRows.filter(function (row) { return row.kind === 'separator'; }).length === sectors.length, 'separator after each sector');
  assert(model.scheduleRows[model.scheduleRows.length - 2].kind === 'separator', 'separator before grand total');
  assert(model.scheduleRows.filter(function (row) { return row.kind === 'grandTotal'; }).length === 1, 'one grand total');
  const healthcareTotal = sectorTotals[0];
  assert(healthcareTotal.values.niy == null && healthcareTotal.values.mwault_year == null, 'yield and empty WAULT stay blank');
  assert(model.unmatchedRows.length === 1, 'orphan unit was not reported');
  assert(model.scheduleRows.every(function (row) { return row.kind === 'sector' || row.values.address !== 'Orphan House'; }), 'orphan entered the schedule');

  const book = buildWorkbookFromModel(model);
  const sheet = book.Sheets['Tenancy Schedule'];
  assert(sheet.A1.v === 'Property Reference' && sheet.B1.v === 'Valuation', 'header order');
  assert(sheet.E1.v === 'Tenure (UXT)', 'tenure heading');
  assert(sheet.M1.v === 'WAULT (Years)', 'WAULT heading');
  assert(model.columns.some(function (column) { return column.id === 'fair_value' && column.label === 'Fair Value September 2026'; }), 'dynamic fair value heading');
  assert(!model.columns.some(function (column) { return column.id === 'fair_value_march_2026' || column.id === 'pct_change_26_april_24' || column.id === 'pct_change_x_ytd'; }), 'comparison columns hidden');
  assert(sheet.A2.v === '', 'blank separator missing');
  assert(sheet.A3.v === 'Healthcare' && sheet['!merges'] && sheet['!merges'].length === 4, 'sector headers');
  assert(sheet.B4.t === 'd' && sheet.B4.z === 'mmm-yy', 'valuation is not a real mmm-yy date');
  assert(sheet.A4.s.fill.fgColor.rgb === 'F2F0ED' && sheet.A4.s.font.bold === true, 'property row style');
  assert(!sheet.A5.s.font.bold, 'unit row style');
  assert(sheet.A1.s.fill.fgColor.rgb === '002060' && sheet.A1.s.font.color.rgb === 'FFFFFF', 'header style');
  assert(sheet.A1.s.font.name === 'Arial' && sheet.A1.s.font.sz === 8, 'header font');
  assert(sheet.A3.s.font.name === 'Arial' && sheet.A3.s.font.sz === 8 && sheet.A3.s.font.bold === true, 'sector font');
  assert(sheet.A4.s.font.name === 'Arial' && sheet.A4.s.font.sz === 8, 'property font');
  assert(sheet.A5.s.font.name === 'Arial' && sheet.A5.s.font.sz === 8, 'unit font');
  assert(sheet['!rows'][0].hpt === 40 && sheet['!rows'][1].hpt === 8, 'header or separator height');
  assert(sheet['!rows'][2].hpt === 18 && sheet['!rows'][3].hpt === 24 && sheet['!rows'][4].hpt === 24, 'sector property unit heights');
  assert(sheet.A1.s.border.bottom.style === 'medium' && sheet.A4.s.border.top.style === 'medium', 'header and property borders');
  assert(sheet.A5.s.border.top.style === 'thin' && sheet.A5.s.alignment.horizontal === 'left', 'unit border or alignment');
  assert(sheet.A4.s.alignment.horizontal === 'left' && sheet.D4.s.alignment.horizontal === 'left', 'property reference and address left');
  assert(sheet.H5.s.alignment.horizontal === 'left', 'tenant name left');
  assert(sheet.B4.s.alignment.horizontal === 'center' && sheet.F4.s.alignment.horizontal === 'center', 'valuation and area centred');
  assert(sheet.C4.s.alignment.horizontal === 'center' && sheet.E5.s.alignment.horizontal === 'center', 'town and tenure centred');
  assert(sheet.A10.v === 'Healthcare TOTAL' && sheet.A10.s.fill.fgColor.rgb === '002060' && sheet.A10.s.font.color.rgb === 'FFFFFF', 'sector total style');
  assert(sheet.A3.s.fill.fgColor.rgb === sheet.A10.s.fill.fgColor.rgb, 'sector total fill matches sector start');
  assert(sheet.A3.s.font.name === 'Arial' && sheet.A10.s.font.name === sheet.A3.s.font.name && sheet.A10.s.font.sz === 8 && sheet.A10.s.font.bold === true, 'sector total font matches sector start');
  assert(sheet.A10.s.alignment.horizontal === 'left' && sheet.F10.s.alignment.horizontal === 'center', 'sector total label left and values centred');
  assert(sheet.A10.s.border.left && sheet.A10.s.border.left.style === 'medium' && !sheet.A10.s.border.right, 'sector total outer left only');
  assert(!sheet.B10.s.border.left && !sheet.B10.s.border.right, 'sector total has no internal vertical borders');
  assert(sheet.A3.s.border.top.style === 'medium' && sheet.A10.s.border.top.style === 'medium' && sheet.A10.s.border.bottom.style === 'medium', 'sector band keeps the outer horizontal edge');
  assert(!sheet['!merges'].some(function (merge) { return merge.s.r === 9; }), 'sector total must stay unmerged');
  assert(sheet['!rows'][9].hpt === 18 && sheet['!rows'][10].hpt === 8, 'total and separator heights');
  assert(sheet['!freeze'].ySplit === 2 && sheet['!autofilter'], 'freeze or filter missing');
  assert(book.SheetNames.indexOf('Unmatched Records') !== -1, 'unmatched sheet missing');

  const writtenPath = resolve(tmpdir(), 'tenancy-regression.xlsx');
  const writtenBytes = Buffer.from(workbookToArrayBuffer(book, { freezeHeaderRows: 2 }));
  const { writeFileSync } = await import('node:fs');
  writeFileSync(writtenPath, writtenBytes);
  const reread = XLSX.read(writtenBytes, { type: 'buffer', cellDates: true, cellNF: true });
  const writtenSheet = reread.Sheets['Tenancy Schedule'];
  assert(writtenSheet.B4.v instanceof Date && writtenSheet.B4.z === 'mmm-yy', 'written valuation format');
  assert(String(writtenBytes).indexOf('ySplit="2"') !== -1, 'frozen header was not written into the workbook');
  assert(String(writtenBytes).indexOf('ht="40"') !== -1 && String(writtenBytes).indexOf('customHeight="1"') !== -1, 'custom row heights were not written');
  assert(String(writtenBytes).indexOf('showGridLines="0"') !== -1, 'worksheet gridlines were not hidden');

  const { waultToExpiry, waultToBreak, resolveBreakEnd } = await vite.ssrLoadModule('/src/logic/wault.js');
  const { asDate } = await vite.ssrLoadModule('/src/logic/yearfrac.js');
  const { updateDraftRule, resetDraft } = await vite.ssrLoadModule('/src/logic/draftState.js');
  const { DEFAULT_SCHEDULE_COLUMNS } = await vite.ssrLoadModule('/src/templates/defaultSchedule.js');
  function closeNumber(value, expected) {
    return typeof value === 'number' && Math.abs(value - expected) < 1e-9;
  }
  const valuation = new Date(2026, 8, 30);
  const simple = waultToExpiry([
    { valuation: valuation, expiry: new Date(2036, 8, 30), rent: 100 },
    { valuation: valuation, expiry: new Date(2031, 8, 30), rent: 100 },
  ]);
  assert(closeNumber(simple, 7.5), 'simple WAULT expiry ' + simple);
  const weighted = waultToExpiry([
    { valuation: valuation, expiry: new Date(2036, 8, 30), rent: 300 },
    { valuation: valuation, expiry: new Date(2031, 8, 30), rent: 100 },
  ]);
  assert(closeNumber(weighted, 8.75), 'weighted WAULT expiry ' + weighted);
  const breakWault = waultToBreak([
    { valuation: valuation, expiry: new Date(2036, 8, 30), breakDate: new Date(2031, 8, 30), rent: 100 },
    { valuation: valuation, expiry: new Date(2036, 8, 30), breakDate: null, rent: 100 },
  ]);
  const expiryAlongside = waultToExpiry([
    { valuation: valuation, expiry: new Date(2036, 8, 30), rent: 100 },
    { valuation: valuation, expiry: new Date(2036, 8, 30), rent: 100 },
  ]);
  assert(closeNumber(breakWault, 7.5), 'WAULT to break ' + breakWault);
  assert(closeNumber(expiryAlongside, 10), 'WAULT to expiry alongside break ' + expiryAlongside);
  const passedBreak = resolveBreakEnd(valuation, new Date(2025, 8, 30), new Date(2031, 8, 30));
  assert(asDate(passedBreak).getFullYear() === 2031, 'passed break falls back to expiry');
  assert(waultToExpiry([{ valuation: valuation, expiry: new Date(2036, 8, 30), rent: 0 }]) == null, 'zero rent excluded');
  assert(waultToBreak([{ valuation: valuation, expiry: null, breakDate: null, rent: 100 }]) == null, 'no qualifying end date');

  const logicBefore = JSON.stringify(DEFAULT_SCHEDULE_COLUMNS.map(function (column) { return column.logic; }));
  updateDraftRule('tenancy_schedule', 'area_sqft', 'sectorTotal', { method: 'blank' });
  const logicAfter = JSON.stringify(DEFAULT_SCHEDULE_COLUMNS.map(function (column) { return column.logic; }));
  assert(logicBefore === logicAfter, 'draft edit changed production logic');
  resetDraft('tenancy_schedule');
  const { buildDraftSpecification } = await vite.ssrLoadModule('/src/logic/exportDraft.js');
  const { getDraftTemplate } = await vite.ssrLoadModule('/src/logic/draftState.js');
  const spec = buildDraftSpecification('tenancy_schedule', getDraftTemplate('tenancy_schedule'));
  const specText = JSON.stringify(spec.document, null, 2);
  assert(specText.indexOf('\n  "schemaVersion"') === 0 || specText.indexOf('  "schemaVersion"') !== -1, 'draft json is pretty printed');
  assert(spec.document.logicIndex.waultToExpiry.length === 1, 'export includes WAULT expiry');
  assert(spec.document.logicIndex.waultToBreak.length === 1, 'export includes WAULT break');
  assert(spec.document.logicIndex.area[0].logic.unit.ref.section === 'Occupied Area (SF)', 'export describes unit occupied area');
  assert(spec.document.logicIndex.area[0].logic.property.scope === 'matchedUnits', 'export describes property area as unit sum');
  assert(spec.document.logicIndex.capitalValue[0].logic.property.calculation === 'divide', 'export describes CV divide');
  assert(spec.document.logicIndex.currentFairValue[0].columnId === 'fair_value', 'export describes current fair value');
  assert(spec.document.logicIndex.disabledComparison.length === 3, 'export lists disabled comparison fields');
  assert(spec.document.logicIndex.summaryAggregation.some(function (rule) { return rule.context === 'sectorTotal' && rule.rule.summary === 'sum'; }), 'export includes sector sum rules');
  assert(spec.document.draftDefinition.columns.some(function (column) { return column.logic && column.logic.sectorTotal && column.logic.grandTotal; }), 'column logic keeps total contexts');
  assert(specText.indexOf('Asset 1') === -1 && specText.indexOf('"sectorTotal"') !== -1, 'export has total logic and no workbook rows');

  const secondProperty = makeWorkbook(propertyRefs, [
    propertyMap('Asset 1', 'Healthcare', { 'Occupied Area|Aug-2027': 100, 'General|Contracted Rent': 1000, 'General|Net Rent': 1300903, 'General|Rental Value (ERV)': 4460000, 'General|Say Value': 5000 }),
    propertyMap('Asset 2', 'Healthcare', { 'Occupied Area|Aug-2027': 300, 'General|Contracted Rent': 3000, 'General|Say Value': 7000 }),
  ]);
  const totalUnits = makeWorkbook(unitRefs, [
    unitMap('Asset 1', 'A', 'Tenant A', { 'General|Lease Expiry': new Date(2036, 8, 30), 'General|Contracted Rent': 100, 'General|Effective Break': new Date(2031, 8, 30), 'Occupied Area (SF)|Aug-2027': 9999 }),
    unitMap('Asset 1', 'B', 'Tenant B', { 'General|Lease Expiry': new Date(2036, 8, 30), 'General|Contracted Rent': 100, 'Occupied Area (SF)|Aug-2027': 9999 }),
    unitMap('Asset 2', 'C', 'Tenant C', { 'General|Lease Expiry': new Date(2031, 8, 30), 'General|Contracted Rent': 100, 'Occupied Area (SF)|Aug-2027': 9999 }),
  ]);
  const totalModel = transformToSchedule(
    await parseBuffer(secondProperty, 'totals-property.xlsx', 'property'),
    await parseBuffer(totalUnits, 'totals-units.xlsx', 'unit')
  );
  assert(totalModel.ok, 'total model failed');
  const asset1 = totalModel.scheduleRows.find(function (row) { return row.kind === 'property' && row.values.property_reference === 'ASSET-1'; });
  const asset1Break = totalModel.scheduleRows.find(function (row) { return row.kind === 'unit' && row.values.unit === 'A'; });
  const healthcare = totalModel.scheduleRows.find(function (row) { return row.kind === 'sectorTotal'; });
  const grand = totalModel.scheduleRows.find(function (row) { return row.kind === 'grandTotal'; });
  assert(closeNumber(asset1.values.mwault_year, 10), 'property expiry WAULT ' + asset1.values.mwault_year);
  assert(closeNumber(asset1.values.break_date, 7.5), 'property break WAULT ' + asset1.values.break_date);
  assert(asset1Break.values.break_date instanceof Date, 'unit break stays a date');
  assert(closeNumber(asset1.values.area_sqft, 19998), 'property area sums matched unit areas ' + asset1.values.area_sqft);
  assert(closeNumber(healthcare.values.area_sqft, 29997), 'sector area sums unit areas ' + healthcare.values.area_sqft);
  assert(closeNumber(grand.values.area_sqft, 29997), 'grand area sums portfolio unit areas ' + grand.values.area_sqft);
  assert(closeNumber(healthcare.values.contracted_rent_pa, 4000), 'sector rent sums properties ' + healthcare.values.contracted_rent_pa);
  assert(closeNumber(healthcare.values.contracted_rent_psf, 4000 / 29997), 'sector psf recalculates ' + healthcare.values.contracted_rent_psf);
  assert(closeNumber(asset1.values.cv_psf, 5000 / 19998), 'property CV is fair value / area ' + asset1.values.cv_psf);
  assert(closeNumber(healthcare.values.cv_psf, 12000 / 29997), 'sector CV is fair value / area ' + healthcare.values.cv_psf);
  assert(healthcare.values.niy == null, 'sector yield stays blank');
  assert(closeNumber(healthcare.values.mwault_year, (10 * 100 + 10 * 100 + 5 * 100) / 300), 'sector expiry WAULT');
  assert(closeNumber(grand.values.break_date, healthcare.values.break_date), 'grand break WAULT uses units');
  const totalBook = buildWorkbookFromModel(totalModel);
  const totalSheet = totalBook.Sheets['Tenancy Schedule'];
  assert(totalSheet.L4.t === 'n' && totalSheet.L4.z === '0.00', 'property break cell is a WAULT number');
  assert(totalSheet.L5.t === 'd' && totalSheet.L5.z === 'dd/mm/yyyy', 'unit break cell stays a date');
  assert(totalSheet.M9.t === 'n' && totalSheet.L9.z === '0.00', 'sector total WAULT formats');
  const annualFormat = '\u00a3#,##0';
  const headerRow = {};
  Object.keys(totalSheet).forEach(function (addr) {
    if (addr[0] === '!') return;
    const cell = XLSX.utils.decode_cell(addr);
    if (cell.r === 0) headerRow[totalSheet[addr].v] = cell.c;
  });
  function formatAt(header, rowIndex) {
    const addr = XLSX.utils.encode_cell({ r: rowIndex, c: headerRow[header] });
    return totalSheet[addr];
  }
  const contracted = formatAt('Contracted Rent (pa)', 3);
  const deemed = formatAt('Deemed Rent (psf)', 3);
  const market = formatAt('Market Rent (pa)', 3);
  const fair = formatAt('Fair Value September 2026', 3);
  const capital = formatAt('CV (psf)', 3);
  assert(contracted.t === 'n' && contracted.z === annualFormat && contracted.v === 1000, 'contracted rent format ' + contracted.z);
  assert(deemed.t === 'n' && deemed.z === annualFormat && deemed.v === 1300903, 'deemed rent format ' + deemed.z + ' ' + deemed.v);
  assert(market.t === 'n' && market.z === annualFormat && market.v === 4460000, 'market rent format ' + market.z);
  assert(fair.t === 'n' && fair.z === annualFormat && fair.v === 5000, 'fair value format ' + fair.z);
  assert(capital.t === 'n' && capital.z === '\u00a3#,##0.00', 'capital value format ' + capital.z);
  assert(totalSheet.A9.v === 'Healthcare TOTAL' && typeof totalSheet.F9.v === 'number', 'sector total values stay in their own cells');
  assert(totalSheet.A11.v === 'GRAND TOTAL', 'grand total label');
  function bandMatches(left, right) {
    return left.s.fill.fgColor.rgb === right.s.fill.fgColor.rgb &&
      left.s.font.name === right.s.font.name &&
      left.s.font.sz === right.s.font.sz &&
      left.s.font.bold === true &&
      right.s.font.bold === true &&
      left.s.font.color.rgb === 'FFFFFF' &&
      right.s.font.color.rgb === 'FFFFFF';
  }
  function noInternalVertical(rowIndex) {
    const range = XLSX.utils.decode_range(totalSheet['!ref']);
    for (let c = 1; c < range.e.c; c += 1) {
      const cell = totalSheet[XLSX.utils.encode_cell({ r: rowIndex, c: c })];
      if (!cell || (cell.s.border && (cell.s.border.left || cell.s.border.right))) return false;
    }
    const first = totalSheet[XLSX.utils.encode_cell({ r: rowIndex, c: 0 })];
    const last = totalSheet[XLSX.utils.encode_cell({ r: rowIndex, c: range.e.c })];
    return first.s.border.left && !first.s.border.right && last.s.border.right && !last.s.border.left;
  }
  assert(bandMatches(totalSheet.A9, totalSheet.A11), 'grand total uses the sector total band');
  assert(noInternalVertical(8) && noInternalVertical(10), 'summary rows have no internal vertical borders');
  assert(totalSheet.A9.s.alignment.horizontal === 'left' && totalSheet.A11.s.alignment.horizontal === 'left', 'summary labels stay left');
  assert(totalSheet.F9.s.alignment.horizontal === 'center' && totalSheet.F11.s.alignment.horizontal === 'center', 'summary values stay centred');
  assert(!totalSheet['!merges'].some(function (merge) { return merge.s.r === 8 || merge.s.r === 10; }), 'summary rows stay unmerged');
  const waultProperty = formatAt('WAULT (Years)', 3);
  const waultSector = formatAt('WAULT (Years)', 8);
  const waultGrand = formatAt('WAULT (Years)', 10);
  const breakSector = formatAt('Break Date', 8);
  const breakGrand = formatAt('Break Date', 10);
  assert(waultProperty.t === 'n' && waultProperty.z === '0.00', 'property WAULT format');
  assert(waultSector.t === 'n' && waultSector.z === '0.00' && waultSector.v !== Math.round(waultSector.v), 'sector WAULT keeps precision and displays 0.00');
  assert(waultGrand.t === 'n' && waultGrand.z === '0.00', 'grand WAULT format');
  assert(breakSector.t === 'n' && breakSector.z === '0.00', 'sector break WAULT format');
  assert(breakGrand.t === 'n' && breakGrand.z === '0.00', 'grand break WAULT format');
  assert(formatAt('Deemed Rent (psf)', 8).z === annualFormat && formatAt('Deemed Rent (psf)', 10).z === annualFormat, 'summary deemed rent format');
  assert(formatAt('Fair Value September 2026', 8).z === annualFormat && formatAt('Fair Value September 2026', 10).z === annualFormat, 'summary fair value format');
  assert(formatAt('CV (psf)', 8).z === '\u00a3#,##0.00' && formatAt('CV (psf)', 10).z === '\u00a3#,##0.00', 'summary capital value format');
  assert(formatAt('Area (sq ft)', 8).z === '#,##0' && formatAt('Contracted Rent (pa)', 8).z === annualFormat, 'summary area and rent formats');
  assert(totalSheet.AA4.t === 'n' && totalSheet.AA4.z === '£#,##0.00', 'capital value format');
  assert(totalSheet.AC4.t === 'n' && totalSheet.AC4.z === '£#,##0', 'fair value format');
  assert(totalSheet.AC1.v === 'Fair Value September 2026', 'written fair value heading');

  const propertyOnly = transformToPropertySummary(propertyParsed);
  assert(propertyOnly.ok && propertyOnly.rows.length === 6, 'property only dropped a valid property');
  assert(propertyOnly.rows.every(function (row) { return row.values.property; }), 'blank property name entered Property Only');
  const propertyBook = buildPropertySummaryWorkbook(propertyOnly);
  assert(propertyBook.SheetNames[0] === 'Property Summary', 'property only sheet');
  const propertyBuffer = XLSX.write(propertyBook, { bookType: 'xlsx', type: 'array' });
  assert(propertyBuffer.byteLength > 100, 'property only workbook was not created');

  const currentBytes = makeWorkbook(unitRefs, [
    unitMap('Asset 1', '1', 'Tenant A', { 'General|Contracted Rent': 100 }),
    unitMap('Asset 1', '2', 'Tenant B', { 'General|Contracted Rent': 200 }),
    unitMap('Asset 1', '2', 'Tenant B duplicate', { 'General|Contracted Rent': 200 }),
    unitMap('Asset 9', '1', 'New Tenant'),
    unitMap('', 'Total', ''),
  ]);
  const previousBytes = makeWorkbook(unitRefs, [
    unitMap('Asset 1', '1', 'Tenant A', { 'General|Contracted Rent': 150 }),
    unitMap('Asset 1', '3', 'Gone Tenant'),
    unitMap('Asset 8', '1', 'Same Tenant', { 'General|Contracted Rent': 80 }),
  ]);
  const unchangedBytes = makeWorkbook(unitRefs, [
    unitMap('Asset 8', '1', 'Same Tenant', { 'General|Contracted Rent': 80 }),
  ]);
  const currentParsed = await parseBuffer(currentBytes, 'current.xlsx', 'unit');
  const previousParsed = await parseBuffer(previousBytes, 'previous.xlsx', 'unit');
  const unchangedParsed = await parseBuffer(unchangedBytes, 'previous-same.xlsx', 'unit');
  assert(currentParsed.dataRowCount === 4, 'compare total row was kept');
  const compared = runTenancyCompare(currentParsed, previousParsed);
  const same = runTenancyCompare(unchangedParsed, unchangedParsed);
  assert(compared.ok && same.ok, 'compare failed without configuration');
  const statuses = {};
  compared.pairs.forEach(function (pair) { statuses[pair.status] = (statuses[pair.status] || 0) + 1; });
  assert(statuses.changed === 1, 'changed status ' + JSON.stringify(statuses));
  assert(statuses.new === 1, 'new status ' + JSON.stringify(statuses));
  assert(statuses.removed === 2, 'removed status ' + JSON.stringify(statuses));
  assert(statuses.ambiguous === 1, 'ambiguous status ' + JSON.stringify(statuses));
  assert(same.summary.unchanged === 1 && same.summary.changed === 0, 'unchanged compare');

  let unrelatedError = '';
  try {
    await parseBuffer(makeWorkbook([{ section: 'Info', field: 'Title' }], [['Hello']]), 'other.xlsx', 'property');
  } catch (err) {
    unrelatedError = err.message;
  }
  assert(unrelatedError.indexOf('Could not find header marker') === 0, 'unrelated workbook error: ' + unrelatedError);
  assert(unrelatedError.toLowerCase().indexOf('undefined') === -1, 'cryptic unrelated error');

  const propertyMissingName = propertyRefs.filter(function (ref) { return fieldKey(ref) !== 'General|Property Name'; });
  let missingNameError = '';
  try {
    await parseBuffer(makeWorkbook(propertyMissingName, [propertyMissingName.map(function () { return 'x'; })]), 'bad-property.xlsx', 'property');
  } catch (err) {
    missingNameError = err.message;
  }
  assert(missingNameError.indexOf('Source format not recognised') === 0, 'missing property name: ' + missingNameError);
  assert(missingNameError.indexOf('General > Property Name') !== -1, 'missing property name detail');

  const unitMissingSuite = unitRefs.filter(function (ref) { return fieldKey(ref) !== 'General|Suite'; });
  let missingSuiteError = '';
  try {
    await parseBuffer(makeWorkbook(unitMissingSuite, [unitMissingSuite.map(function () { return 'x'; })]), 'bad-unit.xlsx', 'unit');
  } catch (err) {
    missingSuiteError = err.message;
  }
  assert(missingSuiteError.indexOf('Source format not recognised') === 0, 'missing suite: ' + missingSuiteError);
  assert(missingSuiteError.indexOf('General > Suite') !== -1, 'missing suite detail');

  let swappedUnit = '';
  try {
    await parseBuffer(unitBytes, 'TS.xlsx', 'property');
  } catch (err) {
    swappedUnit = err.message;
  }
  let swappedProperty = '';
  try {
    await parseBuffer(propertyBytes, 'Property.xlsx', 'unit');
  } catch (err) {
    swappedProperty = err.message;
  }
  assert(swappedUnit.indexOf('Could not find header marker') === 0 || swappedUnit.indexOf('Source format not recognised') === 0, 'TS in property slot: ' + swappedUnit);
  assert(swappedProperty.indexOf('Could not find header marker') === 0 || swappedProperty.indexOf('Source format not recognised') === 0, 'Property in unit slot: ' + swappedProperty);
  assert(swappedUnit !== swappedProperty, 'the two source schemas were not distinguished');

  let runtimeConfig = '';
  walk(resolve(root, 'src'), function (file, text) {
    if (
      text.indexOf('localStorage') !== -1 ||
      text.indexOf('mappingConfig') !== -1 ||
      text.indexOf('defaultConfig') !== -1 ||
      text.indexOf('Property_sample') !== -1 ||
      text.indexOf('ts_sample') !== -1 ||
      text.indexOf('Feb-2024') !== -1
    ) {
      runtimeConfig += file + '\n';
    }
  });
  const indexHtml = readFileSync(resolve(root, 'index.html'), 'utf8');
  if (indexHtml.indexOf('localStorage') !== -1 || indexHtml.indexOf('Property_sample') !== -1) {
    runtimeConfig += 'index.html\n';
  }
  assert(!runtimeConfig, 'runtime still references configuration or samples:\n' + runtimeConfig);

  async function expectParse(bytes, name, fileType) {
    try {
      return await parseBuffer(bytes, name, fileType);
    } catch (err) {
      return err.message;
    }
  }

  function staticPropertyRow(refs) {
    return refs.map(function (ref) {
      if (ref.section === 'Occupied Area') return null;
      if (ref.field === 'External ID') return 'X1';
      if (ref.field === 'Property Name') return 'Asset';
      if (ref.field === 'Sector') return 'Healthcare';
      return 'x';
    });
  }

  function propertySheet(periodLabels, periodValues, options) {
    const dropName = options && options.dropName;
    const nonPeriods = options && options.nonPeriods;
    const labels = nonPeriods || periodLabels;
    const section = 'Occupied Area';
    const base = getRequiredPropertyFields().filter(function (ref) {
      return !dropName || ref.field !== 'Property Name';
    });
    const refs = base.concat(labels.map(function (label) { return { section: section, field: label }; }));
    const row = staticPropertyRow(base).concat(periodValues || labels.map(function () { return 1; }));
    return makeWorkbook(refs, [row]);
  }

  const feb = await expectParse(propertySheet(['Feb-2024', 'Feb-2025', 'Feb-2026', 'Feb-2027', 'Feb-2028']), 'feb.xlsx', 'property');
  const aug = await expectParse(propertySheet(['Aug-2027', 'Aug-2028', 'Aug-2029', 'Aug-2030', 'Aug-2031']), 'aug.xlsx', 'property');
  const mar = await expectParse(propertySheet(['Mar-2030', 'Mar-2031', 'Mar-2032']), 'mar.xlsx', 'property');
  function unitPeriodSheet(periodLabels, periodValues, options) {
    const nonPeriods = options && options.nonPeriods;
    const dropGroup = options && options.dropGroup;
    const base = [UNIT_SOURCE_FIELDS.leaseId].concat(getRequiredUnitFields());
    const labels = dropGroup ? [] : (nonPeriods || periodLabels);
    const section = 'Occupied Area (SF)';
    const refs = base.concat(labels.map(function (label) { return { section: section, field: label }; }));
    const row = base.map(function (ref) {
      if (ref.field === 'Property Name') return 'Asset 1';
      if (ref.field === 'Suite') return '1';
      if (ref.field === 'Tenant Name') return 'Tenant A';
      if (ref.field === 'Lease ID') return 'L1';
      if (ref.field === 'Lease Start' || ref.field === 'Lease Expiry') return new Date(2024, 0, 1);
      return 'x';
    }).concat((periodValues || labels.map(function () { return 1; })));
    return makeWorkbook(refs, [row]);
  }
  const noPeriods = await expectParse(unitPeriodSheet(null, null, { nonPeriods: ['Total', 'Notes'] }), 'notes.xlsx', 'unit');
  const missingGroup = await expectParse(unitPeriodSheet(null, null, { dropGroup: true }), 'nogroup.xlsx', 'unit');
  const missingName = await expectParse(propertySheet(['Aug-2027'], [10], { dropName: true }), 'noname.xlsx', 'property');
  assert(feb && feb.rows, 'Feb periods should be valid: ' + feb);
  assert(aug && aug.rows, 'Aug periods should be valid: ' + aug);
  assert(mar && mar.rows, 'Mar periods should be valid: ' + mar);
  assert(typeof noPeriods === 'string' && noPeriods.indexOf('No period columns found under:') !== -1 && noPeriods.indexOf('Occupied Area (SF)') !== -1, 'blank period group: ' + noPeriods);
  assert(typeof missingGroup === 'string' && missingGroup.indexOf('Occupied Area (SF)') !== -1, 'missing unit area group: ' + missingGroup);
  assert(typeof missingName === 'string' && missingName.indexOf('General > Property Name') !== -1, 'static field still required: ' + missingName);

  async function unitArea(labels) {
    const base = [UNIT_SOURCE_FIELDS.leaseId].concat(getRequiredUnitFields());
    const values = [15321, 15321, 0, 12000, 15321];
    const refs = base.concat(labels.map(function (label) {
      return { section: 'Occupied Area (SF)', field: label };
    }));
    const row = base.map(function (ref) {
      if (ref.field === 'Property Name') return 'Asset 1';
      if (ref.field === 'Suite') return '1';
      if (ref.field === 'Tenant Name') return 'Tenant A';
      if (ref.field === 'Lease ID') return 'L1';
      if (ref.field === 'Lease Start' || ref.field === 'Lease Expiry') return new Date(2024, 0, 1);
      return 'x';
    }).concat(values);
    const parsed = await parseBuffer(makeWorkbook(refs, [row]), 'area.xlsx', 'unit');
    const propertyBase = getRequiredPropertyFields().concat([{ section: 'Occupied Area', field: 'Aug-2027' }]);
    const propertyRow = propertyBase.map(function (ref) {
      if (ref.field === 'Property Name') return 'Asset 1';
      if (ref.field === 'External ID') return 'A1';
      if (ref.field === 'Sector') return 'Healthcare';
      if (ref.section === 'Occupied Area') return 100;
      return 'x';
    });
    const propertyParsedForArea = await parseBuffer(makeWorkbook(propertyBase, [propertyRow]), 'area-property.xlsx', 'property');
    const scheduled = transformToSchedule(propertyParsedForArea, parsed);
    const unitRow = scheduled.scheduleRows.find(function (entry) { return entry.kind === 'unit'; });
    return unitRow ? unitRow.values.area_sqft : null;
  }

  const earliestProperty = await expectParse(
    propertySheet(['Aug-2027', 'Aug-2028', 'Aug-2031'], [10, 50, 30]),
    'earliest.xlsx',
    'property'
  );
  const earliestArea = readSourceValue(
    earliestProperty.rows[0],
    earliestProperty.columns,
    'property',
    PROPERTY_SOURCE_FIELDS.occupiedArea
  );
  assert(earliestArea === 10, 'property area should use the earliest period, got ' + earliestArea);

  const augArea = await unitArea(['Aug-2027', 'Aug-2028', 'Aug-2029', 'Aug-2030', 'Aug-2031']);
  const febArea = await unitArea(['Feb-2032', 'Feb-2033', 'Feb-2034', 'Feb-2035', 'Feb-2036']);
  assert(augArea === 15321, 'Aug unit area ' + augArea);
  assert(febArea === 15321, 'rolled unit area ' + febArea);

  const scopedAreaRefs = unitRefs.concat([
    { section: 'Occupied Area (SF)', field: 'Aug-2028' },
    { section: 'Occupied Area (SF)', field: 'Aug-2029' },
  ]);
  const scopedUnits = makeWorkbook(scopedAreaRefs, [
    valuesRow(scopedAreaRefs, { 'General|Property Name': 'Asset 1', 'General|Suite': '1', 'General|Tenant Name': 'One', 'Occupied Area (SF)|Aug-2027': 10000, 'Occupied Area (SF)|Aug-2028': 10000, 'Occupied Area (SF)|Aug-2029': 8000 }),
    valuesRow(scopedAreaRefs, { 'General|Property Name': 'Asset 1', 'General|Suite': '2', 'General|Tenant Name': 'Two', 'Occupied Area (SF)|Aug-2027': 5000, 'Occupied Area (SF)|Aug-2028': 5500, 'Occupied Area (SF)|Aug-2029': 5500 }),
    valuesRow(scopedAreaRefs, { 'General|Property Name': 'Asset 1', 'General|Suite': '3', 'General|Tenant Name': 'Three', 'Occupied Area (SF)|Aug-2028': 2000, 'Occupied Area (SF)|Aug-2029': 2000 }),
  ]);
  const scopedProperty = makeWorkbook(propertyRefs, [propertyMap('Asset 1', 'Healthcare', { 'General|Say Value': 175000 })]);
  const scopedModel = transformToSchedule(
    await parseBuffer(scopedProperty, 'scoped-property.xlsx', 'property'),
    await parseBuffer(scopedUnits, 'scoped-units.xlsx', 'unit')
  );
  const scopedUnitsOut = scopedModel.scheduleRows.filter(function (row) { return row.kind === 'unit'; });
  const scopedPropertyOut = scopedModel.scheduleRows.find(function (row) { return row.kind === 'property'; });
  const scopedSector = scopedModel.scheduleRows.find(function (row) { return row.kind === 'sectorTotal'; });
  const scopedGrand = scopedModel.scheduleRows.find(function (row) { return row.kind === 'grandTotal'; });
  assert(scopedUnitsOut.map(function (row) { return row.values.area_sqft; }).join(',') === '10000,5500,2000', 'unit areas stay row-scoped');
  assert(closeNumber(scopedPropertyOut.values.area_sqft, 17500), 'property area 17500');
  assert(closeNumber(scopedSector.values.area_sqft, 17500), 'sector area 17500');
  assert(closeNumber(scopedGrand.values.area_sqft, 17500), 'grand area 17500');
  assert(closeNumber(scopedPropertyOut.values.cv_psf, 10), 'property CV 10');
  assert(scopedPropertyOut.values.cv_psf !== scopedPropertyOut.values.fair_value, 'CV is not fair value');

  const mixedDates = makeWorkbook(propertyRefs, [
    propertyMap('Asset 1', 'Healthcare', { 'General|Traditional Valuation Date': new Date(2026, 8, 30) }),
    propertyMap('Asset 2', 'Healthcare', { 'General|Traditional Valuation Date': new Date(2025, 2, 15) }),
  ]);
  const mixedModel = transformToSchedule(
    await parseBuffer(mixedDates, 'mixed-dates.xlsx', 'property'),
    await parseBuffer(scopedUnits, 'scoped-units.xlsx', 'unit')
  );
  assert(mixedModel.columns.some(function (column) { return column.id === 'fair_value' && column.label === 'Fair Value'; }), 'mixed dates use generic Fair Value');
  assert(mixedModel.warnings.some(function (warning) { return warning.indexOf('Multiple valuation dates detected') !== -1; }), 'mixed date warning');

  const propertySample = resolve(root, 'samples/Property_sample.xlsx');
  const unitSample = resolve(root, 'samples/ts_sample.xlsx');
  if (existsSync(propertySample) && existsSync(unitSample)) {
    const sampleProperty = await parseBuffer(readFileSync(propertySample), 'Property_sample.xlsx', 'property');
    const sampleUnit = await parseBuffer(readFileSync(unitSample), 'ts_sample.xlsx', 'unit');
    const sampleModel = transformToSchedule(sampleProperty, sampleUnit);
    const samplePropertyOnly = transformToPropertySummary(sampleProperty);
    const sampleCompare = runTenancyCompare(sampleUnit, sampleUnit);
    assert(sampleModel.ok && samplePropertyOnly.ok && sampleCompare.ok, 'sample fixtures failed');
    assert(sampleModel.scheduleRows.some(function (row) { return row.kind === 'sector'; }), 'sample sector grouping');
    const sampleBook = buildWorkbookFromModel(sampleModel);
    assert(sampleBook.Sheets['Tenancy Schedule'].A1.v === 'Property Reference', 'sample schedule header');
    function close(value, expected) {
      return typeof value === 'number' && Math.abs(value - expected) < 0.02;
    }
    const propertyAreas = sampleModel.scheduleRows.filter(function (row) { return row.kind === 'property'; }).map(function (row) { return row.values.area_sqft; });
    const unitAreas = sampleModel.scheduleRows.filter(function (row) { return row.kind === 'unit'; }).map(function (row) { return row.values.area_sqft; });
    assert(propertyAreas.length === 2, 'sample property count');
    assert(unitAreas.some(function (value) { return close(value, 741); }) && unitAreas.some(function (value) { return close(value, 856); }), 'sample unit areas stay distinct');
    const namedUnit = sampleModel.scheduleRows.find(function (row) { return row.kind === 'unit'; });
    const namedProperty = sampleModel.scheduleRows.find(function (row) { return row.kind === 'property' && row.matchKey === namedUnit.matchKey; });
    assert(close(namedProperty.values.area_sqft, 741 + 856), 'sample property area sums its units');
    assert(sampleProperty.rows.length === 2 && sampleUnit.rows.length === 11, 'sample footer rows were not excluded');
    console.log('sample fixtures parsed', sampleModel.summary.propertiesProcessed, 'properties');
  } else {
    console.log('sample files not present; synthetic fixtures were used');
  }

  console.log(failures.length ? 'REGRESSION FAILED' : 'OK: regression');
} catch (err) {
  fail(err && err.stack ? err.stack : String(err));
} finally {
  await vite.close();
}

if (failures.length) process.exit(1);
