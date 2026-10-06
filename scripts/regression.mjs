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
    { section: 'Capital Cost', field: 'Aug-2027' },
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
  assert(sheet.A1.v === 'Asset Unit Reference' && sheet.B1.v === 'Tenant Reference', 'hidden system headers');
  assert(sheet['!cols'][0].hidden === true && sheet['!cols'][1].hidden === true, 'system columns are hidden');
  assert(sheet.C1.v === 'Property Reference' && sheet.D1.v === 'Valuation' && sheet.E1.v === 'Town' && sheet.F1.v === 'Address', 'required visible order');
  assert(sheet.G1.v === 'Tenure (UXT)', 'tenure heading');
  assert(sheet.O1.v === 'WAULT (Years)', 'WAULT heading');
  assert(model.columns.some(function (column) { return column.id === 'fair_value' && column.label === 'Fair Value September 2026'; }), 'dynamic fair value heading');
  assert(!model.columns.some(function (column) { return column.id === 'fair_value_march_2026' || column.id === 'pct_change_26_april_24' || column.id === 'pct_change_x_ytd'; }), 'comparison columns hidden');
  assert(sheet.A2.v === '', 'blank separator missing');
  assert(sheet.C3.v === 'Healthcare' && sheet['!merges'] && sheet['!merges'].length === 4 && sheet['!merges'][0].s.c === 2, 'sector headers start at the first visible column');
  assert(sheet.D4.t === 'd' && sheet.D4.z === 'mmm-yy', 'valuation is not a real mmm-yy date');
  assert(sheet.C4.s.fill.fgColor.rgb === 'F2F0ED' && sheet.C4.s.font.bold === true, 'property row style');
  assert(!sheet.C5.s.font.bold, 'unit row style');
  assert(sheet.A1.s.fill.fgColor.rgb === '002060' && sheet.A1.s.font.color.rgb === 'FFFFFF', 'header style');
  assert(sheet.A1.s.font.name === 'Arial' && sheet.A1.s.font.sz === 8, 'header font');
  assert(sheet.C3.s.font.name === 'Arial' && sheet.C3.s.font.sz === 8 && sheet.C3.s.font.bold === true, 'sector font');
  assert(sheet.C4.s.font.name === 'Arial' && sheet.C4.s.font.sz === 8, 'property font');
  assert(sheet.C5.s.font.name === 'Arial' && sheet.C5.s.font.sz === 8, 'unit font');
  assert(sheet['!rows'][0].hpt === 40 && sheet['!rows'][1].hpt === 8, 'header or separator height');
  assert(sheet['!rows'][2].hpt === 24 && sheet['!rows'][3].hpt === 24 && sheet['!rows'][4].hpt === 24, 'sector property unit heights');
  assert(sheet.C1.s.border.bottom.style === 'medium' && sheet.C4.s.border.top.style === 'medium', 'header and property borders');
  assert(sheet.C5.s.border.top.style === 'thin' && sheet.C5.s.alignment.horizontal === 'left', 'unit border or alignment');
  assert(sheet.C4.s.alignment.horizontal === 'left' && sheet.F4.s.alignment.horizontal === 'left', 'property reference and address left');
  assert(sheet.J5.s.alignment.horizontal === 'left', 'tenant name left');
  assert(sheet.D4.s.alignment.horizontal === 'center' && sheet.H4.s.alignment.horizontal === 'center', 'valuation and area centred');
  assert(sheet.E4.s.alignment.horizontal === 'center' && sheet.G5.s.alignment.horizontal === 'center', 'town and tenure centred');
  assert(sheet.C10.v === 'Healthcare TOTAL' && sheet.C10.s.fill.fgColor.rgb === '002060' && sheet.C10.s.font.color.rgb === 'FFFFFF', 'sector total style');
  assert(sheet.C3.s.fill.fgColor.rgb === sheet.C10.s.fill.fgColor.rgb, 'sector total fill matches sector start');
  assert(sheet.C3.s.font.name === 'Arial' && sheet.C10.s.font.name === sheet.C3.s.font.name && sheet.C10.s.font.sz === 8 && sheet.C10.s.font.bold === true, 'sector total font matches sector start');
  assert(sheet.C10.s.alignment.horizontal === 'left' && sheet.H10.s.alignment.horizontal === 'center', 'sector total label left and values centred');
  assert(sheet.C10.s.border.left && sheet.C10.s.border.left.style === 'medium' && !sheet.C10.s.border.right, 'sector total outer left only');
  assert(!sheet.D10.s.border.left && !sheet.D10.s.border.right, 'sector total has no internal vertical borders');
  assert(sheet.C3.s.border.top.style === 'medium' && sheet.C10.s.border.top.style === 'medium' && sheet.C10.s.border.bottom.style === 'medium', 'sector band keeps the outer horizontal edge');
  assert(sheet['!autofilter'].ref.indexOf('C1:') === 0, 'autofilter starts at Property Reference');
  assert(!sheet['!merges'].some(function (merge) { return merge.s.r === 9; }), 'sector total must stay unmerged');
  assert(sheet['!rows'][9].hpt === 24 && sheet['!rows'][10].hpt === 8, 'total and separator heights');
  model.scheduleRows.forEach(function (row, index) {
    const height = sheet['!rows'][index + 2].hpt;
    if (row.kind === 'separator') {
      assert(height === 8, 'separator row height ' + height);
      return;
    }
    assert(height === 24, row.kind + ' row height ' + height);
  });
  assert(sheet['!freeze'].ySplit === 2 && sheet['!autofilter'], 'freeze or filter missing');
  assert(book.SheetNames.indexOf('Unmatched Records') !== -1, 'unmatched sheet missing');

  const writtenPath = resolve(tmpdir(), 'tenancy-regression.xlsx');
  const writtenBytes = Buffer.from(workbookToArrayBuffer(book, { freezeHeaderRows: 2 }));
  const { writeFileSync } = await import('node:fs');
  writeFileSync(writtenPath, writtenBytes);
  const reread = XLSX.read(writtenBytes, { type: 'buffer', cellDates: true, cellNF: true });
  const writtenSheet = reread.Sheets['Tenancy Schedule'];
  assert(writtenSheet.D4.v instanceof Date && writtenSheet.D4.z === 'mmm-yy', 'written valuation format');
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
  assert(totalSheet.N4.t === 'n' && totalSheet.N4.z === '0.00', 'property break cell is a WAULT number');
  assert(totalSheet.N5.t === 'd' && totalSheet.N5.z === 'dd/mm/yyyy', 'unit break cell stays a date');
  assert(totalSheet.O9.t === 'n' && totalSheet.N9.z === '0.00', 'sector total WAULT formats');
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
  assert(totalSheet.C9.v === 'Healthcare TOTAL' && typeof totalSheet.H9.v === 'number', 'sector total values stay in their own cells');
  assert(totalSheet.C11.v === 'GRAND TOTAL', 'grand total label');
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
    const firstVisible = 2;
    for (let c = firstVisible + 1; c < range.e.c; c += 1) {
      const cell = totalSheet[XLSX.utils.encode_cell({ r: rowIndex, c: c })];
      if (!cell || (cell.s.border && (cell.s.border.left || cell.s.border.right))) return false;
    }
    const first = totalSheet[XLSX.utils.encode_cell({ r: rowIndex, c: firstVisible })];
    const last = totalSheet[XLSX.utils.encode_cell({ r: rowIndex, c: range.e.c })];
    return first.s.border.left && !first.s.border.right && last.s.border.right && !last.s.border.left;
  }
  assert(bandMatches(totalSheet.C9, totalSheet.C11), 'grand total uses the sector total band');
  assert(noInternalVertical(8) && noInternalVertical(10), 'summary rows have no internal vertical borders');
  assert(totalSheet.C9.s.alignment.horizontal === 'left' && totalSheet.C11.s.alignment.horizontal === 'left', 'summary labels stay left');
  assert(totalSheet.H9.s.alignment.horizontal === 'center' && totalSheet.H11.s.alignment.horizontal === 'center', 'summary values stay centred');
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
  assert(formatAt('CV (psf)', 3).t === 'n' && formatAt('CV (psf)', 3).z === '£#,##0.00', 'capital value format');
  assert(formatAt('Fair Value September 2026', 3).t === 'n' && formatAt('Fair Value September 2026', 3).z === '£#,##0', 'fair value format');
  assert(formatAt('Fair Value September 2026', 0).v === 'Fair Value September 2026', 'written fair value heading');

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
    })).concat([{ section: 'Capital Cost', field: 'Aug-2027' }]);
    const row = base.map(function (ref) {
      if (ref.field === 'Property Name') return 'Asset 1';
      if (ref.field === 'Suite') return '1';
      if (ref.field === 'Tenant Name') return 'Tenant A';
      if (ref.field === 'Lease ID') return 'L1';
      if (ref.field === 'Lease Start' || ref.field === 'Lease Expiry') return new Date(2024, 0, 1);
      return 'x';
    }).concat(values).concat([0]);
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
    assert(sampleBook.Sheets['Tenancy Schedule'].C1.v === 'Property Reference', 'sample schedule header');
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

  const { setReportLayout, resetReportLayout, defaultOptionalIds } = await vite.ssrLoadModule('/src/logic/reportLayout.js');
  const { assetUnitReference } = await vite.ssrLoadModule('/src/standardiser/transform.js');
  assert(assetUnitReference('Copper House', 'Unit 01') === 'Copper House_Unit 01', 'joined asset unit reference');
  assert(assetUnitReference('Copper House', ' ') === 'Copper House', 'address-only reference');
  assert(assetUnitReference('', 'Unit 01') === 'Unit 01', 'unit-only reference');
  const referenceUnit = totalModel.scheduleRows.find(function (row) { return row.kind === 'unit' && row.values.unit === 'A'; });
  assert(referenceUnit.values.asset_unit_reference === 'Asset 1_A', 'unit reference uses that row ' + referenceUnit.values.asset_unit_reference);
  assert(totalModel.scheduleRows.find(function (row) { return row.kind === 'property'; }).values.asset_unit_reference == null, 'property reference column stays blank');
  assert(totalModel.scheduleRows.find(function (row) { return row.kind === 'sectorTotal'; }).values.tenant_reference == null, 'sector tenant reference stays blank');

  setReportLayout(defaultOptionalIds().filter(function (id) { return id !== 'deemed_rent_psf' && id !== 'area_sqft'; }));
  const dependencyModel = transformToSchedule(
    await parseBuffer(secondProperty, 'layout-property.xlsx', 'property'),
    await parseBuffer(totalUnits, 'layout-units.xlsx', 'unit')
  );
  assert(!dependencyModel.columns.some(function (column) { return column.id === 'area_sqft' || column.id === 'deemed_rent_psf'; }), 'removed fields are absent');
  assert(dependencyModel.columns[0].id === 'asset_unit_reference' && dependencyModel.columns[2].id === 'property_reference' && dependencyModel.columns[5].id === 'address', 'fixed columns stay first');
  const dependencyAsset = dependencyModel.scheduleRows.find(function (row) { return row.kind === 'property' && row.values.property_reference === 'ASSET-1'; });
  assert(closeNumber(dependencyAsset.values.cv_psf, 5000 / 19998), 'capital value still uses internal area');
  const ids = defaultOptionalIds().filter(function (id) { return id !== 'mwault_year'; });
  ids.splice(ids.indexOf('area_sqft'), 0, 'mwault_year');
  setReportLayout(ids);
  const reordered = transformToSchedule(
    await parseBuffer(secondProperty, 'order-property.xlsx', 'property'),
    await parseBuffer(totalUnits, 'order-units.xlsx', 'unit')
  );
  const waultAt = reordered.columns.findIndex(function (column) { return column.id === 'mwault_year'; });
  const areaAt = reordered.columns.findIndex(function (column) { return column.id === 'area_sqft'; });
  assert(waultAt !== -1 && waultAt < areaAt, 'WAULT can be placed before Area');
  const reorderedSheet = buildWorkbookFromModel(reordered).Sheets['Tenancy Schedule'];
  const reorderedHeaders = {};
  Object.keys(reorderedSheet).forEach(function (addr) {
    if (addr[0] === '!') return;
    const decoded = XLSX.utils.decode_cell(addr);
    if (decoded.r === 0) reorderedHeaders[reorderedSheet[addr].v] = decoded.c;
  });
  assert(reorderedSheet[XLSX.utils.encode_cell({ r: 3, c: reorderedHeaders['WAULT (Years)'] })].z === '0.00', 'moved WAULT keeps 0.00');
  assert(reorderedSheet[XLSX.utils.encode_cell({ r: 3, c: reorderedHeaders['Area (sq ft)'] })].z === '#,##0', 'moved Area keeps its format');
  const sectorTotal = reordered.scheduleRows.find(function (row) { return row.kind === 'sectorTotal'; });
  assert(closeNumber(sectorTotal.values.mwault_year, (10 * 100 + 10 * 100 + 5 * 100) / 300), 'sector WAULT follows the field after reorder');
  resetReportLayout();

  const currentJune = new Date(2026, 8, 30);
  const previousJune = new Date(2026, 5, 30);
  function comparisonProperty(name, ref, sector, say, erv, date) {
    return propertyMap(name, sector, {
      'General|External ID': ref,
      'General|Say Value': say,
      'General|Rental Value (ERV)': erv,
      'General|Traditional Valuation Date': date,
    });
  }
  function comparisonUnit(name, suite, erv) {
    return unitMap(name, suite, suite, { 'Variable|Rental Value (ERV)': erv });
  }
  const currentPropertyBytes = makeWorkbook(propertyRefs, [
    comparisonProperty('Alpha', 'A', 'Healthcare', 44100000, 5000, currentJune),
    comparisonProperty('Beta', 'B', 'Healthcare', 17820000, 2000, currentJune),
    comparisonProperty('Gamma', 'C', 'Offices', 1000000, 100, currentJune),
    comparisonProperty('Echo', 'E', 'Healthcare', 10000000, 50, currentJune),
  ]);
  const previousPropertyBytes = makeWorkbook(propertyRefs, [
    comparisonProperty('Alpha', 'A', 'Healthcare', 46800000, 4000, previousJune),
    comparisonProperty('Beta', 'B', 'Healthcare', 17720000, 1500, previousJune),
    comparisonProperty('Delta', 'D', 'Healthcare', 999, 10, previousJune),
    comparisonProperty('Echo', 'E', 'Healthcare', 10000000, 80, previousJune),
  ]);
  const currentUnitBytes = makeWorkbook(unitRefs, [
    comparisonUnit('Alpha', 'Unit 01', 10),
    comparisonUnit('Alpha', 'Unit 02', 20),
    comparisonUnit('Alpha', 'Unit 03', 30),
    comparisonUnit('Beta', 'Unit 01', 50),
    comparisonUnit('Echo', 'Unit 01', 1),
    comparisonUnit('Echo', 'Unit 02', 1),
  ]);
  const previousUnitBytes = makeWorkbook(unitRefs, [
    comparisonUnit('Alpha', 'Unit 01', 15),
    comparisonUnit('Alpha', 'Unit 03', 25),
    comparisonUnit('Alpha', 'Unit 04', 99),
    comparisonUnit('Beta', 'Unit 01', 80),
    comparisonUnit('Echo', 'Unit 01', 5000000),
    comparisonUnit('Echo', 'Unit 02', 3000000),
    comparisonUnit('Echo', 'Unit 99', 2000000),
  ]);
  setReportLayout(defaultOptionalIds(), { market_rent_pa: true });
  const valuationCompare = transformToSchedule(
    await parseBuffer(currentPropertyBytes, 'current-property.xlsx', 'property'),
    await parseBuffer(currentUnitBytes, 'current-units.xlsx', 'unit'),
    {
      previousProperty: await parseBuffer(previousPropertyBytes, 'previous-property.xlsx', 'property'),
      previousUnit: await parseBuffer(previousUnitBytes, 'previous-units.xlsx', 'unit'),
    }
  );
  assert(valuationCompare.ok, 'comparison schedule generates');
  assert(!valuationCompare.scheduleRows.some(function (row) { return row.values && row.values.property_reference === 'D'; }), 'previous-only property is ignored');
  assert(!valuationCompare.scheduleRows.some(function (row) { return row.kind === 'unit' && row.values.unit === 'Unit 04'; }), 'previous-only unit is ignored');
  const fairAt = valuationCompare.columns.findIndex(function (column) { return column.id === 'fair_value'; });
  assert(valuationCompare.columns[fairAt + 1].id === 'fair_value__previous' && valuationCompare.columns[fairAt + 2].id === 'fair_value__change', 'fair value comparison columns stay beside the field');
  const rentAt = valuationCompare.columns.findIndex(function (column) { return column.id === 'market_rent_pa'; });
  assert(valuationCompare.columns[rentAt + 1].id === 'market_rent_pa__previous', 'market rent previous column is adjacent');
  assert(!valuationCompare.columns.some(function (column) { return column.id === 'market_rent_pa__change'; }), 'market rent has no change column');
  const alpha = valuationCompare.scheduleRows.find(function (row) { return row.kind === 'property' && row.values.property_reference === 'A'; });
  const beta = valuationCompare.scheduleRows.find(function (row) { return row.kind === 'property' && row.values.property_reference === 'B'; });
  const gamma = valuationCompare.scheduleRows.find(function (row) { return row.kind === 'property' && row.values.property_reference === 'C'; });
  const echo = valuationCompare.scheduleRows.find(function (row) { return row.kind === 'property' && row.values.property_reference === 'E'; });
  assert(alpha.values.fair_value__previous === 46800000 && gamma.values.fair_value__previous == null, 'property match uses External ID and unmatched stays blank');
  assert(closeNumber(alpha.values.fair_value__change, (44100000 / 46800000) - 1), 'fair value change uses current divided by previous');
  const echoUnits = valuationCompare.scheduleRows.filter(function (row) { return row.kind === 'unit' && row.matchKey === echo.matchKey; });
  const echoUnitRent = echoUnits.reduce(function (sum, row) { return sum + (row.values.market_rent_pa__previous || 0); }, 0);
  assert(echo.values.fair_value__previous === 10000000 && echoUnitRent === 8000000, 'previous fair value stays the direct property value when matched units would only reconstruct 8000000');
  const alphaUnits = valuationCompare.scheduleRows.filter(function (row) { return row.kind === 'unit' && row.matchKey === alpha.matchKey; });
  const unit01 = alphaUnits.find(function (row) { return row.values.unit === 'Unit 01'; });
  const unit02 = alphaUnits.find(function (row) { return row.values.unit === 'Unit 02'; });
  const unit03 = alphaUnits.find(function (row) { return row.values.unit === 'Unit 03'; });
  assert(unit01.values.market_rent_pa__previous === 15 && unit03.values.market_rent_pa__previous === 25, 'matched units keep their own previous rent');
  assert(unit02.values.market_rent_pa__previous == null, 'unmatched unit comparison stays blank');
  assert(unit01.values.fair_value__previous == null && unit01.values.fair_value__change == null, 'unit fair value comparison stays blank');
  const betaUnit = valuationCompare.scheduleRows.find(function (row) { return row.kind === 'unit' && row.values.unit === 'Unit 01' && row.values.market_rent_pa === 50; });
  assert(betaUnit.values.market_rent_pa__previous === 80, 'same unit name in another property does not cross-match');
  const valuationHealthcare = valuationCompare.scheduleRows.find(function (row) { return row.kind === 'sectorTotal' && row.sector === 'Healthcare'; });
  assert(valuationHealthcare.values.fair_value__previous === 46800000 + 17720000 + 10000000, 'sector previous fair value sums matched property values');
  const valuationGrand = valuationCompare.scheduleRows.find(function (row) { return row.kind === 'grandTotal'; });
  assert(valuationGrand.values.fair_value__previous === 46800000 + 17720000 + 10000000, 'grand previous fair value ignores previous-only properties');
  const valuationBook = buildWorkbookFromModel(valuationCompare).Sheets['Tenancy Schedule'];
  function cellFor(row, columnId) {
    const column = valuationCompare.columns.findIndex(function (item) { return item.id === columnId; });
    const rowIndex = valuationCompare.scheduleRows.indexOf(row) + 2;
    return valuationBook[XLSX.utils.encode_cell({ r: rowIndex, c: column })];
  }
  assert(cellFor(alpha, 'fair_value').s.font.color.rgb === 'FF0000' && cellFor(alpha, 'fair_value').s.font.bold === true, 'lower current fair value is exact red and stays bold');
  assert(cellFor(alpha, 'fair_value').s.fill.fgColor.rgb !== 'FFFFCC', 'property fair value does not use the unit change fill');
  assert(cellFor(alpha, 'fair_value__previous').s.font.color.rgb === '000000', 'previous fair value stays neutral');
  assert(cellFor(alpha, 'fair_value__change').s.font.color.rgb === 'FF0000' && cellFor(alpha, 'fair_value__change').z === '0.00%', 'negative fair value change is exact red percent');
  assert(cellFor(beta, 'fair_value').s.font.color.rgb === '00B050', 'higher current fair value is exact green');
  assert(cellFor(beta, 'fair_value__change').s.font.color.rgb === '00B050', 'positive fair value change is exact green');
  assert(cellFor(gamma, 'fair_value').s.font.color.rgb === '000000' && cellFor(gamma, 'fair_value').s.fill.fgColor.rgb !== 'FFFFCC', 'missing previous value is not coloured');
  const lowerUnit = cellFor(unit01, 'market_rent_pa');
  const higherUnit = cellFor(unit03, 'market_rent_pa');
  const unmatchedUnit = cellFor(unit02, 'market_rent_pa');
  const previousUnitRent = cellFor(unit01, 'market_rent_pa__previous');
  assert(lowerUnit.s.fill.fgColor.rgb === 'FFFFCC' && lowerUnit.s.font.color.rgb === 'FF0000' && lowerUnit.s.font.bold !== true, 'decreased unit current value is yellow and exact red, not bold');
  assert(higherUnit.s.fill.fgColor.rgb === 'FFFFCC' && higherUnit.s.font.color.rgb === '00B050' && higherUnit.s.font.bold !== true, 'increased unit current value is yellow and exact green, not bold');
  assert(unmatchedUnit.s.fill.fgColor.rgb !== 'FFFFCC' && unmatchedUnit.s.font.color.rgb === '000000', 'unit with no previous value stays unhighlighted');
  assert(previousUnitRent.s.fill.fgColor.rgb !== 'FFFFCC' && previousUnitRent.s.font.color.rgb === '000000', 'previous unit value stays neutral');
  assert(cellFor(valuationHealthcare, 'fair_value').s.font.bold === true && cellFor(valuationGrand, 'fair_value__change').s.font.color.rgb === 'FF0000', 'summary comparison keeps bold text and uses exact red');
  const valuationHeaders = Object.keys(valuationBook).filter(function (addr) { return addr[0] !== '!' && XLSX.utils.decode_cell(addr).r === 0; }).map(function (addr) { return valuationBook[addr].v; });
  assert(valuationHeaders.indexOf('Fair Value\nSeptember 2026') !== -1 && valuationHeaders.indexOf('Fair Value\nJune 2026') !== -1, 'fair value headers use each valuation date');
  assert(valuationHeaders.indexOf('Market Rent (pa)\nSeptember 2026') !== -1 && valuationHeaders.indexOf('Market Rent (pa)\nJune 2026') !== -1, 'compared field headers use both valuation dates');
  const propertyOnlyPrevious = transformToSchedule(
    await parseBuffer(currentPropertyBytes, 'current-property.xlsx', 'property'),
    await parseBuffer(currentUnitBytes, 'current-units.xlsx', 'unit'),
    { previousProperty: await parseBuffer(previousPropertyBytes, 'previous-property.xlsx', 'property') }
  );
  assert(propertyOnlyPrevious.columns.some(function (column) { return column.id === 'fair_value__previous'; }), 'previous property alone still enables fair value comparison');
  assert(propertyOnlyPrevious.scheduleRows.find(function (row) { return row.kind === 'unit' && row.values.unit === 'Unit 01' && row.values.market_rent_pa === 10; }).values.market_rent_pa__previous == null, 'unit comparison stays blank without a previous unit file');
  const plain = transformToSchedule(
    await parseBuffer(currentPropertyBytes, 'current-property.xlsx', 'property'),
    await parseBuffer(currentUnitBytes, 'current-units.xlsx', 'unit')
  );
  assert(!plain.columns.some(function (column) { return column.id === 'fair_value__previous' || column.id === 'fair_value__change'; }), 'fair value comparison stays hidden until a previous property file is loaded');

  const { describeRule, methodLabel } = await vite.ssrLoadModule('/src/logic/describe.js');
  const { getNumberFormat } = await vite.ssrLoadModule('/src/excel/writer.js');
  const capColumn = DEFAULT_SCHEDULE_COLUMNS.find(function (column) { return column.id === 'cap_ex'; });
  assert(capColumn.format === '£#,##0;-£#,##0', 'cap ex uses a normal negative currency format');
  assert(capColumn.format.indexOf('[Red]') === -1, 'cap ex format has no red negative colour');
  assert(getNumberFormat('cap_ex', -15000) === '£#,##0;-£#,##0'.replace(/£/g, '\u00a3'), 'excel cap ex format has no decimals or red');
  assert(methodLabel(capColumn.logic.property) === 'Source Field', 'property cap ex value method');
  assert(methodLabel(capColumn.logic.unit) === 'Aggregated Source / Derived Aggregation', 'unit cap ex value method');
  assert(methodLabel(capColumn.logic.sectorTotal) === 'Summary Aggregation', 'sector cap ex value method');
  assert(describeRule(capColumn.logic.property, 'property', DEFAULT_SCHEDULE_COLUMNS, 'cap_ex') === 'Property › General › Capital Costs\nUses Capital Costs directly from the Property Level source.', 'property cap ex logic text');
  assert(describeRule(capColumn.logic.unit, 'unit', DEFAULT_SCHEDULE_COLUMNS, 'cap_ex') === 'Unit / TS › Capital Cost › All Period Columns\nAggregation: Sum\nSign Convention: Cost / Outflow\nOutput: Negative absolute value\nSums all Capital Cost period values for the current Unit and reports the result as a negative cost.', 'unit cap ex logic text');
  assert(capColumn.logic.unit.sign === 'cost' && !capColumn.logic.property.sign, 'only unit cap ex uses the cost sign');
  assert(describeRule(capColumn.logic.sectorTotal, 'sectorTotal', DEFAULT_SCHEDULE_COLUMNS, 'cap_ex') === 'SUM Property › General › Capital Costs in Sector', 'sector cap ex logic text');
  assert(describeRule(capColumn.logic.grandTotal, 'grandTotal', DEFAULT_SCHEDULE_COLUMNS, 'cap_ex') === 'SUM Property › General › Capital Costs in Portfolio', 'grand cap ex logic text');
  assert(capColumn.logic.unit.ref.section === 'Capital Cost' && capColumn.logic.unit.ref.selection === 'allPeriods' && capColumn.logic.unit.ref.aggregation === 'sum' && capColumn.logic.unit.ref.scope === 'currentUnitRow', 'unit cap ex sums every period on the current row');
  assert(capColumn.logic.property.ref.section === 'General' && capColumn.logic.property.ref.field === 'Capital Costs', 'property cap ex is the static capital costs field');
  assert(defaultOptionalIds().indexOf('cap_ex') !== -1, 'cap ex stays an optional report field');
  const capDraft = spec.document.productionDefinition.columns.find(function (column) { return column.id === 'cap_ex'; });
  assert(capDraft && capDraft.logic.property.ref.field === 'Capital Costs' && capDraft.logic.unit.ref.section === 'Capital Cost' && capDraft.logic.sectorTotal.summary === 'sum' && capDraft.logic.sectorTotal.scope === 'properties', 'draft export carries the shared cap ex definition');

  const capPropertyRefs = propertyRefs.concat([
    { section: 'Capital Expenditure', field: 'Aug-2027' },
    { section: 'Capital Expenditure', field: 'Aug-2028' },
  ]);
  const capUnitRefs = unitRefs.concat([
    { section: 'Capital Cost', field: 'Aug-2028' },
    { section: 'Capital Cost', field: 'Aug-2029' },
  ]);
  const capPropertyBytes = makeWorkbook(capPropertyRefs, [
    valuesRow(capPropertyRefs, { 'General|Property Name': 'Asset A', 'General|Sector': 'Healthcare', 'General|External ID': 'ASSET-A', 'General|Traditional Valuation Date': valuationDate, 'General|City': 'Leeds', 'General|Capital Costs': -130000, 'Capital Expenditure|Aug-2027': 999999, 'Capital Expenditure|Aug-2028': 1 }),
    valuesRow(capPropertyRefs, { 'General|Property Name': 'Asset B', 'General|Sector': 'Healthcare', 'General|External ID': 'ASSET-B', 'General|Traditional Valuation Date': valuationDate, 'General|City': 'Leeds', 'General|Capital Costs': 70000, 'Capital Expenditure|Aug-2027': 500000 }),
    valuesRow(capPropertyRefs, { 'General|Property Name': 'Asset C', 'General|Sector': 'Offices', 'General|External ID': 'ASSET-C', 'General|Traditional Valuation Date': valuationDate, 'General|City': 'Leeds', 'General|Capital Costs': 50000, 'Capital Expenditure|Aug-2027': 800000 }),
    valuesRow(capPropertyRefs, { 'General|Property Name': 'Asset Blank', 'General|Sector': 'Healthcare', 'General|External ID': 'ASSET-BLANK', 'General|Traditional Valuation Date': valuationDate, 'General|City': 'Leeds', 'General|Capital Costs': null }),
  ]);
  const capUnitBytes = makeWorkbook(capUnitRefs, [
    valuesRow(capUnitRefs, { 'General|Property Name': 'Asset A', 'General|Suite': 'A', 'General|Tenant Name': 'Tenant A', 'Capital Cost|Aug-2027': 10000, 'Capital Cost|Aug-2028': 5000, 'Capital Cost|Aug-2029': 0 }),
    valuesRow(capUnitRefs, { 'General|Property Name': 'Asset A', 'General|Suite': 'B', 'General|Tenant Name': 'Tenant B', 'Capital Cost|Aug-2027': 2000, 'Capital Cost|Aug-2028': 3000 }),
    valuesRow(capUnitRefs, { 'General|Property Name': 'Asset B', 'General|Suite': 'C', 'General|Tenant Name': 'Tenant C', 'Capital Cost|Aug-2027': 999999 }),
  ]);
  const capModel = transformToSchedule(
    await parseBuffer(capPropertyBytes, 'cap-property.xlsx', 'property'),
    await parseBuffer(capUnitBytes, 'cap-units.xlsx', 'unit')
  );
  const capPropertyA = capModel.scheduleRows.find(function (row) { return row.kind === 'property' && row.values.address === 'Asset A'; });
  const capPropertyB = capModel.scheduleRows.find(function (row) { return row.kind === 'property' && row.values.address === 'Asset B'; });
  const capPropertyC = capModel.scheduleRows.find(function (row) { return row.kind === 'property' && row.values.address === 'Asset C'; });
  const capPropertyBlank = capModel.scheduleRows.find(function (row) { return row.kind === 'property' && row.values.address === 'Asset Blank'; });
  const capUnitA = capModel.scheduleRows.find(function (row) { return row.kind === 'unit' && row.values.unit === 'A'; });
  const capUnitB = capModel.scheduleRows.find(function (row) { return row.kind === 'unit' && row.values.unit === 'B'; });
  const capHealthcare = capModel.scheduleRows.find(function (row) { return row.kind === 'sectorTotal' && row.sector === 'Healthcare'; });
  const capGrand = capModel.scheduleRows.find(function (row) { return row.kind === 'grandTotal'; });
  assert(capPropertyA && capPropertyA.values.cap_ex === -130000, 'property cap ex keeps the source sign, got ' + (capPropertyA && capPropertyA.values.cap_ex));
  assert(capPropertyB && capPropertyB.values.cap_ex === 70000, 'positive property cap ex is not flipped by the unit rule');
  assert(capPropertyC && capPropertyC.values.cap_ex === 50000, 'property C cap ex');
  assert(capPropertyBlank && capPropertyBlank.values.cap_ex == null, 'blank capital costs stays blank');
  assert(capUnitA && capUnitA.values.cap_ex === -15000, 'unit A cap ex is a negative cost, got ' + (capUnitA && capUnitA.values.cap_ex));
  assert(capUnitB && capUnitB.values.cap_ex === -5000, 'unit B cap ex stays on its own row, got ' + (capUnitB && capUnitB.values.cap_ex));
  const capUnitC = capModel.scheduleRows.find(function (row) { return row.kind === 'unit' && row.values.unit === 'C'; });
  assert(capUnitC && capUnitC.values.cap_ex === -999999, 'unit C cap ex is row-scoped');
  assert(capHealthcare && capHealthcare.values.cap_ex === -60000, 'sector cap ex sums property capital costs once, got ' + (capHealthcare && capHealthcare.values.cap_ex));
  assert(capGrand && capGrand.values.cap_ex === -10000, 'grand cap ex sums property capital costs only, got ' + (capGrand && capGrand.values.cap_ex));
  const capBook = buildWorkbookFromModel(capModel);
  const capSheet = capBook.Sheets['Tenancy Schedule'];
  const capCol = capModel.columns.findIndex(function (column) { return column.id === 'cap_ex'; });
  const capExcelRow = capModel.scheduleRows.findIndex(function (row) { return row.kind === 'property' && row.values.address === 'Asset A'; });
  const capCell = capSheet[XLSX.utils.encode_cell({ r: capExcelRow + 2, c: capCol })];
  const capFormat = '£#,##0;-£#,##0'.replace(/£/g, '\u00a3');
  assert(capCell && capCell.v === -130000 && capCell.z === capFormat && capCell.z.indexOf('[Red]') === -1 && capCell.s.font.color.rgb === '000000', 'negative property cap ex stays black');
  const capUnitExcel = capModel.scheduleRows.findIndex(function (row) { return row.kind === 'unit' && row.values.unit === 'A'; });
  const capUnitCell = capSheet[XLSX.utils.encode_cell({ r: capUnitExcel + 2, c: capCol })];
  assert(capUnitCell && capUnitCell.v === -15000 && capUnitCell.z === capFormat && capUnitCell.s.font.color.rgb === '000000', 'negative unit cap ex stays black');
  const capTotalExcel = capModel.scheduleRows.findIndex(function (row) { return row.kind === 'sectorTotal' && row.sector === 'Healthcare'; });
  const capTotalCell = capSheet[XLSX.utils.encode_cell({ r: capTotalExcel + 2, c: capCol })];
  assert(capTotalCell && capTotalCell.v === -60000 && capTotalCell.s.font.color.rgb === 'FFFFFF', 'sector cap ex keeps the summary font');
  const capGrandExcel = capModel.scheduleRows.findIndex(function (row) { return row.kind === 'grandTotal'; });
  const capGrandCell = capSheet[XLSX.utils.encode_cell({ r: capGrandExcel + 2, c: capCol })];
  assert(capGrandCell && capGrandCell.v === -10000 && capGrandCell.s.font.color.rgb === 'FFFFFF', 'grand cap ex keeps the summary font');
  ['sector', 'property', 'unit', 'sectorTotal', 'grandTotal'].forEach(function (kind) {
    const rowIndex = capModel.scheduleRows.findIndex(function (row) { return row.kind === kind; });
    assert(rowIndex !== -1 && capSheet['!rows'][rowIndex + 2].hpt === 24, kind + ' cap workbook row height');
  });

  const missingCapitalCosts = propertyRefs.filter(function (ref) { return fieldKey(ref) !== 'General|Capital Costs'; });
  const missingCapitalCostGroup = unitRefs.filter(function (ref) { return ref.section !== 'Capital Cost'; });
  const nonPeriodCapitalCost = unitRefs.map(function (ref) {
    return ref.section === 'Capital Cost' ? { section: 'Capital Cost', field: 'Total' } : ref;
  });
  const missingCosts = await expectParse(makeWorkbook(missingCapitalCosts, [missingCapitalCosts.map(function () { return 'x'; })]), 'no-costs.xlsx', 'property');
  const missingGroupCap = await expectParse(makeWorkbook(missingCapitalCostGroup, [missingCapitalCostGroup.map(function () { return 'x'; })]), 'no-cap-group.xlsx', 'unit');
  const nonPeriodCap = await expectParse(makeWorkbook(nonPeriodCapitalCost, [nonPeriodCapitalCost.map(function (ref) {
    if (ref.field === 'Property Name') return 'Asset 1';
    if (ref.field === 'Suite') return '1';
    if (ref.field === 'Lease ID') return 'L1';
    return 'x';
  })]), 'cap-total.xlsx', 'unit');
  assert(typeof missingCosts === 'string' && missingCosts.indexOf('General > Capital Costs') !== -1, 'property validation requires Capital Costs: ' + missingCosts);
  assert(typeof missingGroupCap === 'string' && missingGroupCap.indexOf('Capital Cost') !== -1, 'unit validation requires Capital Cost group: ' + missingGroupCap);
  assert(typeof nonPeriodCap === 'string' && nonPeriodCap.indexOf('No period columns found under:') !== -1 && nonPeriodCap.indexOf('Capital Cost') !== -1, 'unit capital cost accepts any period date: ' + nonPeriodCap);

  const { roundedChangePercentagePoints, fairValueChangeTone } = await vite.ssrLoadModule('/src/logic/comparison.js');
  assert(roundedChangePercentagePoints(0.00004) === 0 && fairValueChangeTone(0.00004) === '', 'fair value change of +0.004% rounds to no change');
  assert(roundedChangePercentagePoints(-0.00004) === 0 && fairValueChangeTone(-0.00004) === '' && Object.is(roundedChangePercentagePoints(-0.00004), 0), 'fair value change of -0.004% rounds to positive zero');
  assert(roundedChangePercentagePoints(0.00006) === 0.01 && fairValueChangeTone(0.00006) === 'up', 'fair value change of +0.006% rounds to an increase');
  assert(roundedChangePercentagePoints(-0.00006) === -0.01 && fairValueChangeTone(-0.00006) === 'down', 'fair value change of -0.006% rounds to a decrease');
  const toleranceRows = [
    { name: 'up-noise', ratio: 0.00004, value: 0, color: '000000' },
    { name: 'down-noise', ratio: -0.00004, value: 0, color: '000000' },
    { name: 'up', ratio: 0.00006, value: 0.00006, color: '00B050' },
    { name: 'down', ratio: -0.00006, value: -0.00006, color: 'FF0000' },
  ];
  const toleranceModel = {
    columns: [
      { id: 'fair_value', label: 'Fair Value', comparisonRole: 'current', parentId: 'fair_value', directionHighlight: false },
      { id: 'fair_value__change', label: 'Change %', comparisonRole: 'change', parentId: 'fair_value', format: '0.00%' },
    ],
    scheduleRows: toleranceRows.map(function (item) {
      return { kind: 'property', values: { fair_value: 1, fair_value__change: item.ratio } };
    }),
    unmatchedRows: [],
  };
  const toleranceSheet = buildWorkbookFromModel(toleranceModel).Sheets['Tenancy Schedule'];
  toleranceRows.forEach(function (item, index) {
    const cell = toleranceSheet[XLSX.utils.encode_cell({ r: index + 2, c: 1 })];
    assert(cell.v === item.value && Object.is(cell.v, item.value) && cell.s.font.color.rgb === item.color && cell.z === '0.00%', 'fair value change ' + item.name + ' uses two-decimal percentage tolerance');
  });
  resetReportLayout();

  console.log(failures.length ? 'REGRESSION FAILED' : 'OK: regression');
} catch (err) {
  fail(err && err.stack ? err.stack : String(err));
} finally {
  await vite.close();
}

if (failures.length) process.exit(1);
