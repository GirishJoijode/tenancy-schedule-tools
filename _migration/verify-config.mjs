/**
 * Confirms src/config/defaultConfig.js is a complete shipped default
 * and that normalisation preserves it value-for-value.
 */
import { createHash } from 'node:crypto';
import { dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '..');

const { defaultConfig } = await import(pathToFileURL(resolve(root, 'src/config/defaultConfig.js')).href);
const { normaliseMappingConfig, cloneMappingConfig } = await import(
  pathToFileURL(resolve(root, 'src/config/mapping.js')).href
);
const { cloneDefaultConfig } = await import(
  pathToFileURL(resolve(root, 'src/config/configManager.js')).href
);

function sha256(obj) {
  return createHash('sha256').update(JSON.stringify(obj)).digest('hex');
}

function fail(message) {
  console.error('FAIL:', message);
  process.exit(1);
}

function assertEqual(label, a, b) {
  if (JSON.stringify(a) !== JSON.stringify(b)) fail(label);
}

const normalised = normaliseMappingConfig(defaultConfig);
assertEqual('propertyMatch', normalised.propertyMatch, defaultConfig.propertyMatch);
assertEqual('unitMatch', normalised.unitMatch, defaultConfig.unitMatch);
assertEqual('columnMappings', normalised.columnMappings, defaultConfig.columnMappings);
assertEqual('propertySummary', normalised.propertySummary, defaultConfig.propertySummary);

if (normalised.columnMappings.length !== 32) {
  fail('expected 32 Default columns, got ' + normalised.columnMappings.length);
}
if (!normalised.propertySummary || normalised.propertySummary.columnMappings.length !== 14) {
  fail('expected 14 Property Only columns');
}

function countSource(list, side) {
  return list.filter(function (entry) {
    return entry[side] && entry[side].type === 'source';
  }).length;
}

const propertyMappings = countSource(normalised.columnMappings, 'propertyMapping');
const unitMappings = countSource(normalised.columnMappings, 'unitMapping');
const ignoredColumns = normalised.columnMappings.filter(function (entry) {
  return entry.ignore === true;
}).length;
const unmappedColumns = normalised.columnMappings.filter(function (entry) {
  const prop = entry.propertyMapping && entry.propertyMapping.type === 'source';
  const unit = entry.unitMapping && entry.unitMapping.type === 'source';
  return entry.ignore !== true && !prop && !unit;
}).length;

const psMapped = countSource(normalised.propertySummary.columnMappings, 'propertyMapping');
const psIgnored = normalised.propertySummary.columnMappings.filter(function (entry) {
  return entry.ignore === true;
}).length;
const psUnmapped = normalised.propertySummary.columnMappings.filter(function (entry) {
  const prop = entry.propertyMapping && entry.propertyMapping.type === 'source';
  return entry.ignore !== true && !prop;
}).length;

if (propertyMappings < 1 || unitMappings < 1) {
  fail('Default template unexpectedly has no source mappings');
}
if (psMapped < 1) {
  fail('Property Only unexpectedly has no source mappings');
}

const resetClone = cloneDefaultConfig();
const resetNormalised = normaliseMappingConfig(resetClone);
assertEqual('Reset propertyMatch', resetNormalised.propertyMatch, normalised.propertyMatch);
assertEqual('Reset unitMatch', resetNormalised.unitMatch, normalised.unitMatch);
assertEqual('Reset columnMappings', resetNormalised.columnMappings, normalised.columnMappings);
assertEqual('Reset propertySummary', resetNormalised.propertySummary, normalised.propertySummary);

const noStorageClone = cloneMappingConfig(cloneDefaultConfig());
assertEqual(
  'No-localStorage default',
  normaliseMappingConfig(noStorageClone).columnMappings,
  normalised.columnMappings
);
assertEqual(
  'No-localStorage Property Only',
  normaliseMappingConfig(noStorageClone).propertySummary,
  normalised.propertySummary
);

const { getActiveOutputColumns } = await import(
  pathToFileURL(resolve(root, 'src/standardiser/transform.js')).href
);
const { getActivePropertySummaryColumns } = await import(
  pathToFileURL(resolve(root, 'src/propertySummary/propertySummary.js')).href
);
const { getUnitMappedCompareColumns } = await import(
  pathToFileURL(resolve(root, 'src/compare/engine.js')).href
);

const defaultOutput = getActiveOutputColumns(normalised);
if (defaultOutput.length !== 32 - ignoredColumns) {
  fail('Default output columns should be template minus ignored');
}

const outputProbe = JSON.parse(JSON.stringify(normalised));
let probeIgnored = 0;
for (let i = 0; i < outputProbe.columnMappings.length; i += 1) {
  const entry = outputProbe.columnMappings[i];
  const propBlank = !entry.propertyMapping || entry.propertyMapping.type !== 'source';
  const unitBlank = !entry.unitMapping || entry.unitMapping.type !== 'source';
  if (propBlank && unitBlank && probeIgnored < 2) {
    entry.ignore = true;
    probeIgnored += 1;
  }
}
const tsOutput = getActiveOutputColumns(outputProbe);
if (tsOutput.length !== 30) {
  fail('expected 30 Default output columns with 2 ignored, got ' + tsOutput.length);
}
const ignoredIds = outputProbe.columnMappings
  .filter(function (entry) {
    return entry.ignore === true;
  })
  .map(function (entry) {
    return entry.targetColumnId;
  });
if (
  tsOutput.some(function (col) {
    return ignoredIds.indexOf(col.id) !== -1;
  })
) {
  fail('ignored Default columns were still written');
}

const compareCols = getUnitMappedCompareColumns(outputProbe);
if (
  compareCols.some(function (col) {
    return ignoredIds.indexOf(col.id) !== -1;
  })
) {
  fail('ignored columns were included in TS Compare');
}

const psProbe = JSON.parse(JSON.stringify(normalised.propertySummary));
psProbe.columnMappings[1].ignore = true;
psProbe.columnMappings[2].ignore = true;
const psOutput = getActivePropertySummaryColumns(psProbe);
if (psOutput.length !== 12) {
  fail('expected 12 Property Only output columns with 2 ignored, got ' + psOutput.length);
}
if (psOutput[0].propertyMapping.type !== 'source') {
  fail('Property Only ignore must preserve stored mappings for included columns');
}

console.log('OK: shipped default configuration is complete and normalisation-stable.');
console.log('  SHA-256              ', sha256(normalised));
console.log('  Default columns      ', normalised.columnMappings.length);
console.log('  Default property     ', propertyMappings);
console.log('  Default unit         ', unitMappings);
console.log('  Default ignored      ', ignoredColumns);
console.log('  Default unmapped     ', unmappedColumns);
console.log('  Property Only cols   ', normalised.propertySummary.columnMappings.length);
console.log('  Property Only mapped ', psMapped);
console.log('  Property Only ignored', psIgnored);
console.log('  Property Only blank  ', psUnmapped);
console.log('  propertyMatch        ', normalised.propertyMatch.sourceKey);
console.log('  unitMatch            ', normalised.unitMatch.sourceKey);
