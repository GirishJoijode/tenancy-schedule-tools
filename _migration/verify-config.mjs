/**
 * Confirms src/config/defaultConfig.js is a value-for-value copy of the
 * extracted standalone configuration.
 */
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '..');

const extracted = JSON.parse(readFileSync(resolve(here, 'extracted-config.json'), 'utf8'));
const { defaultConfig } = await import(pathToFileURL(resolve(root, 'src/config/defaultConfig.js')).href);

function sha256(obj) {
  return createHash('sha256').update(JSON.stringify(obj)).digest('hex');
}

const extractedHash = sha256(extracted);
const migratedHash = sha256(defaultConfig);

if (extractedHash !== migratedHash) {
  console.error('FAIL: defaultConfig.js does not match extracted-config.json');
  console.error('  extracted:', extractedHash);
  console.error('  migrated :', migratedHash);
  process.exit(1);
}

function countSource(side) {
  return defaultConfig.columnMappings.filter(function (entry) {
    return entry[side] && entry[side].type === 'source';
  }).length;
}

const propertyMappings = countSource('propertyMapping');
const unitMappings = countSource('unitMapping');

if (defaultConfig.columnMappings.length !== 32) {
  console.error('FAIL: expected 32 columnMappings, got', defaultConfig.columnMappings.length);
  process.exit(1);
}
if (propertyMappings !== 13) {
  console.error('FAIL: expected 13 property source mappings, got', propertyMappings);
  process.exit(1);
}
if (unitMappings !== 13) {
  console.error('FAIL: expected 13 unit source mappings, got', unitMappings);
  process.exit(1);
}

console.log('OK: defaultConfig.js matches extracted standalone configuration.');
console.log('  SHA-256           ', migratedHash);
console.log('  columnMappings    ', defaultConfig.columnMappings.length);
console.log('  property mappings ', propertyMappings);
console.log('  unit mappings     ', unitMappings);
console.log('  propertyMatch     ', defaultConfig.propertyMatch.sourceKey);
console.log('  unitMatch         ', defaultConfig.unitMatch.sourceKey);
