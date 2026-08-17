/**
 * One-time migration helper: extract config/CSS/HTML/JS from standalone index.html
 */
import fs from 'fs';
import crypto from 'crypto';
import path from 'path';

const root = process.cwd();
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

// Backup original
fs.copyFileSync(path.join(root, 'index.html'), path.join(root, 'index.standalone.backup.html'));

const cfgMatch = html.match(/<script id="embedded-config"[^>]*>([\s\S]*?)<\/script>/);
if (!cfgMatch) throw new Error('No embedded-config');
const cfgRaw = cfgMatch[1].trim();
const cfg = JSON.parse(cfgRaw);
fs.mkdirSync(path.join(root, '_migration'), { recursive: true });
fs.writeFileSync(path.join(root, '_migration', 'extracted-config.json'), JSON.stringify(cfg, null, 2));
fs.writeFileSync(path.join(root, '_migration', 'extracted-config.raw.json'), cfgRaw);

const propMapped = cfg.columnMappings.filter((m) => m.propertyMapping?.type === 'source').length;
const unitMapped = cfg.columnMappings.filter((m) => m.unitMapping?.type === 'source').length;
const hash = crypto.createHash('sha256').update(JSON.stringify(cfg)).digest('hex');

console.log('CONFIG_HASH', hash);
console.log('configVersion', cfg.configVersion);
console.log('applicationVersion', cfg.applicationVersion);
console.log('propertyMapped', propMapped);
console.log('unitMapped', unitMapped);
console.log('propertyMatch', cfg.propertyMatch);
console.log('unitMatch', cfg.unitMatch);
console.log('columnMappings', cfg.columnMappings.length);

// Extract style
const styleMatch = html.match(/<style>([\s\S]*?)<\/style>/);
if (!styleMatch) throw new Error('No style');
fs.writeFileSync(path.join(root, '_migration', 'extracted.css'), styleMatch[1]);
console.log('CSS_LEN', styleMatch[1].length);

// Extract app JS (after marker)
const appMarker = '/* ===== Tenancy Schedule Standardiser application ===== */';
const appStart = html.indexOf(appMarker);
if (appStart < 0) throw new Error('No app marker');
const afterMarker = html.slice(appStart + appMarker.length);
const scriptEnd = afterMarker.indexOf('</script>');
const appJs = afterMarker.slice(0, scriptEnd).trim();
// Strip IIFE wrapper
let body = appJs;
if (body.startsWith('(function () {')) {
  body = body.replace(/^\(function \(\) \{\s*'use strict';\s*/, '');
  body = body.replace(/\}\)\(\);\s*$/, '');
}
// Remove window.XLSX bootstrap
body = body.replace(
  /var XLSX = window\.XLSX;\s*if \(!XLSX[\s\S]*?return;\s*\}\s*/,
  ''
);

fs.writeFileSync(path.join(root, '_migration', 'extracted-app.js'), body);
console.log('APP_JS_LEN', body.length);

// Extract body HTML between <body> and first <script
const bodyMatch = html.match(/<body>([\s\S]*?)<script>/);
if (!bodyMatch) throw new Error('No body');
let bodyHtml = bodyMatch[1];
// Remove embedded-config from body html (we'll load from JS)
bodyHtml = bodyHtml.replace(/<!-- Embedded mapping[\s\S]*?<script id="embedded-config"[\s\S]*?<\/script>\s*/i, '\n');
bodyHtml = bodyHtml.replace(/<script id="embedded-config"[\s\S]*?<\/script>\s*/i, '\n');
fs.writeFileSync(path.join(root, '_migration', 'extracted-body.html'), bodyHtml.trim());
console.log('BODY_LEN', bodyHtml.length);

// List function names
const fns = [...body.matchAll(/function\s+([A-Za-z0-9_]+)\s*\(/g)].map((m) => m[1]);
console.log('FUNCTIONS', fns.length);
fs.writeFileSync(path.join(root, '_migration', 'functions.json'), JSON.stringify(fns, null, 2));
