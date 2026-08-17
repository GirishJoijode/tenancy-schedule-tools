import {
  CONFIG_VERSION,
  CONFIG_OUTPUT_FILENAME,
  STANDARD_TEMPLATE_COLUMNS,
  STORAGE_KEY,
} from './constants.js';
import { defaultConfig } from './defaultConfig.js';
import { ensurePropertySummaryOnConfig } from './propertySummaryColumns.js';

const VALID_TARGET_IDS = new Set(
  STANDARD_TEMPLATE_COLUMNS.map(function (col) {
    return col.id;
  })
);

export function cloneDefaultConfig() {
  return ensurePropertySummaryOnConfig(JSON.parse(JSON.stringify(defaultConfig)));
}

function isValidSideMapping(mapping) {
  if (mapping === null || mapping === undefined) return true;
  if (typeof mapping !== 'object' || Array.isArray(mapping)) return false;
  if (mapping.type === 'blank') return true;
  if (mapping.type === 'source') {
    return typeof mapping.sourceKey === 'string' && mapping.sourceKey.length > 0;
  }
  return false;
}

/**
 * Structural validation for configuration objects coming from localStorage or
 * an imported JSON file. Unknown/extra keys are tolerated; unusable shapes are not.
 */
export function validateConfig(obj) {
  const errors = [];

  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) {
    return { valid: false, errors: ['Configuration must be a JSON object.'] };
  }

  const version = Number(obj.configVersion);
  if (!Number.isFinite(version)) {
    errors.push('configVersion is missing or not a number.');
  } else if (version > CONFIG_VERSION) {
    errors.push(
      'Configuration version ' + version + ' is newer than this application supports (' + CONFIG_VERSION + ').'
    );
  }

  if (!Array.isArray(obj.columnMappings)) {
    errors.push('columnMappings must be an array.');
  } else {
    const unknownIds = [];
    for (let i = 0; i < obj.columnMappings.length; i += 1) {
      const entry = obj.columnMappings[i];
      if (!entry || typeof entry !== 'object' || Array.isArray(entry)) {
        errors.push('columnMappings[' + i + '] must be an object.');
        continue;
      }
      if (typeof entry.targetColumnId !== 'string' || !entry.targetColumnId) {
        errors.push('columnMappings[' + i + '] is missing targetColumnId.');
        continue;
      }
      if (!VALID_TARGET_IDS.has(entry.targetColumnId)) {
        unknownIds.push(entry.targetColumnId);
        continue;
      }
      if (!isValidSideMapping(entry.propertyMapping)) {
        errors.push('Invalid propertyMapping for "' + entry.targetColumnId + '".');
      }
      if (!isValidSideMapping(entry.unitMapping)) {
        errors.push('Invalid unitMapping for "' + entry.targetColumnId + '".');
      }
    }
    if (unknownIds.length) {
      errors.push(
        'Unknown standardised column id(s): ' + unknownIds.slice(0, 8).join(', ') +
          (unknownIds.length > 8 ? ' and ' + (unknownIds.length - 8) + ' more' : '') + '.'
      );
    }
  }

  if (!isValidSideMapping(obj.propertyMatch)) {
    errors.push('propertyMatch has an invalid shape.');
  }
  if (!isValidSideMapping(obj.unitMatch)) {
    errors.push('unitMatch has an invalid shape.');
  }

  return { valid: errors.length === 0, errors: errors };
}

function getStorage() {
  try {
    return window.localStorage;
  } catch (err) {
    console.warn('localStorage is not available; configuration will not persist.', err);
    return null;
  }
}

/**
 * Returns the persisted configuration when it is present and structurally valid,
 * otherwise a fresh clone of the migrated default configuration.
 */
export function loadRuntimeConfig() {
  const storage = getStorage();
  if (!storage) return cloneDefaultConfig();

  let raw = null;
  try {
    raw = storage.getItem(STORAGE_KEY);
  } catch (err) {
    console.warn('Could not read stored configuration.', err);
    return cloneDefaultConfig();
  }

  if (!raw) return cloneDefaultConfig();

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (err) {
    console.warn('Stored configuration is not valid JSON; falling back to defaults.', err);
    return cloneDefaultConfig();
  }

  const check = validateConfig(parsed);
  if (!check.valid) {
    console.warn('Stored configuration failed validation; falling back to defaults.', check.errors);
    return cloneDefaultConfig();
  }

  return ensurePropertySummaryOnConfig(parsed);
}

export function saveRuntimeConfig(config) {
  const storage = getStorage();
  if (!storage) return false;
  try {
    storage.setItem(STORAGE_KEY, JSON.stringify(config));
    return true;
  } catch (err) {
    console.warn('Could not save configuration to localStorage.', err);
    return false;
  }
}

export function clearRuntimeConfig() {
  const storage = getStorage();
  if (!storage) return false;
  try {
    storage.removeItem(STORAGE_KEY);
    return true;
  } catch (err) {
    console.warn('Could not clear stored configuration.', err);
    return false;
  }
}

export function exportConfigJson(config, filename) {
  const name = filename || CONFIG_OUTPUT_FILENAME;
  const text = JSON.stringify(config, null, 2);
  const blob = new Blob([text], { type: 'application/json;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.rel = 'noopener';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(function () {
    URL.revokeObjectURL(url);
  }, 2000);
  return { ok: true, filename: name };
}

function readFileAsText(file) {
  if (typeof file.text === 'function') {
    return file.text();
  }
  return new Promise(function (resolve, reject) {
    const reader = new FileReader();
    reader.onload = function () {
      resolve(String(reader.result || ''));
    };
    reader.onerror = function () {
      reject(new Error('Could not read the selected file.'));
    };
    reader.readAsText(file);
  });
}

/**
 * Reads, parses and validates a configuration JSON file.
 * Resolves with the configuration object, or rejects with a readable Error.
 */
export function importConfigJson(file) {
  if (!file) {
    return Promise.reject(new Error('No configuration file selected.'));
  }

  const name = String(file.name || '');
  if (name && !name.toLowerCase().endsWith('.json')) {
    return Promise.reject(new Error('Please select a .json configuration file.'));
  }

  return readFileAsText(file).then(function (text) {
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (err) {
      throw new Error('The selected file is not valid JSON.');
    }

    const check = validateConfig(parsed);
    if (!check.valid) {
      throw new Error('The configuration file is not valid: ' + check.errors.join(' '));
    }

    return ensurePropertySummaryOnConfig(parsed);
  });
}
