import {
  APP_VERSION,
  CONFIG_VERSION,
  DEFAULT_MATCH_HINTS,
  STANDARD_TEMPLATE_COLUMNS,
} from './constants.js';
import { loadRuntimeConfig, saveRuntimeConfig } from './configManager.js';
import { ensurePropertySummaryOnConfig, normalisePropertySummaryConfig } from './propertySummaryColumns.js';
import { normaliseHeader } from '../excel/parser.js';
import { APP_STATE } from '../state.js';

export function blankMapping() {
  return { type: 'blank' };
}

export function sourceMappingFromColumn(sourceColumn) {
  return {
    type: 'source',
    sourceKey: sourceColumn.key,
    sourceSection: sourceColumn.section,
    sourceField: sourceColumn.field,
    occurrence: sourceColumn.occurrence,
  };
}

export function createEmptyMappingConfig() {
  var config = {
    configVersion: CONFIG_VERSION,
    applicationVersion: APP_VERSION,
    propertyMatch: blankMapping(),
    unitMatch: blankMapping(),
    columnMappings: STANDARD_TEMPLATE_COLUMNS.map(function (col) {
      return {
        targetColumnId: col.id,
        propertyMapping: blankMapping(),
        unitMapping: blankMapping(),
        ignore: false,
      };
    }),
  };
  return ensurePropertySummaryOnConfig(config);
}

export function cloneMappingConfig(config) {
  return JSON.parse(JSON.stringify(config));
}

export function normaliseSideMapping(mapping) {
  if (!mapping || typeof mapping !== 'object') return blankMapping();
  if (mapping.type === 'source' && typeof mapping.sourceKey === 'string' && mapping.sourceKey) {
    return {
      type: 'source',
      sourceKey: mapping.sourceKey,
      sourceSection: mapping.sourceSection != null ? mapping.sourceSection : '',
      sourceField: mapping.sourceField != null ? mapping.sourceField : '',
      occurrence: Number(mapping.occurrence) || 1,
    };
  }
  return blankMapping();
}

export function normaliseMappingConfig(raw) {
  var base = createEmptyMappingConfig();
  if (!raw || typeof raw !== 'object') return base;

  if (raw.propertyMatch) {
    base.propertyMatch = normaliseSideMapping(raw.propertyMatch);
  }
  if (raw.unitMatch) {
    base.unitMatch = normaliseSideMapping(raw.unitMatch);
  }

  var byId = new Map();
  if (Array.isArray(raw.columnMappings)) {
    for (var i = 0; i < raw.columnMappings.length; i += 1) {
      var entry = raw.columnMappings[i];
      if (entry && typeof entry.targetColumnId === 'string') {
        byId.set(entry.targetColumnId, entry);
      }
    }
  }

  base.columnMappings = STANDARD_TEMPLATE_COLUMNS.map(function (col) {
    var existing = byId.get(col.id);
    if (!existing) {
      return {
        targetColumnId: col.id,
        propertyMapping: blankMapping(),
        unitMapping: blankMapping(),
        ignore: false,
      };
    }
    return {
      targetColumnId: col.id,
      propertyMapping: normaliseSideMapping(existing.propertyMapping),
      unitMapping: normaliseSideMapping(existing.unitMapping),
      ignore: existing.ignore === true,
    };
  });

  base.configVersion = Number(raw.configVersion) || CONFIG_VERSION;
  base.applicationVersion = String(raw.applicationVersion || APP_VERSION);
  base.propertySummary = normalisePropertySummaryConfig(raw.propertySummary);

  return base;
}

export function resolveSourceKey(columns, sourceKey) {
  if (!columns || !sourceKey) return null;
  for (var i = 0; i < columns.length; i += 1) {
    if (columns[i].key === sourceKey) return columns[i];
  }
  return null;
}

export function findExactSourceColumn(columns, section, field) {
  if (!columns) return null;
  var sectionKey = normaliseHeader(section);
  var fieldKey = normaliseHeader(field);
  for (var i = 0; i < columns.length; i += 1) {
    var c = columns[i];
    if (
      normaliseHeader(c.section) === sectionKey &&
      normaliseHeader(c.field) === fieldKey &&
      c.occurrence === 1
    ) {
      return c;
    }
  }
  return null;
}

export function applyDefaultMatchFields(config, propertyParsed, unitParsed) {
  if (config.propertyMatch && config.propertyMatch.type !== 'source') {
    var propCol = findExactSourceColumn(
      propertyParsed && propertyParsed.columns,
      DEFAULT_MATCH_HINTS.property.section,
      DEFAULT_MATCH_HINTS.property.field
    );
    if (propCol) config.propertyMatch = sourceMappingFromColumn(propCol);
  }

  if (config.unitMatch && config.unitMatch.type !== 'source') {
    var unitCol = findExactSourceColumn(
      unitParsed && unitParsed.columns,
      DEFAULT_MATCH_HINTS.unit.section,
      DEFAULT_MATCH_HINTS.unit.field
    );
    if (unitCol) config.unitMatch = sourceMappingFromColumn(unitCol);
  }

  return config;
}

export function resolveSideStatus(mapping, columns) {
  if (!mapping || mapping.type !== 'source') return 'blank';
  if (!columns) {
    return 'mapped';
  }
  return resolveSourceKey(columns, mapping.sourceKey) ? 'mapped' : 'missing';
}

export function getColumnMappingStatus(entry, propertyParsed, unitParsed) {
  if (entry && entry.ignore === true) return 'Ignored';

  var propStatus = resolveSideStatus(entry.propertyMapping, propertyParsed && propertyParsed.columns);
  var unitStatus = resolveSideStatus(entry.unitMapping, unitParsed && unitParsed.columns);

  if (propStatus === 'missing' || unitStatus === 'missing') {
    return 'Missing source';
  }

  var propMapped = propStatus === 'mapped';
  var unitMapped = unitStatus === 'mapped';

  if (propMapped && unitMapped) return 'Mapped';
  if (propMapped || unitMapped) return 'Partially mapped';
  return 'Unmapped';
}

export function getMappingSummary(config, propertyParsed, unitParsed) {
  var propertyMappings = 0;
  var unitMappings = 0;
  var missingMappings = 0;
  var mappedColumns = 0;
  var partiallyMapped = 0;
  var unmappedColumns = 0;
  var ignoredColumns = 0;

  for (var i = 0; i < config.columnMappings.length; i += 1) {
    var entry = config.columnMappings[i];
    var status = getColumnMappingStatus(entry, propertyParsed, unitParsed);
    if (entry.propertyMapping && entry.propertyMapping.type === 'source') propertyMappings += 1;
    if (entry.unitMapping && entry.unitMapping.type === 'source') unitMappings += 1;

    if (status === 'Ignored') ignoredColumns += 1;
    else if (status === 'Missing source') missingMappings += 1;
    else if (status === 'Mapped') mappedColumns += 1;
    else if (status === 'Partially mapped') partiallyMapped += 1;
    else unmappedColumns += 1;
  }

  if (resolveSideStatus(config.propertyMatch, propertyParsed && propertyParsed.columns) === 'missing') {
    missingMappings += 1;
  }
  if (resolveSideStatus(config.unitMatch, unitParsed && unitParsed.columns) === 'missing') {
    missingMappings += 1;
  }

  return {
    standardColumns: STANDARD_TEMPLATE_COLUMNS.length,
    propertyMappings: propertyMappings,
    unitMappings: unitMappings,
    missingMappings: missingMappings,
    mappedColumns: mappedColumns,
    partiallyMapped: partiallyMapped,
    unmappedColumns: unmappedColumns,
    ignoredColumns: ignoredColumns,
    blankColumns: unmappedColumns,
  };
}

export function isMappingConfigured(config, propertyParsed, unitParsed) {
  config = config != null ? config : APP_STATE.mappingConfig;
  propertyParsed = propertyParsed != null ? propertyParsed : APP_STATE.property;
  unitParsed = unitParsed != null ? unitParsed : APP_STATE.unit;

  if (!config) return false;

  if (config.propertyMatch && config.propertyMatch.type !== 'source') return false;
  if (config.unitMatch && config.unitMatch.type !== 'source') return false;

  if (!propertyParsed || !unitParsed) return false;

  var summary = getMappingSummary(config, propertyParsed, unitParsed);
  return summary.missingMappings === 0;
}

export function getMainMappingStatusText() {
  var config = APP_STATE.mappingConfig;
  if (!config) return 'Not configured';

  var hasFiles = Boolean(APP_STATE.property && APP_STATE.unit);

  if (!hasFiles) {
    if (
      (config.propertyMatch && config.propertyMatch.type === 'source') ||
      (config.unitMatch && config.unitMatch.type === 'source')
    ) {
      return 'Configuration loaded';
    }
    var hasAny = config.columnMappings.some(function (e) {
      return (
        (e.propertyMapping && e.propertyMapping.type === 'source') ||
        (e.unitMapping && e.unitMapping.type === 'source')
      );
    });
    return hasAny ? 'Configuration loaded' : 'Not configured';
  }

  var summary = getMappingSummary(config, APP_STATE.property, APP_STATE.unit);
  if (summary.missingMappings > 0) {
    return 'Missing sources (' + summary.missingMappings + ')';
  }

  if (
    (config.propertyMatch && config.propertyMatch.type !== 'source') ||
    (config.unitMatch && config.unitMatch.type !== 'source')
  ) {
    return 'Match fields required';
  }

  return '\u2713 Configured';
}

export function toEmbeddableConfig(config) {
  var normalised = normaliseMappingConfig(config);
  return {
    configVersion: CONFIG_VERSION,
    applicationVersion: APP_VERSION,
    propertyMatch: normalised.propertyMatch,
    unitMatch: normalised.unitMatch,
    columnMappings: normalised.columnMappings,
    propertySummary: normalised.propertySummary,
  };
}

/**
 * Seeds APP_STATE from the persisted runtime configuration, falling back to the
 * migrated default configuration when nothing valid is stored.
 */
export function initialiseMappingState() {
  APP_STATE.mappingConfig = normaliseMappingConfig(loadRuntimeConfig());
  return APP_STATE.mappingConfig;
}

export function applyMappingConfig(config) {
  APP_STATE.mappingConfig = normaliseMappingConfig(config);
  APP_STATE.mappingConfigured = isMappingConfigured();
  saveRuntimeConfig(toEmbeddableConfig(APP_STATE.mappingConfig));
  return APP_STATE.mappingConfig;
}

export function getTemplateColumn(id) {
  for (var i = 0; i < STANDARD_TEMPLATE_COLUMNS.length; i += 1) {
    if (STANDARD_TEMPLATE_COLUMNS[i].id === id) return STANDARD_TEMPLATE_COLUMNS[i];
  }
  return null;
}

export function getMappingConfig() {
  return APP_STATE.mappingConfig;
}
