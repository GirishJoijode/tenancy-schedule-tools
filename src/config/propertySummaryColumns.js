function blankMapping() {
  return { type: 'blank' };
}

export const PROPERTY_SUMMARY_VALUATION_TITLE = 'VALUATION SUMMARY';

export const PROPERTY_SUMMARY_OUTPUT_FILENAME = 'Property_Summary.xlsx';
export const PROPERTY_SUMMARY_SHEET_NAME = 'Property Summary';

/**
 * Standardised Property Summary columns. Independent of Tenancy Schedule columns.
 * Edit this list to add, remove, or reorder output fields later.
 */
export const PROPERTY_SUMMARY_COLUMNS = [
  { id: 'property', label: 'Property' },
  { id: 'property_reference', label: 'Property Reference' },
  { id: 'contracted_rent_gross', label: 'Contracted Rent (Gross)' },
  { id: 'contracted_rent_net', label: 'Contracted Rent (Net)' },
  { id: 'market_rent', label: 'Market Rent' },
  { id: 'niy', label: 'NIY' },
  { id: 'ney', label: 'NEY' },
  { id: 'ry', label: 'RY' },
  { id: 'market_value', label: 'Market Value' },
  { id: 'market_value_change', label: 'Market Value Change' },
  { id: 'market_rent_change', label: 'Market Rent Change' },
  { id: 'market_value_change_pct', label: 'Market Value Change %' },
  { id: 'draft_valuation_comments', label: 'DRAFT VALUATION COMMENTS' },
  { id: 'draft_to_final_valuation_comments', label: 'DRAFT TO FINAL VALUATION COMMENTS' },
];

export const PROPERTY_SUMMARY_NUM_FMT = {
  contracted_rent_gross: '\u00a3#,##0',
  contracted_rent_net: '\u00a3#,##0',
  market_rent: '\u00a3#,##0',
  niy: '0.00%',
  ney: '0.00%',
  ry: '0.00%',
  market_value: '\u00a3#,##0',
  market_value_change: '\u00a3#,##0;[Red]-\u00a3#,##0',
  market_rent_change: '\u00a3#,##0;[Red]-\u00a3#,##0',
  market_value_change_pct: '0.00%;[Red]-0.00%',
};

export const PROPERTY_SUMMARY_COL_WIDTHS = {
  property: 36,
  property_reference: 16,
  contracted_rent_gross: 16,
  contracted_rent_net: 16,
  market_rent: 14,
  niy: 8,
  ney: 8,
  ry: 8,
  market_value: 16,
  market_value_change: 16,
  market_rent_change: 16,
  market_value_change_pct: 16,
  draft_valuation_comments: 42,
  draft_to_final_valuation_comments: 42,
};

export const PROPERTY_SUMMARY_WRAP_IDS = {
  draft_valuation_comments: true,
  draft_to_final_valuation_comments: true,
  property: true,
};

export const PROPERTY_SUMMARY_VALUATION_IDS = [
  'contracted_rent_gross',
  'contracted_rent_net',
  'market_rent',
  'niy',
  'ney',
  'ry',
  'market_value',
  'market_value_change',
  'market_rent_change',
  'market_value_change_pct',
];

export function createEmptyPropertySummaryConfig() {
  return {
    columnMappings: PROPERTY_SUMMARY_COLUMNS.map(function (col) {
      return {
        targetColumnId: col.id,
        propertyMapping: blankMapping(),
        ignore: false,
      };
    }),
  };
}

export function normalisePropertySummaryConfig(raw) {
  var byId = new Map();
  var source = raw && typeof raw === 'object' ? raw : null;
  var list = source && Array.isArray(source.columnMappings) ? source.columnMappings : [];

  for (var i = 0; i < list.length; i += 1) {
    var entry = list[i];
    if (entry && typeof entry.targetColumnId === 'string') {
      byId.set(entry.targetColumnId, entry);
    }
  }

  return {
    columnMappings: PROPERTY_SUMMARY_COLUMNS.map(function (col) {
      var existing = byId.get(col.id);
      var mapping = existing && existing.propertyMapping;
      if (mapping && mapping.type === 'source' && typeof mapping.sourceKey === 'string' && mapping.sourceKey) {
        return {
          targetColumnId: col.id,
          propertyMapping: {
            type: 'source',
            sourceKey: mapping.sourceKey,
            sourceSection: mapping.sourceSection != null ? mapping.sourceSection : '',
            sourceField: mapping.sourceField != null ? mapping.sourceField : '',
            occurrence: Number(mapping.occurrence) || 1,
          },
          ignore: existing.ignore === true,
        };
      }
      return {
        targetColumnId: col.id,
        propertyMapping: blankMapping(),
        ignore: existing ? existing.ignore === true : false,
      };
    }),
  };
}

export function ensurePropertySummaryOnConfig(config) {
  if (!config || typeof config !== 'object') return config;
  config.propertySummary = normalisePropertySummaryConfig(config.propertySummary);
  return config;
}

export function getPropertySummaryMappingSummary(psConfig, propertyParsed) {
  var mapped = 0;
  var blank = 0;
  var missing = 0;
  var ignored = 0;
  var columns = (propertyParsed && propertyParsed.columns) || null;
  var mappings = (psConfig && psConfig.columnMappings) || [];

  for (var i = 0; i < PROPERTY_SUMMARY_COLUMNS.length; i += 1) {
    var col = PROPERTY_SUMMARY_COLUMNS[i];
    var entry = null;
    for (var j = 0; j < mappings.length; j += 1) {
      if (mappings[j].targetColumnId === col.id) {
        entry = mappings[j];
        break;
      }
    }
    if (entry && entry.ignore === true) {
      ignored += 1;
      continue;
    }
    var mapping = entry && entry.propertyMapping;
    if (!mapping || mapping.type !== 'source') {
      blank += 1;
      continue;
    }
    if (columns && !columns.some(function (c) { return c.key === mapping.sourceKey; })) {
      missing += 1;
    } else {
      mapped += 1;
    }
  }

  return {
    standardColumns: PROPERTY_SUMMARY_COLUMNS.length,
    propertyMappings: mapped,
    blankColumns: blank,
    unmappedColumns: blank,
    ignoredColumns: ignored,
    missingMappings: missing,
  };
}

export function getPropertySummaryColumn(id) {
  for (var i = 0; i < PROPERTY_SUMMARY_COLUMNS.length; i += 1) {
    if (PROPERTY_SUMMARY_COLUMNS[i].id === id) return PROPERTY_SUMMARY_COLUMNS[i];
  }
  return null;
}
