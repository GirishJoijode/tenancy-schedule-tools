/**
 * Property Only schedule. Column order, labels, sources, and number formats are fixed here.
 */
import { PROPERTY_SOURCE_FIELDS } from '../schema/sourceFields.js';
import { attachPropertyOnlyLogic } from '../logic/attachFieldLogic.js';

const PROPERTY_ONLY_COLUMN_DEFS = [
  {
    id: 'property',
    label: 'Property',
    property: PROPERTY_SOURCE_FIELDS.propertyName,
  },
  {
    id: 'property_reference',
    label: 'Property Reference',
    property: PROPERTY_SOURCE_FIELDS.externalId,
  },
  {
    id: 'contracted_rent_gross',
    label: 'Contracted Rent (Gross)',
    property: PROPERTY_SOURCE_FIELDS.contractedRent,
  },
  {
    id: 'contracted_rent_net',
    label: 'Contracted Rent (Net)',
    property: PROPERTY_SOURCE_FIELDS.netRent,
  },
  {
    id: 'market_rent',
    label: 'Market Rent',
    property: PROPERTY_SOURCE_FIELDS.rentalValue,
  },
  {
    id: 'niy',
    label: 'NIY',
    property: PROPERTY_SOURCE_FIELDS.initialYield,
  },
  {
    id: 'ney',
    label: 'NEY',
    property: PROPERTY_SOURCE_FIELDS.equivalentYield,
  },
  {
    id: 'ry',
    label: 'RY',
    property: PROPERTY_SOURCE_FIELDS.reversionaryYield,
  },
  {
    id: 'market_value',
    label: 'Market Value',
    property: PROPERTY_SOURCE_FIELDS.sayValue,
  },
  {
    id: 'market_value_change',
    label: 'Market Value Change',
    property: null,
  },
  {
    id: 'market_rent_change',
    label: 'Market Rent Change',
    property: null,
  },
  {
    id: 'market_value_change_pct',
    label: 'Market Value Change %',
    property: null,
  },
  {
    id: 'draft_valuation_comments',
    label: 'DRAFT VALUATION COMMENTS',
    property: null,
  },
  {
    id: 'draft_to_final_valuation_comments',
    label: 'DRAFT TO FINAL VALUATION COMMENTS',
    property: null,
  },
];

export const PROPERTY_ONLY_COLUMNS = attachPropertyOnlyLogic(PROPERTY_ONLY_COLUMN_DEFS);

export const PROPERTY_ONLY_OUTPUT_FILENAME = 'Property_Summary.xlsx';
export const PROPERTY_ONLY_SHEET_NAME = 'Property Summary';
export const PROPERTY_ONLY_VALUATION_TITLE = 'VALUATION SUMMARY';

export const PROPERTY_ONLY_NUM_FMT = {
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

export const PROPERTY_ONLY_COL_WIDTHS = {
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

export const PROPERTY_ONLY_WRAP_IDS = {
  draft_valuation_comments: true,
  draft_to_final_valuation_comments: true,
  property: true,
};

export const PROPERTY_ONLY_VALUATION_IDS = [
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
