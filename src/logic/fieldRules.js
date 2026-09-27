import { PROPERTY_SOURCE_FIELDS, UNIT_SOURCE_FIELDS } from '../schema/sourceFields.js';

function source(ref) {
  return { method: 'source', ref: ref };
}

function aggregated(ref) {
  return { method: 'aggregatedSource', ref: ref };
}

function blank() {
  return { method: 'blank' };
}

function sum() {
  return { method: 'summary', summary: 'sum', scope: 'properties' };
}

function sumMatchedUnits(scopeNote) {
  return {
    method: 'summary',
    summary: 'sum',
    scope: 'matchedUnits',
    source: { kind: 'templateField', id: 'area_sqft' },
    scopeNote: scopeNote || 'property',
  };
}

function divideFields(numeratorId, denominatorId, scopeNote) {
  return {
    method: 'derived',
    calculation: 'divide',
    numerator: { kind: 'templateField', id: numeratorId },
    denominator: { kind: 'templateField', id: denominatorId },
    scopeNote: scopeNote || 'property',
  };
}

function recalculate(numeratorId, denominatorId) {
  return {
    method: 'summary',
    summary: 'recalculate',
    operation: 'divide',
    numerator: { kind: 'templateField', id: numeratorId },
    denominator: { kind: 'templateField', id: denominatorId },
  };
}

function sectorLabel() {
  return { method: 'label', label: 'sectorTotal' };
}

function grandLabel() {
  return { method: 'label', label: 'grandTotal' };
}

function contexts(property, unit, sectorTotal, grandTotal) {
  return {
    property: property,
    unit: unit,
    sectorTotal: sectorTotal,
    grandTotal: grandTotal,
  };
}

export function waultExpiryRule(scopeNote) {
  return {
    method: 'derived',
    calculation: 'weightedAverage',
    measure: 'unexpiredLeaseTerm',
    start: { kind: 'propertyContext', field: 'valuation' },
    end: { kind: 'templateField', id: 'expiry_date', dataset: 'unit' },
    weight: { kind: 'templateField', id: 'contracted_rent_pa', dataset: 'unit' },
    scope: 'matchedUnits',
    scopeNote: scopeNote || 'property',
    eligibility: { validEndDate: true, rentGreaterThan: 0 },
    yearFraction: 'actualActual',
  };
}

export function waultBreakRule(scopeNote) {
  return {
    method: 'derived',
    calculation: 'weightedAverage',
    measure: 'unexpiredLeaseTermToBreak',
    start: { kind: 'propertyContext', field: 'valuation' },
    preferredEnd: { kind: 'templateField', id: 'break_date', dataset: 'unit' },
    fallbackEnd: { kind: 'templateField', id: 'expiry_date', dataset: 'unit' },
    breakRule: 'futureAndOnOrBeforeExpiry',
    weight: { kind: 'templateField', id: 'contracted_rent_pa', dataset: 'unit' },
    scope: 'matchedUnits',
    scopeNote: scopeNote || 'property',
    eligibility: { validEndDate: true, rentGreaterThan: 0 },
    yearFraction: 'actualActual',
    formats: { unit: 'dd/mm/yyyy', summary: '0.00' },
  };
}

var blankRow = contexts(blank(), blank(), blank(), blank());

export const DEFAULT_FIELD_LOGIC = {
  property_reference: contexts(source(PROPERTY_SOURCE_FIELDS.externalId), blank(), sectorLabel(), grandLabel()),
  valuation: contexts(source(PROPERTY_SOURCE_FIELDS.traditionalValuationDate), blank(), blank(), blank()),
  town: contexts(source(PROPERTY_SOURCE_FIELDS.city), blank(), blank(), blank()),
  address: contexts(source(PROPERTY_SOURCE_FIELDS.propertyName), blank(), blank(), blank()),
  tenure_uk: contexts(blank(), source(UNIT_SOURCE_FIELDS.tenure), blank(), blank()),
  area_sqft: contexts(
    sumMatchedUnits('property'),
    aggregated(UNIT_SOURCE_FIELDS.occupiedArea),
    sumMatchedUnits('sector'),
    sumMatchedUnits('portfolio')
  ),
  unit: contexts(blank(), source(UNIT_SOURCE_FIELDS.suite), blank(), blank()),
  tenant_name: contexts(blank(), source(UNIT_SOURCE_FIELDS.tenantName), blank(), blank()),
  start_date: contexts(blank(), source(UNIT_SOURCE_FIELDS.leaseStart), blank(), blank()),
  review_date: contexts(blank(), source(UNIT_SOURCE_FIELDS.nextReview), blank(), blank()),
  expiry_date: contexts(blank(), source(UNIT_SOURCE_FIELDS.leaseExpiry), blank(), blank()),
  break_date: contexts(
    waultBreakRule('property'),
    source(UNIT_SOURCE_FIELDS.effectiveBreak),
    waultBreakRule('sector'),
    waultBreakRule('portfolio')
  ),
  mwault_year: contexts(
    waultExpiryRule('property'),
    blank(),
    waultExpiryRule('sector'),
    waultExpiryRule('portfolio')
  ),
  contracted_rent_pa: contexts(
    source(PROPERTY_SOURCE_FIELDS.contractedRent),
    source(UNIT_SOURCE_FIELDS.contractedRent),
    sum(),
    sum()
  ),
  contracted_rent_psf: contexts(blank(), blank(), recalculate('contracted_rent_pa', 'area_sqft'), recalculate('contracted_rent_pa', 'area_sqft')),
  deemed_rent_psf: contexts(
    source(PROPERTY_SOURCE_FIELDS.netRent),
    source(UNIT_SOURCE_FIELDS.currentRent),
    sum(),
    sum()
  ),
  market_rent_pa: contexts(
    source(PROPERTY_SOURCE_FIELDS.rentalValue),
    source(UNIT_SOURCE_FIELDS.rentalValue),
    sum(),
    sum()
  ),
  market_rent_psf: contexts(blank(), blank(), recalculate('market_rent_pa', 'area_sqft'), recalculate('market_rent_pa', 'area_sqft')),
  current_void: contexts(blank(), source(UNIT_SOURCE_FIELDS.voidMonths), blank(), blank()),
  rent_free_month: contexts(blank(), source(UNIT_SOURCE_FIELDS.freeRentMonths), blank(), blank()),
  expiry_void_month: blankRow,
  rent_free_expiry_month: blankRow,
  cap_rate: contexts(blank(), source(UNIT_SOURCE_FIELDS.hardcoreYield), blank(), blank()),
  niy: contexts(source(PROPERTY_SOURCE_FIELDS.initialYield), blank(), blank(), blank()),
  ney: contexts(source(PROPERTY_SOURCE_FIELDS.equivalentYield), blank(), blank(), blank()),
  ry: contexts(source(PROPERTY_SOURCE_FIELDS.reversionaryYield), blank(), blank(), blank()),
  cv_psf: contexts(
    divideFields('fair_value', 'area_sqft', 'property'),
    blank(),
    recalculate('fair_value', 'area_sqft'),
    recalculate('fair_value', 'area_sqft')
  ),
  cap_ex: contexts(blank(), blank(), sum(), sum()),
  fair_value: contexts(source(PROPERTY_SOURCE_FIELDS.sayValue), blank(), sum(), sum()),
  fair_value_march_2026: blankRow,
  pct_change_26_april_24: blankRow,
  pct_change_x_ytd: blankRow,
  comments: blankRow,
};

export const DEFAULT_LOGIC_CONTEXTS = ['property', 'unit', 'sectorTotal', 'grandTotal'];
export const PROPERTY_ONLY_LOGIC_CONTEXTS = ['property'];
