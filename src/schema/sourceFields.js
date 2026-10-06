function sourceField(section, field) {
  return { type: 'field', section: section, field: field, occurrence: 1 };
}

function periodGroup(dataset, section, selection, aggregation) {
  return {
    type: 'periodGroup',
    dataset: dataset,
    section: section,
    selection: selection,
    aggregation: aggregation || null,
  };
}

/** Fixed Property.xlsx fields. Templates and validation reference these objects. */
export const PROPERTY_SOURCE_FIELDS = {
  externalId: sourceField('General', 'External ID'),
  propertyName: sourceField('General', 'Property Name'),
  sector: sourceField('General', 'Sector'),
  traditionalValuationDate: sourceField('General', 'Traditional Valuation Date'),
  city: sourceField('General', 'City'),
  contractedRent: sourceField('General', 'Contracted Rent'),
  netRent: sourceField('General', 'Net Rent'),
  rentalValue: sourceField('General', 'Rental Value (ERV)'),
  initialYield: sourceField('General', 'Initial Yield (at Valuation Date)'),
  equivalentYield: sourceField('General', 'Equivalent Yield (at Valuation Date)'),
  reversionaryYield: sourceField('General', 'Reversionary Yield (at Valuation Date)'),
  sayValue: sourceField('General', 'Say Value'),
  capitalCosts: sourceField('General', 'Capital Costs'),
  occupiedArea: periodGroup('property', 'Occupied Area', 'earliestPeriod'),
};

/** Fixed TS.xlsx fields. Templates, validation, and TS Compare reference these objects. */
export const UNIT_SOURCE_FIELDS = {
  leaseId: sourceField('General', 'Lease ID'),
  propertyName: sourceField('General', 'Property Name'),
  tenantName: sourceField('General', 'Tenant Name'),
  suite: sourceField('General', 'Suite'),
  tenure: sourceField('General', 'Tenure'),
  leaseStart: sourceField('General', 'Lease Start'),
  leaseExpiry: sourceField('General', 'Lease Expiry'),
  effectiveBreak: sourceField('General', 'Effective Break'),
  contractedRent: sourceField('General', 'Contracted Rent'),
  nextReview: sourceField('Variable', 'Next Review'),
  currentRent: sourceField('Variable', 'Current Rent'),
  rentalValue: sourceField('Variable', 'Rental Value (ERV)'),
  voidMonths: sourceField('Variable', 'Void'),
  freeRentMonths: sourceField('Variable', 'Free Rent (Months)'),
  occupiedArea: Object.assign(periodGroup('unit', 'Occupied Area (SF)', 'allPeriods', 'max'), {
    scope: 'currentUnitRow',
  }),
  capitalCost: Object.assign(periodGroup('unit', 'Capital Cost', 'allPeriods', 'sum'), {
    scope: 'currentUnitRow',
  }),
  hardcoreYield: sourceField('Traditional Valuation', 'Hardcore Yield'),
};

export const PROPERTY_UNIT_JOIN = PROPERTY_SOURCE_FIELDS.propertyName;
export const PROPERTY_SECTOR = PROPERTY_SOURCE_FIELDS.sector;

export const COMPARE_IDENTITY = {
  property: UNIT_SOURCE_FIELDS.propertyName,
  unit: UNIT_SOURCE_FIELDS.suite,
};

export const HEADER_MARKERS = {
  property: PROPERTY_SOURCE_FIELDS.externalId.field,
  unit: UNIT_SOURCE_FIELDS.leaseId.field,
};
