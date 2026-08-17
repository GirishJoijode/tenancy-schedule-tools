export const APP_VERSION = '1.0.0';
export const CONFIG_VERSION = 1;

export const HEADER_MARKERS = {
  property: 'External ID',
  unit: 'Lease ID',
};

export const HEADER_SEARCH_ROWS = 25;
export const GROUP_HEADER_OFFSET = 1;
export const ACCEPTED_EXTENSION = '.xlsx';

export const STANDARD_TEMPLATE_COLUMNS = [
  { id: 'property_reference', label: 'Property Reference' },
  { id: 'town', label: 'Town' },
  { id: 'address', label: 'Address' },
  { id: 'tenure_uk', label: 'Tenure (UK)' },
  { id: 'area_sqft', label: 'Area (sq ft)' },
  { id: 'unit', label: 'Unit' },
  { id: 'tenant_name', label: 'Tenant Name' },
  { id: 'start_date', label: 'Start Date' },
  { id: 'review_date', label: 'Review Date' },
  { id: 'expiry_date', label: 'Expiry Date' },
  { id: 'break_date', label: 'Break Date' },
  { id: 'mwault_year', label: 'MWAULT (Year)' },
  { id: 'contracted_rent_pa', label: 'Contracted Rent (pa)' },
  { id: 'contracted_rent_psf', label: 'Contracted Rent (psf)' },
  { id: 'deemed_rent_psf', label: 'Deemed Rent (psf)' },
  { id: 'market_rent_pa', label: 'Market Rent (pa)' },
  { id: 'market_rent_psf', label: 'Market Rent (psf)' },
  { id: 'current_void', label: 'Current Void' },
  { id: 'rent_free_month', label: 'Rent Free (Month)' },
  { id: 'expiry_void_month', label: 'Expiry Void (Month)' },
  { id: 'rent_free_expiry_month', label: 'Rent Free Expiry Month' },
  { id: 'cap_rate', label: 'Cap Rate' },
  { id: 'niy', label: 'NIY' },
  { id: 'ney', label: 'NEY' },
  { id: 'ry', label: 'RY' },
  { id: 'cv_psf', label: 'CV (psf)' },
  { id: 'cap_ex', label: 'Cap Ex' },
  { id: 'fair_value_april_2024', label: 'Fair Value April 2024' },
  { id: 'fair_value_march_2026', label: 'Fair Value March 2026' },
  { id: 'pct_change_26_april_24', label: '% Change 26 - April 24' },
  { id: 'pct_change_x_ytd', label: '% Change x YTD' },
  { id: 'comments', label: 'Comments' },
];

export const DEFAULT_MATCH_HINTS = {
  property: { section: 'General', field: 'Property Name' },
  unit: { section: 'General', field: 'Property Name' },
};

export const COMPARE_IDENTITY = {
  property: { section: 'General', field: 'Property Name' },
  unit: { section: 'General', field: 'Suite' },
};

export const COMPARE_NUMERIC_EPS = 1e-9;

export const DATE_COMPARE_IDS = {
  start_date: true,
  review_date: true,
  expiry_date: true,
  break_date: true,
};

export const OUTPUT_FILENAME = 'Standardised_Tenancy_Schedule.xlsx';
export const COMPARE_OUTPUT_FILENAME = 'Tenancy_Schedule_Comparison.xlsx';
export const CONFIG_OUTPUT_FILENAME = 'Tenancy_Schedule_Config.json';

export const STORAGE_KEY = 'tenancyScheduleTools.config.v1';
