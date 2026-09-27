/**
 * Central registry of Standardiser schedule types.
 * Add a new object here (plus its view/implementation) to extend the selector.
 * `stateKey` is the existing session key used by APP_STATE.scheduleType.
 */
export const SCHEDULE_TYPES = [
  {
    id: 'tenancy_schedule',
    stateKey: 'tenancy',
    viewId: 'schedule-tenancy',
    label: 'Default',
    description: 'Convert Property and Unit level exports into the standard tenancy schedule.',
    summary: 'Property and unit-level schedule',
    category: 'Standard',
    inputs: ['property', 'unit'],
  },
  {
    id: 'property_summary',
    stateKey: 'propertySummary',
    viewId: 'schedule-property-summary',
    label: 'Property Only',
    description: 'Generate a property-level valuation summary from a Property Level export.',
    summary: 'Property-only valuation summary',
    category: 'Standard',
    inputs: ['property'],
  },
];

export const DEFAULT_SCHEDULE_TYPE_ID = 'tenancy_schedule';

export function getScheduleTypeById(id) {
  for (var i = 0; i < SCHEDULE_TYPES.length; i += 1) {
    if (SCHEDULE_TYPES[i].id === id) return SCHEDULE_TYPES[i];
  }
  return null;
}

export function getScheduleTypeByStateKey(stateKey) {
  for (var i = 0; i < SCHEDULE_TYPES.length; i += 1) {
    if (SCHEDULE_TYPES[i].stateKey === stateKey) return SCHEDULE_TYPES[i];
  }
  return null;
}

export function getDefaultScheduleType() {
  return getScheduleTypeById(DEFAULT_SCHEDULE_TYPE_ID) || SCHEDULE_TYPES[0];
}
