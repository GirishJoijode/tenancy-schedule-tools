/**
 * Top-level application tools. The Standardiser / TS Compare switcher
 * renders from this list. Add an entry here when a new primary tool exists.
 */
export const APP_TOOLS = [
  {
    id: 'standardiser',
    mode: 'standardiser',
    buttonId: 'mode-standardiser-btn',
    viewId: 'view-standardiser',
    label: 'Standardiser',
    description: 'Create standardised valuation schedules',
  },
  {
    id: 'compare',
    mode: 'compare',
    buttonId: 'mode-compare-btn',
    viewId: 'view-compare',
    label: 'TS Compare',
    description: 'Compare current and previous tenancy schedules',
  },
];

export const DEFAULT_APP_TOOL_MODE = 'standardiser';
