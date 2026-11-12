/**
 * Default schedule row structure. Generation and the draft export both read this.
 * Group headers and totals are generic. Portfolio Breakdown decides the groups.
 * The separator after each group total is the gap before the next group and,
 * after the last group, the gap before Grand Total.
 */
export const DEFAULT_SCHEDULE_STRUCTURE = {
  headerSeparator: { kind: 'separator', style: 'headerSeparator' },
  groupHeader: { kind: 'group', role: 'presentation', metrics: false },
  perGroup: ['groupHeader', 'properties', 'groupTotal', 'separator'],
  afterGroups: ['grandTotal'],
  ungrouped: ['properties', 'separator', 'grandTotal'],
  separatorStyle: 'headerSeparator',
};
