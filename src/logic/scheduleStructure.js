/**
 * Default schedule row structure. Generation and the draft export both read this.
 * The separator after each sector total is the gap before the next sector and,
 * after the last sector, the gap before Grand Total.
 */
export const DEFAULT_SCHEDULE_STRUCTURE = {
  headerSeparator: { kind: 'separator', style: 'headerSeparator' },
  sectorHeader: { kind: 'sector', role: 'presentation', metrics: false },
  perSector: ['sectorHeader', 'properties', 'sectorTotal', 'separator'],
  afterSectors: ['grandTotal'],
  separatorStyle: 'headerSeparator',
};
