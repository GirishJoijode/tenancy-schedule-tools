export function normaliseHeader(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function normaliseMatchingKey(value) {
  return normaliseHeader(value);
}

export function normaliseJoinValue(value) {
  return normaliseMatchingKey(value);
}

export function buildSourceColumnKey(fileType, section, field, occurrence) {
  return [
    normaliseMatchingKey(fileType),
    normaliseMatchingKey(section) || '(no section)',
    normaliseMatchingKey(field) || '(blank)',
    String(occurrence || 1),
  ].join('|');
}

export function sourceRef(section, field, occurrence) {
  return {
    section: section,
    field: field,
    occurrence: occurrence || 1,
  };
}

export function sourceLabel(ref) {
  if (!ref) return '';
  return ref.section + ' > ' + ref.field;
}
