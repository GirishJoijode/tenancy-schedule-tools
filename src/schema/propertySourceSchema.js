import { DEFAULT_SCHEDULE_COLUMNS } from '../templates/defaultSchedule.js';
import { PROPERTY_ONLY_COLUMNS } from '../templates/propertyOnly.js';
import { PROPERTY_SOURCE_FIELDS } from './sourceFields.js';

var BASELINE_PROPERTY_FIELDS = [
  PROPERTY_SOURCE_FIELDS.externalId,
  PROPERTY_SOURCE_FIELDS.propertyName,
  PROPERTY_SOURCE_FIELDS.sector,
  PROPERTY_SOURCE_FIELDS.traditionalValuationDate,
  PROPERTY_SOURCE_FIELDS.city,
];

function collect(refs) {
  var seenFields = new Set();
  var seenGroups = new Set();
  var fields = [];
  var periodGroups = [];
  for (var i = 0; i < refs.length; i += 1) {
    var ref = refs[i];
    if (!ref) continue;
    if (ref.type === 'periodGroup') {
      var groupKey = String(ref.section || '').toLowerCase();
      if (!groupKey || seenGroups.has(groupKey)) continue;
      seenGroups.add(groupKey);
      periodGroups.push(ref);
      continue;
    }
    if (!ref.field) continue;
    var key = [
      String(ref.section || '').toLowerCase(),
      String(ref.field).toLowerCase(),
      String(ref.occurrence || 1),
    ].join('|');
    if (seenFields.has(key)) continue;
    seenFields.add(key);
    fields.push({
      type: 'field',
      section: ref.section,
      field: ref.field,
      occurrence: ref.occurrence || 1,
    });
  }
  return { fields: fields, periodGroups: periodGroups };
}

function propertyRefs() {
  var refs = BASELINE_PROPERTY_FIELDS.slice();
  for (var i = 0; i < DEFAULT_SCHEDULE_COLUMNS.length; i += 1) {
    refs.push(DEFAULT_SCHEDULE_COLUMNS[i].property);
  }
  for (var j = 0; j < PROPERTY_ONLY_COLUMNS.length; j += 1) {
    refs.push(PROPERTY_ONLY_COLUMNS[j].property);
  }
  return refs;
}

export function getRequiredPropertySources() {
  return collect(propertyRefs());
}

export function getRequiredPropertyFields() {
  return getRequiredPropertySources().fields;
}
