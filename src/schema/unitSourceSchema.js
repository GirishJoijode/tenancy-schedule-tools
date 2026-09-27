import { DEFAULT_SCHEDULE_COLUMNS } from '../templates/defaultSchedule.js';
import { UNIT_SOURCE_FIELDS } from './sourceFields.js';

var BASELINE_UNIT_FIELDS = [
  UNIT_SOURCE_FIELDS.propertyName,
  UNIT_SOURCE_FIELDS.tenantName,
  UNIT_SOURCE_FIELDS.suite,
  UNIT_SOURCE_FIELDS.leaseStart,
  UNIT_SOURCE_FIELDS.leaseExpiry,
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

export function getRequiredUnitSources() {
  var refs = BASELINE_UNIT_FIELDS.slice();
  for (var i = 0; i < DEFAULT_SCHEDULE_COLUMNS.length; i += 1) {
    refs.push(DEFAULT_SCHEDULE_COLUMNS[i].unit);
  }
  return collect(refs);
}

export function getRequiredUnitFields() {
  return getRequiredUnitSources().fields;
}
