import { SCHEDULE_TYPES } from '../config/scheduleTypes.js';
import { DEFAULT_SCHEDULE_COLUMNS } from '../templates/defaultSchedule.js';
import { PROPERTY_ONLY_COLUMNS } from '../templates/propertyOnly.js';
import { DEFAULT_SCHEDULE_STRUCTURE } from './scheduleStructure.js';
import { DEFAULT_LOGIC_CONTEXTS, PROPERTY_ONLY_LOGIC_CONTEXTS } from './fieldRules.js';

var DETAILS = {
  tenancy_schedule: {
    columns: DEFAULT_SCHEDULE_COLUMNS,
    contexts: DEFAULT_LOGIC_CONTEXTS,
    structure: DEFAULT_SCHEDULE_STRUCTURE,
  },
  property_summary: {
    columns: PROPERTY_ONLY_COLUMNS,
    contexts: PROPERTY_ONLY_LOGIC_CONTEXTS,
    structure: null,
  },
};

export function listLogicTemplates() {
  return SCHEDULE_TYPES.map(function (type) {
    var detail = DETAILS[type.id] || { columns: [], contexts: ['property'], structure: null };
    return {
      id: type.id,
      label: type.label,
      columns: detail.columns,
      contexts: detail.contexts,
      structure: detail.structure,
    };
  });
}

export function getLogicTemplate(id) {
  var templates = listLogicTemplates();
  for (var i = 0; i < templates.length; i += 1) {
    if (templates[i].id === id) return templates[i];
  }
  return templates[0];
}
