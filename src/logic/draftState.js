import { getLogicTemplate } from './registry.js';

var drafts = {};

export function cloneDefinition(value) {
  return JSON.parse(JSON.stringify(value));
}

export function getDraftTemplate(templateId) {
  if (!drafts[templateId]) {
    var production = getLogicTemplate(templateId);
    drafts[templateId] = {
      id: production.id,
      label: production.label,
      contexts: production.contexts.slice(),
      structure: cloneDefinition(production.structure),
      columns: cloneDefinition(production.columns),
    };
  }
  return drafts[templateId];
}

export function resetDraft(templateId) {
  delete drafts[templateId];
  return getDraftTemplate(templateId);
}

export function updateDraftRule(templateId, columnId, context, rule) {
  var draft = getDraftTemplate(templateId);
  for (var i = 0; i < draft.columns.length; i += 1) {
    if (draft.columns[i].id !== columnId) continue;
    if (!draft.columns[i].logic) draft.columns[i].logic = {};
    draft.columns[i].logic[context] = rule;
    return draft;
  }
  return draft;
}
