import { APP_STATE } from '../state.js';
import { escapeHtml } from './common.js';
import {
  BREAKDOWN_OPTIONS,
  assessBreakdown,
  buildPortfolioGroups,
  createGroupId,
  getBreakdownMode,
  getCustomConfig,
  listCurrentProperties,
  reconcileCustomAssignments,
  setBreakdownMode,
  setCustomConfig,
  validateCustomGroups,
} from '../logic/portfolioBreakdown.js';

var menuOpen = false;
var highlightIndex = -1;
var priorMode = 'sector';
var draft = { groups: [], assignments: {} };
var searchQuery = '';
var selectedKey = '';
var selectedGroupId = '';
var dragKind = '';
var dragId = '';
var nameIssues = {};

function optionById(id) {
  for (var i = 0; i < BREAKDOWN_OPTIONS.length; i += 1) {
    if (BREAKDOWN_OPTIONS[i].id === id) return BREAKDOWN_OPTIONS[i];
  }
  return BREAKDOWN_OPTIONS[1];
}

function displayLabel(id) {
  var option = optionById(id);
  return option.selectedLabel || option.label;
}

function enabled() {
  return APP_STATE.scheduleType === 'tenancy';
}

function properties() {
  return listCurrentProperties(APP_STATE.property);
}

function assignedCount() {
  var live = {};
  properties().forEach(function (property) { live[property.externalKey] = true; });
  var count = 0;
  Object.keys(draft.assignments).forEach(function (key) {
    if (live[key] && draft.assignments[key]) count += 1;
  });
  return count;
}

function refreshNote() {
  var note = document.getElementById('portfolio-breakdown-note');
  var edit = document.getElementById('portfolio-breakdown-edit');
  if (!note || !edit) return;
  edit.hidden = !enabled() || getBreakdownMode() !== 'custom';
  if (!enabled() || !APP_STATE.property) {
    note.hidden = true;
    note.textContent = '';
    return;
  }
  if (getBreakdownMode() === 'custom') {
    var config = getCustomConfig();
    var customGroups = buildPortfolioGroups(properties(), {
      mode: 'custom',
      customConfig: config,
    });
    note.hidden = false;
    note.className = 'breakdown-note';
    if (!customGroups.effective && properties().length) {
      note.textContent = customGroups.populated
        ? 'Only one custom group is present. The schedule will be generated without group sections.'
        : 'Custom groups are not populated. The schedule will be generated without group sections.';
      return;
    }
    note.textContent = config.groups.length + (config.groups.length === 1 ? ' group' : ' groups') + ' \u2022 ' + properties().length + (properties().length === 1 ? ' property' : ' properties');
    return;
  }
  var assessment = assessBreakdown(APP_STATE.property, getBreakdownMode());
  if (!assessment.message) {
    note.hidden = true;
    note.textContent = '';
    return;
  }
  note.hidden = false;
  note.className = 'breakdown-note';
  note.textContent = assessment.message;
}

export function refreshPortfolioBreakdownChrome() {
  var field = document.getElementById('portfolio-breakdown-field');
  var label = document.getElementById('portfolio-breakdown-trigger-label');
  var trigger = document.getElementById('portfolio-breakdown-trigger');
  if (field) field.classList.toggle('is-disabled', !enabled());
  if (trigger) {
    trigger.disabled = !enabled();
    trigger.setAttribute('aria-disabled', enabled() ? 'false' : 'true');
  }
  if (label) label.textContent = displayLabel(getBreakdownMode());
  renderMenu();
  refreshNote();
}

function renderMenu() {
  var menu = document.getElementById('portfolio-breakdown-menu');
  if (!menu) return;
  var selected = getBreakdownMode();
  menu.innerHTML = BREAKDOWN_OPTIONS.map(function (option, index) {
    var isSelected = option.id === selected;
    return (
      '<li role="option" class="schedule-type-option' +
      (isSelected ? ' is-selected' : '') +
      (index === highlightIndex ? ' is-highlighted' : '') +
      '" data-breakdown-id="' + escapeHtml(option.id) +
      '" aria-selected="' + (isSelected ? 'true' : 'false') + '">' +
      '<span class="schedule-type-option-mark" aria-hidden="true">' + (isSelected ? '\u2713' : '') + '</span>' +
      '<span class="schedule-type-option-label">' + escapeHtml(option.label) + '</span>' +
      '</li>'
    );
  }).join('');
}

function setMenuOpen(next) {
  if (!enabled()) return;
  menuOpen = next;
  highlightIndex = next ? BREAKDOWN_OPTIONS.findIndex(function (option) { return option.id === getBreakdownMode(); }) : -1;
  var root = document.getElementById('portfolio-breakdown-select');
  var menu = document.getElementById('portfolio-breakdown-menu');
  var trigger = document.getElementById('portfolio-breakdown-trigger');
  if (root) root.classList.toggle('is-open', menuOpen);
  if (menu) menu.classList.toggle('hidden', !menuOpen);
  if (trigger) trigger.setAttribute('aria-expanded', menuOpen ? 'true' : 'false');
  renderMenu();
}

function chooseMode(id) {
  setMenuOpen(false);
  if (id === 'custom') {
    openCustomModal();
    return;
  }
  setBreakdownMode(id);
  refreshPortfolioBreakdownChrome();
}

function cloneDraftFromSession() {
  draft = getCustomConfig();
  searchQuery = '';
  selectedKey = '';
  selectedGroupId = draft.groups[0] ? draft.groups[0].id : '';
  nameIssues = {};
}

function visibleProperties(list) {
  var query = searchQuery.trim().toLowerCase();
  if (!query) return list;
  return list.filter(function (property) {
    return (property.externalId + ' ' + property.name).toLowerCase().indexOf(query) !== -1;
  });
}

function propertiesFor(groupId) {
  return properties().filter(function (property) {
    return draft.assignments[property.externalKey] === groupId;
  });
}

function unassignedProperties() {
  var groupIds = {};
  draft.groups.forEach(function (group) { groupIds[group.id] = true; });
  return properties().filter(function (property) {
    var assigned = draft.assignments[property.externalKey];
    return !assigned || !groupIds[assigned];
  });
}

function propertyRow(property, listName) {
  var selected = property.externalKey === selectedKey;
  return (
    '<li class="group-property' + (selected ? ' is-selected' : '') + '" draggable="true" data-drag-kind="property" data-key="' +
    escapeHtml(property.externalKey) + '" data-list="' + escapeHtml(listName) + '" tabindex="0">' +
    '<span class="group-handle" aria-hidden="true">\u283F</span>' +
    '<span class="group-property-copy"><span class="group-property-id">' + escapeHtml(property.externalId || '\u2014') +
    '</span><span class="group-property-name">' + escapeHtml(property.name || 'Property') + '</span></span>' +
    '</li>'
  );
}

function renderGroups() {
  if (!APP_STATE.property) {
    return '<p class="group-empty">Upload a Property Level file to configure custom portfolio groups.</p>';
  }
  if (!draft.groups.length) {
    return '<p class="group-empty">Add a group, then assign properties to it.</p>';
  }
  return draft.groups.map(function (group, index) {
    var members = visibleProperties(propertiesFor(group.id));
    var issue = nameIssues[group.id] || '';
    return (
      '<section class="group-card' + (selectedGroupId === group.id ? ' is-selected' : '') + '" data-drop-group="' + escapeHtml(group.id) + '">' +
      '<header class="group-card-head">' +
      '<button type="button" class="group-handle" draggable="true" data-drag-kind="group" data-group-id="' + escapeHtml(group.id) + '" aria-label="Reorder ' + escapeHtml(group.name || 'group') + '">\u283F</button>' +
      '<input class="group-name" data-group-name="' + escapeHtml(group.id) + '" value="' + escapeHtml(group.name || '') + '" aria-label="Group name">' +
      '<button type="button" data-group-action="up" data-group-id="' + escapeHtml(group.id) + '"' + (index === 0 ? ' disabled' : '') + ' aria-label="Move group up">Up</button>' +
      '<button type="button" data-group-action="down" data-group-id="' + escapeHtml(group.id) + '"' + (index === draft.groups.length - 1 ? ' disabled' : '') + ' aria-label="Move group down">Down</button>' +
      '<button type="button" data-group-action="delete" data-group-id="' + escapeHtml(group.id) + '" aria-label="Delete group">Delete</button>' +
      '</header>' +
      (issue ? '<p class="group-issue">' + escapeHtml(issue) + '</p>' : '') +
      '<ul class="group-list" data-drop-group="' + escapeHtml(group.id) + '">' +
      (members.length ? members.map(function (property) { return propertyRow(property, group.id); }).join('') : '<li class="group-placeholder">Drop properties here</li>') +
      '</ul></section>'
    );
  }).join('');
}

function renderModalBody() {
  var unassigned = visibleProperties(unassignedProperties());
  var moveOptions = '<option value="">Move to\u2026</option><option value="unassigned">Unassigned</option>' +
    draft.groups.map(function (group) {
      return '<option value="' + escapeHtml(group.id) + '">' + escapeHtml(group.name || 'New Group') + '</option>';
    }).join('');
  return (
    '<div class="group-toolbar">' +
    '<label class="group-search">Search properties...<input id="group-property-search" type="search" placeholder="Search properties..." aria-label="Search properties" value="' + escapeHtml(searchQuery) + '"></label>' +
    '<button type="button" class="group-add" data-group-action="add">+ Add Group</button>' +
    '<label class="group-move">Move selected<select id="group-move-to"' + (selectedKey ? '' : ' disabled') + '>' + moveOptions + '</select></label>' +
    '</div>' +
    '<div class="group-panels">' +
    '<section class="group-panel"><header><h3>Custom Groups</h3></header><div class="group-scroll" id="group-scroll">' + renderGroups() + '</div></section>' +
    '<section class="group-panel" data-drop-unassigned="true"><header><h3>Unassigned Properties</h3><p>Unassigned properties will be placed in Unclassified.</p></header>' +
    '<ul class="group-list" data-drop-unassigned="true">' +
    (APP_STATE.property
      ? (unassigned.length ? unassigned.map(function (property) { return propertyRow(property, 'unassigned'); }).join('') : '<li class="group-placeholder">All visible properties are assigned.</li>')
      : '<li class="group-placeholder">No properties loaded.</li>') +
    '</ul></section></div>'
  );
}

function renderModalFooter() {
  var props = properties();
  var unassigned = unassignedProperties().length;
  var summary = draft.groups.length + ' groups \u2022 ' + assignedCount() + ' assigned \u2022 ' + unassigned + ' unassigned';
  if (!APP_STATE.property) summary = 'No Property Level file loaded';
  var invalid = Object.keys(validateCustomGroups(draft.groups)).length > 0;
  return '<p class="report-status">' + escapeHtml(summary) + '</p><div class="report-footer-actions">' +
    '<button type="button" class="report-reset" data-group-action="reset"' + (draft.groups.length || Object.keys(draft.assignments).length ? '' : ' disabled') + '>Reset</button>' +
    '<button type="button" class="report-cancel" data-group-action="cancel">Cancel</button>' +
    '<button type="button" class="report-apply" data-group-action="apply"' + (invalid ? ' disabled' : '') + '>Apply Breakdown</button></div>';
}

function paintModal(focusSelector) {
  var root = document.getElementById('portfolio-breakdown-root');
  var footer = document.getElementById('portfolio-breakdown-footer');
  var search = document.getElementById('group-property-search');
  var searchFocused = !!(search && document.activeElement === search);
  var caret = searchFocused ? search.selectionStart : null;
  var nameInput = document.activeElement && document.activeElement.getAttribute && document.activeElement.getAttribute('data-group-name');
  var nameCaret = nameInput ? document.activeElement.selectionStart : null;
  if (root) root.innerHTML = renderModalBody();
  if (footer) footer.innerHTML = renderModalFooter();
  if (searchFocused) {
    var nextSearch = document.getElementById('group-property-search');
    if (nextSearch) {
      nextSearch.focus({ preventScroll: true });
      if (caret != null) nextSearch.setSelectionRange(caret, caret);
    }
    return;
  }
  if (nameInput) {
    var nextName = document.querySelector('[data-group-name="' + nameInput + '"]');
    if (nextName) {
      nextName.focus({ preventScroll: true });
      if (nameCaret != null) nextName.setSelectionRange(nameCaret, nameCaret);
    }
    return;
  }
  if (!focusSelector) return;
  var modal = document.getElementById('portfolio-breakdown-modal');
  var target = modal ? modal.querySelector(focusSelector) : null;
  if (target && !target.disabled) target.focus({ preventScroll: true });
}

function openCustomModal() {
  priorMode = getBreakdownMode();
  cloneDraftFromSession();
  var modal = document.getElementById('portfolio-breakdown-modal');
  if (!modal) return;
  modal.hidden = false;
  modal.classList.remove('hidden');
  document.body.classList.add('report-open');
  var app = document.getElementById('app');
  if (app) app.setAttribute('inert', '');
  paintModal('#group-property-search');
}

function closeCustomModal(applied) {
  var modal = document.getElementById('portfolio-breakdown-modal');
  if (modal) {
    modal.hidden = true;
    modal.classList.add('hidden');
  }
  document.body.classList.remove('report-open');
  var app = document.getElementById('app');
  if (app) app.removeAttribute('inert');
  if (!applied && priorMode !== 'custom') setBreakdownMode(priorMode);
  refreshPortfolioBreakdownChrome();
  var trigger = document.getElementById('portfolio-breakdown-trigger');
  if (trigger) trigger.focus();
}

function moveGroup(id, direction) {
  var index = draft.groups.findIndex(function (group) { return group.id === id; });
  var next = index + direction;
  if (index < 0 || next < 0 || next >= draft.groups.length) return;
  var moved = draft.groups.splice(index, 1)[0];
  draft.groups.splice(next, 0, moved);
  selectedGroupId = id;
}

function assignProperty(key, groupId) {
  if (!key) return;
  if (!groupId || groupId === 'unassigned') delete draft.assignments[key];
  else draft.assignments[key] = groupId;
  selectedKey = key;
}

function addGroup() {
  var name = 'New Group';
  var taken = {};
  draft.groups.forEach(function (group) { taken[String(group.name || '').trim().toLowerCase()] = true; });
  if (taken['new group']) {
    var n = 2;
    while (taken['new group ' + n]) n += 1;
    name = 'New Group ' + n;
  }
  var id = createGroupId();
  draft.groups.push({ id: id, name: name });
  selectedGroupId = id;
  nameIssues = validateCustomGroups(draft.groups);
  paintModal('[data-group-name="' + id + '"]');
  var input = document.querySelector('[data-group-name="' + id + '"]');
  if (input) input.select();
}

function applyDraft() {
  nameIssues = validateCustomGroups(draft.groups);
  if (Object.keys(nameIssues).length) {
    paintModal();
    return;
  }
  setCustomConfig(draft);
  setBreakdownMode('custom');
  closeCustomModal(true);
}

function clearDrag() {
  dragKind = '';
  dragId = '';
  var modal = document.getElementById('portfolio-breakdown-modal');
  if (!modal) return;
  modal.querySelectorAll('.is-dragging, .is-drop-active').forEach(function (node) {
    node.classList.remove('is-dragging', 'is-drop-active');
  });
}

function autoscroll(event) {
  var scroller = document.getElementById('group-scroll');
  if (!scroller) return;
  var rect = scroller.getBoundingClientRect();
  if (event.clientY < rect.top + 28) scroller.scrollTop -= 12;
  if (event.clientY > rect.bottom - 28) scroller.scrollTop += 12;
}

export function setupPortfolioBreakdown() {
  refreshPortfolioBreakdownChrome();
  var trigger = document.getElementById('portfolio-breakdown-trigger');
  var menu = document.getElementById('portfolio-breakdown-menu');
  var edit = document.getElementById('portfolio-breakdown-edit');
  var modal = document.getElementById('portfolio-breakdown-modal');
  if (trigger) {
    trigger.addEventListener('click', function () {
      setMenuOpen(!menuOpen);
    });
    trigger.addEventListener('keydown', function (event) {
      if (!enabled()) return;
      if (event.key === 'ArrowDown' || event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        if (!menuOpen) setMenuOpen(true);
        else if (event.key === 'ArrowDown') {
          highlightIndex = Math.min(BREAKDOWN_OPTIONS.length - 1, highlightIndex + 1);
          renderMenu();
        } else chooseMode(BREAKDOWN_OPTIONS[highlightIndex].id);
      } else if (event.key === 'ArrowUp' && menuOpen) {
        event.preventDefault();
        highlightIndex = Math.max(0, highlightIndex - 1);
        renderMenu();
      } else if (event.key === 'Escape') setMenuOpen(false);
    });
  }
  if (menu) {
    menu.addEventListener('click', function (event) {
      var option = event.target.closest('[data-breakdown-id]');
      if (!option) return;
      chooseMode(option.getAttribute('data-breakdown-id'));
    });
  }
  if (edit) edit.addEventListener('click', openCustomModal);
  document.addEventListener('pointerdown', function (event) {
    var root = document.getElementById('portfolio-breakdown-select');
    if (!menuOpen || !root || root.contains(event.target)) return;
    setMenuOpen(false);
  });
  if (!modal || modal.dataset.bound) return;
  modal.dataset.bound = 'true';
  modal.addEventListener('click', function (event) {
    if (event.target.closest('[data-breakdown-dismiss]') || event.target.id === 'portfolio-breakdown-close') {
      closeCustomModal(false);
      return;
    }
    var property = event.target.closest('.group-property');
    if (property && !event.target.closest('[data-group-action]')) {
      selectedKey = property.getAttribute('data-key');
      selectedGroupId = property.getAttribute('data-list');
      paintModal();
      return;
    }
    var action = event.target.closest('[data-group-action]');
    if (!action || action.disabled) return;
    var name = action.getAttribute('data-group-action');
    var groupId = action.getAttribute('data-group-id');
    if (name === 'cancel') closeCustomModal(false);
    if (name === 'add') addGroup();
    if (name === 'reset') {
      draft = { groups: [], assignments: {} };
      selectedKey = '';
      selectedGroupId = '';
      nameIssues = {};
      paintModal();
    }
    if (name === 'apply') applyDraft();
    if (name === 'delete') {
      draft.groups = draft.groups.filter(function (group) { return group.id !== groupId; });
      Object.keys(draft.assignments).forEach(function (key) {
        if (draft.assignments[key] === groupId) delete draft.assignments[key];
      });
      nameIssues = validateCustomGroups(draft.groups);
      paintModal();
    }
    if (name === 'up') {
      moveGroup(groupId, -1);
      paintModal();
    }
    if (name === 'down') {
      moveGroup(groupId, 1);
      paintModal();
    }
  });
  modal.addEventListener('input', function (event) {
    if (event.target.id === 'group-property-search') {
      searchQuery = event.target.value;
      paintModal();
      return;
    }
    var groupId = event.target.getAttribute('data-group-name');
    if (!groupId) return;
    draft.groups.forEach(function (group) {
      if (group.id === groupId) group.name = event.target.value;
    });
    nameIssues = validateCustomGroups(draft.groups);
    var footer = document.getElementById('portfolio-breakdown-footer');
    if (footer) footer.innerHTML = renderModalFooter();
    var issue = event.target.parentNode.querySelector('.group-issue');
    if (nameIssues[groupId]) {
      if (!issue) {
        issue = document.createElement('p');
        issue.className = 'group-issue';
        event.target.parentNode.insertAdjacentElement('afterend', issue);
      }
      issue.textContent = nameIssues[groupId];
    } else if (issue) issue.remove();
  });
  modal.addEventListener('change', function (event) {
    if (event.target.id !== 'group-move-to') return;
    assignProperty(selectedKey, event.target.value);
    paintModal();
  });
  modal.addEventListener('dragstart', function (event) {
    var handle = event.target.closest('[data-drag-kind]');
    if (!handle) return;
    dragKind = handle.getAttribute('data-drag-kind');
    dragId = dragKind === 'group' ? handle.getAttribute('data-group-id') : handle.getAttribute('data-key');
    handle.classList.add('is-dragging');
    event.dataTransfer.setData('text/plain', dragKind + ':' + dragId);
    event.dataTransfer.effectAllowed = 'move';
  });
  modal.addEventListener('dragover', function (event) {
    if (!dragKind) return;
    var target = event.target.closest('[data-drop-group], [data-drop-unassigned]');
    if (!target) return;
    event.preventDefault();
    modal.querySelectorAll('.is-drop-active').forEach(function (node) { node.classList.remove('is-drop-active'); });
    target.classList.add('is-drop-active');
    autoscroll(event);
  });
  modal.addEventListener('drop', function (event) {
    var target = event.target.closest('[data-drop-group], [data-drop-unassigned]');
    event.preventDefault();
    var kind = dragKind;
    var id = dragId;
    clearDrag();
    if (!target || !kind) return;
    if (kind === 'property') {
      var groupId = target.getAttribute('data-drop-group') || target.getAttribute('data-group-id');
      if (target.hasAttribute('data-drop-unassigned')) assignProperty(id, 'unassigned');
      else assignProperty(id, groupId);
      paintModal();
      return;
    }
    if (kind === 'group') {
      var before = target.getAttribute('data-drop-group');
      if (!before || before === id) return;
      var moving = draft.groups.filter(function (group) { return group.id === id; })[0];
      draft.groups = draft.groups.filter(function (group) { return group.id !== id; });
      var index = draft.groups.findIndex(function (group) { return group.id === before; });
      if (moving && index >= 0) draft.groups.splice(index, 0, moving);
      selectedGroupId = id;
      paintModal();
    }
  });
  modal.addEventListener('dragend', clearDrag);
  document.addEventListener('keydown', function (event) {
    if (!modal || modal.hidden) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      closeCustomModal(false);
    }
  });
}

export function syncPortfolioBreakdownToProperty(propertyData) {
  reconcileCustomAssignments(propertyData);
  refreshPortfolioBreakdownChrome();
}
