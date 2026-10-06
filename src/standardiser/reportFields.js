import { APP_STATE } from '../state.js';
import { escapeHtml } from '../ui/common.js';
import { resolvePortfolioValuation } from '../logic/fairValue.js';
import { readSourceValue } from '../schema/sourceLookup.js';
import { PROPERTY_SOURCE_FIELDS } from '../schema/sourceFields.js';
import { comparisonColumnActive } from '../logic/comparison.js';
import {
  defaultOptionalIds,
  getComparisonSelection,
  getReportLayout,
  listFixedColumns,
  listOptionalColumns,
  setReportLayout,
} from '../logic/reportLayout.js';

var draftIncluded = [];
var draftAvailable = [];
var draftComparison = {};
var selected = { list: 'included', id: '' };
var dragId = '';
var searchQuery = '';
var prompt = '';
var pendingFocus = '';

function optionalById() {
  var map = {};
  listOptionalColumns().forEach(function (column) {
    map[column.id] = column;
  });
  return map;
}

function fairValueHeading() {
  var propertyData = APP_STATE.property;
  if (!propertyData || !propertyData.rows || !propertyData.rows.length) return 'Fair Value';
  return resolvePortfolioValuation(propertyData.rows.map(function (row) {
    return readSourceValue(row, propertyData.columns, 'property', PROPERTY_SOURCE_FIELDS.traditionalValuationDate);
  })).heading;
}

function labelFor(id) {
  if (id === 'fair_value') return fairValueHeading();
  var column = optionalById()[id];
  return column ? column.label : id;
}

function syncDraftFromSession() {
  var included = getReportLayout();
  var includedSet = {};
  included.forEach(function (id) { includedSet[id] = true; });
  draftIncluded = included.slice();
  draftAvailable = defaultOptionalIds().filter(function (id) { return !includedSet[id]; });
  draftComparison = Object.assign({}, getComparisonSelection());
  selected = { list: 'included', id: draftIncluded[0] || '' };
}

function defaultComparisonMap() {
  var map = {};
  listOptionalColumns().forEach(function (column) {
    if (column.comparison && column.comparison.supported) map[column.id] = !!column.comparison.defaultEnabled;
  });
  return map;
}

function comparisonMatches(left, right) {
  var keys = {};
  Object.keys(left || {}).forEach(function (id) { keys[id] = true; });
  Object.keys(right || {}).forEach(function (id) { keys[id] = true; });
  return Object.keys(keys).every(function (id) {
    return !!left[id] === !!right[id];
  });
}

function sameOrder(left, right) {
  if (left.length !== right.length) return false;
  for (var i = 0; i < left.length; i += 1) {
    if (left[i] !== right[i]) return false;
  }
  return true;
}

function isDefaultDraft() {
  return sameOrder(draftIncluded, defaultOptionalIds()) && comparisonMatches(draftComparison, defaultComparisonMap());
}

function layoutDelta() {
  var session = getReportLayout();
  var sessionSet = {};
  var draftSet = {};
  session.forEach(function (id) { sessionSet[id] = true; });
  draftIncluded.forEach(function (id) { draftSet[id] = true; });
  var added = 0;
  var removed = 0;
  var reordered = 0;
  session.forEach(function (id) { if (!draftSet[id]) removed += 1; });
  draftIncluded.forEach(function (id) { if (!sessionSet[id]) added += 1; });
  var sessionShared = session.filter(function (id) { return draftSet[id]; });
  var draftShared = draftIncluded.filter(function (id) { return sessionSet[id]; });
  for (var i = 0; i < sessionShared.length; i += 1) {
    if (sessionShared[i] !== draftShared[i]) {
      reordered = 1;
      break;
    }
  }
  var selection = getComparisonSelection();
  Object.keys(draftComparison).forEach(function (id) {
    if (!!draftComparison[id] !== !!selection[id]) reordered += 1;
  });
  return added + removed + reordered;
}

function matchesQuery(id) {
  var query = searchQuery.trim().toLowerCase();
  if (!query) return true;
  return labelFor(id).toLowerCase().indexOf(query) !== -1;
}

function visibleIds(ids) {
  return ids.filter(matchesQuery);
}

function move(id, toIncluded, beforeId) {
  if (!id || beforeId === id) {
    selected = { list: toIncluded ? 'included' : 'available', id: id };
    return;
  }
  draftIncluded = draftIncluded.filter(function (item) { return item !== id; });
  draftAvailable = draftAvailable.filter(function (item) { return item !== id; });
  var target = toIncluded ? draftIncluded : draftAvailable;
  if (beforeId && target.indexOf(beforeId) !== -1) target.splice(target.indexOf(beforeId), 0, id);
  else target.push(id);
  selected = { list: toIncluded ? 'included' : 'available', id: id };
}

function shiftSelected(direction) {
  if (selected.list !== 'included' || !selected.id) return;
  var index = draftIncluded.indexOf(selected.id);
  var next = index + direction;
  if (index < 0 || next < 0 || next >= draftIncluded.length) return;
  var copy = draftIncluded.slice();
  var moved = copy.splice(index, 1)[0];
  copy.splice(next, 0, moved);
  draftIncluded = copy;
}

function lockIcon() {
  return '<svg class="report-lock" viewBox="0 0 16 16" aria-hidden="true" focusable="false"><rect x="3" y="7" width="10" height="7" rx="1" fill="none" stroke="currentColor" stroke-width="1.4"></rect><path d="M5.5 7V5a2.5 2.5 0 0 1 5 0v2" fill="none" stroke="currentColor" stroke-width="1.4"></path></svg>';
}

function gripIcon() {
  return '<svg width="10" height="14" viewBox="0 0 10 14" aria-hidden="true" focusable="false"><circle cx="2" cy="2" r="1.1"></circle><circle cx="8" cy="2" r="1.1"></circle><circle cx="2" cy="7" r="1.1"></circle><circle cx="8" cy="7" r="1.1"></circle><circle cx="2" cy="12" r="1.1"></circle><circle cx="8" cy="12" r="1.1"></circle></svg>';
}

function compareIcon() {
  return '<svg class="report-compare-icon" width="12" height="12" viewBox="0 0 12 12" aria-hidden="true" focusable="false"><circle cx="6" cy="6" r="4.25" fill="none" stroke="currentColor" stroke-width="1.1"></circle><path d="M6 3.6V6.2L7.7 7.2" fill="none" stroke="currentColor" stroke-width="1.1" stroke-linecap="round"></path></svg>';
}

function emptyCopy(listName, ids) {
  if (searchQuery.trim() && !visibleIds(ids).length) return 'No matching fields.';
  if (!ids.length && listName === 'available') return 'All available fields are currently included.';
  if (!ids.length) return 'No optional fields are included.';
  return '';
}

function renderItems(ids, listName) {
  var shown = visibleIds(ids);
  var empty = emptyCopy(listName, ids);
  if (empty && !shown.length) return '<li class="report-empty">' + empty + '</li>';
  return shown.map(function (id) {
    var active = selected.list === listName && selected.id === id;
    var column = optionalById()[id];
    var handle = listName === 'included' ? '<span class="report-handle" aria-hidden="true">' + gripIcon() + '</span>' : '';
    var meta = '';
    var compare = '';
    if (listName === 'included' && column && column.comparison && column.comparison.supported) {
      var on = !!draftComparison[id];
      var label = labelFor(id);
      if (on) {
        var childCopy = column.comparison.showPercentChange ? 'Current + Previous + Change %' : 'Current + Previous';
        meta += '<span class="report-compare-sub">' + childCopy + '</span>';
        if (!comparisonColumnActive(column, true, { previousProperty: APP_STATE.previousProperty, previousUnit: APP_STATE.previousUnit })) {
          meta += '<span class="report-compare-note">Awaiting previous data</span>';
        }
      }
      var changeFlag = column.comparison.showPercentChange ? '<span class="report-change-flag">Includes % change</span>' : '';
      compare = '<label class="report-compare" draggable="false"><span class="report-compare-main">' + compareIcon() + '<span>Compare</span><input type="checkbox" data-report-compare="' + escapeHtml(id) + '"' + (on ? ' checked' : '') + ' aria-label="Compare ' + escapeHtml(label) + ' to Previous Valuation"></span>' + changeFlag + '</label>';
    }
    return '<li class="report-item' + (active ? ' is-selected' : '') + '" draggable="true" data-list="' + listName + '" data-id="' + escapeHtml(id) + '" role="option" aria-selected="' + (active ? 'true' : 'false') + '">' + handle + '<span class="report-label"><span class="report-field-name">' + escapeHtml(labelFor(id)) + '</span>' + meta + '</span>' + compare + '</li>';
  }).join('');
}

function renderBody() {
  var required = listFixedColumns('required').map(function (column) {
    return '<li class="report-fixed">' + lockIcon() + '<span>' + escapeHtml(column.label) + '</span></li>';
  }).join('');
  var systems = listFixedColumns('system').map(function (column) {
    return '<li><span>' + escapeHtml(column.label) + '</span><span class="report-hidden-tag">Hidden in Excel</span></li>';
  }).join('');
  var includedIndex = draftIncluded.indexOf(selected.id);
  var upDisabled = selected.list !== 'included' || includedIndex <= 0 ? ' disabled' : '';
  var downDisabled = selected.list !== 'included' || includedIndex < 0 || includedIndex >= draftIncluded.length - 1 ? ' disabled' : '';
  var addDisabled = selected.list !== 'available' || !selected.id ? ' disabled' : '';
  var removeDisabled = selected.list !== 'included' || !selected.id ? ' disabled' : '';
  return (
    '<div class="report-meta">' +
    '<section class="report-fixed-block"><h3>Required Fields</h3><ul class="report-fixed-list">' + required + '</ul></section>' +
    '<section class="report-system-block"><h3>System fields</h3><ul>' + systems + '</ul></section>' +
    '</div>' +
    '<label class="report-search"><input id="report-field-search" type="search" placeholder="Search fields..." aria-label="Search fields" value="' + escapeHtml(searchQuery) + '"></label>' +
    '<div class="report-picker">' +
    '<section class="report-panel" data-drop="available">' +
    '<header class="report-panel-head"><div class="report-panel-copy"><div class="report-panel-title"><h3>Available Fields</h3><span class="report-count">' + draftAvailable.length + '</span></div><p>Not currently included</p></div></header>' +
    '<ul class="report-list" role="listbox" tabindex="0" aria-label="Available fields">' + renderItems(draftAvailable, 'available') + '</ul>' +
    '</section>' +
    '<div class="report-move" aria-label="Transfer fields">' +
    '<button type="button" data-report-action="add" aria-label="Add selected field"' + addDisabled + '>Add →</button>' +
    '<button type="button" data-report-action="remove" aria-label="Remove selected field"' + removeDisabled + '>← Remove</button>' +
    '</div>' +
    '<section class="report-panel" data-drop="included">' +
    '<header class="report-panel-head"><div class="report-panel-copy"><div class="report-panel-title"><h3>Included in Report</h3><span class="report-count">' + draftIncluded.length + '</span></div><p>Drag to reorder. Optional fields appear after the four required columns.</p><p class="report-legend">Numeric increases are shown in green; decreases in red.</p></div>' +
    '<div class="report-order-btns"><button type="button" data-report-action="up" aria-label="Move selected field up"' + upDisabled + '>Up</button>' +
    '<button type="button" data-report-action="down" aria-label="Move selected field down"' + downDisabled + '>Down</button></div></header>' +
    '<ul class="report-list" role="listbox" tabindex="0" aria-label="Included fields">' + renderItems(draftIncluded, 'included') + '</ul>' +
    '</section>' +
    '</div>'
  );
}

function renderFooter() {
  if (prompt === 'discard') {
    return '<p class="report-status">Discard layout changes?</p><div class="report-footer-actions"><button type="button" class="report-cancel" data-report-action="keep">Keep editing</button><button type="button" class="report-apply" data-report-action="discard">Discard</button></div>';
  }
  if (prompt === 'reset') {
    return '<p class="report-status">Restore the default field order?</p><div class="report-footer-actions"><button type="button" class="report-cancel" data-report-action="keep">Keep editing</button><button type="button" class="report-apply" data-report-action="confirm-reset">Restore</button></div>';
  }
  var changes = layoutDelta();
  var status = changes === 0 ? 'No changes' : (changes === 1 ? '1 layout change' : changes + ' layout changes');
  return '<p class="report-status" aria-live="polite">' + status + '</p><div class="report-footer-actions">' +
    '<button type="button" class="report-reset" data-report-action="reset"' + (isDefaultDraft() ? ' disabled' : '') + '>Reset to Default</button>' +
    '<button type="button" class="report-cancel" data-report-action="cancel">Cancel</button>' +
    '<button type="button" class="report-apply" data-report-action="apply"' + (changes === 0 ? ' disabled' : '') + '>Apply Report Layout</button>' +
    '</div>';
}

function captureListScroll() {
  var root = document.getElementById('report-fields-root');
  var state = [];
  if (!root) return state;
  root.querySelectorAll('.report-list').forEach(function (list) {
    state.push({ label: list.getAttribute('aria-label') || '', top: list.scrollTop });
  });
  return state;
}

function applyListScroll(state) {
  var root = document.getElementById('report-fields-root');
  if (!root) return;
  state.forEach(function (item) {
    var lists = root.querySelectorAll('.report-list');
    for (var i = 0; i < lists.length; i += 1) {
      if ((lists[i].getAttribute('aria-label') || '') === item.label) {
        lists[i].scrollTop = item.top;
      }
    }
  });
}

function paint() {
  var scrolls = captureListScroll();
  var search = document.getElementById('report-field-search');
  var searchFocused = !!(search && document.activeElement === search);
  var caret = searchFocused ? search.selectionStart : null;
  var root = document.getElementById('report-fields-root');
  var footer = document.getElementById('report-fields-footer');
  if (root) root.innerHTML = renderBody();
  if (footer) footer.innerHTML = renderFooter();
  applyListScroll(scrolls);
  if (searchFocused) {
    var nextSearch = document.getElementById('report-field-search');
    if (nextSearch) {
      nextSearch.focus({ preventScroll: true });
      if (caret != null) nextSearch.setSelectionRange(caret, caret);
    }
    applyListScroll(scrolls);
    return;
  }
  if (!pendingFocus) return;
  var modal = document.getElementById('report-fields-modal');
  var target = modal ? modal.querySelector(pendingFocus) : null;
  pendingFocus = '';
  if (target && !target.disabled) target.focus({ preventScroll: true });
  applyListScroll(scrolls);
}

function syncCompareRow(id) {
  var items = document.querySelectorAll('#report-fields-root .report-item[data-list="included"]');
  var item = null;
  for (var i = 0; i < items.length; i += 1) {
    if (items[i].getAttribute('data-id') === id) {
      item = items[i];
      break;
    }
  }
  if (!item) {
    paint();
    return;
  }
  var column = optionalById()[id];
  var label = item.querySelector('.report-label');
  if (!label || !column) return;
  label.querySelectorAll('.report-compare-sub, .report-compare-note').forEach(function (node) {
    node.remove();
  });
  if (draftComparison[id]) {
    var sub = document.createElement('span');
    sub.className = 'report-compare-sub';
    sub.textContent = column.comparison && column.comparison.showPercentChange
      ? 'Current + Previous + Change %'
      : 'Current + Previous';
    label.appendChild(sub);
    if (!comparisonColumnActive(column, true, {
      previousProperty: APP_STATE.previousProperty,
      previousUnit: APP_STATE.previousUnit,
    })) {
      var note = document.createElement('span');
      note.className = 'report-compare-note';
      note.textContent = 'Awaiting previous data';
      label.appendChild(note);
    }
  }
  var footer = document.getElementById('report-fields-footer');
  if (footer && !prompt) footer.innerHTML = renderFooter();
}

function clearDragState() {
  dragId = '';
  document.documentElement.classList.remove('report-dragging');
  var modal = document.getElementById('report-fields-modal');
  if (!modal) return;
  modal.querySelectorAll('.is-dragging, .is-insert-before, .is-drop-active').forEach(function (node) {
    node.classList.remove('is-dragging', 'is-insert-before', 'is-drop-active');
  });
}

function openModal() {
  if (APP_STATE.scheduleType !== 'tenancy') return;
  searchQuery = '';
  prompt = '';
  syncDraftFromSession();
  var modal = document.getElementById('report-fields-modal');
  if (!modal) return;
  modal.hidden = false;
  modal.classList.remove('hidden');
  document.body.classList.add('report-open');
  document.getElementById('app').setAttribute('inert', '');
  paint();
  var search = document.getElementById('report-field-search');
  if (search) search.focus();
}

function closeModal() {
  var modal = document.getElementById('report-fields-modal');
  if (!modal) return;
  prompt = '';
  modal.hidden = true;
  modal.classList.add('hidden');
  document.body.classList.remove('report-open');
  document.getElementById('app').removeAttribute('inert');
  var button = document.getElementById('report-fields-btn');
  if (button) button.focus();
}

function requestClose() {
  if (!layoutDelta()) {
    closeModal();
    return;
  }
  prompt = 'discard';
  pendingFocus = '[data-report-action="keep"]';
  paint();
}

function focusable(modal) {
  return Array.prototype.filter.call(modal.querySelectorAll('button, input, [tabindex="0"]'), function (el) {
    return !el.disabled && (el.offsetParent !== null || el === document.activeElement);
  });
}

export function refreshReportFieldsAvailability() {
  var button = document.getElementById('report-fields-btn');
  if (!button) return;
  var enabled = APP_STATE.scheduleType === 'tenancy';
  button.disabled = !enabled;
  button.title = enabled
    ? 'Choose and arrange optional report columns. Required and system fields remain fixed.'
    : 'Custom report fields are not yet available for this template.';
}

export function setupReportFields() {
  refreshReportFieldsAvailability();
  var button = document.getElementById('report-fields-btn');
  var modal = document.getElementById('report-fields-modal');
  if (button) button.addEventListener('click', openModal);
  if (!modal || modal.dataset.bound) return;
  modal.dataset.bound = 'true';
  modal.addEventListener('click', function (event) {
    if (event.target.closest('[data-report-dismiss]') || event.target.id === 'report-fields-close') {
      requestClose();
      return;
    }
    if (event.target.closest('.report-compare')) return;
    var item = event.target.closest('.report-item');
    if (item) {
      selected = { list: item.getAttribute('data-list'), id: item.getAttribute('data-id') };
      paint();
      return;
    }
    var action = event.target.closest('[data-report-action]');
    if (!action || action.disabled) return;
    var name = action.getAttribute('data-report-action');
    if (name === 'cancel') requestClose();
    if (name === 'keep') {
      prompt = '';
      pendingFocus = '[data-report-action="cancel"]';
      paint();
    }
    if (name === 'discard') closeModal();
    if (name === 'reset') {
      prompt = 'reset';
      pendingFocus = '[data-report-action="keep"]';
      paint();
    }
    if (name === 'confirm-reset') {
      draftIncluded = defaultOptionalIds();
      draftAvailable = [];
      selected = { list: 'included', id: draftIncluded[0] || '' };
      searchQuery = '';
      prompt = '';
      draftComparison = defaultComparisonMap();
      pendingFocus = '[data-report-action="apply"]';
      paint();
    }
    if (name === 'apply') {
      if (!layoutDelta()) return;
      setReportLayout(draftIncluded, draftComparison);
      closeModal();
    }
    if (name === 'add' && selected.list === 'available' && selected.id) move(selected.id, true);
    if (name === 'remove' && selected.list === 'included' && selected.id) move(selected.id, false);
    if (name === 'up') shiftSelected(-1);
    if (name === 'down') shiftSelected(1);
    if (name === 'add' || name === 'remove') pendingFocus = '[data-drop="' + selected.list + '"] .report-list';
    if (name === 'up' || name === 'down') pendingFocus = '[data-report-action="' + name + '"]';
    if (name === 'add' || name === 'remove' || name === 'up' || name === 'down') paint();
  });
  modal.addEventListener('change', function (event) {
    var input = event.target.closest('[data-report-compare]');
    if (!input) return;
    draftComparison[input.getAttribute('data-report-compare')] = !!input.checked;
    syncCompareRow(input.getAttribute('data-report-compare'));
  });
  modal.addEventListener('input', function (event) {
    if (event.target.id !== 'report-field-search') return;
    searchQuery = event.target.value;
    paint();
  });
  modal.addEventListener('dragstart', function (event) {
    if (event.target.closest('.report-compare')) {
      event.preventDefault();
      return;
    }
    var item = event.target.closest('.report-item');
    if (!item) return;
    dragId = item.getAttribute('data-id');
    selected = { list: item.getAttribute('data-list'), id: dragId };
    item.classList.add('is-dragging');
    document.documentElement.classList.add('report-dragging');
    event.dataTransfer.setData('text/plain', dragId);
    event.dataTransfer.effectAllowed = 'move';
  });
  modal.addEventListener('dragover', function (event) {
    var panel = event.target.closest('[data-drop]');
    if (!panel) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
    modal.querySelectorAll('[data-drop]').forEach(function (node) { node.classList.remove('is-drop-active'); });
    panel.classList.add('is-drop-active');
    modal.querySelectorAll('.is-insert-before').forEach(function (node) { node.classList.remove('is-insert-before'); });
    var item = event.target.closest('.report-item');
    if (item && panel.contains(item)) item.classList.add('is-insert-before');
  });
  modal.addEventListener('drop', function (event) {
    var panel = event.target.closest('[data-drop]');
    if (!panel) return;
    event.preventDefault();
    var item = event.target.closest('.report-item');
    var id = dragId || event.dataTransfer.getData('text/plain');
    clearDragState();
    if (!id) return;
    move(id, panel.getAttribute('data-drop') === 'included', item ? item.getAttribute('data-id') : '');
    pendingFocus = '[data-drop="' + selected.list + '"] .report-list';
    paint();
  });
  modal.addEventListener('dragend', clearDragState);
  document.addEventListener('keydown', function (event) {
    if (modal.hidden) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      if (prompt) {
        prompt = '';
        pendingFocus = '[data-report-action="cancel"]';
        paint();
        return;
      }
      requestClose();
      return;
    }
    if ((event.key === 'ArrowUp' || event.key === 'ArrowDown') && event.target.closest('[data-drop]')) {
      var panel = event.target.closest('[data-drop]');
      var listName = panel.getAttribute('data-drop');
      var ids = visibleIds(listName === 'included' ? draftIncluded : draftAvailable);
      if (!ids.length) return;
      event.preventDefault();
      var index = selected.list === listName ? ids.indexOf(selected.id) : -1;
      var next = event.key === 'ArrowDown' ? Math.min(ids.length - 1, index + 1) : Math.max(0, index <= 0 ? 0 : index - 1);
      selected = { list: listName, id: ids[next] };
      pendingFocus = '[data-drop="' + listName + '"] .report-list';
      paint();
      return;
    }
    if (event.key !== 'Tab') return;
    var nodes = focusable(modal);
    if (!nodes.length) return;
    var first = nodes[0];
    var last = nodes[nodes.length - 1];
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && (document.activeElement === last || nodes.indexOf(document.activeElement) === -1)) {
      event.preventDefault();
      first.focus();
    }
  });
}
