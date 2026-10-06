import {
  getDefaultScheduleType,
  getScheduleTypeByStateKey,
  SCHEDULE_TYPES,
} from '../config/scheduleTypes.js';
import { APP_STATE } from '../state.js';
import { escapeHtml } from './common.js';
import { refreshReportFieldsAvailability } from '../standardiser/reportFields.js';

var open = false;
var highlightIndex = -1;

function getSelectedType() {
  return getScheduleTypeByStateKey(APP_STATE.scheduleType) || getDefaultScheduleType();
}

function getEls() {
  return {
    root: document.getElementById('schedule-type-select'),
    trigger: document.getElementById('schedule-type-trigger'),
    triggerLabel: document.getElementById('schedule-type-trigger-label'),
    menu: document.getElementById('schedule-type-menu'),
    blurb: document.getElementById('schedule-type-blurb'),
  };
}

function applyScheduleType(type) {
  if (!type) return;
  APP_STATE.scheduleType = type.stateKey;

  for (var i = 0; i < SCHEDULE_TYPES.length; i += 1) {
    var entry = SCHEDULE_TYPES[i];
    var view = document.getElementById(entry.viewId);
    if (view) view.classList.toggle('hidden', entry.stateKey !== type.stateKey);
  }

  var els = getEls();
  if (els.triggerLabel) els.triggerLabel.textContent = type.label;
  if (els.blurb) els.blurb.textContent = type.description;
  refreshReportFieldsAvailability();
  if (els.trigger) els.trigger.setAttribute('aria-label', 'Schedule type, ' + type.label);
  renderMenu();
}

function renderMenu() {
  var els = getEls();
  if (!els.menu) return;
  var selected = getSelectedType();

  els.menu.innerHTML = SCHEDULE_TYPES.map(function (type, index) {
    var isSelected = type.id === selected.id;
    var isHighlighted = index === highlightIndex;
    return (
      '<li role="option" class="schedule-type-option' +
      (isSelected ? ' is-selected' : '') +
      (isHighlighted ? ' is-highlighted' : '') +
      '" data-schedule-id="' +
      escapeHtml(type.id) +
      '" aria-selected="' +
      (isSelected ? 'true' : 'false') +
      '" id="schedule-type-option-' +
      escapeHtml(type.id) +
      '">' +
      '<span class="schedule-type-option-mark" aria-hidden="true">' +
      (isSelected ? '\u2713' : '') +
      '</span>' +
      '<span class="schedule-type-option-copy">' +
      '<span class="schedule-type-option-label">' +
      escapeHtml(type.label) +
      '</span>' +
      '<span class="schedule-type-option-summary">' +
      escapeHtml(type.summary || '') +
      '</span>' +
      '</span>' +
      '</li>'
    );
  }).join('');
}

function setOpen(nextOpen) {
  var els = getEls();
  open = Boolean(nextOpen);
  if (!els.root || !els.trigger || !els.menu) return;

  els.root.classList.toggle('is-open', open);
  els.menu.classList.toggle('hidden', !open);
  els.trigger.setAttribute('aria-expanded', open ? 'true' : 'false');

  if (open) {
    var selected = getSelectedType();
    highlightIndex = Math.max(
      0,
      SCHEDULE_TYPES.findIndex(function (type) {
        return type.id === selected.id;
      })
    );
    renderMenu();
    var option = document.getElementById('schedule-type-option-' + selected.id);
    if (option) {
      els.trigger.setAttribute('aria-activedescendant', option.id);
      option.scrollIntoView({ block: 'nearest' });
    }
  } else {
    highlightIndex = -1;
    els.trigger.removeAttribute('aria-activedescendant');
  }
}

function selectByIndex(index) {
  var type = SCHEDULE_TYPES[index];
  if (!type) return;
  applyScheduleType(type);
  setOpen(false);
  var els = getEls();
  if (els.trigger) els.trigger.focus();
}

export function setScheduleType(stateKey) {
  applyScheduleType(getScheduleTypeByStateKey(stateKey) || getDefaultScheduleType());
}

export function setupScheduleTypeSwitcher() {
  var els = getEls();
  if (!els.root || !els.trigger || !els.menu) return;

  renderMenu();
  applyScheduleType(getDefaultScheduleType());
  setOpen(false);

  els.trigger.addEventListener('click', function (event) {
    event.stopPropagation();
    setOpen(!open);
  });

  els.menu.addEventListener('click', function (event) {
    var option = event.target.closest('[data-schedule-id]');
    if (!option) return;
    var index = SCHEDULE_TYPES.findIndex(function (type) {
      return type.id === option.dataset.scheduleId;
    });
    selectByIndex(index);
  });

  els.trigger.addEventListener('keydown', function (event) {
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault();
      if (!open) {
        setOpen(true);
        return;
      }
      var delta = event.key === 'ArrowDown' ? 1 : -1;
      highlightIndex = (highlightIndex + delta + SCHEDULE_TYPES.length) % SCHEDULE_TYPES.length;
      renderMenu();
      var highlighted = SCHEDULE_TYPES[highlightIndex];
      if (highlighted) {
        els.trigger.setAttribute('aria-activedescendant', 'schedule-type-option-' + highlighted.id);
      }
      return;
    }

    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      if (!open) {
        setOpen(true);
        return;
      }
      selectByIndex(highlightIndex >= 0 ? highlightIndex : 0);
      return;
    }

    if (event.key === 'Escape') {
      if (open) {
        event.preventDefault();
        setOpen(false);
      }
      return;
    }

    if (event.key === 'Home' && open) {
      event.preventDefault();
      highlightIndex = 0;
      renderMenu();
      return;
    }

    if (event.key === 'End' && open) {
      event.preventDefault();
      highlightIndex = SCHEDULE_TYPES.length - 1;
      renderMenu();
    }
  });

  document.addEventListener('click', function (event) {
    if (!open) return;
    if (els.root.contains(event.target)) return;
    setOpen(false);
  });

  document.addEventListener('keydown', function (event) {
    if (event.key === 'Escape' && open) setOpen(false);
  });
}
