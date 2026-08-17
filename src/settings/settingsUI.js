import { APP_VERSION, CONFIG_VERSION, STANDARD_TEMPLATE_COLUMNS } from '../config/constants.js';
import { cloneDefaultConfig, exportConfigJson, importConfigJson, validateConfig } from '../config/configManager.js';
import {
  applyDefaultMatchFields,
  applyMappingConfig,
  blankMapping,
  cloneMappingConfig,
  getColumnMappingStatus,
  getMappingSummary,
  sourceMappingFromColumn,
} from '../config/mapping.js';
import {
  PROPERTY_SUMMARY_COLUMNS,
  getPropertySummaryMappingSummary,
} from '../config/propertySummaryColumns.js';
import {
  getDefaultScheduleType,
  getScheduleTypeById,
  getScheduleTypeByStateKey,
  SCHEDULE_TYPES,
} from '../config/scheduleTypes.js';
import { APP_STATE, PROPERTY_SUMMARY_STATE } from '../state.js';
import { cssEscape, escapeAttr, escapeHtml, openConfirm } from '../ui/common.js';

var draftConfig = null;
var onConfigChanged = null;
var settingsTemplateId = 'tenancy_schedule';
var settingsDirty = false;

export function setupSettingsUI(options) {
  options = options || {};
  onConfigChanged = options.onConfigChanged || null;

  var modal = document.getElementById('settings-modal');
  var openBtn = document.getElementById('settings-btn');

  openBtn.addEventListener('click', function () {
    openSettings();
  });

  modal.querySelectorAll('[data-role="close"]').forEach(function (el) {
    el.addEventListener('click', function () {
      closeSettings();
    });
  });

  document.addEventListener('keydown', function (event) {
    if (event.key === 'Escape' && !modal.classList.contains('hidden')) {
      var confirmModal = document.getElementById('confirm-modal');
      if (confirmModal && !confirmModal.classList.contains('hidden')) {
        document.getElementById('confirm-modal').classList.add('hidden');
        return;
      }
      closeSettings();
    }
  });

  modal.querySelectorAll('[data-settings-tab]').forEach(function (tabBtn) {
    tabBtn.addEventListener('click', function () {
      activateTab(tabBtn.dataset.settingsTab);
    });
  });

  populateSettingsTemplateSelect();
  var templateSelect = document.getElementById('settings-template-select');
  if (templateSelect) {
    templateSelect.addEventListener('change', function () {
      settingsTemplateId = templateSelect.value;
      applySettingsTemplateView();
    });
  }

  document.getElementById('btn-save-config').addEventListener('click', function () {
    saveSettingsChanges();
  });

  document.getElementById('btn-export-config').addEventListener('click', function () {
    syncDraftFromDom();
    applyMappingConfig(draftConfig);
    notifyChanged();
    var result = exportConfigJson(draftConfig);
    if (result.ok) {
      setDirty(false);
      showFeedback('Downloaded ' + result.filename + '.', 'success');
    } else {
      showFeedback('The configuration file could not be downloaded.', 'error');
    }
  });

  var importInput = document.getElementById('config-import-input');
  document.getElementById('btn-import-config').addEventListener('click', function () {
    if (!importInput) return;
    try {
      importInput.click();
    } catch (err) {
      showFeedback('The file picker could not be opened.', 'error');
    }
  });

  if (importInput) {
    importInput.addEventListener('change', function () {
      var file = importInput.files && importInput.files[0];
      importInput.value = '';
      if (!file) return;

      importConfigJson(file)
        .then(function (imported) {
          draftConfig = cloneMappingConfig(imported);
          applyDefaultMatchFields(draftConfig, APP_STATE.property, APP_STATE.unit);
          applyMappingConfig(draftConfig);
          setDirty(false);
          showFeedback('Configuration imported and applied.', 'success');
          notifyChanged();
          renderAll();
        })
        .catch(function (err) {
          showFeedback(
            err && err.message
              ? err.message
              : 'The configuration file could not be imported.',
            'error'
          );
        });
    });
  }

  document.getElementById('btn-reset-mapping').addEventListener('click', function () {
    openConfirm(
      'Reset mapping?',
      'This restores the shipped Default and Property Only mappings. Uploaded Excel files are not affected.',
      function () {
        draftConfig = cloneMappingConfig(cloneDefaultConfig());
        applyMappingConfig(draftConfig);
        setDirty(false);
        showFeedback('Mappings restored to the original defaults.', 'success');
        notifyChanged();
        renderAll();
      },
      'Reset Mapping'
    );
  });

  var filterInput = document.getElementById('mapping-filter');
  if (filterInput) {
    filterInput.addEventListener('input', function () {
      filterMappingRows(filterInput.value);
    });
  }

  var psFilterInput = document.getElementById('ps-mapping-filter');
  if (psFilterInput) {
    psFilterInput.addEventListener('input', function () {
      filterPropertySummaryRows(psFilterInput.value);
    });
  }
}

function notifyChanged() {
  if (typeof onConfigChanged === 'function') onConfigChanged();
}

function populateSettingsTemplateSelect() {
  var select = document.getElementById('settings-template-select');
  if (!select) return;
  select.innerHTML = SCHEDULE_TYPES.map(function (type) {
    return (
      '<option value="' +
      escapeAttr(type.id) +
      '">' +
      escapeHtml(type.label) +
      '</option>'
    );
  }).join('');
}

function getInitialSettingsTemplate() {
  var app = document.getElementById('app');
  if (app && app.classList.contains('app-mode-compare')) {
    return getDefaultScheduleType();
  }
  return getScheduleTypeByStateKey(APP_STATE.scheduleType) || getDefaultScheduleType();
}

function applySettingsTemplateView() {
  var type = getScheduleTypeById(settingsTemplateId) || getDefaultScheduleType();
  var tenancyBlock = document.getElementById('tenancy-mapping-block');
  var psBlock = document.getElementById('ps-mapping-block');
  var matchingTab = document.getElementById('settings-matching-tab');
  var matchingPanel = document.querySelector('[data-settings-panel="matching"]');
  var isTenancy = type.mappingKind !== 'propertySummary';

  if (tenancyBlock) tenancyBlock.classList.toggle('hidden', !isTenancy);
  if (psBlock) psBlock.classList.toggle('hidden', isTenancy);
  if (matchingTab) matchingTab.classList.toggle('hidden', !type.hasPropertyMatching);
  if (matchingPanel && !type.hasPropertyMatching && !matchingPanel.classList.contains('hidden')) {
    activateTab('mapping');
  }
}

function setDirty(dirty) {
  settingsDirty = Boolean(dirty);
  var el = document.getElementById('settings-dirty-state');
  if (!el) return;
  el.classList.remove('is-dirty', 'is-saved');
  if (settingsDirty) {
    el.textContent = 'Unsaved changes';
    el.classList.add('is-dirty');
  } else {
    el.textContent = '';
  }
}

function markDirty() {
  setDirty(true);
}

var savedStateTimer = null;

function saveSettingsChanges() {
  syncDraftFromDom();
  var check = validateConfig(draftConfig);
  if (!check.valid) {
    showFeedback(
      'The configuration is not valid: ' + (check.errors || []).join(' '),
      'error'
    );
    return;
  }
  applyMappingConfig(draftConfig);
  setDirty(false);
  var el = document.getElementById('settings-dirty-state');
  if (el) {
    el.textContent = '\u2713 Configuration saved';
    el.classList.add('is-saved');
  }
  if (savedStateTimer) clearTimeout(savedStateTimer);
  savedStateTimer = setTimeout(function () {
    if (!settingsDirty) setDirty(false);
  }, 2500);
  showFeedback('Configuration saved.', 'success');
  notifyChanged();
  renderAll();
}

function openSettings() {
  draftConfig = cloneMappingConfig(APP_STATE.mappingConfig || cloneDefaultConfig());
  applyDefaultMatchFields(draftConfig, APP_STATE.property, APP_STATE.unit);

  var initial = getInitialSettingsTemplate();
  settingsTemplateId = initial.id;
  var templateSelect = document.getElementById('settings-template-select');
  if (templateSelect) templateSelect.value = settingsTemplateId;

  var modal = document.getElementById('settings-modal');
  modal.classList.remove('hidden');
  activateTab('mapping');
  applySettingsTemplateView();
  setDirty(false);
  clearFeedback();
  renderAll();
}

function closeSettings() {
  document.getElementById('settings-modal').classList.add('hidden');
  draftConfig = null;
}

function activateTab(tabId) {
  var matchingTab = document.getElementById('settings-matching-tab');
  if (tabId === 'matching' && matchingTab && matchingTab.classList.contains('hidden')) {
    tabId = 'mapping';
  }
  var modal = document.getElementById('settings-modal');
  modal.querySelectorAll('[data-settings-tab]').forEach(function (btn) {
    btn.classList.toggle('active', btn.dataset.settingsTab === tabId);
  });
  modal.querySelectorAll('[data-settings-panel]').forEach(function (panel) {
    panel.classList.toggle('hidden', panel.dataset.settingsPanel !== tabId);
  });
}

function renderAll() {
  if (!draftConfig) return;
  renderMappingSummary();
  renderMappingTable();
  renderPropertySummarySummary();
  renderPropertySummaryTable();
  renderMatchFields();
  renderConfigTab();
}

function showFeedback(message, kind) {
  var el = document.querySelector('#settings-modal [data-role="settings-feedback"]');
  if (!el) return;
  el.textContent = message;
  el.classList.remove('hidden', 'feedback-success', 'feedback-error');
  el.classList.add(kind === 'error' ? 'feedback-error' : 'feedback-success');
}

function clearFeedback() {
  var el = document.querySelector('#settings-modal [data-role="settings-feedback"]');
  if (!el) return;
  el.textContent = '';
  el.classList.add('hidden');
  el.classList.remove('feedback-success', 'feedback-error');
}

function renderMappingSummary() {
  var summary = getMappingSummary(draftConfig, APP_STATE.property, APP_STATE.unit);
  var el = document.getElementById('mapping-summary');
  el.innerHTML =
    '<span><strong>' +
    summary.standardColumns +
    '</strong> Standard Columns</span>' +
    '<span><strong>' +
    summary.mappedColumns +
    '</strong> Mapped</span>' +
    '<span><strong>' +
    summary.partiallyMapped +
    '</strong> Partially mapped</span>' +
    '<span><strong>' +
    summary.unmappedColumns +
    '</strong> Unmapped</span>' +
    '<span><strong>' +
    summary.ignoredColumns +
    '</strong> Ignored</span>' +
    (summary.missingMappings
      ? '<span class="summary-missing"><strong>' +
        summary.missingMappings +
        '</strong> Missing source</span>'
      : '');
}

function getPropertyParsedForSettings() {
  return PROPERTY_SUMMARY_STATE.property || APP_STATE.property;
}

function buildSourceOptionsHtml(fileType, selectedMapping, parsedOverride) {
  var parsed = parsedOverride !== undefined ? parsedOverride : APP_STATE[fileType];
  var columns = (parsed && parsed.columns) || [];
  var selectedKey =
    selectedMapping && selectedMapping.type === 'source' ? selectedMapping.sourceKey : '';

  var html = '<option value="">-- Blank --</option>';

  if (selectedKey && !columns.some(function (c) {
    return c.key === selectedKey;
  })) {
    var label = [
      selectedMapping.sourceSection || '(No Section)',
      selectedMapping.sourceField || '(unknown)',
    ].join(' > ');
    html +=
      '<option value="' +
      escapeAttr(selectedKey) +
      '" selected>\n      \u26a0 Missing: ' +
      escapeHtml(label) +
      '\n    </option>';
  }

  for (var i = 0; i < columns.length; i += 1) {
    var col = columns[i];
    var selected = col.key === selectedKey ? ' selected' : '';
    html +=
      '<option value="' +
      escapeAttr(col.key) +
      '"' +
      selected +
      '>' +
      escapeHtml(col.displayLabel) +
      '</option>';
  }

  if (!parsed) {
    html +=
      '<option value="" disabled>Upload a ' +
      (fileType === 'property' ? 'Property' : 'Unit') +
      ' file to see source fields</option>';
  }

  return html;
}

function renderMappingTable() {
  var tbody = document.getElementById('mapping-table-body');
  var filterValue = ((document.getElementById('mapping-filter') &&
    document.getElementById('mapping-filter').value) ||
    ''
  ).trim().toLowerCase();

  var rows = STANDARD_TEMPLATE_COLUMNS.map(function (col) {
    var entry = null;
    for (var i = 0; i < draftConfig.columnMappings.length; i += 1) {
      if (draftConfig.columnMappings[i].targetColumnId === col.id) {
        entry = draftConfig.columnMappings[i];
        break;
      }
    }
    if (!entry) {
      entry = {
        targetColumnId: col.id,
        propertyMapping: blankMapping(),
        unitMapping: blankMapping(),
        ignore: false,
      };
    }

    var ignored = entry.ignore === true;
    var status = getColumnMappingStatus(entry, APP_STATE.property, APP_STATE.unit);
    var statusClass = statusClassName(status);
    var hidden = filterValue && col.label.toLowerCase().indexOf(filterValue) === -1 ? 'hidden' : '';

    return (
      '<tr data-target-id="' +
      escapeAttr(col.id) +
      '" class="' +
      hidden +
      (ignored ? ' is-ignored' : '') +
      '">' +
      '<td class="col-standard">' +
      '<div class="standard-label">' +
      escapeHtml(col.label) +
      '</div>' +
      '<div class="standard-id">' +
      escapeHtml(col.id) +
      '</div>' +
      '</td>' +
      '<td>' +
      '<label class="sr-only" for="prop-map-' +
      escapeAttr(col.id) +
      '">Property Row Source for ' +
      escapeHtml(col.label) +
      '</label>' +
      '<select id="prop-map-' +
      escapeAttr(col.id) +
      '" data-map-side="property" data-target-id="' +
      escapeAttr(col.id) +
      '" class="source-select"' +
      (ignored ? ' disabled' : '') +
      '>' +
      buildSourceOptionsHtml('property', entry.propertyMapping) +
      '</select>' +
      '</td>' +
      '<td>' +
      '<label class="sr-only" for="unit-map-' +
      escapeAttr(col.id) +
      '">Unit Row Source for ' +
      escapeHtml(col.label) +
      '</label>' +
      '<select id="unit-map-' +
      escapeAttr(col.id) +
      '" data-map-side="unit" data-target-id="' +
      escapeAttr(col.id) +
      '" class="source-select"' +
      (ignored ? ' disabled' : '') +
      '>' +
      buildSourceOptionsHtml('unit', entry.unitMapping) +
      '</select>' +
      '</td>' +
      '<td class="col-ignore">' +
      '<label class="ignore-check">' +
      '<input type="checkbox" data-ignore-target="' +
      escapeAttr(col.id) +
      '"' +
      (ignored ? ' checked' : '') +
      ' title="Excluded from generated schedule" />' +
      '<span class="sr-only">Ignore in output</span>' +
      '</label>' +
      '</td>' +
      '<td class="col-status">' +
      '<span class="status-pill ' +
      statusClass +
      '">' +
      escapeHtml(status) +
      '</span>' +
      '</td>' +
      '</tr>'
    );
  }).join('');

  tbody.innerHTML = rows;

  tbody.querySelectorAll('select.source-select').forEach(function (select) {
    select.addEventListener('change', function () {
      var targetId = select.dataset.targetId;
      var side = select.dataset.mapSide;
      var key = select.value;
      updateDraftColumnMapping(targetId, side, key);
      markDirty();
      renderMappingSummary();
      updateRowStatus(targetId);
    });
  });

  tbody.querySelectorAll('input[data-ignore-target]').forEach(function (checkbox) {
    checkbox.addEventListener('change', function () {
      updateDraftIgnore(checkbox.dataset.ignoreTarget, checkbox.checked);
      markDirty();
      renderMappingTable();
      renderMappingSummary();
    });
  });
}

function updateRowStatus(targetId) {
  var entry = null;
  for (var i = 0; i < draftConfig.columnMappings.length; i += 1) {
    if (draftConfig.columnMappings[i].targetColumnId === targetId) {
      entry = draftConfig.columnMappings[i];
      break;
    }
  }
  if (!entry) return;
  var status = getColumnMappingStatus(entry, APP_STATE.property, APP_STATE.unit);
  var row = document.querySelector(
    '#mapping-table-body tr[data-target-id="' + cssEscape(targetId) + '"]'
  );
  if (!row) return;
  var pill = row.querySelector('.status-pill');
  if (pill) {
    pill.textContent = status;
    pill.className = 'status-pill ' + statusClassName(status);
  }
}

function filterMappingRows(query) {
  var q = query.trim().toLowerCase();
  document.querySelectorAll('#mapping-table-body tr[data-target-id]').forEach(function (row) {
    var label = (row.querySelector('.standard-label') &&
      row.querySelector('.standard-label').textContent) ||
      '';
    row.classList.toggle('hidden', Boolean(q) && label.toLowerCase().indexOf(q) === -1);
  });
}

function updateDraftColumnMapping(targetId, side, sourceKey) {
  var entry = null;
  for (var i = 0; i < draftConfig.columnMappings.length; i += 1) {
    if (draftConfig.columnMappings[i].targetColumnId === targetId) {
      entry = draftConfig.columnMappings[i];
      break;
    }
  }
  if (!entry) {
    entry = {
      targetColumnId: targetId,
      propertyMapping: blankMapping(),
      unitMapping: blankMapping(),
      ignore: false,
    };
    draftConfig.columnMappings.push(entry);
  }

  var field = side === 'property' ? 'propertyMapping' : 'unitMapping';
  entry[field] = mappingFromSelectValue(side, sourceKey, entry[field]);
}

function updateDraftIgnore(targetId, ignored) {
  var entry = null;
  for (var i = 0; i < draftConfig.columnMappings.length; i += 1) {
    if (draftConfig.columnMappings[i].targetColumnId === targetId) {
      entry = draftConfig.columnMappings[i];
      break;
    }
  }
  if (!entry) {
    entry = {
      targetColumnId: targetId,
      propertyMapping: blankMapping(),
      unitMapping: blankMapping(),
      ignore: false,
    };
    draftConfig.columnMappings.push(entry);
  }
  entry.ignore = Boolean(ignored);
}

function mappingFromSelectValue(fileType, sourceKey, previous, parsedOverride) {
  if (!sourceKey) return blankMapping();

  var parsed = parsedOverride !== undefined ? parsedOverride : APP_STATE[fileType];
  var cols = (parsed && parsed.columns) || [];
  var col = null;
  for (var i = 0; i < cols.length; i += 1) {
    if (cols[i].key === sourceKey) {
      col = cols[i];
      break;
    }
  }
  if (col) return sourceMappingFromColumn(col);

  if (previous && previous.type === 'source' && previous.sourceKey === sourceKey) {
    return {
      type: 'source',
      sourceKey: sourceKey,
      sourceSection: previous.sourceSection || '',
      sourceField: previous.sourceField || '',
      occurrence: previous.occurrence || 1,
    };
  }

  return {
    type: 'source',
    sourceKey: sourceKey,
    sourceSection: '',
    sourceField: '',
    occurrence: 1,
  };
}

function renderMatchFields() {
  var propSelect = document.getElementById('property-match-select');
  var unitSelect = document.getElementById('unit-match-select');

  propSelect.innerHTML = buildSourceOptionsHtml('property', draftConfig.propertyMatch);
  unitSelect.innerHTML = buildSourceOptionsHtml('unit', draftConfig.unitMatch);

  propSelect.onchange = function () {
    draftConfig.propertyMatch = mappingFromSelectValue(
      'property',
      propSelect.value,
      draftConfig.propertyMatch
    );
    markDirty();
  };
  unitSelect.onchange = function () {
    draftConfig.unitMatch = mappingFromSelectValue(
      'unit',
      unitSelect.value,
      draftConfig.unitMatch
    );
    markDirty();
  };
}

function renderConfigTab() {
  document.getElementById('config-version-value').textContent = String(
    draftConfig.configVersion || CONFIG_VERSION
  );
  document.getElementById('app-version-value').textContent = APP_VERSION;
}

function ensureDraftPropertySummary() {
  if (!draftConfig.propertySummary || !Array.isArray(draftConfig.propertySummary.columnMappings)) {
    draftConfig.propertySummary = {
      columnMappings: PROPERTY_SUMMARY_COLUMNS.map(function (col) {
        return { targetColumnId: col.id, propertyMapping: blankMapping(), ignore: false };
      }),
    };
  }
}

function updateDraftPropertySummaryMapping(targetId, sourceKey) {
  ensureDraftPropertySummary();
  var entry = null;
  for (var i = 0; i < draftConfig.propertySummary.columnMappings.length; i += 1) {
    if (draftConfig.propertySummary.columnMappings[i].targetColumnId === targetId) {
      entry = draftConfig.propertySummary.columnMappings[i];
      break;
    }
  }
  if (!entry) {
    entry = { targetColumnId: targetId, propertyMapping: blankMapping(), ignore: false };
    draftConfig.propertySummary.columnMappings.push(entry);
  }
  entry.propertyMapping = mappingFromSelectValue(
    'property',
    sourceKey,
    entry.propertyMapping,
    getPropertyParsedForSettings()
  );
}

function renderPropertySummarySummary() {
  var el = document.getElementById('ps-mapping-summary');
  if (!el || !draftConfig) return;
  var summary = getPropertySummaryMappingSummary(
    draftConfig.propertySummary,
    getPropertyParsedForSettings()
  );
  el.innerHTML =
    '<span><strong>' +
    summary.standardColumns +
    '</strong> Standard Columns</span>' +
    '<span><strong>' +
    summary.propertyMappings +
    '</strong> Mapped</span>' +
    '<span><strong>' +
    summary.unmappedColumns +
    '</strong> Unmapped</span>' +
    '<span><strong>' +
    summary.ignoredColumns +
    '</strong> Ignored</span>' +
    (summary.missingMappings
      ? '<span class="summary-missing"><strong>' +
        summary.missingMappings +
        '</strong> Missing source</span>'
      : '');
}

function renderPropertySummaryTable() {
  var tbody = document.getElementById('ps-mapping-table-body');
  if (!tbody || !draftConfig) return;
  ensureDraftPropertySummary();

  var filterValue = ((document.getElementById('ps-mapping-filter') &&
    document.getElementById('ps-mapping-filter').value) ||
    ''
  ).trim().toLowerCase();
  var parsed = getPropertyParsedForSettings();

  var rows = PROPERTY_SUMMARY_COLUMNS.map(function (col) {
    var entry = null;
    for (var i = 0; i < draftConfig.propertySummary.columnMappings.length; i += 1) {
      if (draftConfig.propertySummary.columnMappings[i].targetColumnId === col.id) {
        entry = draftConfig.propertySummary.columnMappings[i];
        break;
      }
    }
    if (!entry) {
      entry = { targetColumnId: col.id, propertyMapping: blankMapping(), ignore: false };
    }

    var ignored = entry.ignore === true;
    var mapping = entry.propertyMapping;
    var status = ignored ? 'Ignored' : 'Unmapped';
    if (!ignored && mapping && mapping.type === 'source') {
      var found = parsed && parsed.columns && parsed.columns.some(function (c) {
        return c.key === mapping.sourceKey;
      });
      status = !parsed || found ? 'Mapped' : 'Missing source';
    }
    var hidden = filterValue && col.label.toLowerCase().indexOf(filterValue) === -1 ? 'hidden' : '';

    return (
      '<tr data-ps-target-id="' +
      escapeAttr(col.id) +
      '" class="' +
      hidden +
      (ignored ? ' is-ignored' : '') +
      '">' +
      '<td class="col-standard">' +
      '<div class="standard-label">' +
      escapeHtml(col.label) +
      '</div>' +
      '<div class="standard-id">' +
      escapeHtml(col.id) +
      '</div>' +
      '</td>' +
      '<td>' +
      '<label class="sr-only" for="ps-map-' +
      escapeAttr(col.id) +
      '">Property Level Source for ' +
      escapeHtml(col.label) +
      '</label>' +
      '<select id="ps-map-' +
      escapeAttr(col.id) +
      '" data-ps-target-id="' +
      escapeAttr(col.id) +
      '" class="source-select"' +
      (ignored ? ' disabled' : '') +
      '>' +
      buildSourceOptionsHtml('property', mapping, parsed) +
      '</select>' +
      '</td>' +
      '<td class="col-ignore">' +
      '<label class="ignore-check">' +
      '<input type="checkbox" data-ps-ignore-target="' +
      escapeAttr(col.id) +
      '"' +
      (ignored ? ' checked' : '') +
      ' title="Excluded from generated schedule" />' +
      '<span class="sr-only">Ignore in output</span>' +
      '</label>' +
      '</td>' +
      '<td class="col-status">' +
      '<span class="status-pill ' +
      statusClassName(status) +
      '">' +
      escapeHtml(status) +
      '</span>' +
      '</td>' +
      '</tr>'
    );
  }).join('');

  tbody.innerHTML = rows;

  tbody.querySelectorAll('select.source-select').forEach(function (select) {
    select.addEventListener('change', function () {
      updateDraftPropertySummaryMapping(select.dataset.psTargetId, select.value);
      markDirty();
      renderPropertySummarySummary();
      updatePropertySummaryRowStatus(select.dataset.psTargetId);
    });
  });

  tbody.querySelectorAll('input[data-ps-ignore-target]').forEach(function (checkbox) {
    checkbox.addEventListener('change', function () {
      updateDraftPropertySummaryIgnore(checkbox.dataset.psIgnoreTarget, checkbox.checked);
      markDirty();
      renderPropertySummaryTable();
      renderPropertySummarySummary();
    });
  });
}

function updateDraftPropertySummaryIgnore(targetId, ignored) {
  ensureDraftPropertySummary();
  var entry = null;
  for (var i = 0; i < draftConfig.propertySummary.columnMappings.length; i += 1) {
    if (draftConfig.propertySummary.columnMappings[i].targetColumnId === targetId) {
      entry = draftConfig.propertySummary.columnMappings[i];
      break;
    }
  }
  if (!entry) {
    entry = { targetColumnId: targetId, propertyMapping: blankMapping(), ignore: false };
    draftConfig.propertySummary.columnMappings.push(entry);
  }
  entry.ignore = Boolean(ignored);
}

function updatePropertySummaryRowStatus(targetId) {
  if (!draftConfig || !draftConfig.propertySummary) return;
  var entry = null;
  for (var i = 0; i < draftConfig.propertySummary.columnMappings.length; i += 1) {
    if (draftConfig.propertySummary.columnMappings[i].targetColumnId === targetId) {
      entry = draftConfig.propertySummary.columnMappings[i];
      break;
    }
  }
  if (!entry) return;
  var parsed = getPropertyParsedForSettings();
  var mapping = entry.propertyMapping;
  var status = entry.ignore === true ? 'Ignored' : 'Unmapped';
  if (entry.ignore !== true && mapping && mapping.type === 'source') {
    var found = parsed && parsed.columns && parsed.columns.some(function (c) {
      return c.key === mapping.sourceKey;
    });
    status = !parsed || found ? 'Mapped' : 'Missing source';
  }
  var row = document.querySelector(
    '#ps-mapping-table-body tr[data-ps-target-id="' + cssEscape(targetId) + '"]'
  );
  if (!row) return;
  var pill = row.querySelector('.status-pill');
  if (pill) {
    pill.textContent = status;
    pill.className = 'status-pill ' + statusClassName(status);
  }
}

function filterPropertySummaryRows(query) {
  var q = query.trim().toLowerCase();
  document.querySelectorAll('#ps-mapping-table-body tr[data-ps-target-id]').forEach(function (row) {
    var label = (row.querySelector('.standard-label') &&
      row.querySelector('.standard-label').textContent) ||
      '';
    row.classList.toggle('hidden', Boolean(q) && label.toLowerCase().indexOf(q) === -1);
  });
}

function syncDraftFromDom() {
  if (!draftConfig) return;

  document.querySelectorAll('#mapping-table-body select.source-select').forEach(function (select) {
    updateDraftColumnMapping(select.dataset.targetId, select.dataset.mapSide, select.value);
  });

  document.querySelectorAll('#ps-mapping-table-body select.source-select').forEach(function (select) {
    updateDraftPropertySummaryMapping(select.dataset.psTargetId, select.value);
  });

  document.querySelectorAll('#mapping-table-body input[data-ignore-target]').forEach(function (checkbox) {
    updateDraftIgnore(checkbox.dataset.ignoreTarget, checkbox.checked);
  });
  document.querySelectorAll('#ps-mapping-table-body input[data-ps-ignore-target]').forEach(function (checkbox) {
    updateDraftPropertySummaryIgnore(checkbox.dataset.psIgnoreTarget, checkbox.checked);
  });

  var propSelect = document.getElementById('property-match-select');
  var unitSelect = document.getElementById('unit-match-select');
  if (propSelect) {
    draftConfig.propertyMatch = mappingFromSelectValue(
      'property',
      propSelect.value,
      draftConfig.propertyMatch
    );
  }
  if (unitSelect) {
    draftConfig.unitMatch = mappingFromSelectValue(
      'unit',
      unitSelect.value,
      draftConfig.unitMatch
    );
  }
}

export function refreshSettingsIfOpen() {
  var modal = document.getElementById('settings-modal');
  if (!modal || modal.classList.contains('hidden') || !draftConfig) return;

  syncDraftFromDom();
  applyDefaultMatchFields(draftConfig, APP_STATE.property, APP_STATE.unit);
  renderAll();
}

function statusClassName(status) {
  switch (status) {
    case 'Mapped':
      return 'status-mapped';
    case 'Partially mapped':
      return 'status-partial';
    case 'Missing source':
      return 'status-missing';
    case 'Ignored':
      return 'status-ignored';
    case 'Unmapped':
    default:
      return 'status-blank';
  }
}
