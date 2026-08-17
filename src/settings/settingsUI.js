import { APP_VERSION, CONFIG_VERSION, STANDARD_TEMPLATE_COLUMNS } from '../config/constants.js';
import { cloneDefaultConfig } from '../config/configManager.js';
import { exportConfigJson, importConfigJson } from '../config/configManager.js';
import {
  applyDefaultMatchFields,
  applyMappingConfig,
  blankMapping,
  cloneMappingConfig,
  getColumnMappingStatus,
  getMappingSummary,
  sourceMappingFromColumn,
} from '../config/mapping.js';
import { APP_STATE } from '../state.js';
import { cssEscape, escapeAttr, escapeHtml, openConfirm } from '../ui/common.js';

var draftConfig = null;
var onConfigChanged = null;

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

  document.getElementById('btn-apply-config').addEventListener('click', function () {
    syncDraftFromDom();
    applyMappingConfig(draftConfig);
    showFeedback('Configuration applied and saved in this browser.', 'success');
    notifyChanged();
    renderAll();
  });

  document.getElementById('btn-export-config').addEventListener('click', function () {
    syncDraftFromDom();
    applyMappingConfig(draftConfig);
    notifyChanged();
    var result = exportConfigJson(draftConfig);
    if (result.ok) {
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
      'This restores the original default column mappings and match fields that shipped with the application. Uploaded Excel files are not affected.',
      function () {
        draftConfig = cloneMappingConfig(cloneDefaultConfig());
        applyMappingConfig(draftConfig);
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
}

function notifyChanged() {
  if (typeof onConfigChanged === 'function') onConfigChanged();
}

function openSettings() {
  draftConfig = cloneMappingConfig(APP_STATE.mappingConfig || cloneDefaultConfig());
  applyDefaultMatchFields(draftConfig, APP_STATE.property, APP_STATE.unit);

  var modal = document.getElementById('settings-modal');
  modal.classList.remove('hidden');
  activateTab('mapping');
  clearFeedback();
  renderAll();
}

function closeSettings() {
  document.getElementById('settings-modal').classList.add('hidden');
  draftConfig = null;
}

function activateTab(tabId) {
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
    summary.propertyMappings +
    '</strong> Property mappings</span>' +
    '<span><strong>' +
    summary.unitMappings +
    '</strong> Unit mappings</span>' +
    '<span class="' +
    (summary.missingMappings ? 'summary-missing' : '') +
    '">' +
    '<strong>' +
    summary.missingMappings +
    '</strong> Missing mappings</span>';
}

function buildSourceOptionsHtml(fileType, selectedMapping) {
  var parsed = APP_STATE[fileType];
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
      };
    }

    var status = getColumnMappingStatus(entry, APP_STATE.property, APP_STATE.unit);
    var statusClass = statusClassName(status);
    var hidden = filterValue && col.label.toLowerCase().indexOf(filterValue) === -1 ? 'hidden' : '';

    return (
      '<tr data-target-id="' +
      escapeAttr(col.id) +
      '" class="' +
      hidden +
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
      '" class="source-select">' +
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
      '" class="source-select">' +
      buildSourceOptionsHtml('unit', entry.unitMapping) +
      '</select>' +
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
      renderMappingSummary();
      updateRowStatus(targetId);
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
    };
    draftConfig.columnMappings.push(entry);
  }

  var field = side === 'property' ? 'propertyMapping' : 'unitMapping';
  entry[field] = mappingFromSelectValue(side, sourceKey, entry[field]);
}

function mappingFromSelectValue(fileType, sourceKey, previous) {
  if (!sourceKey) return blankMapping();

  var cols = (APP_STATE[fileType] && APP_STATE[fileType].columns) || [];
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
  };
  unitSelect.onchange = function () {
    draftConfig.unitMatch = mappingFromSelectValue(
      'unit',
      unitSelect.value,
      draftConfig.unitMatch
    );
  };
}

function renderConfigTab() {
  document.getElementById('config-version-value').textContent = String(
    draftConfig.configVersion || CONFIG_VERSION
  );
  document.getElementById('app-version-value').textContent = APP_VERSION;
}

function syncDraftFromDom() {
  if (!draftConfig) return;

  document.querySelectorAll('#mapping-table-body select.source-select').forEach(function (select) {
    updateDraftColumnMapping(select.dataset.targetId, select.dataset.mapSide, select.value);
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
    default:
      return 'status-blank';
  }
}
