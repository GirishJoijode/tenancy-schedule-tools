import { applyDefaultMatchFields, isMappingConfigured } from '../config/mapping.js';
import { parseSourceFile } from '../excel/parser.js';
import { APP_STATE } from '../state.js';
import { escapeHtml, formatCount, qs, setStatusLines } from '../ui/common.js';
import { refreshSettingsIfOpen } from '../settings/settingsUI.js';
import { generateStandardisedSchedule } from './generator.js';
import { analyseMatching } from './transform.js';

function getFileStatusLines(parsed, singular, plural) {
  if (!parsed) return ['No file loaded'];
  var n = parsed.dataRowCount;
  return [
    '\u2713 ' + parsed.fileName,
    formatCount(n) + ' ' + (n === 1 ? singular : plural) + ' detected',
    formatCount(parsed.columns.length) + ' source columns detected',
  ];
}

function getMappingStatusLines() {
  var config = APP_STATE.mappingConfig;
  if (!config) return ['Not configured'];

  var hasFiles = Boolean(APP_STATE.property && APP_STATE.unit);

  if (!hasFiles) {
    if (
      (config.propertyMatch && config.propertyMatch.type === 'source') ||
      (config.unitMatch && config.unitMatch.type === 'source')
    ) {
      return ['Configuration loaded'];
    }
    var hasAny = config.columnMappings.some(function (e) {
      return (
        (e.propertyMapping && e.propertyMapping.type === 'source') ||
        (e.unitMapping && e.unitMapping.type === 'source')
      );
    });
    return [hasAny ? 'Configuration loaded' : 'Not configured'];
  }

  if (APP_STATE.mappingConfigured) {
    return ['\u2713 Configured'];
  }

  return ['\u26a0 Mapping requires attention'];
}

export function setupDropzone(dropzone) {
  var fileType = dropzone.dataset.fileType;
  var input = qs(dropzone, '[data-role="input"]');
  var browseBtn = qs(dropzone, '[data-role="browse"]');
  var idle = qs(dropzone, '[data-role="idle"]');
  var success = qs(dropzone, '[data-role="success"]');
  var errorEl = qs(dropzone, '[data-role="error"]');
  var filenameEl = qs(dropzone, '[data-role="filename"]');
  var columnsEl = qs(dropzone, '[data-role="columns"]');
  var rowsEl = qs(dropzone, '[data-role="rows"]');

  function showIdle() {
    idle.classList.remove('hidden');
    success.classList.add('hidden');
    errorEl.classList.add('hidden');
    dropzone.classList.remove('has-file', 'has-error');
  }

  function showError(message) {
    idle.classList.add('hidden');
    success.classList.add('hidden');
    errorEl.classList.remove('hidden');
    errorEl.textContent = message;
    dropzone.classList.remove('has-file');
    dropzone.classList.add('has-error');
  }

  function showSuccess(parsed) {
    idle.classList.add('hidden');
    errorEl.classList.add('hidden');
    success.classList.remove('hidden');
    filenameEl.textContent = '\u2713 ' + parsed.fileName;
    columnsEl.textContent =
      '\u2713 ' + formatCount(parsed.columns.length) + ' source columns detected';
    rowsEl.textContent =
      '\u2713 ' + formatCount(parsed.dataRowCount) + ' data rows detected';
    dropzone.classList.add('has-file');
    dropzone.classList.remove('has-error');
  }

  function handleFile(file) {
    if (!file) return;

    parseSourceFile(file, fileType)
      .then(function (parsed) {
        APP_STATE[fileType] = parsed;

        if (APP_STATE.mappingConfig) {
          applyDefaultMatchFields(APP_STATE.mappingConfig, APP_STATE.property, APP_STATE.unit);
        }

        showSuccess(parsed);
        clearGenerateResult();
        refreshStatus();
        refreshSettingsIfOpen();
      })
      .catch(function (err) {
        APP_STATE[fileType] = null;
        var message =
          err && typeof err.message === 'string' && err.message
            ? err.message
            : 'The file could not be processed. Please check it is a valid .xlsx export.';
        showError(message);
        clearGenerateResult();
        refreshStatus();
        refreshSettingsIfOpen();
        console.error('[' + fileType + '] parse failed:', err);
      });
  }

  function openPicker(event) {
    if (event) {
      event.stopPropagation();
    }
    // Do not preventDefault — some browsers suppress the file dialog if the initiating click is cancelled.
    if (!input) return;
    try {
      input.click();
    } catch (err) {
      console.error('File picker could not be opened:', err);
    }
  }

  if (browseBtn) {
    browseBtn.addEventListener('click', function (event) {
      event.stopPropagation();
      openPicker(event);
    });
  }

  dropzone.addEventListener('click', function (event) {
    if (browseBtn && (event.target === browseBtn || browseBtn.contains(event.target))) return;
    if (event.target && event.target.tagName === 'INPUT') return;
    openPicker(event);
  });

  dropzone.addEventListener('keydown', function (event) {
    if (event.key === 'Enter' || event.key === ' ') {
      openPicker(event);
    }
  });

  input.addEventListener('change', function () {
    var file = input.files && input.files[0];
    input.value = '';
    handleFile(file);
  });

  dropzone.addEventListener('dragenter', function (event) {
    event.preventDefault();
    dropzone.classList.add('drag-over');
  });

  dropzone.addEventListener('dragover', function (event) {
    event.preventDefault();
    dropzone.classList.add('drag-over');
  });

  dropzone.addEventListener('dragleave', function (event) {
    if (!dropzone.contains(event.relatedTarget)) {
      dropzone.classList.remove('drag-over');
    }
  });

  dropzone.addEventListener('drop', function (event) {
    event.preventDefault();
    dropzone.classList.remove('drag-over');
    var file =
      event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
    handleFile(file);
  });

  showIdle();
}

function clearGenerateResult() {
  var el = document.getElementById('generate-result');
  if (!el) return;
  el.classList.add('hidden');
  el.classList.remove('result-success', 'result-error', 'result-warning');
  el.innerHTML = '';
}

function renderMatchingSummaryPanel(analysis) {
  var el = document.getElementById('matching-summary');
  if (!el) return;

  if (!analysis || !analysis.summary) {
    el.classList.add('hidden');
    el.classList.remove('summary-ok', 'summary-warn');
    el.innerHTML = '';
    return;
  }

  var s = analysis.summary;
  var hasIssues =
    s.propertiesWithoutMatchedUnits > 0 ||
    s.unmatchedUnitRecords > 0 ||
    s.propertiesEmptyMatchValue > 0 ||
    s.unitsEmptyMatchValue > 0 ||
    s.duplicatePropertyKeys > 0;

  el.classList.remove('hidden', 'summary-ok', 'summary-warn');
  el.classList.add(hasIssues ? 'summary-warn' : 'summary-ok');

  var statusLine = hasIssues
    ? '\u26a0 Schedule will be generated with unmatched records'
    : '\u2713 Schedule can be generated';

  var secondary = [];
  if (s.propertiesEmptyMatchValue > 0) {
    secondary.push(
      s.propertiesEmptyMatchValue +
        ' Property row' +
        (s.propertiesEmptyMatchValue === 1 ? ' has' : 's have') +
        ' no match value'
    );
  }
  if (s.unitsEmptyMatchValue > 0) {
    secondary.push(
      s.unitsEmptyMatchValue +
        ' Unit row' +
        (s.unitsEmptyMatchValue === 1 ? ' has' : 's have') +
        ' no match value'
    );
  }
  if (s.duplicatePropertyKeys > 0) {
    secondary.push(
      s.duplicatePropertyKeys +
        ' duplicate Property match key' +
        (s.duplicatePropertyKeys === 1 ? '' : 's')
    );
  }

  el.innerHTML =
    '<h3>Matching Summary</h3>' +
    '<div class="summary-block">' +
    '<strong>Property Level</strong><br>' +
    formatCount(s.propertiesProcessed) +
    ' properties detected<br>' +
    formatCount(s.propertiesMatched) +
    ' properties matched to one or more TS records<br>' +
    formatCount(s.propertiesWithoutMatchedUnits) +
    ' properties with no matched TS records' +
    '</div>' +
    '<div class="summary-block">' +
    '<strong>Unit / TS Level</strong><br>' +
    formatCount(s.unitRecordsProcessed) +
    ' unit records detected<br>' +
    formatCount(s.matchedUnitRecords) +
    ' unit records matched<br>' +
    formatCount(s.unmatchedUnitRecords) +
    ' unit records unmatched' +
    '</div>' +
    '<p class="summary-status">' +
    statusLine +
    '</p>' +
    (secondary.length
      ? '<ul class="summary-secondary">' +
        secondary
          .map(function (line) {
            return '<li>' + escapeHtml(line) + '</li>';
          })
          .join('') +
        '</ul>'
      : '');
}

function showGenerateResult(result) {
  var el = document.getElementById('generate-result');
  if (!el) return;

  el.classList.remove('hidden', 'result-success', 'result-error', 'result-warning');

  var summary = result.summary || {
    propertiesProcessed: 0,
    propertiesMatched: 0,
    propertiesWithoutMatchedUnits: 0,
    unitRecordsProcessed: 0,
    matchedUnitRecords: 0,
    unmatchedUnitRecords: 0,
  };

  if (result.ok) {
    var hasMatchIssues =
      summary.propertiesWithoutMatchedUnits > 0 || summary.unmatchedUnitRecords > 0;
    el.classList.add(hasMatchIssues ? 'result-warning' : 'result-success');

    var warningHtml =
      result.warnings && result.warnings.length
        ? '<div class="result-warning"><p><strong>Notes</strong></p><ul>' +
          result.warnings
            .map(function (w) {
              return '<li>' + escapeHtml(w) + '</li>';
            })
            .join('') +
          '</ul></div>'
        : '';

    el.innerHTML =
      '<p class="result-title">Schedule generated successfully</p>' +
      '<ul class="result-stats">' +
      '<li>Properties written: <strong>' +
      formatCount(summary.propertiesProcessed) +
      '</strong></li>' +
      '<li>Properties with matched units: <strong>' +
      formatCount(summary.propertiesMatched || 0) +
      '</strong></li>' +
      '<li>Properties without matched units: <strong>' +
      formatCount(summary.propertiesWithoutMatchedUnits || 0) +
      '</strong></li>' +
      '<li>Unit records written: <strong>' +
      formatCount(summary.matchedUnitRecords) +
      '</strong></li>' +
      '<li>Unmatched unit records: <strong>' +
      formatCount(summary.unmatchedUnitRecords) +
      '</strong></li>' +
      '</ul>' +
      (result.filename
        ? '<p class="generate-hint">Downloaded ' + escapeHtml(result.filename) + '</p>'
        : '') +
      warningHtml;
    return;
  }

  el.classList.add('result-error');
  var errorHtml =
    result.errors && result.errors.length
      ? '<ul class="result-errors">' +
        result.errors
          .map(function (e) {
            return '<li>' + escapeHtml(e) + '</li>';
          })
          .join('') +
        '</ul>'
      : '<p>Generation could not be completed due to a technical error.</p>';

  el.innerHTML =
    '<p class="result-title">Generation could not complete</p>' + errorHtml;
}

export function setupGenerateButton() {
  var btn = document.getElementById('generate-btn');
  btn.addEventListener('click', function () {
    if (btn.disabled) return;

    clearGenerateResult();

    try {
      var result = generateStandardisedSchedule(
        APP_STATE.property,
        APP_STATE.unit,
        APP_STATE.mappingConfig
      );
      showGenerateResult(result);
    } catch (err) {
      console.error('Generate failed:', err);
      showGenerateResult({
        ok: false,
        errors: ['An unexpected error occurred while generating the schedule.'],
        warnings: [],
        summary: {
          propertiesProcessed: (APP_STATE.property && APP_STATE.property.dataRowCount) || 0,
          unitRecordsProcessed: (APP_STATE.unit && APP_STATE.unit.dataRowCount) || 0,
          matchedUnitRecords: 0,
          unmatchedUnitRecords: 0,
        },
      });
    }
  });
}

export function refreshStatus() {
  var propertyStatus = document.getElementById('property-status');
  var unitStatus = document.getElementById('unit-status');
  var mappingStatus = document.getElementById('mapping-status');
  var generateBtn = document.getElementById('generate-btn');
  var generateHint = document.getElementById('generate-hint');

  setStatusLines(
    propertyStatus,
    getFileStatusLines(APP_STATE.property, 'property', 'properties')
  );
  setStatusLines(unitStatus, getFileStatusLines(APP_STATE.unit, 'record', 'records'));

  APP_STATE.mappingConfigured = isMappingConfigured();
  setStatusLines(mappingStatus, getMappingStatusLines());

  // Incomplete matching must NOT disable Generate - only require both files parsed.
  var canGenerate = Boolean(APP_STATE.property && APP_STATE.unit);
  generateBtn.disabled = !canGenerate;

  if (canGenerate) {
    var analysis = analyseMatching(
      APP_STATE.property,
      APP_STATE.unit,
      APP_STATE.mappingConfig
    );
    renderMatchingSummaryPanel(analysis);
    generateHint.textContent =
      'Ready to generate. Incomplete matching will not block the schedule - unmatched records are reported separately.';
  } else {
    renderMatchingSummaryPanel(null);
    generateHint.textContent =
      'Upload Property Level and Unit / TS Level .xlsx files to enable generation.';
  }
}
