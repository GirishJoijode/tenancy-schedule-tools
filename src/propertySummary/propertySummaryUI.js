import { getPropertySummaryMappingSummary } from '../config/propertySummaryColumns.js';
import { parseSourceFile } from '../excel/parser.js';
import { APP_STATE, PROPERTY_SUMMARY_STATE } from '../state.js';
import { escapeHtml, formatCount, qs, setStatusLines } from '../ui/common.js';
import { refreshSettingsIfOpen } from '../settings/settingsUI.js';
import { generatePropertySummary } from './propertySummary.js';

function getFileStatusLines(parsed) {
  if (!parsed) return ['No file loaded'];
  return [
    '\u2713 ' + parsed.fileName,
    '\u2713 Header row detected',
    '\u2713 ' + formatCount(parsed.columns.length) + ' source columns detected',
    '\u2713 ' + formatCount(parsed.dataRowCount) + ' data rows detected',
  ];
}

function getConfigStatusLines() {
  var config = APP_STATE.mappingConfig && APP_STATE.mappingConfig.propertySummary;
  var summary = getPropertySummaryMappingSummary(config, PROPERTY_SUMMARY_STATE.property);
  var lines = [
    summary.propertyMappings +
      ' of ' +
      summary.standardColumns +
      ' Property Only fields mapped',
  ];
  if (summary.unmappedColumns > 0) {
    lines.push(
      summary.unmappedColumns +
        ' field' +
        (summary.unmappedColumns === 1 ? '' : 's') +
        ' unmapped (included as blank columns)'
    );
  }
  if (summary.ignoredColumns > 0) {
    lines.push(
      summary.ignoredColumns +
        ' field' +
        (summary.ignoredColumns === 1 ? '' : 's') +
        ' ignored (omitted from output)'
    );
  }
  if (summary.missingMappings > 0) {
    lines.push('\u26a0 ' + summary.missingMappings + ' mapped field' + (summary.missingMappings === 1 ? '' : 's') + ' missing from the uploaded file');
  }
  return lines;
}

function clearGenerateResult() {
  var el = document.getElementById('ps-generate-result');
  if (!el) return;
  el.classList.add('hidden');
  el.classList.remove('result-success', 'result-error', 'result-warning');
  el.innerHTML = '';
}

function showGenerateResult(result) {
  var el = document.getElementById('ps-generate-result');
  if (!el) return;

  el.classList.remove('hidden', 'result-success', 'result-error', 'result-warning');
  var summary = result.summary || { propertiesProcessed: 0, columnsWritten: 0 };

  if (result.ok) {
    el.classList.add(result.warnings && result.warnings.length ? 'result-warning' : 'result-success');
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
      '<p class="result-title">Property Only generated successfully</p>' +
      '<ul class="result-stats">' +
      '<li>Properties written: <strong>' +
      formatCount(summary.propertiesProcessed) +
      '</strong></li>' +
      '<li>Columns written: <strong>' +
      formatCount(summary.columnsWritten) +
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
  el.innerHTML = '<p class="result-title">Generation could not complete</p>' + errorHtml;
}

export function refreshPropertySummaryStatus() {
  var propertyStatus = document.getElementById('ps-property-status');
  var mappingStatus = document.getElementById('ps-mapping-status');
  var generateBtn = document.getElementById('ps-generate-btn');
  var generateHint = document.getElementById('ps-generate-hint');

  setStatusLines(propertyStatus, getFileStatusLines(PROPERTY_SUMMARY_STATE.property));
  setStatusLines(mappingStatus, getConfigStatusLines());

  var canGenerate = Boolean(PROPERTY_SUMMARY_STATE.property);
  if (generateBtn) generateBtn.disabled = !canGenerate;

  if (generateHint) {
    generateHint.textContent = canGenerate
      ? 'Ready to generate. Unmapped fields stay in the workbook as blank columns. Ignored fields are omitted.'
      : 'Upload a Property Level .xlsx file to enable generation.';
  }
}

function setupPropertySummaryDropzone(dropzone) {
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
    parseSourceFile(file, 'property')
      .then(function (parsed) {
        PROPERTY_SUMMARY_STATE.property = parsed;
        showSuccess(parsed);
        clearGenerateResult();
        refreshPropertySummaryStatus();
        refreshSettingsIfOpen();
      })
      .catch(function (err) {
        PROPERTY_SUMMARY_STATE.property = null;
        showError(
          err && err.message
            ? err.message
            : 'The file could not be processed. Please check it is a valid .xlsx export.'
        );
        clearGenerateResult();
        refreshPropertySummaryStatus();
        refreshSettingsIfOpen();
        console.error('[propertySummary] parse failed:', err);
      });
  }

  function openPicker(event) {
    if (event) event.stopPropagation();
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
    if (event.key === 'Enter' || event.key === ' ') openPicker(event);
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
    if (!dropzone.contains(event.relatedTarget)) dropzone.classList.remove('drag-over');
  });
  dropzone.addEventListener('drop', function (event) {
    event.preventDefault();
    dropzone.classList.remove('drag-over');
    var file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
    handleFile(file);
  });

  showIdle();
}

export function setupPropertySummaryTool() {
  var dropzone = document.getElementById('ps-property-dropzone');
  if (dropzone) setupPropertySummaryDropzone(dropzone);

  var btn = document.getElementById('ps-generate-btn');
  if (btn) {
    btn.addEventListener('click', function () {
      if (btn.disabled) return;
      clearGenerateResult();
      try {
        var result = generatePropertySummary(
          PROPERTY_SUMMARY_STATE.property,
          APP_STATE.mappingConfig && APP_STATE.mappingConfig.propertySummary
        );
        showGenerateResult(result);
      } catch (err) {
        console.error('Property Summary generate failed:', err);
        showGenerateResult({
          ok: false,
          errors: ['An unexpected error occurred while generating the Property Only schedule.'],
          warnings: [],
          summary: { propertiesProcessed: 0, columnsWritten: 0 },
        });
      }
    });
  }

  refreshPropertySummaryStatus();
}
