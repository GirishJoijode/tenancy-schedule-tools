import { COMPARE_OUTPUT_FILENAME } from '../config/constants.js';
import { parseSourceFile } from '../excel/parser.js';
import { normaliseJoinValue } from '../excel/parser.js';
import { buildCompareWorkbook, downloadWorkbook } from '../excel/writer.js';
import { COMPARE_STATE } from '../state.js';
import { escapeHtml, formatCount, qs } from '../ui/common.js';
import { formatCompareDisplay, runTenancyCompare } from './engine.js';

function setupCompareDropzone(dropzone) {
  var role = dropzone.dataset.compareRole;
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
    parseSourceFile(file, 'unit')
      .then(function (parsed) {
        COMPARE_STATE[role] = parsed;
        COMPARE_STATE.result = null;
        showSuccess(parsed);
        refreshCompareUi();
      })
      .catch(function (err) {
        COMPARE_STATE[role] = null;
        COMPARE_STATE.result = null;
        showError(
          err && err.message
            ? err.message
            : 'The file could not be processed. Please check it is a valid .xlsx export.'
        );
        refreshCompareUi();
        console.error('[compare:' + role + '] parse failed:', err);
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

function refreshCompareUi() {
  var compareBtn = document.getElementById('compare-btn');
  var downloadBtn = document.getElementById('compare-download-btn');
  var hint = document.getElementById('compare-hint');
  var errEl = document.getElementById('compare-error');
  var canCompare = Boolean(COMPARE_STATE.current && COMPARE_STATE.previous);
  compareBtn.disabled = !canCompare;
  downloadBtn.disabled = !(COMPARE_STATE.result && COMPARE_STATE.result.ok);

  if (!COMPARE_STATE.current && !COMPARE_STATE.previous) {
    hint.textContent = 'Upload Current and Previous Unit / TS .xlsx files to compare.';
  } else if (!canCompare) {
    hint.textContent = 'Upload both Current and Previous TS files to enable comparison.';
  } else if (!COMPARE_STATE.result) {
    hint.textContent =
      'Ready to compare the fixed Default unit fields.';
  }

  if (!COMPARE_STATE.result) {
    document.getElementById('compare-results').classList.add('hidden');
    errEl.classList.add('hidden');
    errEl.textContent = '';
  }
}

function renderCompareSummary(summary) {
  var el = document.getElementById('compare-summary');
  var cards = [
    ['Current TS', summary.currentUnits, ''],
    ['Previous TS', summary.previousUnits, ''],
    ['Matched', summary.matched, ''],
    ['Changed', summary.changed, 'stat-changed'],
    ['Unchanged', summary.unchanged, ''],
    ['New Units', summary.newUnits, 'stat-new'],
    ['Removed Units', summary.removedUnits, 'stat-removed'],
    ['Ambiguous', summary.ambiguous, 'stat-ambiguous'],
  ];
  el.innerHTML = cards
    .map(function (c) {
      return (
        '<div class="compare-stat ' +
        c[2] +
        '"><span class="label">' +
        escapeHtml(c[0]) +
        '</span><span class="value">' +
        formatCount(c[1]) +
        '</span></div>'
      );
    })
    .join('');
}

function pairMatchesFilter(pair, filter, search) {
  if (filter && filter !== 'all') {
    if (filter === 'changed' && pair.status !== 'changed') return false;
    if (filter === 'unchanged' && pair.status !== 'unchanged') return false;
    if (filter === 'new' && pair.status !== 'new') return false;
    if (filter === 'removed' && pair.status !== 'removed') return false;
    if (filter === 'ambiguous' && pair.status !== 'ambiguous') return false;
  }
  if (search) {
    var q = normaliseJoinValue(search);
    var hay = [
      pair.propertyName,
      pair.unitName,
      pair.current && pair.current.tenantName,
      pair.previous && pair.previous.tenantName,
      pair.changedLabels && pair.changedLabels.join(' '),
    ]
      .map(function (v) {
        return normaliseJoinValue(v);
      })
      .join(' ');
    if (hay.indexOf(q) === -1) return false;
  }
  return true;
}

function statusBadge(status) {
  var label = status.toUpperCase();
  return '<span class="badge badge-' + status + '">' + escapeHtml(label) + '</span>';
}

function renderCompareTable() {
  var result = COMPARE_STATE.result;
  if (!result || !result.ok) return;

  var head = document.getElementById('compare-table-head');
  var body = document.getElementById('compare-table-body');
  var cols = result.columns;

  var stickyHeaders =
    '<th class="sticky-col sticky-0">Status</th>' +
    '<th class="sticky-col sticky-1">Property</th>' +
    '<th class="sticky-col sticky-2">Unit</th>' +
    '<th class="sticky-col sticky-3">Version</th>' +
    '<th>Changes</th>';
  for (var i = 0; i < cols.length; i += 1) {
    stickyHeaders += '<th>' + escapeHtml(cols[i].label) + '</th>';
  }
  head.innerHTML = '<tr>' + stickyHeaders + '</tr>';

  var filter = COMPARE_STATE.filter || 'changed';
  var search = COMPARE_STATE.search || '';
  var htmlRows = '';
  var lastPropertyKey = null;

  for (var p = 0; p < result.pairs.length; p += 1) {
    var pair = result.pairs[p];
    if (!pairMatchesFilter(pair, filter, search)) continue;

    if (pair.propertyKey !== lastPropertyKey) {
      lastPropertyKey = pair.propertyKey;
      var pname =
        pair.propertyName == null || pair.propertyName === ''
          ? '(Blank Property)'
          : String(pair.propertyName);
      htmlRows +=
        '<tr class="compare-property-band"><td colspan="' +
        (5 + cols.length) +
        '">' +
        escapeHtml(pname) +
        '</td></tr>';
    }

    var changesText =
      pair.status === 'changed'
        ? pair.changedLabels.join(', ')
        : pair.status === 'new'
          ? 'New unit'
          : pair.status === 'removed'
            ? 'Removed unit'
            : pair.status === 'ambiguous'
              ? 'Ambiguous match key'
              : '';

    var rowHtml = function (version, record, isLastOfPair) {
      var trClass =
        'version-' +
        version +
        (pair.status === 'new' || pair.status === 'removed' || pair.status === 'ambiguous'
          ? ' status-' + pair.status
          : '') +
        (isLastOfPair ? ' pair-sep' : '');
      var versionBadge =
        version === 'current'
          ? '<span class="badge badge-current">CURRENT</span>'
          : '<span class="badge badge-previous">PREVIOUS</span>';

      var cells =
        '<td class="sticky-col sticky-0">' +
        (version === 'current' ? statusBadge(pair.status) : '') +
        '</td>' +
        '<td class="sticky-col sticky-1">' +
        escapeHtml(pair.propertyName == null ? '' : String(pair.propertyName)) +
        '</td>' +
        '<td class="sticky-col sticky-2">' +
        escapeHtml(pair.unitName == null ? '' : String(pair.unitName)) +
        '</td>' +
        '<td class="sticky-col sticky-3">' +
        versionBadge +
        '</td>' +
        '<td>' +
        (version === 'current' ? escapeHtml(changesText) : '') +
        '</td>';

      for (var c = 0; c < cols.length; c += 1) {
        var col = cols[c];
        var changed =
          pair.status === 'changed' && pair.changedFields.indexOf(col.id) !== -1;
        var raw = record ? record.values[col.id] : null;
        var display;
        if (!record) {
          display =
            '<span class="compare-empty">' +
            (version === 'current'
              ? 'Not present in current TS'
              : 'Not present in previous TS') +
            '</span>';
        } else {
          display = escapeHtml(formatCompareDisplay(raw, col.id));
        }
        cells +=
          '<td class="' +
          (changed ? 'cell-changed' : '') +
          '">' +
          display +
          '</td>';
      }
      return '<tr class="' + trClass + '">' + cells + '</tr>';
    };

    htmlRows += rowHtml('current', pair.current, false);
    htmlRows += rowHtml('previous', pair.previous, true);
  }

  body.innerHTML =
    htmlRows ||
    '<tr><td colspan="' +
      (5 + cols.length) +
      '" class="compare-empty">No units match the current filter/search.</td></tr>';
}

function showCompareResult(result) {
  var errEl = document.getElementById('compare-error');
  var resultsEl = document.getElementById('compare-results');
  if (!result.ok) {
    resultsEl.classList.add('hidden');
    errEl.classList.remove('hidden');
    errEl.textContent = (result.errors || []).join(' ');
    document.getElementById('compare-download-btn').disabled = true;
    return;
  }
  errEl.classList.add('hidden');
  errEl.textContent = '';
  resultsEl.classList.remove('hidden');
  renderCompareSummary(result.summary);
  renderCompareTable();
  document.getElementById('compare-download-btn').disabled = false;
  if (result.warnings && result.warnings.length) {
    document.getElementById('compare-hint').textContent = result.warnings.join(' ');
  } else {
    document.getElementById('compare-hint').textContent =
      'Comparison complete. Default filter shows Changed units.';
  }
}

function downloadCompareWorkbook(result) {
  if (!result || !result.ok) return;
  var wb = buildCompareWorkbook(result);
  downloadWorkbook(wb, COMPARE_OUTPUT_FILENAME);
}

function clearComparison() {
  COMPARE_STATE.current = null;
  COMPARE_STATE.previous = null;
  COMPARE_STATE.result = null;
  COMPARE_STATE.filter = 'changed';
  COMPARE_STATE.search = '';

  document.querySelectorAll('#view-compare .dropzone').forEach(function (dz) {
    dz.classList.remove('has-file', 'has-error', 'drag-over');
    var idle = dz.querySelector('[data-role="idle"]');
    var success = dz.querySelector('[data-role="success"]');
    var error = dz.querySelector('[data-role="error"]');
    if (idle) idle.classList.remove('hidden');
    if (success) success.classList.add('hidden');
    if (error) {
      error.classList.add('hidden');
      error.textContent = '';
    }
    var input = dz.querySelector('[data-role="input"]');
    if (input) input.value = '';
  });

  document.getElementById('compare-results').classList.add('hidden');
  document.getElementById('compare-error').classList.add('hidden');
  document.getElementById('compare-error').textContent = '';
  document.getElementById('compare-search').value = '';
  document.querySelectorAll('.compare-filter-btn').forEach(function (btn) {
    btn.classList.toggle('active', btn.dataset.compareFilter === 'changed');
  });
  refreshCompareUi();
}

export function setupCompareTool() {
  document.querySelectorAll('#view-compare .dropzone').forEach(setupCompareDropzone);

  document.getElementById('compare-btn').addEventListener('click', function () {
    var result = runTenancyCompare(COMPARE_STATE.current, COMPARE_STATE.previous);
    COMPARE_STATE.result = result;
    showCompareResult(result);
    refreshCompareUi();
  });

  document.getElementById('compare-download-btn').addEventListener('click', function () {
    if (COMPARE_STATE.result && COMPARE_STATE.result.ok) {
      downloadCompareWorkbook(COMPARE_STATE.result);
    }
  });

  document.getElementById('compare-clear-btn').addEventListener('click', clearComparison);

  document.querySelectorAll('.compare-filter-btn').forEach(function (btn) {
    btn.addEventListener('click', function () {
      COMPARE_STATE.filter = btn.dataset.compareFilter;
      document.querySelectorAll('.compare-filter-btn').forEach(function (b) {
        b.classList.toggle('active', b === btn);
      });
      if (COMPARE_STATE.result && COMPARE_STATE.result.ok) renderCompareTable();
    });
  });

  document.getElementById('compare-search').addEventListener('input', function (event) {
    COMPARE_STATE.search = event.target.value || '';
    if (COMPARE_STATE.result && COMPARE_STATE.result.ok) renderCompareTable();
  });

  refreshCompareUi();
}
