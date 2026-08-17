import { APP_TOOLS, DEFAULT_APP_TOOL_MODE } from '../config/appTools.js';
import { escapeHtml } from './common.js';

export function setAppMode(mode) {
  var app = document.getElementById('app');
  var stdView = document.getElementById('view-standardiser');
  var cmpView = document.getElementById('view-compare');
  var stdBtn = document.getElementById('mode-standardiser-btn');
  var cmpBtn = document.getElementById('mode-compare-btn');
  var isCompare = mode === 'compare';

  app.classList.toggle('app-mode-compare', isCompare);
  app.classList.toggle('app-mode-standardiser', !isCompare);
  stdView.classList.toggle('hidden', isCompare);
  cmpView.classList.toggle('hidden', !isCompare);
  if (stdBtn) {
    stdBtn.classList.toggle('active', !isCompare);
    stdBtn.setAttribute('aria-selected', !isCompare ? 'true' : 'false');
    stdBtn.setAttribute('tabindex', isCompare ? '-1' : '0');
  }
  if (cmpBtn) {
    cmpBtn.classList.toggle('active', isCompare);
    cmpBtn.setAttribute('aria-selected', isCompare ? 'true' : 'false');
    cmpBtn.setAttribute('tabindex', isCompare ? '0' : '-1');
  }
  document.title = 'Tenancy Schedule Tools';
}

function renderToolNav() {
  var nav = document.getElementById('app-tools-nav');
  if (!nav) return;

  nav.innerHTML = APP_TOOLS.map(function (tool) {
    return (
      '<button type="button" class="app-tool-btn" role="tab" id="' +
      escapeHtml(tool.buttonId) +
      '" data-app-mode="' +
      escapeHtml(tool.mode) +
      '" aria-controls="' +
      escapeHtml(tool.viewId) +
      '">' +
      '<span class="app-tool-label">' +
      escapeHtml(tool.label) +
      '</span>' +
      '<span class="app-tool-description">' +
      escapeHtml(tool.description) +
      '</span>' +
      '</button>'
    );
  }).join('');
}

export function setupAppModeSwitcher() {
  renderToolNav();

  APP_TOOLS.forEach(function (tool) {
    var btn = document.getElementById(tool.buttonId);
    if (!btn) return;
    btn.addEventListener('click', function () {
      setAppMode(tool.mode);
    });
  });

  var nav = document.getElementById('app-tools-nav');
  if (nav) {
    nav.addEventListener('keydown', function (event) {
      if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
      var buttons = Array.prototype.slice.call(nav.querySelectorAll('.app-tool-btn'));
      var index = buttons.indexOf(document.activeElement);
      if (index < 0) return;
      var next =
        event.key === 'ArrowRight'
          ? (index + 1) % buttons.length
          : (index - 1 + buttons.length) % buttons.length;
      buttons[next].focus();
      buttons[next].click();
      event.preventDefault();
    });
  }

  setAppMode(DEFAULT_APP_TOOL_MODE);
}
