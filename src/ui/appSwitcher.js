import { APP_TOOLS, DEFAULT_APP_TOOL_MODE } from '../config/appTools.js';
import { escapeHtml } from './common.js';

export function setAppMode(mode) {
  var app = document.getElementById('app');
  APP_TOOLS.forEach(function (tool) {
    var active = tool.mode === mode;
    var view = document.getElementById(tool.viewId);
    var button = document.getElementById(tool.buttonId);
    if (view) view.classList.toggle('hidden', !active);
    if (app) app.classList.toggle('app-mode-' + tool.mode, active);
    if (button) {
      button.classList.toggle('active', active);
      button.setAttribute('aria-selected', active ? 'true' : 'false');
      button.setAttribute('tabindex', active ? '0' : '-1');
    }
  });
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
