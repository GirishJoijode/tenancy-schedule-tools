import './styles.css';
import { initialiseMappingState } from './config/mapping.js';
import { setupCompareTool } from './compare/ui.js';
import { setupSettingsUI } from './settings/settingsUI.js';
import { setupDropzone, setupGenerateButton, refreshStatus } from './standardiser/ui.js';
import { setupAppModeSwitcher } from './ui/appSwitcher.js';

function init() {
  initialiseMappingState();
  document.querySelectorAll('#view-standardiser .dropzone').forEach(setupDropzone);
  setupAppModeSwitcher();
  setupCompareTool();
  setupSettingsUI({ onConfigChanged: refreshStatus });
  setupGenerateButton();
  refreshStatus();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
