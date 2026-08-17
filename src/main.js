import './styles.css';
import { initialiseMappingState } from './config/mapping.js';
import { setupCompareTool } from './compare/ui.js';
import {
  setupPropertySummaryTool,
  refreshPropertySummaryStatus,
} from './propertySummary/propertySummaryUI.js';
import { setupScheduleTypeSwitcher } from './ui/scheduleTypeSelect.js';
import { setupSettingsUI } from './settings/settingsUI.js';
import { setupDropzone, setupGenerateButton, refreshStatus } from './standardiser/ui.js';
import { setupAppModeSwitcher } from './ui/appSwitcher.js';

function refreshAllStatus() {
  refreshStatus();
  refreshPropertySummaryStatus();
}

function init() {
  initialiseMappingState();
  document.querySelectorAll('#schedule-tenancy .dropzone').forEach(setupDropzone);
  setupScheduleTypeSwitcher();
  setupPropertySummaryTool();
  setupAppModeSwitcher();
  setupCompareTool();
  setupSettingsUI({ onConfigChanged: refreshAllStatus });
  setupGenerateButton();
  refreshAllStatus();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
