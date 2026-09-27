import './styles.css';
import { setupCompareTool } from './compare/ui.js';
import {
  setupPropertySummaryTool,
  refreshPropertySummaryStatus,
} from './propertySummary/propertySummaryUI.js';
import { setupScheduleTypeSwitcher } from './ui/scheduleTypeSelect.js';
import { setupDropzone, setupGenerateButton, refreshStatus } from './standardiser/ui.js';
import { setupAppModeSwitcher } from './ui/appSwitcher.js';
import { setupConfigurationTool } from './configTool/ui.js';

function init() {
  document.querySelectorAll('#schedule-tenancy .dropzone').forEach(setupDropzone);
  setupScheduleTypeSwitcher();
  setupPropertySummaryTool();
  setupConfigurationTool();
  setupAppModeSwitcher();
  setupCompareTool();
  setupGenerateButton();
  refreshStatus();
  refreshPropertySummaryStatus();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
