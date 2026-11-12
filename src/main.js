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
import { setupReportFields } from './standardiser/reportFields.js';
import { setupUploadGuides } from './standardiser/uploadGuide.js';
import { setupPortfolioBreakdown } from './ui/portfolioBreakdown.js';

function init() {
  document.querySelectorAll('#schedule-tenancy .dropzone').forEach(setupDropzone);
  setupScheduleTypeSwitcher();
  setupPortfolioBreakdown();
  setupReportFields();
  setupUploadGuides();
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
