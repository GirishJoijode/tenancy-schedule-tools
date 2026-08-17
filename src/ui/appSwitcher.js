export function setAppMode(mode) {
  var app = document.getElementById('app');
  var stdView = document.getElementById('view-standardiser');
  var cmpView = document.getElementById('view-compare');
  var title = document.getElementById('app-title');
  var subtitle = document.getElementById('app-subtitle');
  var stdBtn = document.getElementById('mode-standardiser-btn');
  var cmpBtn = document.getElementById('mode-compare-btn');
  var isCompare = mode === 'compare';

  app.classList.toggle('app-mode-compare', isCompare);
  app.classList.toggle('app-mode-standardiser', !isCompare);
  stdView.classList.toggle('hidden', isCompare);
  cmpView.classList.toggle('hidden', !isCompare);
  stdBtn.classList.toggle('active', !isCompare);
  cmpBtn.classList.toggle('active', isCompare);

  if (isCompare) {
    title.textContent = 'Tenancy Schedule Compare';
    subtitle.textContent =
      'Compare a current tenancy schedule against a previous version and identify unit-level changes.';
    document.title = 'Tenancy Schedule Compare';
  } else {
    title.textContent = 'Tenancy Schedule Standardiser';
    subtitle.textContent =
      'Convert Property and Unit level exports into the standard tenancy schedule.';
    document.title = 'Tenancy Schedule Standardiser';
  }
}

export function setupAppModeSwitcher() {
  document.getElementById('mode-standardiser-btn').addEventListener('click', function () {
    setAppMode('standardiser');
  });
  document.getElementById('mode-compare-btn').addEventListener('click', function () {
    setAppMode('compare');
  });
  setAppMode('standardiser');
}
