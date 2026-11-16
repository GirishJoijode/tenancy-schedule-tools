import { APP_STATE } from '../state.js';
import { escapeHtml } from './common.js';
import { buildPortfolio } from '../model/portfolio.js';
import { buildPortfolioGroups, getBreakdownMode, getCustomConfig } from '../logic/portfolioBreakdown.js';
import {
  AM_DEFAULT_HORIZON,
  AM_HORIZONS,
  addCalendarMonths,
  assetQueryGroupOptions,
  canRunAssetQueries,
  deriveAssetManagementView,
  findingsForRegister,
  runAssetManagementReview,
} from '../logic/assetQueries.js';

var DISABLED_REASON = 'Upload both current Property and Unit files to enable Valuation Review.';
var previousFocus = null;
var state = emptyState();

function emptyState() {
  return {
    horizon: AM_DEFAULT_HORIZON,
    eventType: 'all',
    propertyId: '',
    groupId: '',
    measure: 'count',
    viewMode: 'graph',
    granularity: 'month',
    selection: null,
    expandedPropertyId: '',
    drawerId: '',
    review: null,
    groupOptions: [],
  };
}

function workspace() {
  return document.getElementById('am-workspace');
}

export function refreshAssetQueriesButton() {
  var button = document.getElementById('am-queries-btn');
  if (!button) return;
  var enabled = canRunAssetQueries(APP_STATE.property, APP_STATE.unit);
  button.disabled = !enabled;
  if (enabled) button.removeAttribute('title');
  else button.title = DISABLED_REASON;
}

function currentGroups(properties) {
  var mode = getBreakdownMode();
  if (!properties || !properties.length || mode === 'none') return { effective: false, groups: [] };
  return buildPortfolioGroups(properties, {
    mode: mode,
    propertyData: APP_STATE.property,
    customConfig: getCustomConfig(),
  });
}

function analyse() {
  var portfolio = buildPortfolio(APP_STATE.property, APP_STATE.unit);
  var grouping = currentGroups(portfolio.properties);
  state.groupOptions = assetQueryGroupOptions(grouping);
  if (!state.groupOptions.some(function (group) { return group.id === state.groupId; })) state.groupId = '';
  state.review = runAssetManagementReview({
    properties: portfolio.properties,
    horizonMonths: state.horizon,
    groups: grouping.effective ? grouping.groups : [],
  });
  if (state.propertyId && !state.review.properties.some(function (property) { return property.propertyId === state.propertyId; })) {
    state.propertyId = '';
  }
  state.valuations = portfolio.properties.map(function (property, index) {
    var listed = state.review.properties[index] || {};
    var raw = property.standardValues && property.standardValues.valuation;
    var valuation = raw instanceof Date && !Number.isNaN(raw.getTime())
      ? new Date(raw.getFullYear(), raw.getMonth(), raw.getDate())
      : null;
    return { propertyId: listed.propertyId, groupId: listed.groupId, valuation: valuation };
  });
}

function view() {
  return deriveAssetManagementView(state.review, {
    propertyId: state.propertyId,
    groupId: state.groupId,
    eventType: state.eventType,
  });
}

function formatRent(value) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '';
  var rounded = Math.round(value);
  return '£' + rounded.toLocaleString('en-GB');
}

var EVENT_ROWS = [
  { id: 'rent_review', label: 'Rent Reviews' },
  { id: 'break', label: 'Break Options' },
  { id: 'expiry', label: 'Lease Expiries' },
];
var SHORT_MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
var FULL_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

function seriesLabel(eventType) {
  for (var i = 0; i < EVENT_ROWS.length; i += 1) {
    if (EVENT_ROWS[i].id === eventType) return EVENT_ROWS[i].label;
  }
  return 'Events';
}

function selectOptions(items, selected) {
  return items.map(function (item) {
    return '<option value="' + escapeHtml(item.id) + '"' + (item.id === selected ? ' selected' : '') + '>' + escapeHtml(item.label) + '</option>';
  }).join('');
}

function renderChrome() {
  var review = state.review || { properties: [] };
  var counts = document.getElementById('am-portfolio-count');
  if (counts) {
    var properties = review.properties.length;
    var leases = review.properties.reduce(function (sum, property) { return sum + property.leases; }, 0);
    counts.textContent = properties + (properties === 1 ? ' property' : ' properties') + ' \u2022 ' + leases + (leases === 1 ? ' unit' : ' units');
  }
  var filters = document.getElementById('am-filters');
  if (!filters) return;
  var visibleProperties = review.properties.filter(function (property) {
    return !state.groupId || property.groupId === state.groupId;
  });
  var propertyOptions = [{ id: '', label: 'All Properties' }].concat(visibleProperties.map(function (property) {
    var label = property.externalId ? property.externalId + ' \u2014 ' + property.propertyName : property.propertyName;
    return { id: property.propertyId, label: label };
  }));
  var groupField = state.groupOptions.length
    ? '<label class="am-filter am-filter-breakdown">Breakdown<select id="am-group" data-am-filter="group">' +
      selectOptions([{ id: '', label: 'All Portfolio' }].concat(state.groupOptions), state.groupId) +
      '</select></label>'
    : '';
  filters.innerHTML =
    '<label class="am-filter am-filter-portfolio">Portfolio<select id="am-property" data-am-filter="property">' + selectOptions(propertyOptions, state.propertyId) + '</select></label>' +
    groupField +
    '<label class="am-filter am-filter-horizon">Event Horizon<select id="am-horizon" data-am-filter="horizon">' +
    selectOptions(AM_HORIZONS.map(function (months) { return { id: String(months), label: months + ' Months' }; }), String(state.horizon)) +
    '</select></label>' +
    '<label class="am-filter am-filter-event">Event Type<select id="am-event" data-am-filter="event">' +
    selectOptions([
      { id: 'all', label: 'All Events' },
      { id: 'rent_review', label: 'Rent Reviews' },
      { id: 'break', label: 'Breaks' },
      { id: 'expiry', label: 'Expiries' },
    ], state.eventType) +
    '</select></label>';
}

function kpiCard(label, value, title) {
  var tip = title ? ' title="' + escapeHtml(title) + '"' : '';
  return '<div class="am-kpi"' + tip + '><span>' + escapeHtml(label) + '</span><strong>' + escapeHtml(String(value)) + '</strong></div>';
}

function formatCompactRent(value) {
  if (typeof value !== 'number' || !Number.isFinite(value) || value === 0) return '';
  var sign = value < 0 ? '-' : '';
  var abs = Math.abs(value);
  if (abs >= 1000000) {
    var millions = abs / 1000000;
    var text = millions >= 10 ? String(Math.round(millions)) : millions.toFixed(1).replace(/\.0$/, '');
    return sign + '\u00a3' + text + 'm';
  }
  if (abs >= 1000) return sign + '\u00a3' + Math.round(abs / 1000).toLocaleString('en-GB') + 'k';
  return formatRent(value);
}

function intensity(count) {
  if (count >= 7) return 4;
  if (count >= 4) return 3;
  if (count >= 2) return 2;
  if (count === 1) return 1;
  return 0;
}

function fullMonth(key) {
  var parts = String(key || '').split('-');
  var month = Number(parts[1]) - 1;
  if (!FULL_MONTHS[month]) return key;
  return FULL_MONTHS[month] + ' ' + parts[0];
}

var BEFORE_KEY = 'before-valuation';
var BEFORE_LABEL = 'Before Valuation Date';
var BEFORE_WARNING = 'Events recorded before the valuation date require review.';
var BEFORE_EVENT_WARNING = 'Event recorded before the valuation date and may require review.';

function beforeWarning(message) {
  return '<span class="am-before-warn" role="img" aria-label="Before valuation date — review required" title="' + escapeHtml(message || BEFORE_WARNING) + '">' +
    '<svg viewBox="0 0 12 12" width="12" height="12" aria-hidden="true" focusable="false">' +
    '<path fill="currentColor" d="M6 .7 11.3 11H.7L6 .7z"/>' +
    '<path fill="#fff" d="M5.45 4.55h1.1v3.15h-1.1zm0 3.7h1.1V9.4h-1.1z"/>' +
    '</svg></span>';
}

function valuationMetas() {
  return (state.valuations || []).filter(function (item) {
    if (state.propertyId && item.propertyId !== state.propertyId) return false;
    if (state.groupId && item.groupId !== state.groupId) return false;
    return item.valuation instanceof Date;
  });
}

function beforeValuation(finding) {
  return finding.eventDate instanceof Date && finding.valuationDate instanceof Date && finding.eventDate < finding.valuationDate;
}

function forwardMonths() {
  var metas = valuationMetas();
  if (!metas.length) return [];
  var start = null;
  var end = null;
  metas.forEach(function (item) {
    var from = new Date(item.valuation.getFullYear(), item.valuation.getMonth(), 1);
    var to = addCalendarMonths(item.valuation, state.horizon);
    if (!start || from < start) start = from;
    if (!end || to > end) end = to;
  });
  var months = [];
  var cursor = new Date(start.getFullYear(), start.getMonth(), 1);
  var last = new Date(end.getFullYear(), end.getMonth(), 1);
  var guard = 0;
  while (cursor <= last && guard < 240) {
    var key = cursor.getFullYear() + '-' + String(cursor.getMonth() + 1).padStart(2, '0');
    months.push({
      key: key,
      label: SHORT_MONTHS[cursor.getMonth()] + ' ' + String(cursor.getFullYear()).slice(2),
      full: FULL_MONTHS[cursor.getMonth()] + ' ' + cursor.getFullYear(),
      historic: false,
    });
    cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
    guard += 1;
  }
  return months;
}

var SERIES_KEY = { rent_review: 'rentReviews', break: 'breakOptions', expiry: 'leaseExpiries' };

function emptySeries() {
  return { count: 0, contractedRent: 0, findings: [] };
}

function seriesOf(period, eventType) {
  return period[SERIES_KEY[eventType]] || emptySeries();
}

function beforePeriod() {
  return { key: BEFORE_KEY, label: BEFORE_LABEL, full: BEFORE_LABEL, historic: true };
}

function periodAxis() {
  var months = forwardMonths();
  if (state.granularity !== 'year') return [beforePeriod()].concat(months);
  var years = [];
  var seen = {};
  months.forEach(function (month) {
    var year = month.key.slice(0, 4);
    if (seen[year]) return;
    seen[year] = true;
    years.push({ key: year, label: year, full: year, historic: false });
  });
  return [beforePeriod()].concat(years);
}

function aggregateLeaseEvents(options) {
  var findings = options.findings || [];
  var granularity = options.granularity === 'year' ? 'year' : 'month';
  var axis = options.axis || [];
  var periods = axis.map(function (period) {
    return {
      periodKey: period.key,
      periodLabel: period.label,
      periodFull: period.full || period.label,
      historic: !!period.historic,
      rentReviews: emptySeries(),
      breakOptions: emptySeries(),
      leaseExpiries: emptySeries(),
    };
  });
  var index = {};
  periods.forEach(function (period) { index[period.periodKey] = period; });
  findings.forEach(function (finding) {
    var bucketName = SERIES_KEY[finding.eventType];
    if (!bucketName) return;
    var key = beforeValuation(finding)
      ? BEFORE_KEY
      : (granularity === 'year' ? String(finding.monthKey || '').slice(0, 4) : finding.monthKey);
    var period = index[key];
    if (!period) return;
    var series = period[bucketName];
    series.count += 1;
    if (typeof finding.contractedRent === 'number') series.contractedRent += finding.contractedRent;
    series.findings.push(finding);
  });
  return periods;
}

function leasePeriods(dashboard) {
  return aggregateLeaseEvents({
    findings: dashboard.findings,
    granularity: state.granularity,
    axis: periodAxis(),
  });
}

function isPeriodSelected(eventType, periodKey) {
  var selection = state.selection;
  return !!(selection && selection.kind === 'period' && selection.eventType === eventType && selection.periodKey === periodKey && selection.granularity === state.granularity);
}

function tooltipRent(value) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '';
  if (value === 0) return '\u00a30 contracted rent';
  var text = Math.abs(value) >= 1000000 ? formatCompactRent(value) : formatRent(value);
  return text ? text + ' contracted rent' : '';
}

function formatAxisRent(value) {
  if (!value) return '\u00a30';
  var sign = value < 0 ? '-' : '';
  var abs = Math.abs(value);
  if (abs >= 1000000) return sign + '\u00a3' + (abs / 1000000).toFixed(1) + 'm';
  if (abs >= 1000) return sign + '\u00a3' + Math.round(abs / 1000).toLocaleString('en-GB') + 'k';
  return formatRent(value);
}

function niceCeiling(value) {
  if (value <= 0) return 1;
  var pow = Math.pow(10, Math.floor(Math.log10(value)));
  var fraction = value / pow;
  var nice = fraction <= 1 ? 1 : fraction <= 1.5 ? 1.5 : fraction <= 2 ? 2 : fraction <= 2.5 ? 2.5 : fraction <= 5 ? 5 : 10;
  return nice * pow;
}

function chartScale(periods) {
  var max = 0;
  periods.forEach(function (period) {
    var stack = 0;
    EVENT_ROWS.forEach(function (row) {
      var series = seriesOf(period, row.id);
      stack += state.measure === 'rent' ? series.contractedRent : series.count;
    });
    if (stack > max) max = stack;
  });
  if (state.measure === 'rent') {
    var topRent = max > 0 ? niceCeiling(max * 1.18) : 1000;
    if (topRent <= max) topRent = niceCeiling(max * 1.3);
    var rentStep = topRent >= 1000000 && (topRent / 500000) % 1 === 0 ? 500000 : topRent / 4;
    if (topRent / rentStep > 6) rentStep = topRent / 4;
    var rentTicks = [];
    for (var rentTick = topRent; rentTick > rentStep * 0.5; rentTick -= rentStep) rentTicks.push(Math.round(rentTick));
    rentTicks.push(0);
    return { top: topRent, ticks: rentTicks };
  }
  var raw = Math.max(4, Math.ceil((max || 0) * 1.18));
  var step = raw <= 4 ? 1 : raw <= 10 ? 2 : raw <= 20 ? 5 : Math.ceil(raw / 4);
  var top = Math.ceil(raw / step) * step;
  if (top <= max) top += step;
  var ticks = [];
  for (var value = top; value >= 0; value -= step) ticks.push(value);
  return { top: top, ticks: ticks };
}

function formatTick(value) {
  if (state.measure === 'rent') return formatAxisRent(value);
  return String(Math.round(value));
}

function periodTip(row, period, series) {
  var leases = series.count + (series.count === 1 ? ' lease' : ' leases');
  var rentText = tooltipRent(series.contractedRent);
  return row.label + '\n' + period.periodFull + '\n\n' + leases + (rentText ? '\n' + rentText : '') + '\n\nClick to view leases';
}

function segmentControl(label, attribute, options, current) {
  var buttons = options.map(function (option) {
    var pressed = option.id === current;
    return '<button type="button" data-am-' + attribute + '="' + option.id + '" aria-pressed="' + (pressed ? 'true' : 'false') + '">' + escapeHtml(option.label) + '</button>';
  }).join('');
  return '<div class="am-seg-field"><span>' + escapeHtml(label) + '</span><div class="am-segment" role="group" aria-label="' + escapeHtml(label) + '">' + buttons + '</div></div>';
}

function leaseControls() {
  return '<div class="am-panel-tools">' +
    segmentControl('View', 'view', [{ id: 'graph', label: 'Graph' }, { id: 'matrix', label: 'Matrix' }], state.viewMode) +
    segmentControl('Period', 'granularity', [{ id: 'month', label: 'Monthly' }, { id: 'year', label: 'Yearly' }], state.granularity) +
    '<label class="am-filter am-measure">Display by<select id="am-measure" data-am-filter="measure">' +
    selectOptions([{ id: 'count', label: 'Event Count' }, { id: 'rent', label: 'Contracted Rent' }], state.measure) +
    '</select></label></div>';
}

function renderGraph(periods) {
  var scale = chartScale(periods);
  var yearly = state.granularity === 'year';
  var focused = state.selection && state.selection.kind === 'period';
  var horizonStart = false;
  var ticks = scale.ticks.map(function (tick) {
    return '<span>' + escapeHtml(formatTick(tick)) + '</span>';
  }).join('');
  var lines = scale.ticks.map(function () { return '<i></i>'; }).join('');
  var stacks = [];
  var labels = [];
  periods.forEach(function (period) {
    var marker = period.historic ? ' is-historic' : '';
    if (!period.historic && !horizonStart) {
      horizonStart = true;
      marker += ' is-horizon-start';
    }
    var slices = [];
    var total = 0;
    var used = 0;
    EVENT_ROWS.forEach(function (row) {
      var series = seriesOf(period, row.id);
      if (!series.count) return;
      var value = state.measure === 'rent' ? series.contractedRent : series.count;
      var height = scale.top ? (value / scale.top) * 100 : 0;
      if (value > 0 && height < 1.5) height = 1.5;
      used += height;
      total += series.count;
      var selected = isPeriodSelected(row.id, period.periodKey);
      var leases = series.count + (series.count === 1 ? ' lease' : ' leases');
      slices.push({
        row: row,
        height: height,
        selected: selected,
        tip: periodTip(row, period, series),
        label: row.label + ', ' + period.periodFull + ', ' + leases,
      });
    });
    if (used > 82) {
      slices.forEach(function (slice) { slice.height = slice.height * 82 / used; });
    }
    var buttons = slices.map(function (slice, index) {
      var edge = (index === 0 ? ' is-base' : '') + (index === slices.length - 1 ? ' is-cap' : '');
      return '<button type="button" class="am-slice am-seg-' + slice.row.id + edge + (slice.selected ? ' is-selected' : '') + '" style="height:' + slice.height.toFixed(2) + '%" data-am-segment="' + slice.row.id + '" data-am-period="' + escapeHtml(period.periodKey) + '" data-am-period-label="' + escapeHtml(period.periodFull) + '" title="' + escapeHtml(slice.tip) + '" aria-label="' + escapeHtml(slice.label) + '"></button>';
    }).join('');
    var totalLabel = yearly && state.measure === 'count' && total ? '<span class="am-stack-total">' + total + '</span>' : '';
    var beforeClass = period.periodKey === BEFORE_KEY ? ' is-before' : '';
    stacks.push('<div class="am-chart-group' + marker + beforeClass + '"><div class="am-stack-slot">' + totalLabel + '<div class="am-stack">' + buttons + '</div></div></div>');
    var tickLabel = period.periodKey === BEFORE_KEY
      ? beforeWarning() + '<span>Before<br>Valuation Date</span>'
      : escapeHtml(period.periodLabel);
    labels.push('<span class="am-chart-tick' + marker + beforeClass + '">' + tickLabel + '</span>');
  });
  var legend = EVENT_ROWS.map(function (row) {
    return '<li><i class="am-swatch am-seg-' + row.id + '"></i>' + escapeHtml(row.label) + '</li>';
  }).join('');
  var mode = yearly ? 'is-year' : 'is-month';
  var trackWidth = 24;
  if (!yearly) {
    periods.forEach(function (period) { trackWidth += period.periodKey === BEFORE_KEY ? 148 : 52; });
  }
  var track = yearly ? '' : ' style="width:' + trackWidth + 'px"';
  return '<div class="am-chart ' + mode + '"><div class="am-chart-frame"><div class="am-y-axis" aria-hidden="true">' + ticks + '</div><div class="am-plot-view"><div class="am-plot ' + mode + (focused ? ' has-period' : '') + '"' + track + '><div class="am-gridlines" aria-hidden="true">' + lines + '</div>' + stacks.join('') + '</div></div><div class="am-x-axis"><div class="am-x-labels ' + mode + '"' + track + '>' + labels.join('') + '</div></div><div class="am-chart-scrollbar" hidden><div></div></div></div><ul class="am-legend">' + legend + '</ul></div>';
}

function layoutChartScroll(root, saved) {
  var view = root.querySelector('.am-plot-view');
  var axis = root.querySelector('.am-x-axis');
  var bar = root.querySelector('.am-chart-scrollbar');
  var plot = view && view.querySelector('.am-plot');
  if (!view || !axis || !bar || !plot) return;
  var needs = plot.scrollWidth > view.clientWidth + 1;
  bar.hidden = !needs;
  if (!needs) {
    view.scrollLeft = 0;
    axis.scrollLeft = 0;
    return;
  }
  bar.firstElementChild.style.width = plot.scrollWidth + 'px';
  var left = saved || 0;
  bar.scrollLeft = left;
  view.scrollLeft = left;
  axis.scrollLeft = left;
  bar.onscroll = function () {
    view.scrollLeft = bar.scrollLeft;
    axis.scrollLeft = bar.scrollLeft;
  };
}

function renderMatrixTable(periods) {
  var horizonStart = false;
  var headings = periods.map(function (period) {
    var marker = '';
    if (!period.historic && !horizonStart) {
      horizonStart = true;
      marker = ' is-horizon-start';
    }
    var beforeClass = period.periodKey === BEFORE_KEY ? ' is-before' : '';
    var heading = period.periodKey === BEFORE_KEY
      ? beforeWarning() + '<span>' + escapeHtml(period.periodLabel) + '</span>'
      : escapeHtml(period.periodLabel);
    return '<th class="' + (period.historic ? 'is-historic' : 'is-upcoming') + marker + beforeClass + '" scope="col">' + heading + '</th>';
  }).join('');
  var body = EVENT_ROWS.map(function (row) {
    var total = periods.reduce(function (sum, period) { return sum + seriesOf(period, row.id).count; }, 0);
    var cells = periods.map(function (period) {
      var series = seriesOf(period, row.id);
      var beforeClass = period.periodKey === BEFORE_KEY ? ' is-before' : '';
      if (!series.count) return '<td class="am-empty-cell' + (period.historic ? ' is-historic' : '') + beforeClass + '">\u2014</td>';
      var selected = isPeriodSelected(row.id, period.periodKey);
      var leases = series.count + (series.count === 1 ? ' lease' : ' leases');
      var figure = state.measure === 'rent' ? (formatCompactRent(series.contractedRent) || '\u00a30') : String(series.count);
      var secondary = state.measure === 'rent'
        ? leases
        : (series.contractedRent ? formatCompactRent(series.contractedRent) : '');
      return '<td class="' + (period.historic ? 'is-historic' : '') + beforeClass + '"><button type="button" class="am-cell am-cell-' + row.id + ' am-intensity-' + intensity(series.count) + (selected ? ' is-selected' : '') + '" data-am-segment="' + row.id + '" data-am-period="' + escapeHtml(period.periodKey) + '" data-am-period-label="' + escapeHtml(period.periodFull) + '" title="' + escapeHtml(periodTip(row, period, series)) + '" aria-label="' + escapeHtml(row.label + ', ' + period.periodFull + ', ' + leases) + '"><strong>' + escapeHtml(figure) + '</strong>' + (secondary ? '<small>' + escapeHtml(secondary) + '</small>' : '') + '</button></td>';
    }).join('');
    return '<tr><th scope="row" class="am-matrix-label"><span>' + escapeHtml(row.label) + '</span><strong>' + total + '</strong></th>' + cells + '</tr>';
  }).join('');
  var summary = EVENT_ROWS.map(function (row) {
    var total = periods.reduce(function (sum, period) { return sum + seriesOf(period, row.id).count; }, 0);
    return row.label + ' ' + total;
  }).join('. ');
  return '<div class="am-matrix-scroll"><table class="am-matrix"><caption class="visually-hidden">' + escapeHtml(summary) + '</caption><thead><tr><th class="am-matrix-label" scope="col">Event</th>' + headings + '</tr></thead><tbody>' + body + '</tbody></table></div>';
}

function renderLeaseEvents(dashboard) {
  var periods = leasePeriods(dashboard);
  var visual = state.viewMode === 'matrix' ? renderMatrixTable(periods) : renderGraph(periods);
  return '<section class="am-panel am-lease-panel"><div class="am-panel-head am-lease-head"><div><h3>Lease Events</h3><p>Rent reviews, break options and expiries relative to each property\u2019s valuation date.</p></div>' +
    leaseControls() + '</div>' + visual + '</section>';
}

function renderCategories(dashboard) {
  var rows = dashboard.categoryCounts;
  var max = rows.reduce(function (peak, row) { return Math.max(peak, state.measure === 'rent' ? row.rent : row.count); }, 0) || 1;
  var bars = rows.map(function (row) {
    var amount = state.measure === 'rent' ? row.rent : row.count;
    var width = amount ? Math.max(4, Math.round((amount / max) * 100)) : 0;
    var selected = state.selection && state.selection.kind === 'category' && state.selection.eventType === row.id;
    var value = state.measure === 'rent' ? formatRent(row.rent) || '£0' : String(row.count);
    return '<button type="button" class="am-bar' + (selected ? ' is-selected' : '') + '" data-am-category="' + row.id + '"><span>' + escapeHtml(row.label) + '</span><span class="am-bar-track"><span class="am-bar-fill am-seg-' + row.id + '" style="width:' + width + '%"></span></span><strong>' + escapeHtml(value) + '</strong></button>';
  }).join('');
  return '<section class="am-panel"><div class="am-panel-head"><div><h3>Queries by Event Type</h3><p>Click a category to list those leases.</p></div></div><div class="am-bars">' + bars + '</div></section>';
}

function eventTypeLabel(eventType) {
  if (eventType === 'rent_review') return 'Rent Review';
  if (eventType === 'break') return 'Break Option';
  if (eventType === 'expiry') return 'Lease Expiry';
  return seriesLabel(eventType);
}

function reviewMark() {
  return '<span class="am-review">' + beforeWarning(BEFORE_EVENT_WARNING) + '<span>Review</span></span>';
}

function assetChevron() {
  return '<svg class="am-chevron" viewBox="0 0 12 12" width="12" height="12" aria-hidden="true" focusable="false"><path d="M4.2 2.3 8 6 4.2 9.7" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg>';
}

function assetFindings(dashboard, propertyId) {
  return dashboard.findings.filter(function (finding) {
    return finding.propertyId === propertyId;
  }).sort(function (left, right) {
    var leftBefore = beforeValuation(left) ? 0 : 1;
    var rightBefore = beforeValuation(right) ? 0 : 1;
    if (leftBefore !== rightBefore) return leftBefore - rightBefore;
    var date = left.eventDate - right.eventDate;
    if (date) return date;
    var type = String(left.eventType).localeCompare(String(right.eventType));
    if (type) return type;
    var unit = String(left.unitName || '').localeCompare(String(right.unitName || ''));
    if (unit) return unit;
    return String(left.id).localeCompare(String(right.id));
  });
}

function renderAssetEvent(finding) {
  var rent = typeof finding.contractedRent === 'number' ? formatRent(finding.contractedRent) : '\u2014';
  var selected = state.drawerId === finding.id ? ' is-selected' : '';
  var review = beforeValuation(finding) ? '<span class="am-asset-event-status">' + reviewMark() + '</span>' : '';
  return '<button type="button" class="am-asset-event' + selected + '" data-am-finding="' + escapeHtml(finding.id) + '">' +
    '<i class="am-event-mark am-seg-' + escapeHtml(finding.eventType) + '" aria-hidden="true"></i>' +
    '<span class="am-asset-event-type"><span>' + escapeHtml(eventTypeLabel(finding.eventType)) + '</span></span>' +
    '<span class="am-asset-event-who"><strong>' + escapeHtml(finding.unitName) + '</strong><span>' + escapeHtml(finding.tenantName || 'No tenant') + '</span></span>' +
    '<span class="am-asset-event-when"><strong>' + escapeHtml(finding.eventLabel) + '</strong><span>' + escapeHtml(finding.timing) + '</span></span>' +
    '<span class="am-asset-event-rent">' + escapeHtml(rent) + '</span>' +
    review +
    '</button>';
}

function renderAttention(dashboard) {
  if (!dashboard.propertyAttention.length) {
    return '<section class="am-panel"><div class="am-panel-head"><div><h3>Assets Requiring Attention</h3></div></div><p class="am-empty-copy">No assets currently have queries in this selection.</p></section>';
  }
  var clear = state.propertyId ? '<button type="button" class="am-text-btn" data-am-clear-property>Clear property filter</button>' : '';
  var items = dashboard.propertyAttention.map(function (asset) {
    var open = state.expandedPropertyId === asset.propertyId;
    var identity = asset.externalId ? asset.externalId + ' \u2014 ' + asset.propertyName : asset.propertyName;
    var panel = open
      ? '<div class="am-asset-panel"><div class="am-asset-events">' + assetFindings(dashboard, asset.propertyId).map(renderAssetEvent).join('') + '</div></div>'
      : '';
    var events = asset.total + (asset.total === 1 ? ' event' : ' events');
    var review = asset.review ? '<em>' + asset.review + (asset.review === 1 ? ' requires review' : ' require review') + '</em>' : '';
    return '<div class="am-asset-block' + (open ? ' is-open' : '') + '">' +
      '<button type="button" class="am-asset' + (open ? ' is-open' : '') + '" data-am-property="' + escapeHtml(asset.propertyId) + '" aria-expanded="' + (open ? 'true' : 'false') + '">' +
      '<span class="am-asset-id"><strong>' + escapeHtml(asset.propertyName) + '</strong><small>' + escapeHtml(asset.externalId || 'No external ID') + '</small></span>' +
      '<span class="am-asset-counts">' + events + review + '</span>' +
      assetChevron() +
      '<span class="visually-hidden">' + escapeHtml(identity) + '</span></button>' +
      panel + '</div>';
  }).join('');
  return '<section class="am-panel"><div class="am-panel-head"><div><h3>Assets Requiring Attention</h3><p>Properties with the highest number of identified lease events.</p></div>' + clear + '</div><div class="am-assets">' + items + '</div></section>';
}

function selectedFindings(dashboard) {
  var selection = state.selection;
  if (!selection) return [];
  if (selection.kind === 'all') return dashboard.findings;
  if (selection.kind === 'property') {
    return dashboard.findings.filter(function (finding) { return finding.propertyId === selection.propertyId; });
  }
  if (selection.kind === 'category') {
    return findingsForRegister(dashboard.findings, { eventType: selection.eventType, monthKey: '' });
  }
  if (selection.kind !== 'period' || selection.granularity !== state.granularity) return [];
  var period = leasePeriods(dashboard).filter(function (item) { return item.periodKey === selection.periodKey; })[0];
  var series = period && period[SERIES_KEY[selection.eventType]];
  return series ? series.findings : [];
}

function selectionHeading(selection) {
  if (!selection) return '';
  if (selection.kind === 'all') return 'All queries';
  if (selection.kind === 'property') return selection.propertyName + ' \u2014 All Queries';
  if (selection.kind === 'category') return seriesLabel(selection.eventType);
  return seriesLabel(selection.eventType) + ' \u2014 ' + (selection.periodLabel || fullMonth(selection.periodKey));
}

function renderEventRow(finding) {
  var rent = typeof finding.contractedRent === 'number' ? formatRent(finding.contractedRent) : '\u2014';
  var review = finding.status === 'historic' ? '<span class="am-event-status">' + reviewMark() + '</span>' : '<span class="am-event-status"></span>';
  return '<button type="button" class="am-event' + (state.drawerId === finding.id ? ' is-selected' : '') + '" data-am-finding="' + escapeHtml(finding.id) + '">' +
    '<span class="am-event-who"><strong>' + escapeHtml(finding.propertyName) + '</strong><span>' + escapeHtml(finding.externalId || '') + '</span></span>' +
    '<span class="am-event-unit"><strong>' + escapeHtml(finding.unitName) + '</strong><span>' + escapeHtml(finding.tenantName || 'No tenant') + '</span></span>' +
    '<span class="am-event-when"><strong>' + escapeHtml(finding.eventLabel) + '</strong><span>' + escapeHtml(finding.timing) + '</span></span>' +
    '<span class="am-event-rent">' + escapeHtml(rent) + '</span>' +
    review +
    '</button>';
}

function renderSelected(dashboard) {
  var selection = state.selection;
  if (!selection) {
    return '<section id="am-selected" class="am-panel am-selected is-idle"><p>Select a lease-event bar, cell, or category to view the underlying leases.</p><button type="button" class="am-text-btn" data-am-view-all>View all queries</button></section>';
  }
  var rows = selectedFindings(dashboard);
  var countLabel = rows.length + (rows.length === 1 ? ' lease event' : ' lease events');
  var list = rows.length
    ? '<div class="am-events">' + rows.map(renderEventRow).join('') + '</div>'
    : '<p class="am-empty-copy">No leases remain in this selection.</p>';
  return '<section id="am-selected" class="am-panel am-selected"><div class="am-panel-head"><div><h3>' + escapeHtml(selectionHeading(selection)) + '</h3><p>' + escapeHtml(countLabel) + '</p></div><button type="button" class="am-text-btn" data-am-clear-selection>Clear selection</button></div>' + list + '</section>';
}

function renderBody() {
  var body = document.getElementById('am-body');
  if (!body || !state.review) return;
  var heldScroll = body.scrollTop;
  var scroller = body.querySelector('.am-matrix-scroll');
  var matrixScroll = scroller ? scroller.scrollLeft : 0;
  var plot = body.querySelector('.am-chart-scrollbar');
  var plotScroll = plot && !plot.hidden ? plot.scrollLeft : 0;
  var dashboard = view();
  var kpis = '<div class="am-kpis">' +
    kpiCard('Properties reviewed', dashboard.summary.propertiesReviewed) +
    kpiCard('Leases reviewed', dashboard.summary.leasesReviewed) +
    kpiCard('Potential queries', dashboard.summary.queries) +
    kpiCard('Upcoming events', dashboard.summary.upcoming) +
    kpiCard('Historic events', dashboard.summary.historic, 'Lease events recorded before the valuation date.') +
    '</div>';
  if (!dashboard.findings.length) {
    state.selection = null;
    state.expandedPropertyId = '';
    state.drawerId = '';
    body.innerHTML = kpis + '<section class="am-panel am-empty"><h3>No lease events were identified for this selection.</h3><p>Historic events stay included. Adjust the portfolio, event type, or event horizon.</p></section>';
    body.scrollTop = heldScroll;
    renderDrawer(null);
    return;
  }
  if (state.expandedPropertyId && !dashboard.propertyAttention.some(function (asset) { return asset.propertyId === state.expandedPropertyId; })) {
    state.expandedPropertyId = '';
  }
  var drawerFinding = dashboard.findings.filter(function (finding) { return finding.id === state.drawerId; })[0] || null;
  if (state.drawerId && !drawerFinding) state.drawerId = '';
  body.innerHTML = kpis + '<div class="am-visuals">' + renderLeaseEvents(dashboard) + renderCategories(dashboard) + '</div>' + renderSelected(dashboard) + renderAttention(dashboard);
  body.scrollTop = heldScroll;
  var nextScroller = body.querySelector('.am-matrix-scroll');
  if (nextScroller) nextScroller.scrollLeft = matrixScroll;
  layoutChartScroll(body, plotScroll);
  renderDrawer(drawerFinding);
}

function moneyText(value) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '\u2014';
  return formatRent(value);
}

function monthsText(value) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '\u2014';
  var months = Math.round(value);
  return months + (Math.abs(months) === 1 ? ' month' : ' months');
}

function eventTypeName(eventType) {
  if (eventType === 'rent_review') return 'Rent Review';
  if (eventType === 'break') return 'Break Option';
  if (eventType === 'expiry') return 'Lease Expiry';
  return 'Lease event';
}

function drawerSection(title, rows) {
  return '<section class="am-drawer-section"><h4>' + escapeHtml(title) + '</h4>' + rows.map(function (row) {
    return detailRow(row[0], row[1]);
  }).join('') + '</section>';
}

function eventDetails(finding) {
  if (finding.eventType === 'rent_review') {
    return drawerSection('Rent Review Details', [
      ['Review date', finding.eventLabel],
      ['Contracted rent', moneyText(finding.contractedRent)],
      ['Current rent', moneyText(finding.currentRent)],
      ['Updated / reviewed rent', moneyText(finding.reviewedRent)],
      ['Market rent', moneyText(finding.marketRent)],
    ]);
  }
  if (finding.eventType === 'break') {
    return drawerSection('Break Details', [
      ['Break date', finding.eventLabel],
      ['Contracted rent', moneyText(finding.contractedRent)],
      ['Current rent', moneyText(finding.currentRent)],
      ['Market rent', moneyText(finding.marketRent)],
      ['Rent free', monthsText(finding.rentFreeMonths)],
      ['Void period', monthsText(finding.breakVoidMonths)],
    ]);
  }
  return drawerSection('Lease Expiry Details', [
    ['Expiry date', finding.eventLabel],
    ['Contracted rent', moneyText(finding.contractedRent)],
    ['Current rent', moneyText(finding.currentRent)],
    ['Market rent', moneyText(finding.marketRent)],
    ['Rent free', monthsText(finding.rentFreeMonths)],
    ['Expiry void', monthsText(finding.expiryVoidMonths)],
  ]);
}

function renderDrawer(finding) {
  var drawer = document.getElementById('am-drawer');
  if (!drawer) return;
  if (!finding) {
    drawer.hidden = true;
    drawer.innerHTML = '';
    return;
  }
  drawer.hidden = false;
  drawer.innerHTML =
    '<div class="am-drawer-head"><h3>Lease Event Detail</h3><button type="button" class="am-drawer-close" data-am-close-drawer aria-label="Close lease event detail">\u00d7</button></div>' +
    detailRow('Property', finding.propertyName) +
    detailRow('External ID', finding.externalId || '\u2014') +
    detailRow('Unit', finding.unitName) +
    detailRow('Tenant', finding.tenantName || '\u2014') +
    detailRow('Event', eventTypeName(finding.eventType)) +
    detailRow('Event date', finding.eventLabel) +
    detailRow('Valuation date', finding.valuationLabel) +
    detailRow('Timing', finding.timing) +
    (finding.status === 'historic' ? detailRow('Status', 'Review') : '') +
    eventDetails(finding) +
    '<section class="am-drawer-section"><h4>Source</h4><p class="am-source">' + escapeHtml(finding.source) + '</p></section>';
}

function detailRow(label, value) {
  return '<p class="am-detail"><span>' + escapeHtml(label) + '</span><strong>' + escapeHtml(value) + '</strong></p>';
}

function refresh() {
  analyse();
  renderChrome();
  renderBody();
}

function openWorkspace() {
  if (!canRunAssetQueries(APP_STATE.property, APP_STATE.unit)) return;
  state = emptyState();
  var root = workspace();
  if (!root) return;
  previousFocus = document.activeElement;
  refresh();
  root.hidden = false;
  document.body.classList.add('am-open');
  var close = document.getElementById('am-close');
  if (close) close.focus();
}

function closeWorkspace() {
  var root = workspace();
  if (!root || root.hidden) return;
  root.hidden = true;
  document.body.classList.remove('am-open');
  state.drawerId = '';
  if (previousFocus && previousFocus.focus) previousFocus.focus();
}

function closeDrawer() {
  state.drawerId = '';
  renderBody();
}

function onFilterChange(event) {
  var filter = event.target.getAttribute('data-am-filter');
  if (!filter) return;
  if (filter !== 'measure') {
    state.selection = null;
    state.drawerId = '';
  }
  if (filter === 'horizon') {
    state.horizon = Number(event.target.value) || AM_DEFAULT_HORIZON;
    refresh();
    return;
  }
  if (filter === 'property') state.propertyId = event.target.value;
  if (filter === 'group') {
    state.groupId = event.target.value;
    if (state.propertyId) {
      var property = (state.review.properties || []).filter(function (item) { return item.propertyId === state.propertyId; })[0];
      if (!property || (state.groupId && property.groupId !== state.groupId)) state.propertyId = '';
    }
  }
  if (filter === 'event') state.eventType = event.target.value;
  if (filter === 'measure') state.measure = event.target.value;
  if (filter === 'group' || filter === 'property') renderChrome();
  renderBody();
}

function setGranularity(next) {
  if (next !== 'year' && next !== 'month') return;
  if (next === state.granularity) return;
  var selection = state.selection;
  if (selection && selection.kind === 'period') {
    if (selection.periodKey === BEFORE_KEY) {
      state.selection = {
        kind: 'period',
        eventType: selection.eventType,
        periodKey: BEFORE_KEY,
        periodLabel: BEFORE_LABEL,
        granularity: next,
      };
    } else if (next === 'year') {
      var year = String(selection.periodKey || '').slice(0, 4);
      state.selection = year ? {
        kind: 'period',
        eventType: selection.eventType,
        periodKey: year,
        periodLabel: year,
        granularity: 'year',
      } : null;
      if (!year) state.drawerId = '';
    } else {
      state.selection = null;
      state.drawerId = '';
    }
  }
  state.granularity = next;
  renderBody();
}

function clearSelection() {
  state.selection = null;
  state.drawerId = '';
  renderBody();
}

function samePeriodSelection(eventType, periodKey) {
  var selection = state.selection;
  return !!(selection && selection.kind === 'period' && selection.eventType === eventType && selection.periodKey === periodKey && selection.granularity === state.granularity);
}

function emptyPlotClick(target) {
  if (!target || !target.closest) return false;
  if (target.closest('.am-slice, .am-stack-total')) return false;
  return !!target.closest('.am-plot');
}

function onClick(event) {
  var viewButton = event.target.closest('[data-am-view]');
  if (viewButton) {
    state.viewMode = viewButton.getAttribute('data-am-view') === 'matrix' ? 'matrix' : 'graph';
    renderBody();
    return;
  }
  var grainButton = event.target.closest('[data-am-granularity]');
  if (grainButton) {
    setGranularity(grainButton.getAttribute('data-am-granularity') === 'year' ? 'year' : 'month');
    return;
  }
  var segment = event.target.closest('[data-am-segment]');
  if (segment) {
    var eventType = segment.getAttribute('data-am-segment');
    var periodKey = segment.getAttribute('data-am-period');
    if (samePeriodSelection(eventType, periodKey)) {
      clearSelection();
      return;
    }
    state.selection = {
      kind: 'period',
      eventType: eventType,
      periodKey: periodKey,
      periodLabel: segment.getAttribute('data-am-period-label') || fullMonth(periodKey),
      granularity: state.granularity,
    };
    state.drawerId = '';
    renderBody();
    return;
  }
  if (emptyPlotClick(event.target)) {
    if (state.selection) clearSelection();
    return;
  }
  var category = event.target.closest('[data-am-category]');
  if (category) {
    state.selection = { kind: 'category', eventType: category.getAttribute('data-am-category') };
    state.drawerId = '';
    renderBody();
    return;
  }
  var asset = event.target.closest('[data-am-property]');
  if (asset) {
    var propertyId = asset.getAttribute('data-am-property');
    state.expandedPropertyId = state.expandedPropertyId === propertyId ? '' : propertyId;
    renderBody();
    return;
  }
  if (event.target.closest('[data-am-clear-property]')) {
    state.propertyId = '';
    state.selection = null;
    state.drawerId = '';
    renderChrome();
    renderBody();
    return;
  }
  if (event.target.closest('[data-am-view-all]')) {
    state.selection = { kind: 'all' };
    state.drawerId = '';
    renderBody();
    return;
  }
  if (event.target.closest('[data-am-clear-selection]')) {
    clearSelection();
    return;
  }
  if (event.target.closest('[data-am-close-drawer]')) {
    closeDrawer();
    return;
  }
  var row = event.target.closest('[data-am-finding]');
  if (row) openFinding(row.getAttribute('data-am-finding'));
}

function openFinding(id) {
  state.drawerId = id;
  renderBody();
  var close = workspace().querySelector('[data-am-close-drawer]');
  if (close) close.focus({ preventScroll: true });
}

function focusable(root) {
  return Array.prototype.filter.call(root.querySelectorAll('button, select, tr[tabindex="0"]'), function (element) {
    return !element.disabled && element.offsetParent !== null;
  });
}

function onKeyDown(event) {
  var root = workspace();
  if (!root || root.hidden) return;
  if (event.key === 'Enter' && event.target && event.target.closest && event.target.closest('[data-am-finding]') === event.target) {
    openFinding(event.target.getAttribute('data-am-finding'));
    return;
  }
  if (event.key === 'Tab') {
    var items = focusable(root);
    if (!items.length) return;
    var first = items[0];
    var last = items[items.length - 1];
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault();
      first.focus();
    }
    return;
  }
  if (event.key !== 'Escape' || !state.drawerId) return;
  event.preventDefault();
  closeDrawer();
}

export function setupAssetQueries() {
  var button = document.getElementById('am-queries-btn');
  var root = workspace();
  if (!button || !root) return;
  button.addEventListener('click', openWorkspace);
  document.getElementById('am-close').addEventListener('click', closeWorkspace);
  root.addEventListener('change', onFilterChange);
  root.addEventListener('click', onClick);
  document.addEventListener('keydown', onKeyDown);
  refreshAssetQueriesButton();
}
