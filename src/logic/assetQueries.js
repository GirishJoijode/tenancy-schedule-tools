export var AM_HORIZONS = [6, 12, 24, 36];
export var AM_DEFAULT_HORIZON = 36;

var EVENT_LABELS = {
  rent_review: 'Rent Review',
  break: 'Break Option',
  expiry: 'Lease Expiry',
};

var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function calendarDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === 'number' && Number.isFinite(value) && value > 20000 && value < 80000) {
    var utc = new Date(Math.round((value - 25569) * 86400 * 1000));
    return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
  }
  if (typeof value === 'string' && value.trim()) {
    var match = value.trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (!match) return null;
    var day = Number(match[1]);
    var month = Number(match[2]);
    var year = Number(match[3]);
    var parsed = new Date(year, month - 1, day);
    if (parsed.getFullYear() === year && parsed.getMonth() === month - 1 && parsed.getDate() === day) return parsed;
  }
  return null;
}

export function addCalendarMonths(date, months) {
  var monthIndex = date.getMonth() + months;
  var year = date.getFullYear() + Math.floor(monthIndex / 12);
  var month = monthIndex % 12;
  if (month < 0) {
    month += 12;
    year -= 1;
  }
  var lastDay = new Date(year, month + 1, 0).getDate();
  return new Date(year, month, Math.min(date.getDate(), lastDay));
}

function daysBetween(from, to) {
  var ms = Date.UTC(to.getFullYear(), to.getMonth(), to.getDate()) - Date.UTC(from.getFullYear(), from.getMonth(), from.getDate());
  return Math.round(ms / 86400000);
}

function wholeMonthsBetween(from, to) {
  var start = from <= to ? from : to;
  var end = from <= to ? to : from;
  var months = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth());
  if (end.getDate() < start.getDate()) months -= 1;
  return Math.max(0, months);
}

export function timingLabel(valuationDate, eventDate) {
  var days = daysBetween(valuationDate, eventDate);
  if (days === 0) return 'on the valuation date';
  var months = wholeMonthsBetween(valuationDate, eventDate);
  var amount = months >= 1
    ? months + (months === 1 ? ' month' : ' months')
    : Math.abs(days) + (Math.abs(days) === 1 ? ' day' : ' days');
  return days < 0 ? amount + ' before the valuation date' : amount + ' after the valuation date';
}

function inReviewScope(valuationDate, eventDate, horizonMonths) {
  if (eventDate < valuationDate) return true;
  return eventDate <= addCalendarMonths(valuationDate, horizonMonths);
}

function monthKey(date) {
  return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
}

function monthLabel(date) {
  return MONTHS[date.getMonth()] + ' ' + String(date.getFullYear()).slice(2);
}

function formatDisplayDate(date) {
  return String(date.getDate()).padStart(2, '0') + '/' + String(date.getMonth() + 1).padStart(2, '0') + '/' + date.getFullYear();
}

function numericRent(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.trim()) {
    var parsed = Number(value.replace(/,/g, '').trim());
    if (Number.isFinite(parsed)) return parsed;
  }
  return null;
}

function textValue(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(/\s+/g, ' ').trim();
}

function eventTitle(eventType, status) {
  if (eventType === 'rent_review') return status === 'historic' ? 'Historic Rent Review' : 'Upcoming Rent Review';
  if (eventType === 'break') return status === 'historic' ? 'Historic Break' : 'Upcoming Break';
  return status === 'historic' ? 'Lease Expiry Passed' : 'Upcoming Lease Expiry';
}

function leaseFacts(unitValues) {
  var values = unitValues || {};
  return {
    contractedRent: numericRent(values.contracted_rent_pa),
    currentRent: numericRent(values.deemed_rent_psf), // Unit Variable → Current Rent, stored on deemed_rent_psf.
    reviewedRent: null,
    marketRent: numericRent(values.market_rent_pa),
    rentFreeMonths: numericRent(values.rent_free_month),
    expiryVoidMonths: numericRent(values.expiry_void_month),
  };
}

export var AM_QUERY_RULES = [
  {
    id: 'rent_review',
    name: 'Rent Review',
    category: 'Rent Review',
    eventType: 'rent_review',
    dateField: 'review_date',
    source: 'Unit / TS → Variable → Next Review',
  },
  {
    id: 'break',
    name: 'Break Option',
    category: 'Break Option',
    eventType: 'break',
    dateField: 'break_date',
    source: 'Unit / TS → General → Effective Break',
  },
  {
    id: 'expiry',
    name: 'Lease Expiry',
    category: 'Lease Expiry',
    eventType: 'expiry',
    dateField: 'expiry_date',
    source: 'Unit / TS → General → Lease Expiry',
  },
];

function groupMembership(properties, groups) {
  var membership = new Map();
  (groups || []).forEach(function (group) {
    (group.properties || []).forEach(function (property) {
      membership.set(property, { id: String(group.id), label: group.label || 'Group' });
    });
  });
  properties.forEach(function (property) {
    if (!membership.has(property)) membership.set(property, { id: '', label: '' });
  });
  return membership;
}

export function canRunAssetQueries(propertyData, unitData) {
  return Boolean(propertyData && unitData);
}

export function assetQueryGroupOptions(grouping) {
  if (!grouping || !grouping.effective || !grouping.groups || grouping.groups.length < 2) return [];
  return grouping.groups.map(function (group) {
    return { id: String(group.id), label: group.label || 'Group' };
  });
}

export function runAssetManagementReview(options) {
  var request = options || {};
  var properties = request.properties || [];
  var horizonMonths = request.horizonMonths || AM_DEFAULT_HORIZON;
  var membership = groupMembership(properties, request.groups);
  var findings = [];
  var catalogue = [];

  properties.forEach(function (property, propertyIndex) {
    var values = property.standardValues || {};
    var valuationDate = calendarDate(values.valuation);
    var propertyName = textValue(values.address) || 'Property';
    var externalId = textValue(values.property_reference);
    var propertyId = externalId || 'property:' + propertyIndex;
    var group = membership.get(property) || { id: '', label: '' };
    var units = property.units || [];
    catalogue.push({
      propertyId: propertyId,
      propertyName: propertyName,
      externalId: externalId,
      groupId: group.id,
      leases: units.length,
    });
    if (!valuationDate) return;

    units.forEach(function (unit, unitIndex) {
      var unitValues = unit.standardValues || {};
      var unitName = textValue(unitValues.unit) || 'Unit';
      var tenantName = textValue(unitValues.tenant_name);
      var unitId = propertyId + ':' + (textValue(unitValues.unit) || 'unit') + ':' + unitIndex;
      var facts = leaseFacts(unitValues);

      AM_QUERY_RULES.forEach(function (rule) {
        var eventDate = calendarDate(unitValues[rule.dateField]);
        if (!eventDate || !inReviewScope(valuationDate, eventDate, horizonMonths)) return;
        var status = eventDate < valuationDate ? 'historic' : 'upcoming';
        var timing = timingLabel(valuationDate, eventDate);
        findings.push({
          id: rule.id + '|' + unitId,
          ruleId: rule.id,
          category: rule.category,
          eventType: rule.eventType,
          status: status,
          propertyId: propertyId,
          propertyName: propertyName,
          externalId: externalId,
          groupId: group.id,
          groupLabel: group.label,
          unitId: unitId,
          unitName: unitName,
          tenantName: tenantName,
          valuationDate: valuationDate,
          eventDate: eventDate,
          daysFromValuation: daysBetween(valuationDate, eventDate),
          monthKey: monthKey(eventDate),
          monthLabel: monthLabel(eventDate),
          valuationLabel: formatDisplayDate(valuationDate),
          eventLabel: formatDisplayDate(eventDate),
          timing: timing,
          contractedRent: facts.contractedRent,
          currentRent: facts.currentRent,
          reviewedRent: facts.reviewedRent,
          marketRent: facts.marketRent,
          rentFreeMonths: facts.rentFreeMonths,
          expiryVoidMonths: facts.expiryVoidMonths,
          breakVoidMonths: null,
          title: eventTitle(rule.eventType, status),
          source: rule.source,
        });
      });
    });
  });

  return {
    horizonMonths: horizonMonths,
    properties: catalogue,
    findings: findings,
  };
}

function findingSort(left, right) {
  var leftHistoric = left.status === 'historic' ? 0 : 1;
  var rightHistoric = right.status === 'historic' ? 0 : 1;
  if (leftHistoric !== rightHistoric) return leftHistoric - rightHistoric;
  if (left.eventDate - right.eventDate) return left.eventDate - right.eventDate;
  var property = String(left.propertyName || '').localeCompare(String(right.propertyName || ''));
  if (property) return property;
  var unit = String(left.unitName || '').localeCompare(String(right.unitName || ''));
  if (unit) return unit;
  return String(left.id || '').localeCompare(String(right.id || ''));
}

function matchesDashboard(finding, filters) {
  if (filters.propertyId && finding.propertyId !== filters.propertyId) return false;
  if (filters.groupId && finding.groupId !== filters.groupId) return false;
  if (filters.eventType && filters.eventType !== 'all' && finding.eventType !== filters.eventType) return false;
  return true;
}

function emptyMeasure() {
  return { rent_review: 0, break: 0, expiry: 0 };
}

function propertyInScope(property, filters) {
  if (filters.propertyId && property.propertyId !== filters.propertyId) return false;
  if (filters.groupId && property.groupId !== filters.groupId) return false;
  return true;
}

export function deriveAssetManagementView(review, filters) {
  var source = review || { findings: [], properties: [] };
  var active = filters || {};
  var scoped = (source.properties || []).filter(function (property) {
    return propertyInScope(property, active);
  });
  var findings = (source.findings || []).filter(function (finding) {
    return matchesDashboard(finding, active);
  }).sort(findingSort);

  var summary = {
    propertiesReviewed: scoped.length,
    leasesReviewed: scoped.reduce(function (sum, property) { return sum + property.leases; }, 0),
    queries: findings.length,
    upcoming: 0,
    historic: 0,
  };

  var months = [];
  var monthIndex = {};
  findings.forEach(function (finding) {
    summary[finding.status] += 1;
    if (!monthIndex[finding.monthKey]) {
      monthIndex[finding.monthKey] = {
        key: finding.monthKey,
        label: finding.monthLabel,
        counts: emptyMeasure(),
        rent: emptyMeasure(),
      };
      months.push(monthIndex[finding.monthKey]);
    }
    var bucket = monthIndex[finding.monthKey];
    bucket.counts[finding.eventType] += 1;
    if (typeof finding.contractedRent === 'number') bucket.rent[finding.eventType] += finding.contractedRent;
  });
  months.sort(function (left, right) { return left.key < right.key ? -1 : left.key > right.key ? 1 : 0; });

  var filled = [];
  if (months.length) {
    var cursor = months[0].key.split('-');
    var end = months[months.length - 1].key.split('-');
    var year = Number(cursor[0]);
    var month = Number(cursor[1]);
    var endYear = Number(end[0]);
    var endMonth = Number(end[1]);
    while (year < endYear || (year === endYear && month <= endMonth)) {
      var key = year + '-' + String(month).padStart(2, '0');
      filled.push(monthIndex[key] || {
        key: key,
        label: monthLabel(new Date(year, month - 1, 1)),
        counts: emptyMeasure(),
        rent: emptyMeasure(),
      });
      month += 1;
      if (month > 12) {
        month = 1;
        year += 1;
      }
    }
  }

  var categoryCounts = ['expiry', 'rent_review', 'break'].map(function (eventType) {
    return {
      id: eventType,
      label: EVENT_LABELS[eventType],
      count: findings.filter(function (finding) { return finding.eventType === eventType; }).length,
      rent: findings.reduce(function (sum, finding) {
        if (finding.eventType !== eventType || typeof finding.contractedRent !== 'number') return sum;
        return sum + finding.contractedRent;
      }, 0),
    };
  });

  var sourceOrder = {};
  (source.properties || []).forEach(function (property, index) {
    sourceOrder[property.propertyId] = index;
  });
  var attentionMap = {};
  findings.forEach(function (finding) {
    if (!attentionMap[finding.propertyId]) {
      attentionMap[finding.propertyId] = {
        propertyId: finding.propertyId,
        propertyName: finding.propertyName,
        externalId: finding.externalId,
        total: 0,
        review: 0,
        sourceOrder: sourceOrder[finding.propertyId] == null ? 100000 : sourceOrder[finding.propertyId],
      };
    }
    attentionMap[finding.propertyId].total += 1;
    if (finding.status === 'historic') attentionMap[finding.propertyId].review += 1;
  });
  var propertyAttention = Object.keys(attentionMap).map(function (key) { return attentionMap[key]; });
  propertyAttention.sort(function (left, right) {
    if (right.total !== left.total) return right.total - left.total;
    if (right.review !== left.review) return right.review - left.review;
    return left.sourceOrder - right.sourceOrder;
  });

  return {
    summary: summary,
    findings: findings,
    monthlyEvents: filled,
    categoryCounts: categoryCounts,
    propertyAttention: propertyAttention,
  };
}

export function findingsForRegister(findings, focus) {
  var rows = findings || [];
  if (!focus || (!focus.eventType && !focus.monthKey)) return rows;
  return rows.filter(function (finding) {
    if (focus.eventType && finding.eventType !== focus.eventType) return false;
    if (focus.monthKey && finding.monthKey !== focus.monthKey) return false;
    return true;
  });
}
