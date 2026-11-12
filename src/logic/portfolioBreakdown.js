import { readSourceValue } from '../schema/sourceLookup.js';
import { PROPERTY_SOURCE_FIELDS } from '../schema/sourceFields.js';
import { normaliseComparisonPart } from './comparison.js';

export var UNCLASSIFIED_GROUP = 'Unclassified';

export var BREAKDOWN_OPTIONS = [
  { id: 'none', label: 'No Breakdown' },
  { id: 'sector', label: 'Sector' },
  { id: 'region_country', label: 'Region / Country' },
  { id: 'custom', label: 'Custom\u2026', selectedLabel: 'Custom' },
];

var SOURCE_FIELDS = {
  sector: PROPERTY_SOURCE_FIELDS.sector,
  region_country: PROPERTY_SOURCE_FIELDS.regionCountry,
};

var mode = 'sector';
var customConfig = { groups: [], assignments: {} };
var groupSerial = 1;

function cloneConfig(config) {
  var source = config || { groups: [], assignments: {} };
  return {
    groups: (source.groups || []).map(function (group) {
      return { id: group.id, name: group.name };
    }),
    assignments: Object.assign({}, source.assignments || {}),
  };
}

export function collapseGroupText(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
}

export function classifyGroupValue(value) {
  var text = collapseGroupText(value);
  if (!text || text.toLowerCase() === 'unclassified') {
    return { unclassified: true, key: 'unclassified', label: UNCLASSIFIED_GROUP };
  }
  return { unclassified: false, key: text.toLowerCase(), label: text };
}

export function externalAssignmentKey(value) {
  return normaliseComparisonPart(value);
}

export function getBreakdownMode() {
  return mode;
}

export function setBreakdownMode(next) {
  if (next === 'none' || next === 'sector' || next === 'region_country' || next === 'custom') mode = next;
  return mode;
}

export function getCustomConfig() {
  return cloneConfig(customConfig);
}

export function setCustomConfig(next) {
  customConfig = cloneConfig(next);
  return getCustomConfig();
}

export function resetPortfolioBreakdown() {
  mode = 'sector';
  customConfig = { groups: [], assignments: {} };
  groupSerial = 1;
}

export function createGroupId() {
  groupSerial += 1;
  return 'group-' + groupSerial;
}

export function breakdownSpecification() {
  if (mode === 'sector' || mode === 'region_country') {
    var ref = SOURCE_FIELDS[mode];
    return {
      mode: mode,
      source: { section: ref.section, field: ref.field },
    };
  }
  if (mode === 'custom') return { mode: 'custom' };
  return { mode: 'none' };
}

export function propertyIdentity(sourceRow, columns) {
  var externalRaw = columns
    ? readSourceValue(sourceRow, columns, 'property', PROPERTY_SOURCE_FIELDS.externalId)
    : sourceRow && sourceRow.standardValues
      ? sourceRow.standardValues.property_reference
      : null;
  var nameRaw = columns
    ? readSourceValue(sourceRow, columns, 'property', PROPERTY_SOURCE_FIELDS.propertyName)
    : sourceRow && sourceRow.standardValues
      ? sourceRow.standardValues.address
      : null;
  var externalId = collapseGroupText(externalRaw);
  var name = collapseGroupText(nameRaw);
  var externalKey = externalAssignmentKey(externalId);
  if (!externalKey && sourceRow && sourceRow.rowIndex != null) externalKey = 'row:' + sourceRow.rowIndex;
  if (!externalKey && sourceRow && sourceRow.sourceRow) externalKey = 'row:' + sourceRow.sourceRow.rowIndex;
  return {
    externalKey: externalKey,
    externalId: externalId,
    name: name,
    label: externalId && name ? externalId + ' - ' + name : externalId || name || 'Property',
  };
}

export function listCurrentProperties(propertyData) {
  if (!propertyData || !propertyData.rows) return [];
  var listed = [];
  for (var i = 0; i < propertyData.rows.length; i += 1) {
    var row = propertyData.rows[i];
    var identity = propertyIdentity(row, propertyData.columns);
    if (!identity.name && !identity.externalId) continue;
    listed.push({
      externalKey: identity.externalKey,
      externalId: identity.externalId,
      name: identity.name,
      label: identity.label,
      order: i,
    });
  }
  return listed;
}

function identityFromPortfolioProperty(property) {
  if (property && property.externalKey) return { externalKey: property.externalKey };
  var fromValues = propertyIdentity({
    standardValues: property.standardValues || {},
    sourceRow: property.sourceRow,
    rowIndex: property.sourceRow ? property.sourceRow.rowIndex : null,
  }, null);
  if (fromValues.externalKey) return fromValues;
  return propertyIdentity(property.sourceRow, null);
}

export function reconcileCustomAssignments(propertyData) {
  var live = {};
  listCurrentProperties(propertyData).forEach(function (property) {
    if (property.externalKey) live[property.externalKey] = true;
  });
  var next = {};
  Object.keys(customConfig.assignments || {}).forEach(function (key) {
    if (live[key]) next[key] = customConfig.assignments[key];
  });
  customConfig.assignments = next;
  return getCustomConfig();
}

function readGroupRaw(property, propertyData, ref) {
  if (!property) return null;
  if (propertyData && property.sourceRow) {
    return readSourceValue(property.sourceRow, propertyData.columns, 'property', ref);
  }
  return null;
}

function bucketByField(properties, propertyData, ref) {
  var order = [];
  var buckets = new Map();
  var labels = {};
  var unclassified = [];
  var populated = 0;
  for (var i = 0; i < properties.length; i += 1) {
    var property = properties[i];
    var classified = classifyGroupValue(readGroupRaw(property, propertyData, ref));
    if (classified.unclassified) {
      unclassified.push(property);
      continue;
    }
    populated += 1;
    if (!buckets.has(classified.key)) {
      buckets.set(classified.key, []);
      labels[classified.key] = classified.label;
      order.push(classified.key);
    }
    buckets.get(classified.key).push(property);
  }
  var groups = order.map(function (key) {
    return { id: key, label: labels[key], properties: buckets.get(key) };
  });
  if (unclassified.length) {
    groups.push({ id: 'unclassified', label: UNCLASSIFIED_GROUP, properties: unclassified });
  }
  return {
    groups: groups,
    populated: populated,
    unclassified: unclassified.length,
  };
}

function bucketCustom(properties, config) {
  var saved = cloneConfig(config);
  var groups = saved.groups.map(function (group) {
    return { id: group.id, label: collapseGroupText(group.name) || 'New Group', properties: [] };
  });
  var byId = {};
  groups.forEach(function (group) { byId[group.id] = group; });
  var unclassified = [];
  for (var i = 0; i < properties.length; i += 1) {
    var property = properties[i];
    var key = identityFromPortfolioProperty(property).externalKey;
    var groupId = key ? saved.assignments[key] : '';
    if (groupId && byId[groupId]) byId[groupId].properties.push(property);
    else unclassified.push(property);
  }
  var visible = groups.filter(function (group) { return group.properties.length; });
  if (unclassified.length) {
    visible.push({ id: 'unclassified', label: UNCLASSIFIED_GROUP, properties: unclassified });
  }
  return { groups: visible, populated: properties.length - unclassified.length, unclassified: unclassified.length };
}

function describeGrouping(modeName, groups, properties, populated, unclassified) {
  return {
    mode: modeName,
    groups: groups,
    properties: properties,
    populated: populated,
    unclassified: unclassified,
    effective: groups.length > 1,
  };
}

export function buildPortfolioGroups(properties, options) {
  var request = options || {};
  var selected = request.mode || 'sector';
  var rows = properties || [];
  if (selected === 'none') {
    return describeGrouping('none', [], rows.slice(), rows.length, 0);
  }
  if (selected === 'custom') {
    var custom = bucketCustom(rows, request.customConfig || customConfig);
    return describeGrouping('custom', custom.groups, rows.slice(), custom.populated, custom.unclassified);
  }
  var ref = SOURCE_FIELDS[selected] || SOURCE_FIELDS.sector;
  var fielded = bucketByField(rows, request.propertyData, ref);
  return describeGrouping(selected, fielded.groups, rows.slice(), fielded.populated, fielded.unclassified);
}

export function assessBreakdown(propertyData, selectedMode) {
  var selected = selectedMode || mode;
  if (!propertyData || !propertyData.rows || !propertyData.rows.length) return { status: 'idle' };
  if (selected !== 'sector' && selected !== 'region_country') return { status: 'ok' };
  var properties = propertyData.rows.map(function (row) {
    return { sourceRow: row };
  });
  var built = buildPortfolioGroups(properties, {
    mode: selected,
    propertyData: propertyData,
  });
  var label = selected === 'region_country' ? 'Region / Country' : 'Sector';
  if (!built.effective) {
    if (!built.populated) {
      return {
        status: 'unpopulated',
        message: label + ' is not populated. The schedule will be generated without group sections.',
      };
    }
    return {
      status: 'single',
      message: 'Only one ' + label + ' group is present. The schedule will be generated without group sections.',
    };
  }
  if (built.unclassified) {
    var count = built.unclassified;
    return {
      status: 'partial',
      count: count,
      message: count + (count === 1 ? ' property' : ' properties') + ' will be grouped as Unclassified.',
    };
  }
  return { status: 'ok' };
}

export function validateCustomGroups(groups) {
  var issues = {};
  var seen = {};
  (groups || []).forEach(function (group) {
    var name = collapseGroupText(group.name);
    if (!name) {
      issues[group.id] = 'Enter a group name.';
      return;
    }
    if (name.toLowerCase() === 'unclassified') {
      issues[group.id] = 'Unclassified is reserved.';
      return;
    }
    var key = name.toLowerCase();
    if (seen[key]) {
      issues[group.id] = 'Group names must be unique.';
      issues[seen[key]] = 'Group names must be unique.';
      return;
    }
    seen[key] = group.id;
  });
  return issues;
}
