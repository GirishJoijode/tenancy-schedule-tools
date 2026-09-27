import { findSourceColumn, isBlankSourceValue } from './sourceLookup.js';
import { normaliseHeader, sourceLabel } from './text.js';
import { describePeriodGroup, getPeriodColumns } from './periods.js';
import { getRequiredPropertySources } from './propertySourceSchema.js';
import { getRequiredUnitSources } from './unitSourceSchema.js';
import { PROPERTY_UNIT_JOIN } from '../templates/defaultSchedule.js';

function requiredSources(fileType) {
  return fileType === 'property' ? getRequiredPropertySources() : getRequiredUnitSources();
}

export function assessSourceSchema(fileType, columns) {
  var sources = requiredSources(fileType);
  var kind = fileType === 'property' ? 'Property' : 'Unit';
  var problems = [];

  for (var i = 0; i < sources.fields.length; i += 1) {
    var field = sources.fields[i];
    if (!findSourceColumn(columns, fileType, field)) {
      problems.push('Missing required ' + kind + ' field:\n' + sourceLabel(field));
    }
  }

  for (var g = 0; g < sources.periodGroups.length; g += 1) {
    var group = sources.periodGroups[g];
    var wanted = normaliseHeader(group.section);
    var inSection = (columns || []).some(function (column) {
      return normaliseHeader(column.section) === wanted;
    });
    if (!inSection) {
      problems.push('Missing required ' + kind + ' source group:\n' + group.section);
      continue;
    }
    if (!getPeriodColumns(columns, group.section).length) {
      problems.push('No period columns found under:\n' + group.section);
    }
  }

  return problems;
}

export function formatUnrecognisedSource(problems) {
  return ['Source format not recognised', ''].concat(problems).join('\n');
}

export function assertRecognisedSource(fileType, columns) {
  var problems = assessSourceSchema(fileType, columns);
  if (problems.length) {
    throw new Error(formatUnrecognisedSource(problems));
  }
  if (typeof import.meta !== 'undefined' && import.meta.env && import.meta.env.DEV) {
    var groups = requiredSources(fileType).periodGroups;
    for (var i = 0; i < groups.length; i += 1) {
      console.debug(describePeriodGroup(columns, groups[i].section));
    }
  }
}

export function omitBlankPropertyNameRows(extracted, columns, fileType) {
  var nameCol = findSourceColumn(columns, fileType, PROPERTY_UNIT_JOIN);
  var rows = [];
  var sourceRows = extracted.rows || [];
  for (var i = 0; i < sourceRows.length; i += 1) {
    var row = sourceRows[i];
    var raw = nameCol && row.values ? row.values[nameCol.key] : null;
    if (isBlankSourceValue(raw)) continue;
    rows.push(row);
  }
  return { rows: rows, dataRowCount: rows.length };
}
