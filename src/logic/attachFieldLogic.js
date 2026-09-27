function refOf(rule) {
  if (!rule) return null;
  if (rule.method === 'source' || rule.method === 'aggregatedSource') return rule.ref || null;
  return null;
}

function sameRef(left, right) {
  if (!left && !right) return true;
  if (!left || !right) return false;
  return (
    left.type === right.type &&
    left.section === right.section &&
    (left.field || '') === (right.field || '') &&
    (left.selection || '') === (right.selection || '') &&
    (left.aggregation || '') === (right.aggregation || '') &&
    (left.occurrence || 1) === (right.occurrence || 1)
  );
}

export function attachFieldLogic(columns, logicById) {
  return columns.map(function (column) {
    var logic = logicById[column.id];
    if (!logic) throw new Error('Missing production logic for ' + column.id);
    if (!sameRef(column.property, refOf(logic.property))) {
      throw new Error('Property logic does not match the source ref for ' + column.id);
    }
    if (!sameRef(column.unit, refOf(logic.unit))) {
      throw new Error('Unit logic does not match the source ref for ' + column.id);
    }
    return Object.assign({}, column, { logic: logic });
  });
}

export function attachPropertyOnlyLogic(columns) {
  return columns.map(function (column) {
    var rule = column.property ? { method: 'source', ref: column.property } : { method: 'blank' };
    return Object.assign({}, column, { logic: { property: rule } });
  });
}
