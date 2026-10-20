/** Preserve native SharePoint refinement operands and legacy unquoted text. */
function toRefinementOperand(value: string): string {
  if (value.startsWith('"') && value.endsWith('"')) {
    return value;
  }
  if (/^ǂǂ[0-9a-f]+$/i.test(value)) {
    return '"' + value + '"';
  }
  // RefinementToken can be a numeric/date scalar or an executable FQL bin
  // expression such as range(...). SharePoint owns expression validation.
  if (/^[a-z][a-z0-9-]*\s*\(/i.test(value) ||
      /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(value) ||
      /^\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z?)?$/.test(value)) {
    return value;
  }
  const escaped = value.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  return 'equals("' + escaped + '")';
}

/** Build one property group for OR, or independent AND restrictions. */
export function buildRefinementFilters(
  refinersToApply: Map<string, string[]>,
  operator: 'and' | 'or' = 'and'
): string[] {
  const filters: string[] = [];
  refinersToApply.forEach((values, name) => {
    const operands = values.map(toRefinementOperand);
    if (operator === 'or' && operands.length > 1) {
      filters.push(name + ':or(' + operands.join(',') + ')');
    } else {
      operands.forEach(operand => filters.push(name + ':' + operand));
    }
  });
  return filters;
}
