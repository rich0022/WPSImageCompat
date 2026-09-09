import type { Finding } from './types';

/** Mask quoted literals/names and structured references without changing token positions. */
function mask(expression: string, namesAndBrackets: boolean): string {
  let result = '';
  let quote = '';
  let depth = 0;
  for (let i = 0; i < expression.length; i++) {
    const char = expression[i]!;
    if (quote) {
      const preserveName = quote === "'" && !namesAndBrackets;
      result += preserveName ? char : ' ';
      if (char === quote) {
        if (expression[i + 1] === quote) { result += preserveName ? char : ' '; i++; }
        else quote = '';
      }
    } else if (char === '"' || char === "'") {
      quote = char; result += char === "'" && !namesAndBrackets ? char : ' ';
    } else if (namesAndBrackets && char === '[') { depth++; result += ' '; }
    else if (namesAndBrackets && char === ']' && depth) { depth--; result += ' '; }
    else result += depth ? ' ' : char;
  }
  return result;
}

/** Only real formulas and host-typed errors are inputs; constants are never inferred to be errors. */
export function analyzeFormula(input: {
  worksheetName: string; address: string; formula: unknown; value: unknown; valueType: string;
}): Finding[] {
  if (typeof input.formula !== 'string' || !input.formula.startsWith('=') || input.formula === input.value) return [];
  const formula = input.formula;
  const base = { worksheetName: input.worksheetName, address: input.address, formula, source: 'liveCells' as const };
  const findings: Finding[] = [];
  const add = (code: Finding['code'], evidence: string, category: Finding['category'] = 'formulas', status: Finding['status'] = 'risk') =>
    findings.push({ ...base, category, code, evidence, status });
  if (input.valueType === 'Error') add('formulaError', String(input.value), 'formulas', 'confirmed');
  const code = mask(formula, true);
  const functions = [...code.matchAll(/(?<![\w.])([A-Za-z_][\w.]*)\s*\(/g)].map(match => match[1]!);
  const markers = [...new Set(functions.filter(name => /^_(xlfn|xlws)\./i.test(name)))];
  if (markers.length) add('functionMarker', markers.join(', '));
  if (functions.some(name => name.replace(/^(?:_(?:xlfn|xlws)\.)+/i, '').toUpperCase() === 'DISPIMG')) add('dispimg', 'DISPIMG');
  if (/#REF!/i.test(code)) add('brokenReference', '#REF!', 'externalLinks', 'confirmed');
  const references = mask(formula, false);
  // Match workbook-qualified sheet references, never Table[Column] or quoted string literals.
  const external = /'(?:[^']|'')*\[[^\]\r\n]+\](?:[^']|'')*'!|(?<![\w\]])\[[^\]\r\n]+\][^\s!+*\/(),:;=<>"'\[\]]*!/g;
  for (const reference of new Set(references.match(external) ?? [])) add('externalFormula', reference, 'externalLinks');
  return findings;
}
