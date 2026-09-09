import type { CompatibilityReport } from './compatibility/types';

export interface FormulaConversionCandidate {
  worksheetName: string;
  address: string;
  formula: string;
  reasons: Array<'unsupportedFunction' | 'externalReference'>;
}
export interface FormulaConversionPlan {
  candidates: FormulaConversionCandidate[];
  brokenReferenceCount: number;
  blockedByDetailLimit: boolean;
}

/** Only plans reversible-to-audit value freezing. It never invents a missing reference or formula equivalent. */
export function planFormulaConversions(report: CompatibilityReport, options: {
  freezeUnsupported: boolean; freezeExternal: boolean;
}): FormulaConversionPlan {
  const byCell = new Map<string, FormulaConversionCandidate>();
  let detailedRelevant = 0;
  for (const finding of report.findings) {
    if (finding.source !== 'liveCells' || !finding.worksheetName || !finding.address || !finding.formula) continue;
    const reason = finding.code === 'functionMarker' && options.freezeUnsupported ? 'unsupportedFunction'
      : finding.code === 'externalFormula' && options.freezeExternal ? 'externalReference' : undefined;
    if (!reason) continue;
    detailedRelevant++;
    const key = `${finding.worksheetName}\u0000${finding.address}`;
    const candidate = byCell.get(key) ?? { worksheetName: finding.worksheetName, address: finding.address,
      formula: finding.formula, reasons: [] };
    if (!candidate.reasons.includes(reason)) candidate.reasons.push(reason);
    byCell.set(key, candidate);
  }
  const expectedRelevant = report.summaries.filter(summary =>
    (summary.code === 'functionMarker' && options.freezeUnsupported) ||
    (summary.code === 'externalFormula' && options.freezeExternal),
  ).reduce((total, summary) => total + summary.count, 0);
  const brokenReferenceCount = report.summaries.filter(summary => summary.code === 'brokenReference')
    .reduce((total, summary) => total + summary.count, 0);
  return { candidates: [...byCell.values()], brokenReferenceCount,
    blockedByDetailLimit: expectedRelevant > detailedRelevant };
}
