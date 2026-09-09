import type { CompatibilityReport, FindingCode } from '../core/compatibility/types';
import { translate, type Locale, type MessageKey } from './i18n';

const titles: Record<FindingCode, MessageKey> = {
  formulaError: 'compatFormulaError', functionMarker: 'compatFunctionMarker', dispimg: 'compatDispimg',
  brokenReference: 'compatBrokenReference', externalFormula: 'compatExternalFormula', externalRecord: 'compatExternalRecord',
};
const categories = { formulas: 'compatFormulas', externalLinks: 'compatLinks', dates: 'compatDates' } as const;
const states = { confirmed: 'compatConfirmed', risk: 'compatRisk', clear: 'compatClear', unchecked: 'compatUnchecked' } as const;
const completion = { complete: 'compatComplete', partial: 'compatPartial', unavailable: 'compatUnavailable' } as const;
export function renderCompatibility(container: HTMLElement, report: CompatibilityReport | undefined, locale: Locale): void {
  container.replaceChildren();
  if (!report) return;
  const t = (key: MessageKey) => translate(locale, key);
  const paragraph = (text: string) => { const p = document.createElement('p'); p.textContent = text; container.append(p); };
  paragraph(translate(locale, 'compatCounts', { cells: report.cells.scannedCells, sheets: report.cells.scannedWorksheets,
    total: report.cells.totalWorksheets, findings: report.findingCount }));
  for (const check of report.checks) {
    paragraph(t(categories[check.category]) + ' · ' + t(states[check.status]) + ' · ' + t(completion[check.completion]) +
      (check.reasonCode ? ' (' + check.reasonCode + ')' : ''));
  }
  if (report.metadata?.dateSystem && report.metadata.dateSystem !== 'unknown') {
    paragraph(translate(locale, 'compatDateSystem', { system: report.metadata.dateSystem }));
  }
  paragraph(t('compatDateHint'));
  paragraph(t('compatScope'));
  if (report.omittedFindings) paragraph(translate(locale, 'compatOmitted', { count: report.omittedFindings }));
  if (report.findings.length > 100) paragraph(t('compatFirst100'));
  const list = document.createElement('ul');
  for (const finding of report.findings.slice(0, 100)) {
    const item = document.createElement('li');
    item.textContent = t(titles[finding.code]) + ' · ' + t(states[finding.status]) + ' · ' +
      (finding.worksheetName ? finding.worksheetName + '!' + finding.address : t('workbook'));
    const details = document.createElement('details');
    const summary = document.createElement('summary'); summary.textContent = t('details');
    const evidence = document.createElement('p');
    evidence.textContent = t(finding.source === 'liveCells' ? 'compatLive' : 'compatSnapshot') + ': ' + finding.evidence;
    details.append(summary, evidence);
    if (finding.formula) { const code = document.createElement('pre'); code.textContent = finding.formula; details.append(code); }
    item.append(details); list.append(item);
  }
  container.append(list);
}
