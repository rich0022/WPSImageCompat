import type { CompatibilityReport, FindingCode, FindingSummary } from '../core/compatibility/types';
import { translate, type Locale, type MessageKey } from './i18n';

const titles: Record<FindingCode, MessageKey> = {
  formulaError: 'compatFormulaError', functionMarker: 'compatFunctionMarker', dispimg: 'compatDispimg',
  brokenReference: 'compatBrokenReference', externalFormula: 'compatExternalFormula', externalRecord: 'compatExternalRecord',
};
const categories = { formulas: 'compatFormulas', externalLinks: 'compatLinks', dates: 'compatDates' } as const;
const states = { confirmed: 'compatConfirmed', risk: 'compatRisk', clear: 'compatClear', unchecked: 'compatUnchecked' } as const;
const completion = { complete: 'compatComplete', partial: 'compatPartial', unavailable: 'compatUnavailable' } as const;

/** A concise UI summary; the detailed local download still holds each finding. */
export function summarizeCompatibilityFindings(report: CompatibilityReport): FindingSummary[] {
  if (report.summaries.length) return report.summaries;
  const groups = new Map<string, FindingSummary>();
  for (const finding of report.findings) {
    const key = `${finding.code}:${finding.status}`;
    const group = groups.get(key) ?? { code: finding.code, status: finding.status, count: 0 };
    group.count++;
    if (!group.firstLocation && finding.worksheetName && finding.address) {
      group.firstLocation = { worksheetName: finding.worksheetName, address: finding.address };
    }
    groups.set(key, group);
  }
  return [...groups.values()];
}

export function renderCompatibility(container: HTMLElement, report: CompatibilityReport | undefined, locale: Locale,
  onNavigate?: (location: { worksheetName: string; address: string }) => void): void {
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
  if (report.findings.length) paragraph(t('compatSummaryHint'));
  const list = document.createElement('ul');
  for (const finding of summarizeCompatibilityFindings(report)) {
    const item = document.createElement('li');
    item.append(t(titles[finding.code]) + ' · ' + t(states[finding.status]) + ' · ' +
      translate(locale, 'compatTypeCount', { count: finding.count }));
    if (finding.firstLocation && onNavigate) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'location-button';
      button.textContent = translate(locale, 'compatGoTo', { sheet: finding.firstLocation.worksheetName, address: finding.firstLocation.address });
      button.addEventListener('click', () => onNavigate(finding.firstLocation!));
      item.append(' ', button);
    }
    list.append(item);
  }
  container.append(list);
}
