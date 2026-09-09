import { throwIfCancelled } from '../../utils/cancellation';
import { WorkbookError } from '../../utils/errors';
import { readWorkbook } from '../workbook-reader';
import { scanCompatibilityCells } from './scanner';
import { parseCompatibilityMetadata } from './metadata-parser';
import { MAX_FINDINGS, type Check, type CompatibilityReport, type Finding, type FindingSummary, type Metadata } from './types';

const defaults = { scan: scanCompatibilityCells, read: readWorkbook, parse: parseCompatibilityMetadata };
function mergeSummaries(summaries: FindingSummary[]): FindingSummary[] {
  const merged = new Map<string, FindingSummary>();
  for (const summary of summaries) {
    const key = `${summary.code}:${summary.status}`;
    const existing = merged.get(key);
    if (existing) existing.count += summary.count;
    else merged.set(key, { ...summary });
  }
  return [...merged.values()];
}
export async function checkCompatibility(signal?: AbortSignal, onProgress?: (count: number) => void,
  services = defaults): Promise<CompatibilityReport> {
  throwIfCancelled(signal);
  const { findings: cellFindings, summaries: cellSummaries, ...cells } = await services.scan(signal, onProgress);
  throwIfCancelled(signal);
  let metadata: Metadata | undefined;
  let metadataError: string | undefined;
  try { metadata = await services.parse(await services.read({ signal }), signal); }
  catch (error) {
    throwIfCancelled(signal);
    if (error instanceof WorkbookError && error.code === 'OPERATION_CANCELLED') throw error;
    metadataError = error instanceof WorkbookError ? error.code : 'METADATA_READ_FAILED';
  }
  throwIfCancelled(signal);
  const externalIds = metadata?.externalReferenceIds ?? [];
  const records: Finding[] = externalIds.slice(0, Math.max(0, MAX_FINDINGS - cellFindings.length)).map(id => ({
    category: 'externalLinks', code: 'externalRecord', status: 'risk', source: 'workbookSnapshot', evidence: id,
  }));
  const all = [...cellFindings, ...records];
  const summaries = mergeSummaries([...cellSummaries,
    ...(externalIds.length ? [{ code: 'externalRecord' as const, status: 'risk' as const, count: externalIds.length }] : [])]);
  const checks: Check[] = (['formulas', 'externalLinks', 'dates'] as const).map(category => {
    const reason = category === 'dates' ? metadataError ?? metadata?.dateReason : category === 'externalLinks'
      ? cells.reasonCode ?? metadataError ?? metadata?.externalReason : cells.reasonCode;
    const complete = category === 'dates' ? !!metadata && !reason : cells.complete && !reason;
    const found = category === 'externalLinks'
      ? summaries.filter(summary => ['brokenReference', 'externalFormula', 'externalRecord'].includes(summary.code))
      : summaries.filter(summary => category === 'formulas'
        ? ['formulaError', 'functionMarker', 'dispimg'].includes(summary.code) : false);
    // When the detail cap is hit, absent category details cannot establish a clean result.
    const detailsOmitted = category !== 'dates' && cells.findingCount > cellFindings.length;
    return { category, completion: complete ? 'complete' : (category === 'dates' || !cells.scannedCells ? 'unavailable' : 'partial'),
      status: found.some(f => f.status === 'confirmed') ? 'confirmed' : found.length ? 'risk' : complete && !detailsOmitted ? 'clear' : 'unchecked',
      ...(reason ? { reasonCode: reason } : detailsOmitted ? { reasonCode: 'DETAIL_LIMIT' } : {}) };
  });
  const findings = all.slice(0, MAX_FINDINGS);
  const findingCount = cells.findingCount + externalIds.length;
  return { schemaVersion: 1, createdAt: new Date().toISOString(), cells,
    metadata: metadata ? { dateSystem: metadata.dateSystem, dateReason: metadata.dateReason,
      externalReason: metadata.externalReason, externalReferenceCount: metadata.externalReferenceIds.length } : undefined, checks,
    findings, summaries, findingCount, omittedFindings: findingCount - findings.length, externalTargetsChecked: false };
}
