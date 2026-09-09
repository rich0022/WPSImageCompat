import { throwIfCancelled } from '../../utils/cancellation';
import { WorkbookError } from '../../utils/errors';
import { readWorkbook } from '../workbook-reader';
import { scanCompatibilityCells } from './scanner';
import { parseCompatibilityMetadata } from './metadata-parser';
import { MAX_FINDINGS, type Check, type CompatibilityReport, type Finding, type Metadata } from './types';

const defaults = { scan: scanCompatibilityCells, read: readWorkbook, parse: parseCompatibilityMetadata };
export async function checkCompatibility(signal?: AbortSignal, onProgress?: (count: number) => void,
  services = defaults): Promise<CompatibilityReport> {
  throwIfCancelled(signal);
  const { findings: cellFindings, ...cells } = await services.scan(signal, onProgress);
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
  const records: Finding[] = (metadata?.externalReferenceIds ?? []).map(id => ({
    category: 'externalLinks', code: 'externalRecord', status: 'risk', source: 'workbookSnapshot', evidence: id,
  }));
  const all = [...cellFindings, ...records];
  const checks: Check[] = (['formulas', 'externalLinks', 'dates'] as const).map(category => {
    const reason = category === 'dates' ? metadataError ?? metadata?.dateReason : category === 'externalLinks'
      ? cells.reasonCode ?? metadataError ?? metadata?.externalReason : cells.reasonCode;
    const complete = category === 'dates' ? !!metadata && !reason : cells.complete && !reason;
    const found = all.filter(finding => finding.category === category);
    // When the detail cap is hit, absent category details cannot establish a clean result.
    const detailsOmitted = category !== 'dates' && cells.findingCount > cellFindings.length;
    return { category, completion: complete ? 'complete' : (category === 'dates' || !cells.scannedCells ? 'unavailable' : 'partial'),
      status: found.some(f => f.status === 'confirmed') ? 'confirmed' : found.length ? 'risk' : complete && !detailsOmitted ? 'clear' : 'unchecked',
      ...(reason ? { reasonCode: reason } : detailsOmitted ? { reasonCode: 'DETAIL_LIMIT' } : {}) };
  });
  const findings = all.slice(0, MAX_FINDINGS);
  const findingCount = cells.findingCount + records.length;
  return { schemaVersion: 1, createdAt: new Date().toISOString(), cells,
    metadata: metadata ? { dateSystem: metadata.dateSystem, dateReason: metadata.dateReason,
      externalReason: metadata.externalReason, externalReferenceCount: metadata.externalReferenceIds.length } : undefined, checks,
    findings, findingCount, omittedFindings: findingCount - findings.length, externalTargetsChecked: false };
}
