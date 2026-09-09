import type { DetectionResult } from './workbook-detector';
import type { PreviewResult } from './image-renderer';
export interface HostCapabilities {
  platform: string;
  version: string;
  scan: boolean;
  compressedFile: boolean;
  shapes: boolean;
  mergedCells: boolean;
}
export interface DiagnosticInput {
  host: HostCapabilities;
  operation: string;
  detection?: DetectionResult;
  preview?: PreviewResult;
  errorCode?: string;
}

/** Intentionally excludes filenames, sheet names, addresses, formulas, IDs and image bytes. */
export function createDiagnosticReport(input: DiagnosticInput) {
  const { detection, preview } = input;
  return {
    formatVersion: 1,
    recordedAt: new Date().toISOString(),
    phase: 4,
    host: input.host,
    operation: input.operation,
    errorCode: input.errorCode ?? detection?.resourceErrorCode ?? null,
    scan: detection ? {
      worksheetsScanned: detection.scan.scannedWorksheetCount,
      worksheetsWithImages: detection.scan.worksheetCount,
      imageCells: detection.scan.cells.length,
      resourcesChecked: detection.mappings !== undefined,
      found: detection.mappings?.filter(item => item.status === 'found').length ?? null,
      missing: detection.mappings?.filter(item => item.status === 'missing').length ?? null,
      errors: detection.mappings?.filter(item => item.status === 'error').length ?? null,
      hasCellImages: detection.parsed?.hasCellImages ?? null,
      issueCodes: [...new Set(detection.parsed?.issues.map(issue => issue.code) ?? [])],
    } : null,
    preview: preview ? { inserted: preview.inserted, existing: preview.existing,
      removed: preview.removed, skipped: preview.skipped, issueCount: preview.issues.length } : null,
  };
}
export function readHostCapabilities(): HostCapabilities {
  const supports = (set: string, version: string) => Office.context.requirements.isSetSupported(set, version);
  return {
    platform: String(Office.context.platform), version: Office.context.diagnostics?.version ?? 'unknown',
    scan: supports('ExcelApi', '1.4'), compressedFile: supports('CompressedFile', '1.1'),
    shapes: supports('ExcelApi', '1.10'), mergedCells: supports('ExcelApi', '1.13'),
  };
}
