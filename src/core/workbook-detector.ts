import { throwIfCancelled } from '../utils/cancellation';
import { scanWorkbook } from './dispimg-scanner';
import type { ScanProgress } from './dispimg-scanner';
import { readWorkbook } from './workbook-reader';
import { parseWpsImages } from './wps-image-parser';
import { mapImages } from './image-mapper';
import type { ScanResult, ImageMapping, WpsImageParseResult } from '../types/wps';
import { WorkbookError } from '../utils/errors';

export interface DetectionResult {
  scan: ScanResult;
  mappings?: ImageMapping[];
  parsed?: WpsImageParseResult;
  resourceError?: string;
  resourceErrorCode?: string;
}
export const detectionServices = { scan: scanWorkbook, read: readWorkbook, parse: parseWpsImages };

/** Coordinates the stages, keeping UI and platform-independent parser separate. */
export async function detectWorkbook(
  onProgress?: (progress: ScanProgress) => void,
  onReading?: () => void,
  services = detectionServices,
  signal?: AbortSignal,
): Promise<DetectionResult> {
  throwIfCancelled(signal);
  const scan = await services.scan(onProgress, signal);
  throwIfCancelled(signal);
  if (scan.cells.length === 0) return { scan, mappings: [] };
  onReading?.();
  try {
    const bytes = await services.read({ signal });
    throwIfCancelled(signal);
    const parsed = await services.parse(bytes);
    throwIfCancelled(signal);
    return { scan, parsed, mappings: mapImages(scan.cells, parsed) };
  } catch (error) {
    if (error instanceof WorkbookError && error.code === 'OPERATION_CANCELLED') throw error;
    return { scan, resourceError: error instanceof WorkbookError ? error.message : 'Workbook image resources could not be read.',
      resourceErrorCode: error instanceof WorkbookError ? error.code : 'RESOURCE_READ_FAILED' };
  }
}
