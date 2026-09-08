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
}
export const detectionServices = { scan: scanWorkbook, read: readWorkbook, parse: parseWpsImages };

/** Coordinates the stages, keeping UI and platform-independent parser separate. */
export async function detectWorkbook(
  onProgress?: (progress: ScanProgress) => void,
  onReading?: () => void,
  services = detectionServices,
): Promise<DetectionResult> {
  const scan = await services.scan(onProgress);
  if (scan.cells.length === 0) return { scan, mappings: [] };
  onReading?.();
  try {
    const parsed = await services.parse(await services.read());
    return { scan, parsed, mappings: mapImages(scan.cells, parsed) };
  } catch (error) {
    return { scan, resourceError: error instanceof WorkbookError ? error.message : 'Workbook image resources could not be read.' };
  }
}
