import { detectWorkbook } from './workbook-detector';
import { renderImages } from './image-renderer';
import type { PreviewMode } from './image-renderer';
import type { PreviewSettings } from './image-layout';
import { WorkbookError } from '../utils/errors';

const services = { detect: detectWorkbook, render: renderImages };
/** Always rescan; never use stale cached mappings from a prior workbook snapshot. */
export async function showWorkbookImages(mode: PreviewMode, settings: PreviewSettings,
  onStatus?: (status: string) => void, dependencies = services) {
  onStatus?.('Scanning workbook and reading WPS images…');
  const detection = await dependencies.detect();
  if (!detection.mappings || detection.resourceError) {
    throw new WorkbookError('RESOURCES_UNAVAILABLE', detection.resourceError ?? 'Image resources could not be checked. Existing previews were kept.');
  }
  onStatus?.(mode === 'refresh' ? 'Preparing replacement previews…' : 'Showing images…');
  const preview = await dependencies.render(detection.mappings, settings, mode);
  return { detection, preview };
}
