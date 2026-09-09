import { throwIfCancelled } from '../utils/cancellation';
import { detectWorkbook } from './workbook-detector';
import { renderImages } from './image-renderer';
import type { PreviewMode } from './image-renderer';
import type { PreviewSettings } from './image-layout';
import { WorkbookError } from '../utils/errors';

const services = { detect: detectWorkbook, render: renderImages };
interface PreviewOperationOptions {
  signal?: AbortSignal;
  /** Called before any Shape writes; cancellation is no longer available after this boundary. */
  onRendering?: () => void;
}
export type PreviewStage = 'reading' | 'refreshing' | 'rendering';
/** Always rescan; never use stale cached mappings from a prior workbook snapshot. */
export async function showWorkbookImages(mode: PreviewMode, settings: PreviewSettings,
  onStatus?: (status: PreviewStage) => void, dependencies = services, options: PreviewOperationOptions = {}) {
  const { signal } = options;
  throwIfCancelled(signal);
  onStatus?.('reading');
  const detection = await dependencies.detect(undefined, undefined, undefined, signal);
  throwIfCancelled(signal);
  if (!detection.mappings || detection.resourceError) {
    throw new WorkbookError('RESOURCES_UNAVAILABLE', detection.resourceError ?? 'Image resources could not be checked. Existing previews were kept.');
  }
  options.onRendering?.();
  onStatus?.(mode === 'refresh' ? 'refreshing' : 'rendering');
  const preview = await dependencies.render(detection.mappings, settings, mode);
  return { detection, preview };
}
