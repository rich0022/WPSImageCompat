import type { WpsImageResource } from '../types/wps';

/**
 * Keeps image bytes only while this task pane remains open. Excel may remove
 * WPS-specific OOXML parts after a save, but a just-converted image must still
 * be recoverable when it is manually placed in a cell.
 */
const resources = new Map<string, WpsImageResource>();

export function rememberImageResources(items: Iterable<WpsImageResource | undefined>): void {
  for (const item of items) {
    if (!item?.imageId || !item.base64) continue;
    resources.set(item.imageId, { ...item });
  }
}

export function cachedImageResource(imageId: string): WpsImageResource | undefined {
  const item = resources.get(imageId);
  return item && { ...item };
}

export function clearImageCache(): void { resources.clear(); }
