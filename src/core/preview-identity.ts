import type { CellBox } from './image-layout';
export const PREVIEW_PREFIX = 'WPSIMG_';
export const PREVIEW_MARKER = 'WPS Image Compat preview v1';
export interface PreviewShapeInfo extends CellBox {
  id: string; name: string; altTextTitle: string; altTextDescription: string;
}
export function isPreviewShape(shape: Pick<PreviewShapeInfo, 'name' | 'altTextTitle'>): boolean {
  return shape.name.startsWith(PREVIEW_PREFIX) && shape.altTextTitle === PREVIEW_MARKER;
}
export function previewImageId(shape: PreviewShapeInfo): string | undefined {
  if (!isPreviewShape(shape)) return undefined;
  try {
    const data: unknown = JSON.parse(shape.altTextDescription);
    if (data && typeof data === 'object' && 'imageId' in data && typeof data.imageId === 'string') return data.imageId;
  } catch { /* A user-edited description is not a match. */ }
  return undefined;
}
/** Uses the current anchor position, so inserted rows do not cause duplicate previews. */
export function matchesPreview(shape: PreviewShapeInfo, imageId: string, cell: CellBox): boolean {
  if (previewImageId(shape) !== imageId) return false;
  // Native-size previews can extend over adjacent cells. Their center is not
  // a reliable anchor, especially when Excel clamps top/left to zero.
  try {
    const data = JSON.parse(shape.altTextDescription) as Record<string, unknown>;
    if (data.fitInsideCell === false) {
      return typeof data.offsetLeft === 'number' && typeof data.offsetTop === 'number' &&
        Math.abs(shape.left - data.offsetLeft - cell.left) < 0.5 &&
        Math.abs(shape.top - data.offsetTop - cell.top) < 0.5;
    }
  } catch { return false; }
  const x = shape.left + shape.width / 2, y = shape.top + shape.height / 2;
  return cell.width > 0 && cell.height > 0 && x >= cell.left && x < cell.left + cell.width &&
    y >= cell.top && y < cell.top + cell.height;
}
export async function previewName(imageId: string, address: string, runId: string): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(imageId));
  const hash = Array.from(new Uint8Array(digest)).slice(0, 12).map(byte => byte.toString(16).padStart(2, '0')).join('');
  return `${PREVIEW_PREFIX}${hash}_${address}_${runId}`;
}
