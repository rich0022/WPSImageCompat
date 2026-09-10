import type { CellBox } from './image-layout';
export const PREVIEW_PREFIX = 'WPSIMG_';
export const PREVIEW_MARKER = 'WPS Image Compat preview v1';
export const CONVERTED_PREFIX = 'WPSCONVERT_';
export const CONVERTED_MARKER = 'WPS Image Compat converted v1';
export interface ShapeMetadata {
  imageId: string; address: string; fitInsideCell?: boolean; offsetLeft?: number; offsetTop?: number;
  original?: { left: number; top: number; width: number; height: number; placement: 'TwoCell' | 'Absolute' };
  zoomed?: boolean;
}
export interface PreviewShapeInfo extends CellBox { id: string; name: string; altTextTitle: string; altTextDescription: string; }
const encodedPrefix = 'v2.';
function toBase64(value: string): string {
  const bytes = new TextEncoder().encode(value); let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/u, '');
}
function fromBase64(value: string): string | undefined {
  try {
    const padded = value.replaceAll('-', '+').replaceAll('_', '/') + '='.repeat((4 - value.length % 4) % 4);
    return new TextDecoder().decode(Uint8Array.from(atob(padded), char => char.charCodeAt(0)));
  } catch { return undefined; }
}
function encodedMetadata(name: string, prefix: string): ShapeMetadata | undefined {
  if (!name.startsWith(prefix + encodedPrefix)) return undefined;
  const value = fromBase64(name.slice((prefix + encodedPrefix).length).split('.')[0] ?? '');
  try {
    return normalizeMetadata(value && JSON.parse(value));
  } catch { return undefined; }
}
function normalizeMetadata(data: unknown): ShapeMetadata | undefined {
  if (!data || typeof data !== 'object') return undefined;
  const raw = data as Record<string, unknown>;
  const imageId = typeof raw.i === 'string' ? raw.i : typeof raw.imageId === 'string' ? raw.imageId : undefined;
  if (!imageId) return undefined;
  const original = Array.isArray(raw.o) && raw.o.length === 5 && typeof raw.o[0] === 'number' && typeof raw.o[1] === 'number' &&
    typeof raw.o[2] === 'number' && typeof raw.o[3] === 'number' && (raw.o[4] === 'TwoCell' || raw.o[4] === 'Absolute') ?
    { left: raw.o[0], top: raw.o[1], width: raw.o[2], height: raw.o[3], placement: raw.o[4] } : undefined;
  return { imageId, address: typeof raw.a === 'string' ? raw.a : typeof raw.address === 'string' ? raw.address : '',
    fitInsideCell: typeof raw.f === 'boolean' ? raw.f : typeof raw.fitInsideCell === 'boolean' ? raw.fitInsideCell : undefined,
    offsetLeft: typeof raw.x === 'number' ? raw.x : typeof raw.offsetLeft === 'number' ? raw.offsetLeft : undefined,
    offsetTop: typeof raw.y === 'number' ? raw.y : typeof raw.offsetTop === 'number' ? raw.offsetTop : undefined,
    original, zoomed: raw.z === 1 || raw.zoomed === true };
}
/** Reads versioned metadata from the name, with a legacy fallback only for existing add-in shapes. */
export function shapeMetadata(shape: Pick<PreviewShapeInfo, 'name' | 'altTextTitle' | 'altTextDescription'>): ShapeMetadata | undefined {
  const prefix = shape.name.startsWith(PREVIEW_PREFIX) ? PREVIEW_PREFIX : shape.name.startsWith(CONVERTED_PREFIX) ? CONVERTED_PREFIX : undefined;
  if (!prefix) return undefined;
  return encodedMetadata(shape.name, prefix) ?? (() => {
    try {
      const data: unknown = JSON.parse(shape.altTextDescription);
      return normalizeMetadata(data);
    } catch { return undefined; }
  })();
}
export function isPreviewShape(shape: Pick<PreviewShapeInfo, 'name' | 'altTextTitle'>): boolean { return shape.name.startsWith(PREVIEW_PREFIX) && shape.altTextTitle === PREVIEW_MARKER; }
export function isConvertedShape(shape: Pick<PreviewShapeInfo, 'name' | 'altTextTitle'>): boolean { return shape.name.startsWith(CONVERTED_PREFIX) && shape.altTextTitle === CONVERTED_MARKER; }
export function previewImageId(shape: PreviewShapeInfo): string | undefined { return isPreviewShape(shape) ? shapeMetadata(shape)?.imageId : undefined; }
/** Uses the current anchor position, so inserted rows do not cause duplicate previews. */
export function matchesPreview(shape: PreviewShapeInfo, imageId: string, cell: CellBox): boolean {
  const data = shapeMetadata(shape);
  if (!isPreviewShape(shape) || data?.imageId !== imageId) return false;
  if (data.zoomed) return true;
  if (data.fitInsideCell === false) return typeof data.offsetLeft === 'number' && typeof data.offsetTop === 'number' &&
    Math.abs(shape.left - data.offsetLeft - cell.left) < 0.5 && Math.abs(shape.top - data.offsetTop - cell.top) < 0.5;
  const x = shape.left + shape.width / 2, y = shape.top + shape.height / 2;
  return cell.width > 0 && cell.height > 0 && x >= cell.left && x < cell.left + cell.width && y >= cell.top && y < cell.top + cell.height;
}
function makeName(prefix: string, metadata: ShapeMetadata): string {
  // Excel limits Shape names. Normal WPS IDs are short; unusually long IDs fall back to
  // position-based matching after reopening rather than exceeding the host limit.
  const imageId = metadata.imageId.length > 96 ? metadata.imageId.slice(0, 24) : metadata.imageId;
  const compact = { i: imageId, a: metadata.address, f: metadata.fitInsideCell, x: metadata.offsetLeft, y: metadata.offsetTop,
    o: metadata.original && [metadata.original.left, metadata.original.top, metadata.original.width, metadata.original.height, metadata.original.placement], z: metadata.zoomed ? 1 : undefined };
  return `${prefix}${encodedPrefix}${toBase64(JSON.stringify(compact))}.${crypto.randomUUID()}`;
}
export function previewName(imageId: string, address: string, _runId: string, metadata: Partial<ShapeMetadata> = {}): string { return makeName(PREVIEW_PREFIX, { imageId, address, ...metadata }); }
export function convertedName(imageId: string, address: string, _runId: string, metadata: Partial<ShapeMetadata> = {}): string { return makeName(CONVERTED_PREFIX, { imageId, address, ...metadata }); }
export function matchesConverted(shape: PreviewShapeInfo, imageId: string, address: string): boolean {
  const data = shapeMetadata(shape);
  return isConvertedShape(shape) && data?.imageId === imageId && data.address === address;
}
export function readableDescription(kind: 'preview' | 'converted', address: string): string { return kind === 'preview' ? `WPS Image Compat preview for ${address}.` : `WPS Image Compat converted image for ${address}.`; }
