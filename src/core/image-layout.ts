export interface ImageSize { width: number; height: number }
export interface CellBox extends ImageSize { left: number; top: number }
export interface PreviewSettings { keepAspectRatio: boolean; fitInsideCell: boolean }
export const DEFAULT_PREVIEW_SETTINGS: PreviewSettings = { keepAspectRatio: true, fitInsideCell: true };

/** All coordinates are Excel points (not CSS pixels); zero-size cells are hidden. */
export function imageLayout(cell: CellBox, natural: ImageSize, settings: PreviewSettings): CellBox | undefined {
  if (![cell.left, cell.top, cell.width, cell.height, natural.width, natural.height].every(Number.isFinite) ||
      cell.left < 0 || cell.top < 0 || cell.width < 0 || cell.height < 0 || natural.width <= 0 || natural.height <= 0) {
    throw new Error('Invalid cell or image dimensions.');
  }
  if (cell.width === 0 || cell.height === 0) return undefined;
  const padding = Math.min(1, cell.width / 10, cell.height / 10);
  let { width, height } = natural;
  if (settings.fitInsideCell) {
    const availableWidth = cell.width - 2 * padding, availableHeight = cell.height - 2 * padding;
    if (settings.keepAspectRatio) {
      const scale = Math.min(availableWidth / width, availableHeight / height);
      width *= scale;
      height *= scale;
    } else { width = availableWidth; height = availableHeight; }
  }
  return { width, height,
    left: Math.max(0, cell.left + (cell.width - width) / 2),
    top: Math.max(0, cell.top + (cell.height - height) / 2) };
}
