import type { DispimgCell, ImageMapping, WpsImageParseResult } from '../types/wps';

export function mapImages(cells: DispimgCell[], parsed: WpsImageParseResult): ImageMapping[] {
  const errors = new Set(parsed.issues.filter(issue => issue.kind === 'error').map(issue => issue.imageId));
  return cells.map(cell => {
    if (errors.has(cell.imageId)) return { cell, status: 'error' };
    const resource = parsed.resources.get(cell.imageId);
    return { cell, resource, status: resource ? 'found' : 'missing' };
  });
}
