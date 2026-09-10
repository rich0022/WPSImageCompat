import { WorkbookError } from '../utils/errors';
import type { PreviewSettings } from './image-layout';
import { scanDamagedImageCells, damagedImageCell, type DamagedImageCell } from './legacy-image-cell-metadata';
import { readWorkbook } from './workbook-reader';
import { parseWpsImages } from './wps-image-parser';
import { mapImages } from './image-mapper';
import { renderImages, type PreviewResult } from './image-renderer';

export interface DamagedImageRecoveryResult { detected: number; restored: number; cleared: number; preview: PreviewResult }

async function clearRecoveredCells(cells: DamagedImageCell[]): Promise<number> {
  if (!cells.length) return 0;
  return Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    const byName = new Map(sheets.items.map(sheet => [sheet.name, sheet]));
    let cleared = 0;
    for (const cell of cells) {
      const sheet = byName.get(cell.worksheetName);
      if (!sheet) continue;
      const range = sheet.getRange(cell.address);
      range.load('formulas,values');
      await context.sync();
      const value = range.values[0]?.[0];
      if (range.formulas[0]?.[0] !== value || !damagedImageCell(value, cell.address)?.imageId) continue;
      range.clear(Excel.ClearApplyTo.contents);
      await context.sync();
      cleared++;
    }
    return cleared;
  });
}

/** Restores only cells containing verified old metadata, and clears them only after a permanent picture exists. */
export async function recoverDamagedImageCells(settings: PreviewSettings): Promise<DamagedImageRecoveryResult> {
  const damaged = await scanDamagedImageCells();
  const empty: PreviewResult = { inserted: 0, existing: 0, removed: 0, skipped: 0, issues: [] };
  if (!damaged.length) return { detected: 0, restored: 0, cleared: 0, preview: empty };
  const bytes = await readWorkbook();
  const parsed = await parseWpsImages(bytes);
  if (!parsed.hasCellImages) throw new WorkbookError('RESOURCES_UNAVAILABLE', 'This workbook no longer contains the WPS image resources needed to repair metadata cells.');
  const mappings = mapImages(damaged.map(cell => ({ ...cell, formula: '' })), parsed);
  const preview = await renderImages(mappings, settings, 'recover');
  const rendered = preview.renderedCells ?? [];
  const restored = new Set(rendered.map(cell => `${cell.worksheetName}\u0000${cell.address}\u0000${cell.imageId}`));
  const cleared = await clearRecoveredCells(damaged.filter(cell => restored.has(`${cell.worksheetName}\u0000${cell.address}\u0000${cell.imageId}`)));
  return { detected: damaged.length, restored: rendered.length, cleared, preview };
}
