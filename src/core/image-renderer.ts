import type { ImageMapping } from '../types/wps';
import type { CellBox, PreviewSettings } from './image-layout';
import { imageLayout } from './image-layout';
import { CONVERTED_MARKER, convertedName, isPreviewShape, matchesConverted, matchesPreview, PREVIEW_MARKER, previewName, readableDescription } from './preview-identity';
import { parseDispimgFormula } from './dispimg-formula';
import { WorkbookError } from '../utils/errors';
import { damagedImageCell } from './legacy-image-cell-metadata';

export interface PreviewIssue { location: string; message: string; code?: string }
export interface PreviewResult { inserted: number; existing: number; removed: number; skipped: number; issues: PreviewIssue[]; renderedCells?: ImageMapping['cell'][] }
export type PreviewMode = 'show' | 'refresh' | 'convert' | 'recover';
export function canRenderImages(): boolean {
  return typeof Office !== 'undefined' && Office.context.requirements.isSetSupported('ExcelApi', '1.10');
}
function requireShapes(): void {
  if (!canRenderImages()) throw new WorkbookError('SHAPES_UNSUPPORTED', 'Image previews require ExcelApi 1.10. Update desktop Excel.');
}
function message(error: unknown): string {
  return error instanceof WorkbookError ? error.message : 'Excel could not complete the shape operation. Check worksheet protection and retry.';
}
interface Target { mapping: ImageMapping; sheet: Excel.Worksheet; cell: Excel.Range; box: CellBox; name: string; previewToReplace?: Excel.Shape }
interface Created { shape: Excel.Shape; target: Target }

/** Excel Shape operations only. Formulas are rechecked but never written. */
export async function renderImages(
  mappings: ImageMapping[], settings: PreviewSettings, mode: PreviewMode = 'show',
  onProgress?: (completed: number, total: number) => void,
): Promise<PreviewResult> {
  requireShapes();
  if (mode === 'refresh' && mappings.some(item => item.status !== 'found' || !item.resource?.base64)) {
    throw new WorkbookError('REFRESH_UNRESOLVED', 'Refresh stopped: some image resources are missing or invalid. Existing previews were kept.');
  }
  return Excel.run(async context => {
    const result: PreviewResult = { inserted: 0, existing: 0, removed: 0, skipped: 0, issues: [] };
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    const byName = new Map(sheets.items.map(sheet => [sheet.name, sheet]));
    const relevant = mode === 'refresh' ? sheets.items : sheets.items.filter(sheet => mappings.some(item => item.cell.worksheetName === sheet.name));
    for (const sheet of relevant) {
      sheet.shapes.load('items/id,items/name,items/altTextTitle,items/altTextDescription,items/left,items/top,items/width,items/height');
      sheet.protection.load('protected');
    }
    await context.sync();
    const oldShapes = relevant.flatMap(sheet => sheet.shapes.items.filter(isPreviewShape));
    if (mode === 'refresh' && relevant.some(sheet => sheet.protection.protected &&
        (sheet.shapes.items.some(isPreviewShape) || mappings.some(item => item.cell.worksheetName === sheet.name)))) {
      throw new WorkbookError('PROTECTED_SHEET', 'Refresh stopped: an affected worksheet is protected. Existing previews were kept.');
    }
    const runId = crypto.randomUUID();
    const permanent = mode === 'convert' || mode === 'recover';
    const targets: Target[] = [];
    const claimed = new Set<string>();
    const mergedSupported = Office.context.requirements.isSetSupported('ExcelApi', '1.13');
    for (const mapping of mappings) {
      const location = `${mapping.cell.worksheetName}!${mapping.cell.address}`;
      try {
        if (mapping.status !== 'found' || !mapping.resource?.base64) throw new WorkbookError('MISSING_IMAGE', 'No readable image resource.');
        if (!['image/png', 'image/jpeg'].includes(mapping.resource.mimeType.toLowerCase())) {
          throw new WorkbookError('UNSUPPORTED_IMAGE', 'Preview supports PNG and JPEG only.');
        }
        const sheet = byName.get(mapping.cell.worksheetName);
        if (!sheet) throw new WorkbookError('SHEET_CHANGED', 'Worksheet no longer exists. Scan again.');
        if (sheet.protection.protected) throw new WorkbookError('PROTECTED_SHEET', 'Worksheet is protected.');
        const cell = sheet.getRange(mapping.cell.address);
        cell.load('left,top,width,height,formulas,values,rowHidden,columnHidden');
        const merged = mergedSupported ? cell.getMergedAreasOrNullObject() : undefined;
        if (merged) merged.load('areas/items/left,areas/items/top,areas/items/width,areas/items/height');
        await context.sync();
        const expectedCell = mode === 'recover'
          ? damagedImageCell(cell.values[0]?.[0], mapping.cell.address)?.imageId === mapping.cell.imageId
          : cell.formulas[0]?.[0] !== cell.values[0]?.[0] && parseDispimgFormula(cell.formulas[0]?.[0]) === mapping.cell.imageId;
        if (!expectedCell) {
          throw new WorkbookError('CELL_CHANGED', 'The DISPIMG cell changed. Scan again.');
        }
        const bounds = merged && !merged.isNullObject ? merged.areas.items[0]! : cell;
        const box = { left: bounds.left, top: bounds.top, width: bounds.width, height: bounds.height };
        if (cell.rowHidden || cell.columnHidden || box.width <= 0 || box.height <= 0) {
          result.skipped++;
          result.issues.push({ location, code: 'HIDDEN_CELL', message: 'Hidden or zero-size cell was skipped. Unhide it and refresh to display its image.' });
          continue;
        }
        const existing = permanent
          ? sheet.shapes.items.find(shape => !claimed.has(`${sheet.name}:${shape.id}`) && matchesConverted(shape, mapping.cell.imageId, mapping.cell.address))
          : sheet.shapes.items.find(shape => !claimed.has(`${sheet.name}:${shape.id}`) && matchesPreview(shape, mapping.cell.imageId, box));
        if (mode === 'show' && existing) { claimed.add(`${sheet.name}:${existing.id}`); result.existing++; continue; }
        if (permanent && existing) {
          claimed.add(`${sheet.name}:${existing.id}`); result.existing++;
          (result.renderedCells ??= []).push(mapping.cell);
          continue;
        }
        const previewToReplace = mode === 'convert'
          ? sheet.shapes.items.find(shape => !claimed.has(`${sheet.name}:${shape.id}`) && matchesPreview(shape, mapping.cell.imageId, box)) : undefined;
        targets.push({ mapping, sheet, cell, box,
          name: permanent ? await convertedName(mapping.cell.imageId, mapping.cell.address, runId) : await previewName(mapping.cell.imageId, mapping.cell.address, runId),
          previewToReplace });
      } catch (error) {
        if (mode === 'refresh') throw new WorkbookError('REFRESH_PREFLIGHT', `Refresh stopped at ${location}: ${message(error)} Existing previews were kept.`);
        result.skipped++;
        result.issues.push({ location, code: error instanceof WorkbookError ? error.code : 'OPERATION_FAILED', message: message(error) });
      }
    }
    const created: Created[] = [];
    const cleanup = async (items: Created[]): Promise<boolean> => {
      let complete = true;
      for (const item of items) {
        try { item.shape.delete(); await context.sync(); }
        catch { complete = false; }
      }
      return complete;
    };
    let completed = 0;
    for (const target of targets) {
      let added: Created | undefined;
      try {
        // Recheck immediately before writing, after potentially slow workbook preparation.
        target.cell.load('formulas,values');
        await context.sync();
        const expectedCell = mode === 'recover'
          ? damagedImageCell(target.cell.values[0]?.[0], target.mapping.cell.address)?.imageId === target.mapping.cell.imageId
          : target.cell.formulas[0]?.[0] !== target.cell.values[0]?.[0] &&
            parseDispimgFormula(target.cell.formulas[0]?.[0]) === target.mapping.cell.imageId;
        if (!expectedCell) {
          throw new WorkbookError('CELL_CHANGED', 'The DISPIMG cell changed during preview preparation.');
        }
        const shape = target.sheet.shapes.addImage(target.mapping.resource!.base64!);
        added = { shape, target };
        created.push(added);
        shape.name = target.name;
        shape.altTextTitle = permanent ? CONVERTED_MARKER : PREVIEW_MARKER;
        shape.altTextDescription = readableDescription(permanent ? 'converted' : 'preview', target.mapping.cell.address);
        shape.visible = false;
        shape.load('width,height');
        await context.sync();
        const layout = imageLayout(target.box, shape, settings);
        if (!layout) throw new WorkbookError('HIDDEN_CELL', 'The target cell has no visible size.');
        shape.lockAspectRatio = false;
        shape.width = layout.width;
        shape.height = layout.height;
        shape.left = layout.left;
        shape.top = layout.top;
        const metadata = { fitInsideCell: settings.fitInsideCell, offsetLeft: layout.left - target.box.left, offsetTop: layout.top - target.box.top,
          original: { left: layout.left, top: layout.top, width: layout.width, height: layout.height, placement: 'TwoCell' as const } };
        shape.name = permanent ? convertedName(target.mapping.cell.imageId, target.mapping.cell.address, runId, metadata) :
          previewName(target.mapping.cell.imageId, target.mapping.cell.address, runId, metadata);
        shape.altTextDescription = readableDescription(permanent ? 'converted' : 'preview', target.mapping.cell.address);
        shape.placement = Excel.Placement.twoCell;
        shape.lockAspectRatio = settings.keepAspectRatio;
        shape.visible = true;
        await context.sync();
        result.inserted++;
        if (permanent) {
          (result.renderedCells ??= []).push(target.mapping.cell);
          if (target.previewToReplace) { target.previewToReplace.delete(); await context.sync(); result.removed++; }
        }
      } catch (error) {
        if (mode === 'refresh') {
          const cleaned = await cleanup(created);
          throw new WorkbookError('REFRESH_FAILED', `${message(error)} Old previews were kept. ${cleaned ? 'New previews were removed.' : 'Some new previews could not be removed; use Remove Preview Images after fixing worksheet access.'}`);
        }
        const cleaned = added ? await cleanup([added]) : true;
        result.skipped++;
        result.issues.push({ location: `${target.sheet.name}!${target.mapping.cell.address}`,
          code: !cleaned ? 'CLEANUP_FAILED' : error instanceof WorkbookError ? error.code : 'OPERATION_FAILED',
          message: `${message(error)}${cleaned ? '' : ' A temporary preview could not be removed.'}` });
      }
      onProgress?.(++completed, targets.length);
    }
    // Only after all replacements succeed do we retire the old snapshots.
    if (mode === 'refresh') {
      for (const shape of oldShapes) {
        try { shape.delete(); await context.sync(); result.removed++; }
        catch { result.issues.push({ location: 'Workbook', code: 'CLEANUP_FAILED', message: 'An old preview could not be removed. Run Refresh again after fixing worksheet access.' }); }
      }
    }
    return result;
  });
}

/** Requires both reserved prefix and ownership marker; never deletes arbitrary user pictures. */
export async function removePreviewImages(): Promise<PreviewResult> {
  requireShapes();
  return Excel.run(async context => {
    const result: PreviewResult = { inserted: 0, existing: 0, removed: 0, skipped: 0, issues: [] };
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    for (const sheet of sheets.items) sheet.shapes.load('items/name,items/altTextTitle');
    await context.sync();
    for (const sheet of sheets.items) {
      for (const shape of sheet.shapes.items.filter(isPreviewShape)) {
        try { shape.delete(); await context.sync(); result.removed++; }
        catch { result.skipped++; result.issues.push({ location: sheet.name, code: 'CLEANUP_FAILED', message: 'A preview could not be removed. Check worksheet protection.' }); }
      }
    }
    return result;
  });
}
