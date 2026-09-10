import { CONVERTED_MARKER, convertedName, isConvertedShape, isPreviewShape, PREVIEW_MARKER, previewName, readableDescription, shapeMetadata } from './preview-identity';
import { WorkbookError } from '../utils/errors';

export interface ManagedWorksheetImage { worksheetName: string; shapeId: string; address: string; kind: 'preview' | 'converted'; zoomed: boolean; }
type ShapeKind = ManagedWorksheetImage['kind'];
function supported(): boolean { return typeof Office !== 'undefined' && Office.context.requirements.isSetSupported('ExcelApi', '1.10'); }
function kind(shape: Pick<Excel.Shape, 'name' | 'altTextTitle'>): ShapeKind | undefined {
  return isPreviewShape(shape) ? 'preview' : isConvertedShape(shape) ? 'converted' : undefined;
}
function requireShapes(): void { if (!supported()) throw new WorkbookError('SHAPES_UNSUPPORTED', 'Image previews require ExcelApi 1.10. Update desktop Excel.'); }
export async function listManagedWorksheetImages(): Promise<ManagedWorksheetImage[]> {
  requireShapes();
  return Excel.run(async context => {
    const sheets = context.workbook.worksheets; sheets.load('items/name'); await context.sync();
    for (const sheet of sheets.items) sheet.shapes.load('items/id,items/name,items/altTextTitle,items/altTextDescription');
    await context.sync();
    return sheets.items.flatMap(sheet => sheet.shapes.items.flatMap(shape => {
      const shapeKind = kind(shape); const metadata = shapeMetadata(shape);
      return shapeKind && metadata ? [{ worksheetName: sheet.name, shapeId: shape.id, address: metadata.address, kind: shapeKind, zoomed: !!metadata.zoomed }] : [];
    }));
  });
}
/** Enlarges one owned image over the worksheet, or restores its saved cell-sized geometry. */
export async function toggleWorksheetImageZoom(item: ManagedWorksheetImage): Promise<boolean> {
  requireShapes();
  return Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(item.worksheetName);
    sheet.shapes.load('items/id,items/name,items/altTextTitle,items/altTextDescription,items/left,items/top,items/width,items/height,items/placement');
    await context.sync();
    const shape = sheet.shapes.items.find(candidate => candidate.id === item.shapeId);
    const shapeKind = shape && kind(shape); const metadata = shape && shapeMetadata(shape);
    if (!shape || !shapeKind || !metadata) throw new WorkbookError('IMAGE_NOT_FOUND', 'The selected WPS Image Compat picture is no longer available. Refresh the list.');
    const original = metadata.original ?? { left: shape.left, top: shape.top, width: shape.width, height: shape.height,
      placement: shape.placement === Excel.Placement.absolute ? 'Absolute' as const : 'TwoCell' as const };
    const next = { ...metadata, original, zoomed: !metadata.zoomed };
    if (metadata.zoomed) {
      shape.lockAspectRatio = false; shape.width = original.width; shape.height = original.height; shape.left = original.left; shape.top = original.top;
      shape.placement = original.placement === 'Absolute' ? Excel.Placement.absolute : Excel.Placement.twoCell;
    } else {
      const scale = Math.max(2, Math.min(4, Math.max(360 / Math.max(shape.width, 1), 260 / Math.max(shape.height, 1))));
      const width = shape.width * scale, height = shape.height * scale;
      shape.placement = Excel.Placement.absolute; shape.lockAspectRatio = false; shape.width = width; shape.height = height;
      shape.left = Math.max(0, shape.left - (width - original.width) / 2); shape.top = Math.max(0, shape.top - (height - original.height) / 2);
      shape.lockAspectRatio = true; shape.setZOrder(Excel.ShapeZOrder.bringToFront);
    }
    shape.name = shapeKind === 'preview' ? previewName(metadata.imageId, metadata.address, crypto.randomUUID(), next) :
      convertedName(metadata.imageId, metadata.address, crypto.randomUUID(), next);
    shape.altTextTitle = shapeKind === 'preview' ? PREVIEW_MARKER : CONVERTED_MARKER;
    shape.altTextDescription = readableDescription(shapeKind, metadata.address);
    await context.sync();
    return next.zoomed;
  });
}
