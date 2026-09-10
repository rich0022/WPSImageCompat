import { cellAddress } from './dispimg-formula';

export interface DamagedImageCell { worksheetName: string; address: string; imageId?: string; value: string; kind: 'json' | 'description' }

/** Recognizes only the exact metadata format emitted by older add-in releases. */
export function damagedImageCell(value: unknown, address: string): { imageId: string } | undefined {
  if (typeof value !== 'string') return undefined;
  try {
    const data: unknown = JSON.parse(value);
    if (!data || typeof data !== 'object') return undefined;
    const raw = data as Record<string, unknown>;
    return typeof raw.imageId === 'string' && raw.imageId.startsWith('ID_') && raw.address === address &&
      typeof raw.fitInsideCell === 'boolean' ? { imageId: raw.imageId } : undefined;
  } catch { return undefined; }
}
export function damagedImageDescriptionCell(value: unknown, address: string): boolean {
  return value === `WPS Image Compat converted image for ${address}.` || value === `WPS Image Compat preview for ${address}.`;
}
/**
 * A native in-cell image does not expose its backing metadata consistently:
 * different Excel builds may return it through formulas, values, or display
 * text. Accept only an exact, address-bound payload from any of those fields.
 */
export function damagedImageMetadataValue(candidates: readonly unknown[], address: string):
  { imageId?: string; value: string; kind: DamagedImageCell['kind'] } | undefined {
  for (const candidate of candidates) {
    const metadata = damagedImageCell(candidate, address);
    if (metadata && typeof candidate === 'string') return { ...metadata, value: candidate, kind: 'json' };
    if (damagedImageDescriptionCell(candidate, address) && typeof candidate === 'string') {
      return { value: candidate, kind: 'description' };
    }
  }
  return undefined;
}

/** Read-only scan for cells accidentally replaced by old add-in image metadata. */
export async function scanDamagedImageCells(): Promise<DamagedImageCell[]> {
  return Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    const cells: DamagedImageCell[] = [];
    for (const sheet of sheets.items) {
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load('rowIndex,columnIndex,rowCount,columnCount');
      await context.sync();
      if (used.isNullObject) continue;
      const batchRows = Math.max(1, Math.floor(5000 / used.columnCount));
      for (let offset = 0; offset < used.rowCount; offset += batchRows) {
        const rowCount = Math.min(batchRows, used.rowCount - offset);
        const batch = sheet.getRangeByIndexes(used.rowIndex + offset, used.columnIndex, rowCount, used.columnCount);
        batch.load('formulas,values,text');
        await context.sync();
        for (let row = 0; row < rowCount; row++) for (let column = 0; column < used.columnCount; column++) {
          const address = cellAddress(used.rowIndex + offset + row, used.columnIndex + column);
          const metadata = damagedImageMetadataValue([
            batch.formulas[row]?.[column], batch.values[row]?.[column], batch.text[row]?.[column],
          ], address);
          if (metadata) cells.push({ worksheetName: sheet.name, address, ...metadata });
        }
      }
    }
    return cells;
  });
}
