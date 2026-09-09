import { throwIfCancelled } from '../utils/cancellation';
import type { DispimgCell, ScanResult } from '../types/wps';
import { cellAddress, parseDispimgFormula } from './dispimg-formula';

export interface ScanProgress { worksheetName: string; scannedCells: number }

/** Read-only scan including hidden sheets. Bounded batches avoid one huge payload. */
export async function scanWorkbook(
  onProgress?: (progress: ScanProgress) => void,
  signal?: AbortSignal,
): Promise<ScanResult> {
  throwIfCancelled(signal);
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load('items/name');
    await context.sync();
    const cells: DispimgCell[] = [];
    let scannedCells = 0;
    for (const sheet of worksheets.items) {
      throwIfCancelled(signal);
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load('rowIndex,columnIndex,rowCount,columnCount');
      await context.sync();
      if (used.isNullObject) continue;
      const batchRows = Math.max(1, Math.floor(5000 / used.columnCount));
      for (let offset = 0; offset < used.rowCount; offset += batchRows) {
        throwIfCancelled(signal);
        const rowCount = Math.min(batchRows, used.rowCount - offset);
        const batch = sheet.getRangeByIndexes(
          used.rowIndex + offset, used.columnIndex, rowCount, used.columnCount,
        );
        batch.load('formulas,values');
        await context.sync();
        throwIfCancelled(signal);
        for (let row = 0; row < batch.formulas.length; row++) {
          const values = batch.formulas[row]!;
          for (let column = 0; column < values.length; column++) {
            const formula: unknown = values[column];
            // Constants are returned unchanged in Range.formulas. Exclude text
            // that merely looks like a formula (including apostrophe-prefixed input).
            if (formula === batch.values[row]?.[column]) continue;
            const imageId = parseDispimgFormula(formula);
            if (imageId !== undefined && typeof formula === 'string') {
              cells.push({ worksheetName: sheet.name,
                address: cellAddress(used.rowIndex + offset + row, used.columnIndex + column),
                formula, imageId });
            }
          }
        }
        scannedCells += rowCount * used.columnCount;
        onProgress?.({ worksheetName: sheet.name, scannedCells });
      }
    }
    return { cells, worksheetCount: new Set(cells.map(cell => cell.worksheetName)).size,
      scannedWorksheetCount: worksheets.items.length };
  });
}
