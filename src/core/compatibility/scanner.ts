import { throwIfCancelled } from '../../utils/cancellation';
import { cellAddress } from '../dispimg-formula';
import { analyzeFormula } from './formula-rules';
import { MAX_FINDINGS, type CellScan, type Finding, type FindingSummary } from './types';

export const MAX_SCAN_CELLS = 1_000_000;
/** Both dimensions are tiled. No range write or calculation is performed. */
export async function scanCompatibilityCells(signal?: AbortSignal,
  onProgress?: (count: number) => void): Promise<CellScan> {
  throwIfCancelled(signal);
  const result: CellScan = { findings: [], summaries: [], findingCount: 0, confirmedCount: 0, riskCount: 0,
    scannedCells: 0, scannedWorksheets: 0, totalWorksheets: 0, complete: false };
  const summaries = new Map<string, FindingSummary>();
  const record = (finding: Finding): void => {
    result.findingCount++;
    if (finding.status === 'confirmed') result.confirmedCount++; else result.riskCount++;
    if (result.findings.length < MAX_FINDINGS) result.findings.push(finding);
    const key = `${finding.code}:${finding.status}`;
    const summary = summaries.get(key) ?? { code: finding.code, status: finding.status, count: 0 };
    summary.count++;
    if (!summary.firstLocation && finding.worksheetName && finding.address) {
      summary.firstLocation = { worksheetName: finding.worksheetName, address: finding.address };
    }
    summaries.set(key, summary);
  };
  try {
    await Excel.run(async context => {
      const worksheets = context.workbook.worksheets;
      worksheets.load('items/name');
      await context.sync();
      result.totalWorksheets = worksheets.items.length;
      for (const sheet of worksheets.items) {
        throwIfCancelled(signal);
        const used = sheet.getUsedRangeOrNullObject(true);
        used.load('rowIndex,columnIndex,rowCount,columnCount');
        await context.sync();
        throwIfCancelled(signal);
        if (!used.isNullObject) {
          for (let column = 0; column < used.columnCount; column += 100) {
            const width = Math.min(100, used.columnCount - column);
            const rows = Math.floor(5000 / width);
            for (let row = 0; row < used.rowCount; row += rows) {
              throwIfCancelled(signal);
              const height = Math.min(rows, used.rowCount - row);
              if (result.scannedCells + height * width > MAX_SCAN_CELLS) {
                result.reasonCode = 'CELL_LIMIT'; return;
              }
              const range = sheet.getRangeByIndexes(used.rowIndex + row, used.columnIndex + column, height, width);
              range.load('formulas,values,valueTypes');
              await context.sync();
              throwIfCancelled(signal);
              for (let r = 0; r < height; r++) for (let c = 0; c < width; c++) {
                if (range.formulas[r]?.[c] === undefined || range.valueTypes[r]?.[c] === undefined) throw new Error('Incomplete range');
                const findings = analyzeFormula({ worksheetName: sheet.name,
                  address: cellAddress(used.rowIndex + row + r, used.columnIndex + column + c),
                  formula: range.formulas[r]![c], value: range.values[r]?.[c], valueType: range.valueTypes[r]![c]! });
                for (const finding of findings) record(finding);
              }
              result.scannedCells += width * height;
              onProgress?.(result.scannedCells);
            }
          }
        }
        result.scannedWorksheets++;
      }
      result.complete = true;
    });
  } catch {
    throwIfCancelled(signal);
    result.reasonCode = 'CELL_READ_FAILED';
  }
  result.summaries = [...summaries.values()];
  throwIfCancelled(signal);
  return result;
}
