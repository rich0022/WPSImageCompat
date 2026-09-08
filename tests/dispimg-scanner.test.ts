import { test } from 'node:test';
import assert from 'node:assert/strict';
import { scanWorkbook } from '../src/core/dispimg-scanner';

test('scans offset and hidden sheets, skips empty sheets and formula-like text, remains read-only', async () => {
  const formula = '=DISPIMG("ID_A",1)';
  const compatibility = '=@_xlfn.DISPIMG("ID_A",1)';
  function sheet(name: string, rowIndex: number, columnIndex: number,
    formulas: unknown[][], values: unknown[][]) {
    return Object.freeze({
      name,
      getUsedRangeOrNullObject: () => ({
        isNullObject: formulas.length === 0, rowIndex, columnIndex,
        rowCount: formulas.length, columnCount: formulas[0]?.length ?? 0,
        load() {},
      }),
      getRangeByIndexes: (row: number, column: number, count: number) => {
        assert.equal(column, columnIndex);
        return Object.freeze({
          formulas: formulas.slice(row - rowIndex, row - rowIndex + count),
          values: values.slice(row - rowIndex, row - rowIndex + count),
          load() {},
        });
      },
    });
  }
  const mock = { run: async (callback: (context: unknown) => unknown) => callback({
    workbook: { worksheets: { load() {}, items: [
      sheet('Empty', 0, 0, [], []),
      sheet("中文's sheet", 3, 26, [[formula, formula]], [['#NAME?', formula]]),
      sheet('Hidden', 4, 2, [[compatibility]], [['#NAME?']]),
    ] } }, sync: async () => {},
  }) };
  const previous = Object.getOwnPropertyDescriptor(globalThis, 'Excel');
  Object.defineProperty(globalThis, 'Excel', { configurable: true, value: mock });
  try {
    const result = await scanWorkbook();
    assert.equal(result.scannedWorksheetCount, 3);
    assert.equal(result.worksheetCount, 2);
    assert.deepEqual(result.cells.map(cell => [cell.worksheetName, cell.address, cell.imageId]),
      [["中文's sheet", 'AA4', 'ID_A'], ['Hidden', 'C5', 'ID_A']]);
    assert.deepEqual(await scanWorkbook(), result);
  } finally {
    if (previous) Object.defineProperty(globalThis, 'Excel', previous);
    else Reflect.deleteProperty(globalThis, 'Excel');
  }
});
