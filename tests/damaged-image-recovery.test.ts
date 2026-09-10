import { test } from 'node:test';
import assert from 'node:assert/strict';
import { scanDamagedImageCells } from '../src/core/legacy-image-cell-metadata';

function restore(key: 'Excel', descriptor: PropertyDescriptor | undefined): void {
  if (descriptor) Object.defineProperty(globalThis, key, descriptor);
  else Reflect.deleteProperty(globalThis, key);
}

test('finds a valid moved in-cell recovery payload when Excel exposes it only through formulas', async () => {
  const oldExcel = Object.getOwnPropertyDescriptor(globalThis, 'Excel');
  const value = JSON.stringify({ imageId: 'ID_A', address: 'C5', fitInsideCell: true, offsetLeft: 4, offsetTop: 1 });
  const used = {
    isNullObject: false, rowIndex: 0, columnIndex: 0, rowCount: 2, columnCount: 3, load() {},
  };
  const sheet = {
    name: 'Sheet1',
    getUsedRangeOrNullObject() { return used; },
    getRangeByIndexes() { return { formulas: [['', '', ''], ['', '', value]], values: [['', '', ''], ['', '', 'Picture']], text: [['', '', ''], ['', '', 'Picture']], load() {} }; },
  };
  Object.defineProperty(globalThis, 'Excel', { configurable: true, value: {
    run: async (callback: (context: unknown) => unknown) => callback({
      workbook: { worksheets: { items: [sheet], load() {} } }, sync: async () => {},
    }),
  } });
  try {
    assert.deepEqual(await scanDamagedImageCells(), [{ worksheetName: 'Sheet1', address: 'C2', imageId: 'ID_A', value, kind: 'json' }]);
  } finally { restore('Excel', oldExcel); }
});
