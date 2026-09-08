import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readCompressedFile, readWorkbook } from '../src/core/workbook-reader';
import type { WorkbookFile } from '../src/core/workbook-reader';

function mockFile(overrides: Partial<WorkbookFile> = {}) {
  let closes = 0;
  return {
    file: { size: 4, sliceCount: 2,
      getSlice: async (index: number) => ({ index, size: 2, data: index ? [3, 4] : [1, 2] }),
      close: async () => { closes++; }, ...overrides },
    closes: () => closes,
  };
}
test('assembles slices in order and closes the file exactly once', async () => {
  const mock = mockFile();
  assert.deepEqual(await readCompressedFile(async () => mock.file), new Uint8Array([1, 2, 3, 4]));
  assert.equal(mock.closes(), 1);
});
test('closes on slice failure and preserves the original error', async () => {
  const mock = mockFile({ getSlice: async () => { throw new Error('slice failed'); } });
  await assert.rejects(readCompressedFile(async () => mock.file), /slice failed/);
  assert.equal(mock.closes(), 1);
  await assert.rejects(readCompressedFile(async () => ({ ...mock.file,
    close: async () => { throw new Error('close failed'); } })), /slice failed/);
});
test('rejects inconsistent size, order, bytes, incomplete file and still closes', async () => {
  for (const overrides of [
    { size: 101 * 1024 * 1024 }, { sliceCount: 0 }, { size: 5 },
    { getSlice: async () => ({ index: 1, size: 2, data: [1, 2] }) },
    { getSlice: async () => ({ index: 0, size: 1, data: [256] }) },
    { getSlice: async () => ({ index: 0, size: 2, data: 'ab' }) },
  ]) {
    const mock = mockFile(overrides);
    await assert.rejects(readCompressedFile(async () => mock.file));
    assert.equal(mock.closes(), 1);
  }
});
test('propagates acquisition and close failures', async () => {
  await assert.rejects(readCompressedFile(async () => { throw new Error('open failed'); }), /open failed/);
  const mock = mockFile({ close: async () => { throw new Error('close failed'); } });
  await assert.rejects(readCompressedFile(async () => mock.file), /close failed/);
});
test('Office adapter requests Compressed, forwards slices and closes its handle', async () => {
  let closed = 0;
  type Callback = (result: unknown) => void;
  const previous = Object.getOwnPropertyDescriptor(globalThis, 'Office');
  Object.defineProperty(globalThis, 'Office', { configurable: true, value: {
    FileType: { Compressed: 'compressed' }, AsyncResultStatus: { Succeeded: 'succeeded' },
    context: {
      requirements: { isSetSupported: (name: string) => name === 'CompressedFile' },
      document: { getFileAsync: (type: string, options: { sliceSize: number }, callback: Callback) => {
        assert.equal(type, 'compressed'); assert.equal(options.sliceSize, 1048576);
        callback({ status: 'succeeded', value: {
          size: 2, sliceCount: 1,
          getSliceAsync: (index: number, done: Callback) => done({ status: 'succeeded', value: { index, size: 2, data: [80, 75] } }),
          closeAsync: (done: Callback) => { closed++; done({ status: 'succeeded' }); },
        } });
      } },
    },
  } });
  try {
    assert.deepEqual(await readWorkbook(), new Uint8Array([80, 75]));
    assert.equal(closed, 1);
  } finally {
    if (previous) Object.defineProperty(globalThis, 'Office', previous);
    else Reflect.deleteProperty(globalThis, 'Office');
  }
});
