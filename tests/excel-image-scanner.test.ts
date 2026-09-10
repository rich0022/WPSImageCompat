import { test } from 'node:test';
import assert from 'node:assert/strict';
import { canScanExcelImages, scanExcelImages } from '../src/core/excel-image-scanner';
import { CONVERTED_MARKER } from '../src/core/preview-identity';
import { normalizeManagedWorksheetImageMetadata } from '../src/core/worksheet-image-zoom';

function restore(key: 'Excel' | 'Office', descriptor: PropertyDescriptor | undefined): void {
  if (descriptor) Object.defineProperty(globalThis, key, descriptor);
  else Reflect.deleteProperty(globalThis, key);
}

test('counts only native Excel image shapes without changing them', async () => {
  const oldExcel = Object.getOwnPropertyDescriptor(globalThis, 'Excel');
  const oldOffice = Object.getOwnPropertyDescriptor(globalThis, 'Office');
  const first = { name: 'First', shapes: { items: [{ type: 'Image' }, { type: 'Line' }], load() {} } };
  const second = { name: 'Second', shapes: { items: [{ type: 'Image' }, { type: 'Image' }], load() {} } };
  Object.defineProperty(globalThis, 'Office', { configurable: true, value: {
    context: { requirements: { isSetSupported: () => true } },
  } });
  Object.defineProperty(globalThis, 'Excel', { configurable: true, value: {
    run: async (callback: (context: unknown) => unknown) => callback({
      workbook: { worksheets: { items: [first, second], load() {} } }, sync: async () => {},
    }),
  } });
  try {
    assert.equal(canScanExcelImages(), true);
    assert.deepEqual(await scanExcelImages(), { imageCount: 3, worksheetCount: 2, supported: true });
    assert.equal(first.shapes.items.length, 2);
  } finally {
    restore('Excel', oldExcel);
    restore('Office', oldOffice);
  }
});

test('returns an unavailable result on hosts without picture-shape support', async () => {
  const oldOffice = Object.getOwnPropertyDescriptor(globalThis, 'Office');
  Object.defineProperty(globalThis, 'Office', { configurable: true, value: {
    context: { requirements: { isSetSupported: () => false } },
  } });
  try {
    assert.equal(canScanExcelImages(), false);
    assert.deepEqual(await scanExcelImages(), { imageCount: 0, worksheetCount: 0, supported: false });
  } finally { restore('Office', oldOffice); }
});

test('migrates legacy add-in image metadata into its name and clears the description', async () => {
  const oldExcel = Object.getOwnPropertyDescriptor(globalThis, 'Excel');
  const oldOffice = Object.getOwnPropertyDescriptor(globalThis, 'Office');
  const shape = {
    name: 'WPSCONVERT_legacy', altTextTitle: CONVERTED_MARKER,
    altTextDescription: JSON.stringify({ imageId: 'ID_A', address: 'B2', fitInsideCell: true, offsetLeft: 2, offsetTop: 1 }),
    left: 2, top: 1, width: 90, height: 50, placement: 'TwoCell',
  };
  const sheet = { name: 'Sheet1', shapes: { items: [shape], load() {} } };
  Object.defineProperty(globalThis, 'Office', { configurable: true, value: {
    context: { requirements: { isSetSupported: () => true } },
  } });
  Object.defineProperty(globalThis, 'Excel', { configurable: true, value: {
    Placement: { absolute: 'Absolute' },
    run: async (callback: (context: unknown) => unknown) => callback({
      workbook: { worksheets: { items: [sheet], load() {} } }, sync: async () => {},
    }),
  } });
  try {
    await normalizeManagedWorksheetImageMetadata();
    assert.ok(shape.name.startsWith('WPSCONVERT_v2.'));
    assert.equal(shape.altTextDescription, '');
  } finally {
    restore('Excel', oldExcel);
    restore('Office', oldOffice);
  }
});
