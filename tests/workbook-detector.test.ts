import { test } from 'node:test';
import assert from 'node:assert/strict';
import { detectWorkbook } from '../src/core/workbook-detector';
import { parseWpsImages } from '../src/core/wps-image-parser';
import { fixture } from './fixtures/wps-workbook';
const cell = { worksheetName: 'Sheet1', address: 'A1', formula: '=DISPIMG("ID_A",1)', imageId: 'ID_A' };
const scan = async () => ({ cells: [cell], worksheetCount: 1, scannedWorksheetCount: 1 });
test('coordinates formula scan, ZIP read, parse and mapping', async () => {
  const result = await detectWorkbook(undefined, undefined, {
    scan, read: async () => fixture().generateAsync({ type: 'uint8array' }), parse: parseWpsImages,
  });
  assert.equal(result.mappings?.[0]?.status, 'found');
});
test('resource failures keep formula counts but never report resources as missing', async () => {
  const result = await detectWorkbook(undefined, undefined, {
    scan, read: async () => { throw new Error('failed'); }, parse: parseWpsImages,
  });
  assert.equal(result.scan.cells.length, 1);
  assert.equal(result.mappings, undefined);
  assert.ok(result.resourceError);
});
test('skips binary acquisition when no image formula is found', async () => {
  const result = await detectWorkbook(undefined, undefined, {
    scan: async () => ({ cells: [], worksheetCount: 0, scannedWorksheetCount: 2 }),
    read: async () => { throw new Error('must not be called'); }, parse: parseWpsImages,
  });
  assert.deepEqual(result.mappings, []);
  assert.equal(result.resourceError, undefined);
});
