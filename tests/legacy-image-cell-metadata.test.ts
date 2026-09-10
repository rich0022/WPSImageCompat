import { test } from 'node:test';
import assert from 'node:assert/strict';
import { damagedImageCell, damagedImageDescriptionCell } from '../src/core/legacy-image-cell-metadata';

test('recognizes a recovery payload at its original or moved cell', () => {
  const value = JSON.stringify({ imageId: 'ID_A', address: 'B2', fitInsideCell: true, offsetLeft: 2, offsetTop: 1 });
  assert.deepEqual(damagedImageCell(value, 'B2'), { imageId: 'ID_A' });
  assert.deepEqual(damagedImageCell(value, 'B3'), { imageId: 'ID_A' });
  assert.equal(damagedImageCell(JSON.stringify({ imageId: 'ID_A', address: 'B2' }), 'B2'), undefined);
  assert.equal(damagedImageCell('ordinary text', 'B2'), undefined);
});
test('recognizes the old shape description only at its matching cell', () => {
  assert.equal(damagedImageDescriptionCell('WPS Image Compat converted image for B2.', 'B2'), true);
  assert.equal(damagedImageDescriptionCell('WPS Image Compat converted image for B2.', 'B3'), false);
});
