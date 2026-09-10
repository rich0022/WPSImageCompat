import { test } from 'node:test';
import assert from 'node:assert/strict';
import { damagedImageCell } from '../src/core/legacy-image-cell-metadata';

test('recognizes only the legacy metadata value at its recorded cell', () => {
  const value = JSON.stringify({ imageId: 'ID_A', address: 'B2', fitInsideCell: true, offsetLeft: 2, offsetTop: 1 });
  assert.deepEqual(damagedImageCell(value, 'B2'), { imageId: 'ID_A' });
  assert.equal(damagedImageCell(value, 'B3'), undefined);
  assert.equal(damagedImageCell(JSON.stringify({ imageId: 'ID_A', address: 'B2' }), 'B2'), undefined);
  assert.equal(damagedImageCell('ordinary text', 'B2'), undefined);
});
