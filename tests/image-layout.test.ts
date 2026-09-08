import { test } from 'node:test';
import assert from 'node:assert/strict';
import { imageLayout, DEFAULT_PREVIEW_SETTINGS } from '../src/core/image-layout';
import { isPreviewShape, matchesPreview, PREVIEW_MARKER, previewName } from '../src/core/preview-identity';
const box = { left: 100, top: 50, width: 100, height: 60 };
test('fits wide and tall images proportionally and centers inside cell with padding', () => {
  for (const natural of [{ width: 400, height: 100 }, { width: 100, height: 400 }]) {
    const layout = imageLayout(box, natural, DEFAULT_PREVIEW_SETTINGS)!;
    assert.ok(Math.abs(layout.width / layout.height - natural.width / natural.height) < 0.00001);
    assert.equal(layout.left + layout.width / 2, box.left + box.width / 2);
    assert.equal(layout.top + layout.height / 2, box.top + box.height / 2);
    assert.ok(layout.left >= box.left + 1 && layout.top >= box.top + 1);
    assert.ok(layout.width <= 98 && layout.height <= 58);
  }
});
test('optional stretch fills the cell; native size can exceed it but never uses negative anchors', () => {
  assert.deepEqual(imageLayout(box, { width: 400, height: 100 }, { keepAspectRatio: false, fitInsideCell: true }),
    { left: 101, top: 51, width: 98, height: 58 });
  assert.deepEqual(imageLayout({ ...box, left: 0, top: 0 }, { width: 400, height: 100 },
    { keepAspectRatio: true, fitInsideCell: false }), { left: 0, top: 0, width: 400, height: 100 });
});
test('handles very small or hidden cells and rejects invalid geometry', () => {
  assert.equal(imageLayout({ ...box, width: 0 }, { width: 10, height: 10 }, DEFAULT_PREVIEW_SETTINGS), undefined);
  assert.ok(imageLayout({ ...box, width: 0.1 }, { width: 10, height: 10 }, DEFAULT_PREVIEW_SETTINGS)!.width > 0);
  assert.throws(() => imageLayout(box, { width: NaN, height: 10 }, DEFAULT_PREVIEW_SETTINGS));
});
test('ownership requires reserved prefix and marker; user shapes are protected', () => {
  assert.equal(isPreviewShape({ name: 'WPSIMG_user-photo', altTextTitle: '' }), false);
  assert.equal(isPreviewShape({ name: 'Photo', altTextTitle: PREVIEW_MARKER }), false);
  assert.equal(isPreviewShape({ name: 'WPSIMG_preview', altTextTitle: PREVIEW_MARKER }), true);
});
test('names distinguish ID, cell and execution, without unsafe ID characters', async () => {
  const name = await previewName('ID_"/中文'.repeat(100), 'AA3', 'run1');
  assert.ok(name.startsWith('WPSIMG_'));
  assert.ok(name.length < 255);
  assert.notEqual(name, await previewName('other', 'AA3', 'run1'));
  assert.notEqual(name, await previewName('ID_"/中文'.repeat(100), 'AA4', 'run1'));
});
test('deduplicates by current position and image ID after rows move', () => {
  const shape = { id: '1', name: 'WPSIMG_old_A1', altTextTitle: PREVIEW_MARKER,
    altTextDescription: JSON.stringify({ imageId: 'ID_A' }), ...box };
  assert.equal(matchesPreview(shape, 'ID_A', box), true);
  assert.equal(matchesPreview(shape, 'ID_B', box), false);
  assert.equal(matchesPreview(shape, 'ID_A', { ...box, top: 500 }), false);
});
