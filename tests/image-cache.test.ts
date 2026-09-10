import { test } from 'node:test';
import assert from 'node:assert/strict';
import { cachedImageResource, clearImageCache, rememberImageResources } from '../src/core/image-cache';

test('keeps only readable image resources for the current task-pane session', () => {
  clearImageCache();
  rememberImageResources([
    { imageId: 'ID_A', relationshipId: 'rId1', mediaPath: 'xl/media/a.png', mimeType: 'image/png', base64: 'IMAGE' },
    { imageId: 'ID_B', relationshipId: 'rId2', mediaPath: 'xl/media/b.png', mimeType: 'image/png' },
  ]);
  assert.deepEqual(cachedImageResource('ID_A'),
    { imageId: 'ID_A', relationshipId: 'rId1', mediaPath: 'xl/media/a.png', mimeType: 'image/png', base64: 'IMAGE' });
  assert.equal(cachedImageResource('ID_B'), undefined);
  clearImageCache();
  assert.equal(cachedImageResource('ID_A'), undefined);
});
