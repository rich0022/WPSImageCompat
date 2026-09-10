import assert from 'node:assert/strict';
import test from 'node:test';
import { collectViewableImages } from '../src/taskpane/image-viewer';
import type { ImageMapping } from '../src/types/wps';

function mapping(status: ImageMapping['status'], mimeType = 'image/png', base64 = 'AA=='): ImageMapping {
  return { status, cell: { worksheetName: 'Images', address: 'B3', imageId: 'ID_1', formula: '=DISPIMG("ID_1",1)' },
    resource: { imageId: 'ID_1', relationshipId: 'rId1', mediaPath: 'xl/media/a.png', mimeType, base64 } };
}

test('collects only unique, locally parsed PNG and JPEG images for viewing', () => {
  const images = collectViewableImages([mapping('found'), mapping('found'), mapping('missing'), mapping('found', 'image/gif')]);
  assert.deepEqual(images, [{ key: 'Images\u0000B3\u0000ID_1', label: 'Images!B3 · ID_1', src: 'data:image/png;base64,AA==' }]);
});

test('does not make a viewer image from a missing resource or absent mapping list', () => {
  assert.deepEqual(collectViewableImages([mapping('found', 'image/jpeg', ''), mapping('error')]), []);
  assert.deepEqual(collectViewableImages(undefined), []);
});
