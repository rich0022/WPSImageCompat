import { test } from 'node:test';
import assert from 'node:assert/strict';
import worker, { noStoreResponse } from '../src/worker';

test('update metadata response is never cached', async () => {
  const response = noStoreResponse(new Response('{"version":"0.8.0"}', {
    status: 200, headers: { 'Content-Type': 'application/json', ETag: '"original"' },
  }));
  assert.equal(response.status, 200);
  assert.equal(response.headers.get('Cache-Control'), 'no-store, no-cache, max-age=0, must-revalidate');
  assert.equal(response.headers.get('ETag'), '"original"');
  assert.equal(await response.text(), '{"version":"0.8.0"}');
});

test('manifest requests use the current asset revision instead of a stale edge key', async () => {
  let assetUrl = '';
  const response = await worker.fetch(new Request('https://wpsimagecompat.fogce.workers.dev/manifest.xml'), {
    ASSETS: {
      async fetch(request: Request): Promise<Response> {
        assetUrl = request.url;
        return new Response('<OfficeApp/>');
      },
    },
  });
  assert.equal(new URL(assetUrl).searchParams.get('__manifestRevision'), '1.5.1.0');
  assert.equal(response.headers.get('Cache-Control'), 'no-store, no-cache, max-age=0, must-revalidate');
});
