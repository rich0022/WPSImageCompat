import { test } from 'node:test';
import assert from 'node:assert/strict';
import { noStoreResponse } from '../src/worker';

test('update metadata response is never cached', async () => {
  const response = noStoreResponse(new Response('{"version":"0.8.0"}', {
    status: 200, headers: { 'Content-Type': 'application/json', ETag: '"original"' },
  }));
  assert.equal(response.status, 200);
  assert.equal(response.headers.get('Cache-Control'), 'no-store, no-cache, max-age=0, must-revalidate');
  assert.equal(response.headers.get('ETag'), '"original"');
  assert.equal(await response.text(), '{"version":"0.8.0"}');
});
