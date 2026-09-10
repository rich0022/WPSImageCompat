import { test } from 'node:test';
import assert from 'node:assert/strict';
import { canApplyUpdate, createUpdateChecker, parseRelease } from '../src/core/update-checker';
import type { ReleaseInfo } from '../src/core/update-checker';
const current = { version: '0.5.0', buildId: 'build-a' };
test('validates release metadata and rejects malformed values or unexpected data', () => {
  assert.deepEqual(parseRelease({ ...current, workbook: 'never-forward-this' }), current);
  for (const value of [null, {}, [], { version: '<script>', buildId: 'x' }, { ...current, buildId: '../x' }]) {
    assert.equal(parseRelease(value), undefined);
  }
});
test('detects same-version rebuilds, clears a stale notice, and uses a cache-busting same-origin request', async () => {
  const updates: (ReleaseInfo | undefined)[] = [];
  let remote = { ...current, buildId: 'build-b' };
  const fetcher: typeof fetch = async (url, options) => {
    assert.equal(url, '/version.json?update=123');
    assert.equal(options?.cache, 'no-store');
    assert.equal(options?.credentials, 'omit');
    assert.equal(options?.body, undefined);
    return new Response(JSON.stringify(remote));
  };
  const check = createUpdateChecker(current, release => updates.push(release), fetcher, 10000, () => 123);
  await check();
  remote = current;
  await check();
  assert.deepEqual(updates, [{ version: '0.5.0', buildId: 'build-b' }, undefined]);
});
test('network errors, absent metadata, bad JSON and timeout do not disrupt workbook operations', async t => {
  for (const fetcher of [
    async () => { throw new Error('offline'); },
    async () => new Response('', { status: 404 }),
    async () => new Response('not json'),
    async () => new Response('{}'),
  ]) await createUpdateChecker(current, () => assert.fail('no release expected'), fetcher)();
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const check = createUpdateChecker(current, () => assert.fail('must not notify'), async (_url, options) =>
    new Promise((_resolve, reject) => options?.signal?.addEventListener('abort', () => reject(new Error('timeout')))), 10);
  const pending = check();
  t.mock.timers.tick(10);
  await pending;
});
test('overlapping checks are coalesced and updates cannot be applied during an operation', async () => {
  let finish!: (response: Response) => void;
  let calls = 0;
  const check = createUpdateChecker(current, () => {}, async () => {
    calls++; return new Promise(resolve => { finish = resolve; });
  });
  const pending = check();
  await check();
  assert.equal(calls, 1);
  finish(new Response(JSON.stringify(current)));
  await pending;
  assert.equal(canApplyUpdate(true, current), false);
  assert.equal(canApplyUpdate(false), false);
  assert.equal(canApplyUpdate(false, current), true);
});
