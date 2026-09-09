import { afterEach, beforeEach, test } from 'node:test';
import assert from 'node:assert/strict';
import { officeAsync } from '../src/utils/office-async';
import { readWorkbook } from '../src/core/workbook-reader';

let previous: PropertyDescriptor | undefined;
beforeEach(() => {
  previous = Object.getOwnPropertyDescriptor(globalThis, 'Office');
  Object.defineProperty(globalThis, 'Office', { configurable: true, value: {
    AsyncResultStatus: { Succeeded: 'succeeded' }, FileType: { Compressed: 'compressed' },
  } });
});
afterEach(() => {
  if (previous) Object.defineProperty(globalThis, 'Office', previous);
  else Reflect.deleteProperty(globalThis, 'Office');
});
const success = <T>(value: T) => ({ status: 'succeeded', value }) as unknown as Office.AsyncResult<T>;

test('callback success, host failure and synchronous failure settle without a later timeout', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  assert.equal(await officeAsync<number>(done => done(success(7))), 7);
  await assert.rejects(officeAsync(done => done({ status: 'failed', error: { code: 42,
    message: 'private workbook name' } } as unknown as Office.AsyncResult<unknown>)), {
    code: 'OFFICE_FILE', message: 'Excel file operation failed (42).',
  });
  await assert.rejects(officeAsync(() => { throw new Error('sync failure'); }), /sync failure/);
  t.mock.timers.tick(60000);
});
test('a missing callback times out and a late acquired handle is released exactly once', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  let callback!: (result: Office.AsyncResult<string>) => void;
  const released: string[] = [];
  const pending = officeAsync<string>(done => { callback = done; }, {
    timeoutMs: 10, onLateSuccess: value => { released.push(value); },
  });
  const rejected = assert.rejects(pending, { code: 'OFFICE_TIMEOUT' });
  t.mock.timers.tick(10);
  await rejected;
  callback(success('file handle'));
  callback(success('file handle'));
  assert.deepEqual(released, ['file handle']);
});
test('pre-cancelled operations are never invoked; cancellation releases a late result', async () => {
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(officeAsync(() => assert.fail('must not start'), { signal: controller.signal }),
    { code: 'OPERATION_CANCELLED' });
  const active = new AbortController();
  let callback!: (result: Office.AsyncResult<number>) => void;
  let released = 0;
  const pending = officeAsync<number>(done => { callback = done; }, {
    signal: active.signal, onLateSuccess: value => { released += value; },
  });
  const rejected = assert.rejects(pending, { code: 'OPERATION_CANCELLED' });
  active.abort();
  await rejected;
  callback(success(1));
  assert.equal(released, 1);
});
test('Excel adapter closes a file arriving after cancellation without requesting any slice', async () => {
  const controller = new AbortController();
  let callback!: (result: Office.AsyncResult<Office.File>) => void;
  let closed = 0;
  Object.assign(Office, { context: {
    requirements: { isSetSupported: () => true },
    document: { getFileAsync: (_type: unknown, _options: unknown, done: typeof callback) => { callback = done; } },
  } });
  const pending = readWorkbook({ signal: controller.signal });
  const rejected = assert.rejects(pending, { code: 'OPERATION_CANCELLED' });
  controller.abort();
  await rejected;
  callback(success({
    size: 2, sliceCount: 1,
    getSliceAsync: () => assert.fail('must not read a cancelled file'),
    closeAsync: (done: (result: Office.AsyncResult<void>) => void) => { closed++; done(success(undefined)); },
  } as Office.File));
  await Promise.resolve();
  assert.equal(closed, 1);
});
