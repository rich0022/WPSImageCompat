import { test } from 'node:test';
import assert from 'node:assert/strict';
import { showWorkbookImages } from '../src/core/preview-controller';
import { DEFAULT_PREVIEW_SETTINGS } from '../src/core/image-layout';
const scan = { cells: [], worksheetCount: 0, scannedWorksheetCount: 1 };
test('cancelled detection never mutates Shapes; rendering announces its non-cancellable boundary', async () => {
  const controller = new AbortController();
  await assert.rejects(showWorkbookImages('refresh', DEFAULT_PREVIEW_SETTINGS, undefined, {
    detect: async () => { controller.abort(); return { scan, mappings: [] }; },
    render: async () => assert.fail('must not change Shapes'),
  }, { signal: controller.signal }), { code: 'OPERATION_CANCELLED' });
  const order: string[] = [];
  await showWorkbookImages('show', DEFAULT_PREVIEW_SETTINGS, undefined, {
    detect: async () => ({ scan, mappings: [] }),
    render: async () => { order.push('render'); return { inserted: 0, existing: 0, removed: 0, skipped: 0, issues: [] }; },
  }, { onRendering: () => { order.push('lock cancellation'); } });
  assert.deepEqual(order, ['lock cancellation', 'render']);
});
test('freshly detects on each Show/Refresh and forwards settings', async () => {
  let scans = 0, renders = 0;
  const services = {
    detect: async () => { scans++; return { scan, mappings: [] }; },
    render: async (_mappings: unknown, settings: unknown, mode: unknown) => {
      renders++; assert.equal(settings, DEFAULT_PREVIEW_SETTINGS); assert.ok(mode === 'show' || mode === 'refresh');
      return { inserted: 0, existing: 0, removed: 0, skipped: 0, issues: [] };
    },
  };
  await showWorkbookImages('show', DEFAULT_PREVIEW_SETTINGS, undefined, services);
  await showWorkbookImages('refresh', DEFAULT_PREVIEW_SETTINGS, undefined, services);
  assert.equal(scans, 2); assert.equal(renders, 2);
});
test('failed resource read never invokes shape operations', async () => {
  await assert.rejects(showWorkbookImages('refresh', DEFAULT_PREVIEW_SETTINGS, undefined, {
    detect: async () => ({ scan, resourceError: 'unavailable' }),
    render: async () => { throw new Error('must not render'); },
  }), { code: 'RESOURCES_UNAVAILABLE' });
});
