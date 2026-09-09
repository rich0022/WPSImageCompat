import { test } from 'node:test';
import assert from 'node:assert/strict';
import { navigateToCompatibilityCell } from '../src/core/compatibility/navigator';

test('invalid navigation locations fail before accessing Excel', async () => {
  await assert.rejects(navigateToCompatibilityCell({ worksheetName: 'Sheet 1', address: 'A0' }), { code: 'INVALID_LOCATION' });
});

test('navigation activates only the reported worksheet and selects its reported cell', async () => {
  const events: string[] = [];
  const sheet = { isNullObject: false, load: () => events.push('load'), activate: () => events.push('activate'),
    getRange: (address: string) => ({ select: () => events.push('select:' + address) }) };
  const previous = Object.getOwnPropertyDescriptor(globalThis, 'Excel');
  Object.defineProperty(globalThis, 'Excel', { configurable: true, value: { run: async (callback: (context: unknown) => unknown) => callback({
    workbook: { worksheets: { getItemOrNullObject: (name: string) => { events.push('sheet:' + name); return sheet; } }, },
    sync: async () => events.push('sync'),
  }) } });
  try {
    await navigateToCompatibilityCell({ worksheetName: '报价', address: 'B3' });
    assert.deepEqual(events, ['sheet:报价', 'load', 'sync', 'activate', 'select:B3', 'sync']);
  } finally {
    if (previous) Object.defineProperty(globalThis, 'Excel', previous); else Reflect.deleteProperty(globalThis, 'Excel');
  }
});
