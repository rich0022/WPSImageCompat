import { test } from 'node:test';
import assert from 'node:assert/strict';
import { renderImages, removePreviewImages } from '../src/core/image-renderer';
import { DEFAULT_PREVIEW_SETTINGS } from '../src/core/image-layout';
import { CONVERTED_MARKER, PREVIEW_MARKER } from '../src/core/preview-identity';
import type { ImageMapping } from '../src/types/wps';

function mapping(sheet = 'Sheet1', address = 'A1', imageId = 'ID_A'): ImageMapping {
  return { cell: { worksheetName: sheet, address, imageId, formula: `=DISPIMG("${imageId}",1)` }, status: 'found',
    resource: { imageId, relationshipId: 'rId1', mediaPath: 'xl/media/p.png', mimeType: 'image/png', base64: 'PNG' } };
}
interface MockShape {
  id: string; name: string; altTextTitle: string; altTextDescription: string;
  left: number; top: number; width: number; height: number;
  lockAspectRatio: boolean; visible: boolean; placement: string;
  load(): void; delete(): void;
}
function makeSheet(name = 'Sheet1') {
  let sequence = 0;
  const shapes: MockShape[] = [];
  const cells = new Map<string, ReturnType<typeof makeCell>>();
  const events: string[] = [];
  const controls = { protected: false, failImage: '', failPlacement: false, mutateOnSecondRead: false, failSyncAfterAdd: false };
  function makeShape(shapeName = 'Picture', marked = false): MockShape {
    const id = String(++sequence);
    let placement = 'Absolute';
    const shape = { id, name: shapeName, altTextTitle: marked ? PREVIEW_MARKER : '',
      altTextDescription: JSON.stringify({ imageId: 'ID_A' }),
      left: 1, top: 1, width: 100, height: 50, lockAspectRatio: true, visible: true,
      load() {},
      get placement() { return placement; },
      set placement(value: string) {
        if (controls.failPlacement) throw new Error('placement failure');
        placement = value;
      },
      delete() {
        if (controls.protected) throw new Error('protected');
        events.push(`delete:${id}`);
        const index = shapes.indexOf(shape);
        if (index >= 0) shapes.splice(index, 1);
      },
    };
    shapes.push(shape);
    return shape;
  }
  function makeCell(imageId = 'ID_A', top = 0) {
    let reads = 0;
    const cell = { left: 0, top, width: 100, height: 50, rowHidden: false, columnHidden: false,
      formulas: [[`=DISPIMG("${imageId}",1)`]], values: [['#NAME?']],
      load() { if (++reads === 2 && controls.mutateOnSecondRead) cell.formulas = [['=1']]; },
      getMergedAreasOrNullObject() { return { isNullObject: true, load() {}, areas: { items: [] as unknown[] } }; },
    };
    return cell;
  }
  cells.set('A1', makeCell());
  const sheet = { name, protection: { load() {}, get protected() { return controls.protected; } },
    shapes: { get items() { return shapes; }, load() {}, addImage(base64: string) {
      if (controls.failImage === base64) throw new Error('unsupported image');
      const shape = makeShape(); events.push(`add:${shape.id}`); return shape;
    } },
    getRange(address: string) { const cell = cells.get(address); if (!cell) throw new Error('missing cell'); return cell; },
  };
  return { sheet, shapes, cells, makeShape, makeCell, controls, events };
}
async function withExcel(sheets: ReturnType<typeof makeSheet>[], run: () => Promise<void>, supported = true) {
  const oldExcel = Object.getOwnPropertyDescriptor(globalThis, 'Excel');
  const oldOffice = Object.getOwnPropertyDescriptor(globalThis, 'Office');
  Object.defineProperty(globalThis, 'Office', { configurable: true, value: {
    context: { requirements: { isSetSupported: () => supported } },
  } });
  Object.defineProperty(globalThis, 'Excel', { configurable: true, value: {
    Placement: { twoCell: 'TwoCell' },
    run: async (callback: (context: unknown) => Promise<unknown>) => callback({
      workbook: { worksheets: { items: sheets.map(item => item.sheet), load() {} } }, sync: async () => {
        const failing = sheets.find(item => item.controls.failSyncAfterAdd && item.shapes.some(shape => !shape.visible));
        if (failing) { failing.controls.failSyncAfterAdd = false; throw new Error('host sync failed'); }
      },
    }),
  } });
  try { await run(); }
  finally {
    for (const [key, descriptor] of [['Excel', oldExcel], ['Office', oldOffice]] as const) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor); else Reflect.deleteProperty(globalThis, key);
    }
  }
}
test('inserts positioned image with TwoCell placement, aspect lock and untouched formulas', async () => {
  const sheet = makeSheet();
  await withExcel([sheet], async () => {
    const result = await renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS);
    assert.equal(result.inserted, 1);
    const shape = sheet.shapes[0]!;
    assert.equal(shape.placement, 'TwoCell');
    assert.equal(shape.lockAspectRatio, true);
    assert.equal(shape.altTextTitle, PREVIEW_MARKER);
    assert.ok(shape.name.startsWith('WPSIMG_'));
    assert.equal(shape.visible, true);
    assert.equal(shape.left + shape.width / 2, 50);
    assert.equal(shape.top + shape.height / 2, 25);
    assert.deepEqual(sheet.cells.get('A1')!.formulas, [['=DISPIMG("ID_A",1)']]);
    assert.equal((await renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS)).existing, 1);
    assert.equal(sheet.shapes.length, 1);
  });
});
test('conversion creates a permanent shape, replaces only its preview, and can resume without a duplicate', async () => {
  const sheet = makeSheet();
  await withExcel([sheet], async () => {
    assert.equal((await renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS)).inserted, 1);
    const converted = await renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS, 'convert');
    assert.equal(converted.inserted, 1);
    assert.equal(converted.removed, 1);
    assert.deepEqual(converted.renderedCells?.map(cell => cell.address), ['A1']);
    assert.equal(sheet.shapes.length, 1);
    assert.equal(sheet.shapes[0]!.altTextTitle, CONVERTED_MARKER);
    assert.ok(sheet.shapes[0]!.name.startsWith('WPSCONVERT_'));
    const repeated = await renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS, 'convert');
    assert.equal(repeated.existing, 1);
    assert.equal(sheet.shapes.length, 1);
    assert.equal((await removePreviewImages()).removed, 0);
  });
});
test('recovery renders a permanent image for a verified legacy metadata cell', async () => {
  const sheet = makeSheet();
  const legacy = JSON.stringify({ imageId: 'ID_A', address: 'A1', fitInsideCell: true, offsetLeft: 0, offsetTop: 0 });
  sheet.cells.get('A1')!.formulas = [[legacy]];
  sheet.cells.get('A1')!.values = [[legacy]];
  await withExcel([sheet], async () => {
    const result = await renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS, 'recover');
    assert.equal(result.inserted, 1);
    assert.equal(result.renderedCells?.[0]?.address, 'A1');
    assert.equal(sheet.shapes[0]!.altTextTitle, CONVERTED_MARKER);
  });
});
test('repeated IDs across cells/sheets remain separate and deduplicate after row movement', async () => {
  const first = makeSheet(), second = makeSheet('Hidden sheet');
  first.cells.set('A2', first.makeCell('ID_A', 50));
  await withExcel([first, second], async () => {
    const input = [mapping(), mapping('Sheet1', 'A2'), mapping('Hidden sheet')];
    assert.equal((await renderImages(input, DEFAULT_PREVIEW_SETTINGS)).inserted, 3);
    assert.equal((await renderImages(input, DEFAULT_PREVIEW_SETTINGS)).existing, 3);
    first.cells.get('A1')!.top += 100;
    first.cells.get('A2')!.top += 100;
    first.shapes.forEach(shape => { shape.top += 100; });
    assert.equal((await renderImages(input, DEFAULT_PREVIEW_SETTINGS)).existing, 3);
  });
});
test('refresh builds replacements before deleting old previews and leaves user photos alone', async () => {
  const sheet = makeSheet();
  const old = sheet.makeShape('WPSIMG_old', true), user = sheet.makeShape('Vacation');
  await withExcel([sheet], async () => {
    const result = await renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS, 'refresh');
    assert.equal(result.inserted, 1); assert.equal(result.removed, 1);
    assert.ok(sheet.events[0]!.startsWith('add:'));
    assert.ok(sheet.shapes.includes(user)); assert.ok(!sheet.shapes.includes(old));
    assert.equal((await renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS)).existing, 1);
  });
});
test('failed refresh rolls back new previews and preserves old previews', async () => {
  const sheet = makeSheet();
  const old = sheet.makeShape('WPSIMG_old', true);
  sheet.cells.set('A2', sheet.makeCell('ID_A', 50));
  sheet.controls.failImage = 'BAD';
  const bad = mapping('Sheet1', 'A2'); bad.resource!.base64 = 'BAD';
  await withExcel([sheet], async () => {
    await assert.rejects(renderImages([mapping(), bad], DEFAULT_PREVIEW_SETTINGS, 'refresh'), { code: 'REFRESH_FAILED' });
    assert.deepEqual(sheet.shapes, [old]);
  });
});
test('refresh with missing resources or changed formulas does not delete old images', async () => {
  const sheet = makeSheet(); sheet.makeShape('WPSIMG_old', true);
  await withExcel([sheet], async () => {
    await assert.rejects(renderImages([{ ...mapping(), status: 'missing' }], DEFAULT_PREVIEW_SETTINGS, 'refresh'), { code: 'REFRESH_UNRESOLVED' });
    sheet.cells.get('A1')!.formulas = [['=1']];
    await assert.rejects(renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS, 'refresh'), { code: 'REFRESH_PREFLIGHT' });
    assert.equal(sheet.shapes.length, 1); assert.deepEqual(sheet.events, []);
  });
});
test('show skips hidden cells, protected sheets, and unsupported image formats', async () => {
  const sheet = makeSheet();
  await withExcel([sheet], async () => {
    sheet.cells.get('A1')!.rowHidden = true;
    assert.equal((await renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS)).skipped, 1);
    sheet.cells.get('A1')!.rowHidden = false; sheet.controls.protected = true;
    assert.equal((await renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS)).skipped, 1);
    sheet.controls.protected = false;
    const image = mapping(); image.resource!.mimeType = 'image/svg+xml';
    assert.equal((await renderImages([image], DEFAULT_PREVIEW_SETTINGS)).skipped, 1);
    assert.equal(sheet.shapes.length, 0);
  });
});
test('layout failure cleans up newly created shape', async () => {
  const sheet = makeSheet(); sheet.controls.failPlacement = true;
  await withExcel([sheet], async () => {
    const result = await renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS);
    assert.equal(result.inserted, 0); assert.equal(result.skipped, 1);
    assert.equal(sheet.shapes.length, 0);
  });
});
test('rechecks formula just before writing after preparation', async () => {
  const sheet = makeSheet(); sheet.controls.mutateOnSecondRead = true;
  await withExcel([sheet], async () => {
    assert.equal((await renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS)).skipped, 1);
    assert.equal(sheet.shapes.length, 0);
  });
});
test('remove requires both ownership markers and covers hidden sheets', async () => {
  const first = makeSheet(), second = makeSheet('Hidden');
  first.makeShape('WPSIMG_owned', true); second.makeShape('WPSIMG_owned', true);
  const user = first.makeShape('User photo'), coincidental = first.makeShape('WPSIMG_my_own_photo');
  await withExcel([first, second], async () => {
    assert.equal((await removePreviewImages()).removed, 2);
    assert.deepEqual(first.shapes, [user, coincidental]); assert.deepEqual(second.shapes, []);
    assert.equal((await removePreviewImages()).removed, 0);
  });
});
test('empty refresh removes stale previews and protected removal reports partial failure', async () => {
  const sheet = makeSheet(); sheet.makeShape('WPSIMG_old', true);
  await withExcel([sheet], async () => {
    sheet.controls.protected = true;
    assert.equal((await removePreviewImages()).skipped, 1);
    assert.equal(sheet.shapes.length, 1);
    sheet.controls.protected = false;
    assert.equal((await renderImages([], DEFAULT_PREVIEW_SETTINGS, 'refresh')).removed, 1);
  });
});
test('unsupported Excel host blocks all shape mutations', async () => {
  const sheet = makeSheet();
  await withExcel([sheet], async () => {
    await assert.rejects(renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS), { code: 'SHAPES_UNSUPPORTED' });
    await assert.rejects(removePreviewImages(), { code: 'SHAPES_UNSUPPORTED' });
    assert.equal(sheet.shapes.length, 0);
  }, false);
});
test('uses merged-cell bounds when supported', async () => {
  const sheet = makeSheet();
  sheet.cells.get('A1')!.getMergedAreasOrNullObject = () => ({ isNullObject: false, load() {},
    areas: { items: [{ left: 0, top: 0, width: 200, height: 100 }] } });
  await withExcel([sheet], async () => {
    await renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS);
    const shape = sheet.shapes[0]!;
    assert.equal(shape.left + shape.width / 2, 100);
    assert.equal(shape.top + shape.height / 2, 50);
  });
});
test('native-size overlapping images remain attached to their own cells without duplicates', async () => {
  const sheet = makeSheet();
  sheet.cells.get('A1')!.width = 20;
  sheet.cells.get('A1')!.height = 10;
  sheet.cells.set('A2', { ...sheet.makeCell('ID_A', 10), width: 20, height: 10 });
  await withExcel([sheet], async () => {
    const settings = { keepAspectRatio: true, fitInsideCell: false };
    const input = [mapping(), mapping('Sheet1', 'A2')];
    assert.equal((await renderImages(input, settings)).inserted, 2);
    assert.equal((await renderImages(input, settings)).existing, 2);
    assert.equal(sheet.shapes.length, 2);
  });
});
test('asynchronous host failure after shape creation rolls back the staged preview', async () => {
  const sheet = makeSheet(); const old = sheet.makeShape('WPSIMG_old', true);
  sheet.controls.failSyncAfterAdd = true;
  await withExcel([sheet], async () => {
    await assert.rejects(renderImages([mapping()], DEFAULT_PREVIEW_SETTINGS, 'refresh'), { code: 'REFRESH_FAILED' });
    assert.deepEqual(sheet.shapes, [old]);
  });
});
