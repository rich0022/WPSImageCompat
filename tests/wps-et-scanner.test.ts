import { test } from 'node:test';
import assert from 'node:assert/strict';
import { scanWpsEtWorkbook, selectWpsEtPicture } from '../wps-et/src/wps-et-scanner';

function host() {
  const firstShape = { Name: 'Reference photo', ImageUrl: 'https://example.test/image.png', TopLeftCell: { Address: '$C$2' }, selected: false,
    Select() { this.selected = true; } };
  const firstSheet = { Name: 'Pictures', UsedRange: { Row: 2, Column: 2, Formula: [['', '=@_xlfn.DISPIMG("ID_ABC",1)'], ['=SUM(1,2)', '']] },
    Shapes: { Count: 1, Item: () => firstShape }, activated: false, Activate() { this.activated = true; } };
  const secondSheet = { Name: 'Blank', UsedRange: { Row: 1, Column: 1, Formula: '' }, Shapes: { Count: 0, Item: () => undefined } };
  return { Sheets: { Count: 2, Item: (index: number) => [firstSheet, secondSheet][index - 1] }, firstShape, firstSheet };
}

test('reads WPS ET formulas and exposed picture anchors without workbook writes', async () => {
  const fixture = host();
  const result = await scanWpsEtWorkbook(fixture);
  assert.equal(result.worksheetCount, 2);
  assert.equal(result.dispimgCells.length, 1);
  assert.deepEqual(result.dispimgCells[0], { worksheetName: 'Pictures', address: 'C2', formula: '=@_xlfn.DISPIMG("ID_ABC",1)', imageId: 'ID_ABC' });
  assert.deepEqual(result.pictures, [{ worksheetName: 'Pictures', shapeIndex: 1, name: 'Reference photo', address: '$C$2' }]);
  assert.equal(fixture.firstShape.selected, false);
});
test('selects only the picture returned by the scan', async () => {
  const fixture = host();
  const result = await scanWpsEtWorkbook(fixture);
  await selectWpsEtPicture(fixture, result.pictures[0]!);
  assert.equal(fixture.firstSheet.activated, true);
  assert.equal(fixture.firstShape.selected, true);
});
