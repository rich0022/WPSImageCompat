import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { parseExcelEmbeddedImages, resolveOoxmlPart } from '../src/core/excel-drawing-parser';
import { WorkbookError } from '../src/utils/errors';

const relationship = (id: string, type: string, target: string, extra = '') =>
  `<Relationship Id="${id}" Type="${type}" Target="${target}" ${extra}/>`;
async function workbook(missingMedia = false): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('xl/workbook.xml', '<workbook xmlns:r="relationships"><sheets><sheet name="Pictures" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships>${relationship('rId1', 'officeDocument/worksheet', 'worksheets/sheet1.xml')}</Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns:r="relationships"><drawing r:id="rId3"/></worksheet>');
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<Relationships>${relationship('rId3', 'officeDocument/drawing', '../drawings/drawing1.xml')}</Relationships>`);
  zip.file('xl/drawings/drawing1.xml', `<xdr:wsDr xmlns:xdr="drawing" xmlns:r="relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:row>1</xdr:row></xdr:from><xdr:pic><xdr:nvPicPr><xdr:cNvPr name="Product"/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:a="drawing" r:embed="rId7"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor></xdr:wsDr>`);
  zip.file('xl/drawings/_rels/drawing1.xml.rels', `<Relationships>${relationship('rId7', 'officeDocument/image', '../media/image1.png')}</Relationships>`);
  if (!missingMedia) zip.file('xl/media/image1.png', new Uint8Array([137, 80, 78, 71]));
  return zip.generateAsync({ type: 'uint8array' });
}

test('maps Excel drawing anchors to worksheet cells and media paths', async () => {
  const result = await parseExcelEmbeddedImages(await workbook());
  assert.deepEqual(result.images, [{ worksheetName: 'Pictures', address: 'C2', drawingPath: 'xl/drawings/drawing1.xml',
    relationshipId: 'rId7', mediaPath: 'xl/media/image1.png', name: 'Product' }]);
  assert.deepEqual(result.issues, []);
});
test('reports missing drawing media without selecting an arbitrary replacement', async () => {
  const result = await parseExcelEmbeddedImages(await workbook(true));
  assert.equal(result.images.length, 0);
  assert.equal(result.issues[0]?.code, 'MISSING_MEDIA');
});
test('resolves only safe internal OOXML targets below xl', () => {
  assert.equal(resolveOoxmlPart('xl/worksheets/sheet1.xml', '../drawings/drawing1.xml'), 'xl/drawings/drawing1.xml');
  assert.equal(resolveOoxmlPart('xl/drawings/drawing1.xml', '../media/image1.png'), 'xl/media/image1.png');
  assert.equal(resolveOoxmlPart('xl/workbook.xml', '../../private.xml'), undefined);
  assert.equal(resolveOoxmlPart('xl/workbook.xml', 'https://example.com/image.png'), undefined);
});
test('stops decompression when a drawing XML part exceeds the safety limit', async () => {
  const zip = new JSZip();
  zip.file('xl/workbook.xml', `<workbook>${'x'.repeat(4 * 1024 * 1024)}</workbook>`);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  await assert.rejects(() => parseExcelEmbeddedImages(bytes), (error: unknown) =>
    error instanceof WorkbookError && error.code === 'RESOURCE_LIMIT');
});
