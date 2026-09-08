import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { parseWpsImages, resolveMediaPath } from '../src/core/wps-image-parser';
import { mapImages } from '../src/core/image-mapper';
import { fixture, imagesXml, cellImage, relationship, relationshipsXml, PNG } from './fixtures/wps-workbook';
const parse = async (zip = fixture()) => parseWpsImages(await zip.generateAsync({ type: 'uint8array' }));

test('maps WPS IDs through relationships to exact base64 and content type', async () => {
  const parsed = await parse();
  assert.equal(parsed.hasCellImages, true);
  assert.deepEqual(parsed.issues, []);
  assert.deepEqual(parsed.resources.get('ID_A'), { imageId: 'ID_A', relationshipId: 'rId23',
    mediaPath: 'xl/media/image17.png', mimeType: 'image/png', base64: PNG });
});
test('ordinary Excel without cellimages is a normal empty result', async () => {
  const result = await parse(new JSZip().file('xl/workbook.xml', '<workbook/>'));
  assert.equal(result.hasCellImages, false);
  assert.equal(result.resources.size, 0);
  assert.deepEqual(result.issues, []);
});
test('accepts different namespace prefixes, self-closing root and decoded attributes', async () => {
  const zip = fixture().file('xl/cellimages.xml', imagesXml(cellImage('ID_A&amp;B')).replaceAll('etc:', 'wps:').replace('xmlns:etc', 'xmlns:wps'));
  assert.ok((await parse(zip)).resources.has('ID_A&B'));
  assert.equal((await parse(fixture().file('xl/cellimages.xml', '<cellImages/>'))).resources.size, 0);
});
test('missing relationships and media are reported per image without crashing', async () => {
  for (const path of ['xl/_rels/cellimages.xml.rels', 'xl/media/image17.png']) {
    const zip = fixture(); zip.remove(path);
    const result = await parse(zip);
    assert.equal(result.resources.size, 0);
    assert.ok(result.issues.some(issue => issue.imageId === 'ID_A' && issue.kind === 'missing'));
  }
});
test('duplicate image IDs and relationship IDs are errors, never arbitrarily selected', async () => {
  const duplicateImages = fixture().file('xl/cellimages.xml', imagesXml(cellImage() + cellImage()));
  const duplicateRels = fixture().file('xl/_rels/cellimages.xml.rels', relationshipsXml(relationship() + relationship()));
  for (const zip of [duplicateImages, duplicateRels]) {
    const parsed = await parse(zip);
    assert.equal(parsed.resources.size, 0);
    assert.ok(parsed.issues.some(issue => issue.kind === 'error'));
    assert.equal(mapImages([{ worksheetName: 'Sheet1', address: 'A1', formula: '', imageId: 'ID_A' }], parsed)[0]?.status, 'error');
  }
});
test('resolves relative, root-relative and percent-encoded OPC media paths', () => {
  for (const target of ['media/image17.png', '/xl/media/image17.png', '../xl/media/image17.png']) {
    assert.equal(resolveMediaPath(target), 'xl/media/image17.png');
  }
  assert.equal(resolveMediaPath('media/my%20image.png'), 'xl/media/my image.png');
  for (const target of ['../../outside.png', 'https://example.com/a.png', '//example.com/a.png',
    'media\\a.png', 'media/a.png?x=1', 'media/%2e%2e/%2e%2e/private.png', '%zz']) {
    assert.equal(resolveMediaPath(target), undefined);
  }
});
test('rejects external and invalid targets without fetching them', async () => {
  for (const item of [relationship('https://example.com/a.png', 'TargetMode="External"'), relationship('../../outside.png')]) {
    const parsed = await parse(fixture().file('xl/_rels/cellimages.xml.rels', relationshipsXml(item)));
    assert.equal(parsed.resources.size, 0);
    assert.equal(parsed.issues[0]?.kind, 'error');
  }
});
test('rejects ZIP entries JSZip normalized from traversal paths', async () => {
  const zip = fixture(); zip.remove('xl/media/image17.png');
  zip.file('bad/../xl/media/image17.png', PNG, { base64: true });
  await assert.rejects(parse(zip), { code: 'UNSAFE_ZIP_PATH' });
});
test('invalid ZIP, malformed XML, wrong roots and DTDs fail with clear errors', async () => {
  await assert.rejects(parseWpsImages(new Uint8Array([1, 2, 3])), { code: 'INVALID_XLSX' });
  for (const xml of ['<cellImages>', '<wrong/>', '<!DOCTYPE cellImages [<!ENTITY x "hello">]><cellImages/>']) {
    await assert.rejects(parse(fixture().file('xl/cellimages.xml', xml)), { code: 'INVALID_XML' });
  }
});
test('supports UTF-16 XML and Content Types overrides', async () => {
  const zip = fixture();
  zip.file('xl/cellimages.xml', Buffer.concat([Buffer.from([255, 254]), Buffer.from(imagesXml(), 'utf16le')]));
  zip.file('[Content_Types].xml', '<Types><Override PartName="/xl/media/image17.png" ContentType="image/custom"/></Types>');
  assert.equal((await parse(zip)).resources.get('ID_A')?.mimeType, 'image/custom');
});
test('empty media is an error while other valid images remain available', async () => {
  const zip = fixture().file('xl/cellimages.xml', imagesXml(cellImage() + cellImage('ID_B', 'rId24')));
  zip.file('xl/_rels/cellimages.xml.rels', relationshipsXml(relationship() + relationship('media/empty.png').replace('rId23', 'rId24')));
  zip.file('xl/media/empty.png', new Uint8Array());
  const parsed = await parse(zip);
  assert.ok(parsed.resources.has('ID_A'));
  assert.equal(parsed.issues[0]?.code, 'EMPTY_MEDIA');
});
test('bounds decompressed image size including highly compressed ZIP entries', async () => {
  const zip = fixture().file('xl/media/image17.png', new Uint8Array(20 * 1024 * 1024 + 1));
  const binary = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  await assert.rejects(parseWpsImages(binary), { code: 'RESOURCE_LIMIT' });
});
test('mapping preserves repeated cell locations and identifies absent IDs', async () => {
  const parsed = await parse();
  const cells = ['ID_A', 'ID_A', 'ID_ABSENT'].map((imageId, i) => ({
    worksheetName: 'Sheet1', address: `A${i + 1}`, formula: '=DISPIMG(...)', imageId,
  }));
  const mapped = mapImages(cells, parsed);
  assert.deepEqual(mapped.map(item => item.status), ['found', 'found', 'missing']);
  assert.equal(mapped[0]?.resource, mapped[1]?.resource);
});
