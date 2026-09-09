import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { analyzeFormula } from '../src/core/compatibility/formula-rules';
import { parseCompatibilityMetadata } from '../src/core/compatibility/metadata-parser';
import { checkCompatibility } from '../src/core/compatibility/controller';
import { scanCompatibilityCells, MAX_SCAN_CELLS } from '../src/core/compatibility/scanner';
import { WorkbookError } from '../src/utils/errors';
import type { CellScan } from '../src/core/compatibility/types';
const rules = (formula: string, value: unknown = 1, valueType = 'Double') => analyzeFormula({
  worksheetName: 'Sheet 1', address: 'B3', formula, value, valueType,
});
const archive = async (xml: string) => new JSZip().file('xl/workbook.xml', xml).generateAsync({ type: 'uint8array' });
const cleanScan = (): CellScan => ({ findings: [], findingCount: 0, confirmedCount: 0, riskCount: 0,
  scannedCells: 10, scannedWorksheets: 1, totalWorksheets: 1, complete: true });

test('formula evidence distinguishes an observed error from a function compatibility risk', () => {
  const found = rules('=@_xlfn.DISPIMG("ID_a",1)', '#NAME?', 'Error');
  assert.deepEqual(found.map(f => [f.code, f.status]), [
    ['formulaError', 'confirmed'], ['functionMarker', 'risk'], ['dispimg', 'risk'],
  ]);
  assert.equal(found[0]?.address, 'B3');
  assert.equal(rules('=_xlfn.XLOOKUP(A1,B:B,C:C)').length, 1);
  assert.equal(rules('=1/0', '#DIV/0!', 'Error')[0]?.code, 'formulaError');
  assert.deepEqual(rules('="#NAME?"', '#NAME?', 'String'), []);
  assert.deepEqual(rules('=DISPIMG("ID_a",1)', '=DISPIMG("ID_a",1)', 'String'), []);
});

test('quoted text, worksheet names, table columns and implicit intersection do not cause false positives', () => {
  for (const formula of ['=@SUM(A1:A3)', '=SUM(Table1[DISPIMG(foo)])', '=Table1[_xlfn.SUM(foo)]',
    '="_xlfn.DISPIMG(""ID_a"",1) #REF! [book.xlsx]Sheet!A1"',
    "='DISPIMG(foo) #REF!'!A1", '=SUM(Table1[[#Headers],[Column]])']) assert.deepEqual(rules(formula), [], formula);
});

test('external workbook qualifiers are reported without claiming targets are broken', () => {
  for (const formula of ["='C:\\folder\\[Book.xlsx]My Sheet'!A1", '=[1]Sheet1!A1', `='[Book.xlsx]Sheet "quoted"'!A1`, "='[O''Brien.xlsx]Sheet 1'!A1"]) {
    assert.deepEqual(rules(formula).map(f => [f.code, f.status]), [['externalFormula', 'risk']], formula);
  }
  assert.deepEqual(rules('=SUM(Table1[Column])'), []);
  assert.deepEqual(rules('=INDIRECT("[Book.xlsx]Sheet1!A1")'), []);
  assert.equal(rules('=IFERROR(#REF!,0)')[0]?.code, 'brokenReference');
});

test('ordinary workbook defaults to 1900; valid 1904 is metadata, not an error', async () => {
  for (const [property, expected] of [['', '1900'], ['<workbookPr/>', '1900'],
    ['<workbookPr date1904="0"/>', '1900'], ['<workbookPr date1904="false"/>', '1900'],
    ['<workbookPr date1904="true"/>', '1904'], ['<workbookPr date1904="1"/>', '1904']]) {
    const metadata = await parseCompatibilityMetadata(await archive(`<workbook>${property}</workbook>`));
    assert.equal(metadata.dateSystem, expected);
    assert.deepEqual(metadata.externalReferenceIds, []);
  }
});

test('namespaced external records are inventory only; invalid date and duplicate records are inconclusive', async () => {
  const parsed = await parseCompatibilityMetadata(await archive('<workbook xmlns:r="relationships"><workbookPr date1904="invalid"/><externalReferences><externalReference r:id="rId1"/><externalReference r:id="rId1"/></externalReferences></workbook>'));
  assert.equal(parsed.dateSystem, 'unknown');
  assert.equal(parsed.dateReason, 'INVALID_DATE_SYSTEM');
  assert.equal(parsed.externalReason, 'INVALID_EXTERNAL_RECORDS');
  assert.deepEqual(parsed.externalReferenceIds, ['rId1', 'rId1']);
  assert.equal((await parseCompatibilityMetadata(await archive('<workbook><workbookPr date1904=""/></workbook>'))).dateSystem, 'unknown');
  const malformed = await parseCompatibilityMetadata(await archive('<workbook><externalReferences><externalReference/></externalReferences></workbook>'));
  assert.equal(malformed.externalReason, 'INVALID_EXTERNAL_RECORDS');
});

test('missing, malformed, prohibited, oversized and cancelled metadata exit explicitly', async () => {
  await assert.rejects(parseCompatibilityMetadata(new Uint8Array([1, 2])), { code: 'INVALID_XLSX' });
  await assert.rejects(parseCompatibilityMetadata(await new JSZip().generateAsync({ type: 'uint8array' })), { code: 'WORKBOOK_METADATA_MISSING' });
  for (const xml of ['<workbook>', '<other/>', '<!DOCTYPE workbook><workbook/>']) {
    await assert.rejects(parseCompatibilityMetadata(await archive(xml)), { code: 'INVALID_XML' });
  }
  const bomb = await new JSZip().file('xl/workbook.xml', '<workbook>' + ' '.repeat(4 * 1024 * 1024) + '</workbook>')
    .generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  await assert.rejects(parseCompatibilityMetadata(bomb), { code: 'RESOURCE_LIMIT' });
  const signal = AbortSignal.abort();
  await assert.rejects(parseCompatibilityMetadata(new Uint8Array(), signal), { code: 'OPERATION_CANCELLED' });
});

test('snapshot failure preserves completed formula scan and marks dependent checks incomplete', async () => {
  const report = await checkCompatibility(undefined, undefined, {
    scan: async () => cleanScan(), read: async () => { throw new WorkbookError('UNSUPPORTED_HOST', 'No file API'); },
    parse: parseCompatibilityMetadata,
  });
  assert.equal(report.checks[0]?.status, 'clear');
  assert.equal(report.checks[1]?.completion, 'partial');
  assert.equal(report.checks[1]?.status, 'unchecked');
  assert.equal(report.checks[2]?.completion, 'unavailable');
  assert.equal(report.externalTargetsChecked, false);
});

test('report combines live evidence and snapshot inventory without calling external targets', async () => {
  const data = await archive('<workbook xmlns:r="r"><workbookPr date1904="1"/><externalReferences><externalReference r:id="rId2"/></externalReferences></workbook>');
  const report = await checkCompatibility(undefined, undefined, {
    scan: async () => cleanScan(), read: async () => data, parse: parseCompatibilityMetadata,
  });
  assert.equal(report.findingCount, 1);
  assert.equal(report.findings[0]?.source, 'workbookSnapshot');
  assert.equal(report.checks[1]?.status, 'risk');
  assert.equal(report.checks[2]?.status, 'clear');
  assert.equal(report.metadata?.dateSystem, '1904');
});

test('partial scans and omitted detail never imply a clean formula/external check', async () => {
  for (const scan of [{ ...cleanScan(), complete: false, reasonCode: 'CELL_LIMIT' },
    { ...cleanScan(), findingCount: 6000, riskCount: 6000 }]) {
    const report = await checkCompatibility(undefined, undefined, { scan: async () => scan,
      read: async () => archive('<workbook/>'), parse: parseCompatibilityMetadata });
    assert.equal(report.checks[0]?.status, 'unchecked');
    assert.equal(report.checks[1]?.status, 'unchecked');
  }
});

test('cancellation between live scan and snapshot prevents the file read', async () => {
  const controller = new AbortController();
  await assert.rejects(checkCompatibility(controller.signal, undefined, {
    scan: async () => { controller.abort(); return cleanScan(); },
    read: async () => { assert.fail('read after cancellation'); }, parse: parseCompatibilityMetadata,
  }), { code: 'OPERATION_CANCELLED' });
});

async function withExcel(rows: number, columns: number, fn: () => Promise<void>, errorCells = false, failAfter = Infinity) {
  let syncs = 0;
  const sheet = Object.freeze({ name: 'Hidden sheet', getUsedRangeOrNullObject: () => Object.freeze({
    rowIndex: 2, columnIndex: 0, rowCount: rows, columnCount: columns, isNullObject: false, load() {},
  }), getRangeByIndexes: (_row: number, _col: number, height: number, width: number) => {
    assert.ok(height * width <= 5000);
    const matrix = (value: string | number) => Object.freeze(Array.from({ length: height }, () => Object.freeze(Array(width).fill(value))));
    return Object.freeze({ formulas: matrix(errorCells ? '=1/0' : '=1'), values: matrix(errorCells ? '#DIV/0!' : 1),
      valueTypes: matrix(errorCells ? 'Error' : 'Double'), load() {} });
  } });
  const old = Object.getOwnPropertyDescriptor(globalThis, 'Excel');
  Object.defineProperty(globalThis, 'Excel', { configurable: true, value: {
    run: async (callback: (context: unknown) => unknown) => callback({
      workbook: { worksheets: Object.freeze({ items: Object.freeze([sheet]), load() {} }) },
      sync: async () => { if (++syncs > failAfter) throw new Error('Host failed'); },
    }),
  } });
  try { await fn(); } finally {
    if (old) Object.defineProperty(globalThis, 'Excel', old); else Reflect.deleteProperty(globalThis, 'Excel');
  }
}

test('read-only host scan tiles wide ranges, caps details and preserves actual counts/addresses', async () => {
  await withExcel(1, 6001, async () => {
    const result = await scanCompatibilityCells();
    assert.equal(result.complete, true);
    assert.equal(result.scannedCells, 6001);
    assert.equal(result.findings.length, 5000);
    assert.equal(result.findingCount, 6001);
    assert.equal(result.confirmedCount, 6001);
    assert.equal(result.findings[0]?.address, 'A3');
    assert.deepEqual(await scanCompatibilityCells(), result);
  }, true);
});

test('host scan enforces overall cell limit and reports partial failure, never writes', async () => {
  await withExcel(10001, 101, async () => {
    const result = await scanCompatibilityCells();
    assert.equal(result.complete, false);
    assert.equal(result.reasonCode, 'CELL_LIMIT');
    assert.ok(result.scannedCells <= MAX_SCAN_CELLS);
  });
  await withExcel(200, 100, async () => {
    const result = await scanCompatibilityCells();
    assert.equal(result.complete, false);
    assert.equal(result.reasonCode, 'CELL_READ_FAILED');
    assert.equal(result.scannedCells, 5000);
  }, false, 3);
  await assert.rejects(scanCompatibilityCells(AbortSignal.abort()), { code: 'OPERATION_CANCELLED' });
});


test('detail cap applies to external inventory downloads without an uncapped metadata copy', async () => {
  const report = await checkCompatibility(undefined, undefined, { scan: async () => cleanScan(),
    read: async () => new Uint8Array(), parse: async () => ({ dateSystem: '1900',
      externalReferenceIds: Array.from({ length: 6000 }, (_, i) => 'rId' + i) }) });
  assert.equal(report.findingCount, 6000);
  assert.equal(report.findings.length, 5000);
  assert.equal(report.omittedFindings, 1000);
  assert.equal(report.metadata?.externalReferenceCount, 6000);
  assert.equal('externalReferenceIds' in report.metadata!, false);
});
