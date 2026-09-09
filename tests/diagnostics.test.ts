import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createDiagnosticReport } from '../src/core/diagnostics';
import { conversionAvailability, workbookConverter } from '../src/core/workbook-converter';
const host = { platform: 'Mac', version: 'test', scan: true, compressedFile: true, shapes: true, mergedCells: false };
test('diagnostics export counts and capabilities without workbook data or raw errors', () => {
  const secret = 'PRIVATE_CANARY';
  const cell = { worksheetName: secret, address: secret, formula: secret, imageId: secret };
  const report = createDiagnosticReport({ host, operation: 'show', detection: {
    scan: { cells: [cell], worksheetCount: 1, scannedWorksheetCount: 2 },
    mappings: [{ cell, status: 'found', resource: { imageId: secret, relationshipId: secret,
      mediaPath: secret, mimeType: 'image/png', base64: secret } }],
    resourceError: secret,
    parsed: { hasCellImages: true, resources: new Map(), issues: [{ code: 'MISSING_MEDIA',
      kind: 'missing', message: secret, imageId: secret }] },
  }, preview: { inserted: 1, existing: 0, removed: 0, skipped: 0,
    issues: [{ location: secret, message: secret }] } });
  assert.equal(JSON.stringify(report).includes(secret), false);
  assert.equal(report.scan?.found, 1);
  assert.equal(report.preview?.inserted, 1);
  assert.equal(report.preview?.issueCount, 1);
});
test('unchecked resources remain null rather than being reported as missing', () => {
  const report = createDiagnosticReport({ host, operation: 'scan', detection: {
    scan: { cells: [], worksheetCount: 0, scannedWorksheetCount: 1 },
    resourceError: 'not available', resourceErrorCode: 'OFFICE_TIMEOUT',
  } });
  assert.equal(report.scan?.resourcesChecked, false);
  assert.equal(report.scan?.missing, null);
  assert.equal(report.errorCode, 'OFFICE_TIMEOUT');
});
test('reserved conversion rejects explicitly and does not require an Excel host', async () => {
  assert.equal(conversionAvailability.available, false);
  await assert.rejects(workbookConverter.convert({ output: 'new-workbook', imageKind: 'standard-excel-image' }),
    { code: 'CONVERSION_NOT_IMPLEMENTED' });
});
