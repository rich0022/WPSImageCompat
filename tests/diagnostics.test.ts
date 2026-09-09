import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createDiagnosticReport } from '../src/core/diagnostics';
import { convertWorkbook } from '../src/core/workbook-converter';
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
test('conversion requires a user-selected action before accessing an Excel host', async () => {
  await assert.rejects(convertWorkbook({ convertDispimg: false, freezeUnsupported: false, freezeExternal: false,
    keepAspectRatio: true, fitInsideCell: true }), { code: 'NO_CONVERSION_SELECTED' });
});
test('a WPS image resource failure is reported without touching Excel or pretending the image converted', async () => {
  const result = await convertWorkbook({ convertDispimg: true, freezeUnsupported: false, freezeExternal: false,
    keepAspectRatio: true, fitInsideCell: true }, undefined, {
    detect: async () => ({ scan: { cells: [], worksheetCount: 0, scannedWorksheetCount: 1 },
      resourceError: 'Snapshot unavailable', resourceErrorCode: 'RESOURCES_UNAVAILABLE' }),
    render: async () => assert.fail('render must not run without resources'),
    compatibility: async () => assert.fail('compatibility must not run when no formula action is selected'),
  });
  assert.equal(result.convertedImages, 0);
  assert.equal(result.clearedDispimgFormulas, 0);
  assert.equal(result.issues[0]?.code, 'RESOURCES_UNAVAILABLE');
});
