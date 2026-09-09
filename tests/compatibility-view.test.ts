import { test } from 'node:test';
import assert from 'node:assert/strict';
import { summarizeCompatibilityFindings } from '../src/taskpane/compatibility-view';
import type { CompatibilityReport } from '../src/core/compatibility/types';

const report = (summaries: CompatibilityReport['summaries']): CompatibilityReport => ({
  schemaVersion: 1, createdAt: '2026-09-09T00:00:00.000Z', checks: [],
  cells: { findingCount: 0, confirmedCount: 0, riskCount: 0, scannedCells: 0, scannedWorksheets: 0, totalWorksheets: 0, complete: true },
  findings: [], summaries, findingCount: 0, omittedFindings: 0, externalTargetsChecked: false,
});

test('compatibility UI summaries keep one row per issue type with all counts and a first location', () => {
  const result = summarizeCompatibilityFindings(report([{ code: 'functionMarker', status: 'risk', count: 87,
    firstLocation: { worksheetName: '报价', address: 'B3' } }, { code: 'externalRecord', status: 'risk', count: 2 }]));
  assert.deepEqual(result, [{ code: 'functionMarker', status: 'risk', count: 87,
    firstLocation: { worksheetName: '报价', address: 'B3' } }, { code: 'externalRecord', status: 'risk', count: 2 }]);
});
