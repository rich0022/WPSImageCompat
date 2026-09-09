import { test } from 'node:test';
import assert from 'node:assert/strict';
import { planFormulaConversions } from '../src/core/conversion-planner';
import type { CompatibilityReport } from '../src/core/compatibility/types';

const base = (findings: CompatibilityReport['findings'], summaries: CompatibilityReport['summaries']): CompatibilityReport => ({
  schemaVersion: 1, createdAt: '2026-09-09T00:00:00.000Z', checks: [],
  cells: { findingCount: findings.length, confirmedCount: 0, riskCount: findings.length,
    scannedCells: 10, scannedWorksheets: 1, totalWorksheets: 1, complete: true },
  findings, summaries, findingCount: findings.length, omittedFindings: 0, externalTargetsChecked: false,
});
const cell = (code: 'functionMarker' | 'externalFormula' | 'brokenReference', formula: string) => ({
  category: code === 'functionMarker' ? 'formulas' as const : 'externalLinks' as const, code,
  status: code === 'brokenReference' ? 'confirmed' as const : 'risk' as const,
  source: 'liveCells' as const, worksheetName: 'Sheet 1', address: 'B3', formula, evidence: code,
});

test('plans one value freeze per cell with explicit selected reasons', () => {
  const formula = "=_xlfn.XLOOKUP(A1,'[Book.xlsx]Sheet1'!A:A,'[Book.xlsx]Sheet1'!B:B)";
  const report = base([cell('functionMarker', formula), cell('externalFormula', formula)], [
    { code: 'functionMarker', status: 'risk', count: 1, firstLocation: { worksheetName: 'Sheet 1', address: 'B3' } },
    { code: 'externalFormula', status: 'risk', count: 1, firstLocation: { worksheetName: 'Sheet 1', address: 'B3' } },
  ]);
  assert.deepEqual(planFormulaConversions(report, { freezeUnsupported: true, freezeExternal: true }), {
    candidates: [{ worksheetName: 'Sheet 1', address: 'B3', formula,
      reasons: ['unsupportedFunction', 'externalReference'] }], brokenReferenceCount: 0, blockedByDetailLimit: false,
  });
});

test('never plans unselected categories or guesses a broken reference', () => {
  const report = base([cell('functionMarker', '=_xlfn.XLOOKUP(A1,A:A,B:B)'), cell('externalFormula', '=[Book.xlsx]Sheet1!A1'), cell('brokenReference', '=A1+#REF!')], [
    { code: 'functionMarker', status: 'risk', count: 1 }, { code: 'externalFormula', status: 'risk', count: 1 },
    { code: 'brokenReference', status: 'confirmed', count: 1 },
  ]);
  assert.deepEqual(planFormulaConversions(report, { freezeUnsupported: false, freezeExternal: true }), {
    candidates: [{ worksheetName: 'Sheet 1', address: 'B3', formula: '=[Book.xlsx]Sheet1!A1', reasons: ['externalReference'] }],
    brokenReferenceCount: 1, blockedByDetailLimit: false,
  });
});

test('refuses a partial formula conversion when findings exceed report detail capacity', () => {
  const report = base([cell('functionMarker', '=_xlfn.XLOOKUP(A1,A:A,B:B)')], [
    { code: 'functionMarker', status: 'risk', count: 5001 },
  ]);
  assert.equal(planFormulaConversions(report, { freezeUnsupported: true, freezeExternal: false }).blockedByDetailLimit, true);
});
