export type Category = 'formulas' | 'externalLinks' | 'dates';
export type FindingCode = 'formulaError' | 'functionMarker' | 'dispimg' | 'brokenReference' | 'externalFormula' | 'externalRecord';
export interface Finding {
  category: Category;
  code: FindingCode;
  status: 'confirmed' | 'risk';
  source: 'liveCells' | 'workbookSnapshot';
  worksheetName?: string;
  address?: string;
  formula?: string;
  evidence: string;
}
export interface Check {
  category: Category;
  completion: 'complete' | 'partial' | 'unavailable';
  status: 'confirmed' | 'risk' | 'clear' | 'unchecked';
  reasonCode?: string;
}
export interface CellScan {
  findings: Finding[];
  findingCount: number;
  confirmedCount: number;
  riskCount: number;
  scannedCells: number;
  scannedWorksheets: number;
  totalWorksheets: number;
  complete: boolean;
  reasonCode?: string;
}
export interface Metadata {
  dateSystem: '1900' | '1904' | 'unknown';
  dateReason?: string;
  externalReferenceIds: string[];
  externalReason?: string;
}
export interface CompatibilityReport {
  schemaVersion: 1;
  createdAt: string;
  checks: Check[];
  cells: Omit<CellScan, 'findings'>;
  metadata?: Omit<Metadata, 'externalReferenceIds'> & { externalReferenceCount: number };
  findings: Finding[];
  findingCount: number;
  omittedFindings: number;
  externalTargetsChecked: false;
}
export const MAX_FINDINGS = 5000;
