export interface DispimgCell {
  worksheetName: string;
  /** Local A1 address; worksheetName is stored separately. */
  address: string;
  formula: string;
  imageId: string;
}

export interface WpsImageResource {
  imageId: string;
  relationshipId: string;
  mediaPath: string;
  mimeType: string;
  base64?: string;
}

export interface ImageMapping {
  cell: DispimgCell;
  resource?: WpsImageResource;
  status: 'found' | 'missing' | 'rendered' | 'error';
}

export interface ScanResult {
  cells: DispimgCell[];
  worksheetCount: number;
  scannedWorksheetCount: number;
}

export interface ImageParseIssue {
  code: string;
  message: string;
  imageId?: string;
  kind: 'missing' | 'error';
}

export interface WpsImageParseResult {
  hasCellImages: boolean;
  resources: Map<string, WpsImageResource>;
  issues: ImageParseIssue[];
}
