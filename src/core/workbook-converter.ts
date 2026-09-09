import { WorkbookError } from '../utils/errors';

export interface WorkbookConversionOptions {
  /** Future conversion must produce a separate workbook, never overwrite its source. */
  output: 'new-workbook';
  imageKind: 'standard-excel-image';
}
export interface WorkbookConversionResult {
  convertedImages: number;
  outputName: string;
}
export interface WorkbookConverter {
  convert(options: WorkbookConversionOptions): Promise<WorkbookConversionResult>;
}
export const conversionAvailability = {
  available: false,
  reason: 'Conversion is reserved for a future release. Show Images creates removable previews and always preserves DISPIMG formulas.',
} as const;

/** Explicit non-implementation, never a fake success or a preview relabeled as conversion. */
export const workbookConverter: WorkbookConverter = {
  async convert(_options) {
    throw new WorkbookError('CONVERSION_NOT_IMPLEMENTED', conversionAvailability.reason);
  },
};
