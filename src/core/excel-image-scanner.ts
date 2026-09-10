/** A read-only inventory of ordinary Excel picture shapes already in a workbook. */
export interface ExcelImageScanResult {
  imageCount: number;
  worksheetCount: number;
  supported: boolean;
}

export function canScanExcelImages(): boolean {
  return typeof Office !== 'undefined' && Office.context.requirements.isSetSupported('ExcelApi', '1.10');
}

/**
 * Counts native Excel image shapes. This deliberately does not read image bytes,
 * alter shapes, or include formulas such as IMAGE(), whose source can be remote.
 */
export async function scanExcelImages(): Promise<ExcelImageScanResult> {
  if (!canScanExcelImages()) return { imageCount: 0, worksheetCount: 0, supported: false };
  return Excel.run(async context => {
    const worksheets = context.workbook.worksheets;
    worksheets.load('items/name');
    await context.sync();
    for (const worksheet of worksheets.items) worksheet.shapes.load('items/type');
    await context.sync();
    const counts = worksheets.items.map(worksheet =>
      worksheet.shapes.items.filter(shape => shape.type === 'Image').length);
    return {
      imageCount: counts.reduce((total, count) => total + count, 0),
      worksheetCount: counts.filter(count => count > 0).length,
      supported: true,
    };
  });
}
