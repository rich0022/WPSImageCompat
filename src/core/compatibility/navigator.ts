import { WorkbookError } from '../../utils/errors';

export interface CellLocation { worksheetName: string; address: string }

/** Changes only the active sheet and selection; no workbook values, formulas, or shapes are changed. */
export async function navigateToCompatibilityCell(location: CellLocation): Promise<void> {
  if (!location.worksheetName || !/^[A-Z]+[1-9][0-9]*$/i.test(location.address)) {
    throw new WorkbookError('INVALID_LOCATION', 'The reported cell location is invalid.');
  }
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(location.worksheetName);
    sheet.load('isNullObject');
    await context.sync();
    if (sheet.isNullObject) throw new WorkbookError('LOCATION_NOT_FOUND', 'The reported worksheet no longer exists.');
    sheet.activate();
    sheet.getRange(location.address).select();
    await context.sync();
  });
}
