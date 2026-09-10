import { readWorkbook } from './workbook-reader';
import { parseExcelEmbeddedImages } from './excel-drawing-parser';

/** Reads the current XLSX snapshot and inventories standard Excel drawing-layer pictures. */
export async function inspectExcelEmbeddedImages(signal?: AbortSignal) {
  return parseExcelEmbeddedImages(await readWorkbook({ signal }));
}
