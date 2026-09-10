export interface WpsEtCell { worksheetName: string; address: string; formula: string; imageId?: string }
export interface WpsEtPicture { worksheetName: string; shapeIndex: number; name?: string; address?: string }
export interface WpsEtIssue { worksheetName?: string; message: string }
export interface WpsEtScanResult {
  worksheetCount: number; scannedCells: number; dispimgCells: WpsEtCell[]; pictures: WpsEtPicture[]; issues: WpsEtIssue[];
}

export interface WpsEtHost { Sheets?: unknown; ActiveWorkbook?: { Sheets?: unknown } }
const MAX_FORMULA_CELLS = 10_000;
const dispimg = /^\s*=\s*(?:@\s*)?(?:_xlfn\.)?DISPIMG\s*\(\s*"((?:[^"]|"")+)"/i;

async function value<T>(item: T | Promise<T>): Promise<T> { return item; }
function record(value: unknown): Record<string, unknown> | undefined {
  return value !== null && typeof value === 'object' ? value as Record<string, unknown> : undefined;
}
async function property(source: unknown, name: string): Promise<unknown> {
  const item = record(source)?.[name];
  return item === undefined ? undefined : value(item);
}
async function item(collection: unknown, index: number): Promise<unknown> {
  const method = record(collection)?.Item;
  return typeof method === 'function' ? value((method as (index: number) => unknown)(index)) : undefined;
}
function integer(value: unknown): number | undefined {
  const number = typeof value === 'number' ? value : Number(value);
  return Number.isInteger(number) && number >= 0 ? number : undefined;
}
function columnName(column: number): string {
  let name = '';
  for (let value = column; value > 0; value = Math.floor((value - 1) / 26)) name = String.fromCharCode(65 + (value - 1) % 26) + name;
  return name;
}
function asRows(source: unknown): unknown[][] {
  if (!Array.isArray(source)) return [[source]];
  return source.every(Array.isArray) ? source as unknown[][] : [source];
}

/** Pure adapter over documented WPS ET collections; it never writes cells or shapes. */
export async function scanWpsEtWorkbook(application: WpsEtHost): Promise<WpsEtScanResult> {
  const issues: WpsEtIssue[] = [], dispimgCells: WpsEtCell[] = [], pictures: WpsEtPicture[] = [];
  const workbook = await property(application, 'ActiveWorkbook');
  const sheets = await property(application, 'Sheets') ?? await property(workbook, 'Sheets');
  const count = integer(await property(sheets, 'Count'));
  if (count === undefined) throw new Error('WPS 表格未提供工作表集合。请在 WPS 表格中打开此工具。');
  let scannedCells = 0;
  for (let sheetIndex = 1; sheetIndex <= count; sheetIndex++) {
    const sheet = await item(sheets, sheetIndex);
    const worksheetName = String(await property(sheet, 'Name') ?? `Sheet${sheetIndex}`);
    try {
      const usedRange = await property(sheet, 'UsedRange');
      const startRow = integer(await property(usedRange, 'Row')) ?? 1;
      const startColumn = integer(await property(usedRange, 'Column')) ?? 1;
      const formulas = asRows(await property(usedRange, 'Formula'));
      for (let row = 0; row < formulas.length; row++) {
        const values = formulas[row] ?? [];
        for (let column = 0; column < values.length; column++) {
          if (scannedCells >= MAX_FORMULA_CELLS) { issues.push({ worksheetName, message: `公式扫描达到 ${MAX_FORMULA_CELLS.toLocaleString()} 个单元格上限。` }); break; }
          scannedCells++;
          const cellValue = values[column];
          const formula = typeof cellValue === 'string' ? cellValue : '';
          const match = formula.match(dispimg);
          if (match) dispimgCells.push({ worksheetName, address: `${columnName(startColumn + column)}${startRow + row}`, formula, imageId: match[1]?.replaceAll('""', '"') });
        }
        if (scannedCells >= MAX_FORMULA_CELLS) break;
      }
    } catch (error) { issues.push({ worksheetName, message: error instanceof Error ? error.message : '无法读取工作表公式。' }); }
    try {
      const shapes = await property(sheet, 'Shapes');
      const shapeCount = integer(await property(shapes, 'Count')) ?? 0;
      for (let shapeIndex = 1; shapeIndex <= shapeCount; shapeIndex++) {
        const shape = await item(shapes, shapeIndex);
        const imageUrl = await property(shape, 'ImageUrl');
        if (typeof imageUrl !== 'string' || !imageUrl) continue;
        const topLeftCell = await property(shape, 'TopLeftCell');
        const address = await property(topLeftCell, 'Address');
        const name = await property(shape, 'Name');
        pictures.push({ worksheetName, shapeIndex, name: typeof name === 'string' ? name : undefined, address: typeof address === 'string' ? address.replace(/^.*!/, '') : undefined });
      }
    } catch (error) { issues.push({ worksheetName, message: error instanceof Error ? error.message : '无法读取工作表图片。' }); }
  }
  return { worksheetCount: count, scannedCells, dispimgCells, pictures, issues };
}

/** Selection is explicit and limited to the picture reported by the last read-only scan. */
export async function selectWpsEtPicture(application: WpsEtHost, picture: WpsEtPicture): Promise<void> {
  const workbook = await property(application, 'ActiveWorkbook');
  const sheets = await property(application, 'Sheets') ?? await property(workbook, 'Sheets');
  const count = integer(await property(sheets, 'Count')) ?? 0;
  for (let index = 1; index <= count; index++) {
    const sheet = await item(sheets, index);
    if (String(await property(sheet, 'Name')) !== picture.worksheetName) continue;
    const activate = record(sheet)?.Activate;
    if (typeof activate === 'function') await value((activate as () => unknown).call(sheet));
    const shapes = await property(sheet, 'Shapes');
    const shape = await item(shapes, picture.shapeIndex);
    const select = record(shape)?.Select;
    if (typeof select !== 'function') throw new Error('当前 WPS 版本无法选中该图片。');
    await value((select as () => unknown).call(shape));
    return;
  }
  throw new Error('图片所在工作表已不存在。请重新扫描。');
}
