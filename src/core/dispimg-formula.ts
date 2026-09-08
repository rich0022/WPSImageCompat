/** Recognize a direct WPS image formula, never a text value or nested expression.
 * OOXML uses commas; semicolons are accepted for pasted/localized formulas.
 * Literal string IDs are decoded according to Excel's doubled-quote escaping.
 */
export function parseDispimgFormula(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const match = /^\s*=\s*@?\s*(?:_xlfn\.)?DISPIMG\s*\(\s*"((?:[^"]|"")+)"\s*[,;]\s*1\s*\)\s*$/i.exec(value);
  return match?.[1]?.replace(/""/g, '"');
}

export function cellAddress(rowIndex: number, columnIndex: number): string {
  if (!Number.isInteger(rowIndex) || !Number.isInteger(columnIndex) ||
      rowIndex < 0 || rowIndex >= 1048576 || columnIndex < 0 || columnIndex >= 16384) {
    throw new RangeError('Cell coordinates are outside an Excel worksheet.');
  }
  let column = columnIndex + 1;
  let letters = '';
  while (column > 0) {
    column--;
    letters = String.fromCharCode(65 + column % 26) + letters;
    column = Math.floor(column / 26);
  }
  return `${letters}${rowIndex + 1}`;
}
