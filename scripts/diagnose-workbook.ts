/** Read-only local fixture audit. This script never writes the input or prints workbook content. */
import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { parseWpsImages } from '../src/core/wps-image-parser';
import { parseDispimgFormula } from '../src/core/dispimg-formula';
import { parseXml, nodes } from '../src/utils/xml';

const source = process.argv[2];
const destination = process.argv[3];
if (!source) throw new Error('Usage: npm run diagnose:file -- <workbook.xlsx> [report.json]');
if (destination && resolve(source) === resolve(destination)) throw new Error('Report output must not overwrite the workbook.');
const binary = await readFile(source);
const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
const before = hash(binary);
const started = performance.now();
const parsed = await parseWpsImages(binary);
const zip = await JSZip.loadAsync(binary);
const worksheetResults = [];
let unrecognized = 0;
for (const path of Object.keys(zip.files).filter(path => /^xl\/worksheets\/[^/]+\.xml$/.test(path))) {
  const root = parseXml(await zip.file(path)!.async('uint8array'), 'worksheet');
  let imageCells = 0, found = 0;
  for (const row of nodes(nodes(root.sheetData)[0]?.row)) for (const cell of nodes(row.c)) {
    const formula = typeof cell.f === 'string' ? cell.f : nodes(cell.f)[0]?.['#text'];
    if (typeof formula !== 'string' || !formula.toUpperCase().includes('DISPIMG')) continue;
    const id = parseDispimgFormula(`=${formula}`);
    if (!id) { unrecognized++; continue; }
    imageCells++;
    if (parsed.resources.has(id)) found++;
  }
  // Numbered parts only, not user-facing sheet names or cell addresses.
  worksheetResults.push({ part: path, imageCells, found, missing: imageCells - found });
}
const unchanged = hash(await readFile(source)) === before;
if (!unchanged) throw new Error('Source changed during diagnostics; retry without editing it.');
const report = { testKind: 'read-only-ooxml-parser', bytes: binary.length, sha256: before,
  originalUnchanged: unchanged, durationMs: Math.round(performance.now() - started),
  hasCellImages: parsed.hasCellImages, resourceCount: parsed.resources.size,
  issueCodes: [...new Set(parsed.issues.map(issue => issue.code))],
  unrecognizedFormulaCount: unrecognized, worksheetResults,
  imageTypes: [...new Set([...parsed.resources.values()].map(resource => resource.mimeType))],
  nativeExcelTested: false };
const output = JSON.stringify(report, null, 2);
if (destination) await writeFile(destination, output, { flag: 'wx' });
console.log(output);
