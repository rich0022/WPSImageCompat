import JSZip from 'jszip';
import type { JSZipObject } from 'jszip';
import { attribute, nodes, parseXml, type XmlNode } from '../utils/xml';
import { WorkbookError } from '../utils/errors';
import { cellAddress } from './dispimg-formula';

const XML_LIMIT = 4 * 1024 * 1024;
const MAX_WORKBOOK_BYTES = 100 * 1024 * 1024;
const TOTAL_XML_LIMIT = 32 * 1024 * 1024;

export interface ExcelEmbeddedImage {
  worksheetName: string;
  address?: string;
  drawingPath: string;
  relationshipId: string;
  mediaPath: string;
  name?: string;
}
export interface ExcelEmbeddedImageIssue { code: string; message: string; location?: string }
export interface ExcelEmbeddedImageParseResult { images: ExcelEmbeddedImage[]; issues: ExcelEmbeddedImageIssue[] }
interface Relationship { target: string; external: boolean; type?: string }

function relationshipPath(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  return `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`;
}

/** Resolves internal OOXML targets without allowing a relationship to escape xl/. */
export function resolveOoxmlPart(sourcePath: string, target: string): string | undefined {
  let decoded: string;
  try { decoded = decodeURIComponent(target); } catch { return undefined; }
  if (!decoded || /[\\\u0000-\u001f?#:]/.test(decoded) || decoded.startsWith('//')) return undefined;
  const sourceParts = sourcePath.split('/').slice(0, -1);
  const parts = decoded.startsWith('/') ? [] : sourceParts;
  for (const part of decoded.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') { if (!parts.length) return undefined; parts.pop(); }
    else parts.push(part);
  }
  const path = parts.join('/');
  return path.startsWith('xl/') ? path : undefined;
}

/** Stop decompression once the XML safety limit is crossed. */
async function readEntry(entry: JSZipObject, limit: number): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    let size = 0;
    let stopped = false;
    const stream = (entry as JSZipObject & {
      internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array>;
    }).internalStream('uint8array');
    stream.on('data', chunk => {
      if (stopped) return;
      size += chunk.length;
      if (size > limit) {
        stopped = true;
        stream.pause();
        chunks.length = 0;
        reject(new WorkbookError('RESOURCE_LIMIT', 'Excel drawing XML exceeds the safety limit.'));
      } else chunks.push(chunk);
    });
    stream.on('error', () => reject(new WorkbookError('INVALID_ZIP_ENTRY', 'An Excel drawing XML part cannot be decompressed.')));
    stream.on('end', () => {
      if (stopped) return;
      const bytes = new Uint8Array(size);
      let offset = 0;
      for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
      resolve(bytes);
    });
    stream.resume();
  });
}

async function readXml(zip: JSZip, path: string, root: string, usage: { total: number }): Promise<XmlNode | undefined> {
  const entry = zip.file(path);
  if (!entry) return undefined;
  const original = (entry as JSZipObject & { unsafeOriginalName?: string }).unsafeOriginalName;
  if (original && original !== path) throw new WorkbookError('UNSAFE_ZIP_PATH', 'A ZIP entry contains an ambiguous path.');
  const bytes = await readEntry(entry, Math.min(XML_LIMIT, TOTAL_XML_LIMIT - usage.total));
  usage.total += bytes.length;
  return parseXml(bytes, root);
}

async function readRelationships(zip: JSZip, sourcePath: string, usage: { total: number }): Promise<Map<string, Relationship>> {
  const root = await readXml(zip, relationshipPath(sourcePath), 'Relationships', usage);
  const relationships = new Map<string, Relationship>();
  for (const node of nodes(root?.Relationship)) {
    const id = attribute(node, 'Id'), target = attribute(node, 'Target');
    if (!id || !target || relationships.has(id)) continue;
    relationships.set(id, { target, type: attribute(node, 'Type'), external: (attribute(node, 'TargetMode') ?? 'Internal').toLowerCase() !== 'internal' });
  }
  return relationships;
}

function numberValue(value: unknown): number | undefined {
  const text = typeof value === 'string' || typeof value === 'number' ? String(value) :
    typeof nodes(value)[0]?.['#text'] === 'string' ? nodes(value)[0]?.['#text'] : undefined;
  const number = text === undefined ? NaN : Number(text);
  return Number.isInteger(number) ? number : undefined;
}
function anchorAddress(anchor: XmlNode): string | undefined {
  const from = nodes(anchor.from)[0];
  const column = numberValue(from?.col), row = numberValue(from?.row);
  return column !== undefined && row !== undefined && column >= 0 && row >= 0 ? cellAddress(row, column) : undefined;
}

function picture(anchor: XmlNode): { relationshipId?: string; name?: string } | undefined {
  const pic = nodes(anchor.pic)[0];
  const blip = nodes(nodes(pic?.blipFill)[0]?.blip)[0];
  const properties = nodes(nodes(pic?.nvPicPr)[0]?.cNvPr)[0];
  if (!pic || !blip) return undefined;
  return { relationshipId: attribute(blip, 'embed'), name: attribute(properties ?? {}, 'name') };
}

/** Pure parser for regular Excel drawing-layer pictures. It does not mutate Excel or read image bytes. */
export async function parseExcelEmbeddedImages(binary: Uint8Array): Promise<ExcelEmbeddedImageParseResult> {
  if (!binary.length || binary.length > MAX_WORKBOOK_BYTES) throw new WorkbookError('FILE_SIZE', 'Workbook is empty or exceeds the 100 MiB limit.');
  let zip: JSZip;
  try { zip = await JSZip.loadAsync(binary); }
  catch { throw new WorkbookError('INVALID_XLSX', 'The workbook is not a readable, unencrypted XLSX ZIP archive.'); }
  if (Object.keys(zip.files).length > 20000) throw new WorkbookError('RESOURCE_LIMIT', 'Workbook has too many ZIP entries.');
  const result: ExcelEmbeddedImageParseResult = { images: [], issues: [] };
  const usage = { total: 0 };
  const workbookPath = 'xl/workbook.xml';
  const workbook = await readXml(zip, workbookPath, 'workbook', usage);
  if (!workbook) return result;
  const workbookRelationships = await readRelationships(zip, workbookPath, usage);
  const sheets = nodes(nodes(workbook.sheets)[0]?.sheet);
  for (const sheetNode of sheets) {
    const worksheetName = attribute(sheetNode, 'name') ?? 'Unnamed worksheet';
    const workbookRelationship = workbookRelationships.get(attribute(sheetNode, 'id') ?? '');
    const worksheetPath = workbookRelationship && !workbookRelationship.external ? resolveOoxmlPart(workbookPath, workbookRelationship.target) : undefined;
    if (!worksheetPath) { result.issues.push({ code: 'MISSING_SHEET', message: 'A worksheet drawing relationship cannot be resolved.', location: worksheetName }); continue; }
    const worksheet = await readXml(zip, worksheetPath, 'worksheet', usage);
    if (!worksheet) continue;
    const worksheetRelationships = await readRelationships(zip, worksheetPath, usage);
    for (const drawingNode of nodes(worksheet.drawing)) {
      const drawingRelationship = worksheetRelationships.get(attribute(drawingNode, 'id') ?? '');
      const drawingPath = drawingRelationship && !drawingRelationship.external ? resolveOoxmlPart(worksheetPath, drawingRelationship.target) : undefined;
      if (!drawingPath) { result.issues.push({ code: 'MISSING_DRAWING', message: 'A worksheet drawing cannot be resolved.', location: worksheetName }); continue; }
      const drawing = await readXml(zip, drawingPath, 'wsDr', usage);
      if (!drawing) continue;
      const drawingRelationships = await readRelationships(zip, drawingPath, usage);
      const anchors = [...nodes(drawing.twoCellAnchor), ...nodes(drawing.oneCellAnchor), ...nodes(drawing.absoluteAnchor)];
      for (const anchor of anchors) {
        const item = picture(anchor);
        if (!item?.relationshipId) continue;
        const imageRelationship = drawingRelationships.get(item.relationshipId);
        const mediaPath = imageRelationship && !imageRelationship.external && /\/image$/.test(imageRelationship.type ?? '') ?
          resolveOoxmlPart(drawingPath, imageRelationship.target) : undefined;
        if (!mediaPath || !zip.file(mediaPath)) {
          result.issues.push({ code: 'MISSING_MEDIA', message: 'An Excel drawing picture has no readable embedded media resource.', location: worksheetName });
          continue;
        }
        result.images.push({ worksheetName, address: anchorAddress(anchor), drawingPath, relationshipId: item.relationshipId, mediaPath, name: item.name });
      }
    }
  }
  return result;
}
