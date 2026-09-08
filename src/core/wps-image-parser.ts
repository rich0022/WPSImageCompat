import JSZip from 'jszip';
import type { JSZipObject } from 'jszip';
import type { ImageParseIssue, WpsImageParseResult } from '../types/wps';
import { attribute, nodes, parseXml } from '../utils/xml';
import { toBase64 } from '../utils/base64';
import { WorkbookError } from '../utils/errors';
import { MAX_WORKBOOK_BYTES } from '../utils/limits';

const CELL_IMAGES = 'xl/cellimages.xml';
const RELATIONSHIPS = 'xl/_rels/cellimages.xml.rels';
const XML_LIMIT = 4 * 1024 * 1024;
const IMAGE_LIMIT = 20 * 1024 * 1024;
const TOTAL_LIMIT = 128 * 1024 * 1024;

/** Resolve OPC relationship targets relative to xl/cellimages.xml, never fetch URLs. */
export function resolveMediaPath(target: string): string | undefined {
  let decoded: string;
  try { decoded = decodeURIComponent(target); } catch { return undefined; }
  if (/[\\\u0000-\u001f?#:]/.test(decoded) || decoded.startsWith('//')) return undefined;
  const parts = decoded.startsWith('/') ? [] : ['xl'];
  for (const part of decoded.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') { if (!parts.length) return undefined; parts.pop(); }
    else parts.push(part);
  }
  const path = parts.join('/');
  return path.startsWith('xl/media/') && path.length > 'xl/media/'.length ? path : undefined;
}

/** Pause as soon as the decompressed size crosses a limit instead of allocating it all. */
async function readEntry(entry: JSZipObject, limit: number): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    let size = 0;
    let stopped = false;
    // JSZip 3.10 exposes this API but omits it from JSZipObject's types.
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
        reject(new WorkbookError('RESOURCE_LIMIT', 'Decompressed image resources exceed the safety limit.'));
      } else chunks.push(chunk);
    });
    stream.on('error', () => reject(new WorkbookError('INVALID_ZIP_ENTRY', 'An image resource cannot be decompressed.')));
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

const mimeByExtension: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
  bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', svg: 'image/svg+xml',
  emf: 'image/x-emf', wmf: 'image/x-wmf', webp: 'image/webp',
};

/** Pure OOXML parser: no Office.js operations, no workbook writes, no network. */
export async function parseWpsImages(binary: Uint8Array): Promise<WpsImageParseResult> {
  if (!binary.length || binary.length > MAX_WORKBOOK_BYTES) {
    throw new WorkbookError('FILE_SIZE', 'Workbook is empty or exceeds the 100 MiB limit.');
  }
  let zip: JSZip;
  try { zip = await JSZip.loadAsync(binary); }
  catch { throw new WorkbookError('INVALID_XLSX', 'The workbook is not a readable, unencrypted XLSX ZIP archive.'); }
  const result: WpsImageParseResult = { hasCellImages: false, resources: new Map(), issues: [] };
  if (Object.keys(zip.files).length > 20000) throw new WorkbookError('RESOURCE_LIMIT', 'Workbook has too many ZIP entries.');
  const entry = (path: string): JSZipObject | null => {
    const file = zip.file(path);
    const original = (file as (JSZipObject & { unsafeOriginalName?: string }) | null)?.unsafeOriginalName;
    if (original && original !== path) throw new WorkbookError('UNSAFE_ZIP_PATH', 'A ZIP entry contains an ambiguous path.');
    return file;
  };
  const cellImages = entry(CELL_IMAGES);
  if (!cellImages) return result;
  result.hasCellImages = true;
  let totalBytes = 0;
  const read = async (file: JSZipObject, limit: number) => {
    const bytes = await readEntry(file, Math.min(limit, TOTAL_LIMIT - totalBytes));
    totalBytes += bytes.length;
    return bytes;
  };
  const root = parseXml(await read(cellImages, XML_LIMIT), 'cellImages');
  const relFile = entry(RELATIONSHIPS);
  const relationships = new Map<string, { target?: string; external: boolean; image: boolean }>();
  const duplicateRels = new Set<string>();
  if (relFile) {
    const relRoot = parseXml(await read(relFile, XML_LIMIT), 'Relationships');
    for (const rel of nodes(relRoot.Relationship)) {
      const id = attribute(rel, 'Id');
      if (!id) continue;
      if (relationships.has(id)) duplicateRels.add(id);
      relationships.set(id, {
        target: attribute(rel, 'Target'),
        external: (attribute(rel, 'TargetMode') ?? 'Internal').toLowerCase() !== 'internal',
        image: /\/image$/.test(attribute(rel, 'Type') ?? ''),
      });
    }
  } else {
    result.issues.push({ code: 'MISSING_RELS', kind: 'missing', message: 'WPS image relationships are missing.' });
  }
  const contentTypes = new Map<string, string>();
  const defaults = new Map<string, string>();
  const typesFile = entry('[Content_Types].xml');
  if (typesFile) {
    const typesRoot = parseXml(await read(typesFile, XML_LIMIT), 'Types');
    for (const item of nodes(typesRoot.Default)) {
      const ext = attribute(item, 'Extension'), type = attribute(item, 'ContentType');
      if (ext && type) defaults.set(ext.toLowerCase(), type);
    }
    for (const item of nodes(typesRoot.Override)) {
      const path = attribute(item, 'PartName'), type = attribute(item, 'ContentType');
      if (path && type) contentTypes.set(path.replace(/^\//, ''), type);
    }
  }
  const seen = new Set<string>();
  const mediaCache = new Map<string, string>();
  const issue = (imageId: string | undefined, code: string, message: string, kind: ImageParseIssue['kind'] = 'error') => {
    result.issues.push({ imageId, code, message, kind });
  };
  for (const cellImage of nodes(root.cellImage)) {
    const pic = nodes(cellImage.pic)[0];
    const metadata = nodes(nodes(pic?.nvPicPr)[0]?.cNvPr)[0];
    const imageId = metadata && attribute(metadata, 'name');
    if (!imageId) { issue(undefined, 'INVALID_IMAGE', 'A WPS image has no ID.'); continue; }
    if (seen.has(imageId)) {
      result.resources.delete(imageId);
      issue(imageId, 'DUPLICATE_ID', 'The WPS image ID is ambiguous.');
      continue;
    }
    seen.add(imageId);
    const blips = nodes(nodes(pic?.blipFill)[0]?.blip);
    const relationshipId = blips.length === 1 ? attribute(blips[0]!, 'embed') : undefined;
    if (!relationshipId) { issue(imageId, 'INVALID_IMAGE', 'The WPS image has no unique embedded relationship.'); continue; }
    if (duplicateRels.has(relationshipId)) { issue(imageId, 'DUPLICATE_REL', 'The image relationship is ambiguous.'); continue; }
    const relationship = relationships.get(relationshipId);
    if (!relationship) { issue(imageId, 'MISSING_REL', 'The image relationship cannot be found.', 'missing'); continue; }
    if (relationship.external || !relationship.image) { issue(imageId, 'UNSUPPORTED_REL', 'Only internal image relationships are supported.'); continue; }
    const mediaPath = relationship.target && resolveMediaPath(relationship.target);
    if (!mediaPath) { issue(imageId, 'INVALID_TARGET', 'The image target is outside xl/media or is invalid.'); continue; }
    const media = entry(mediaPath);
    if (!media) { issue(imageId, 'MISSING_MEDIA', 'The embedded image file cannot be found.', 'missing'); continue; }
    const extension = mediaPath.split('.').pop()!.toLowerCase();
    const mimeType = contentTypes.get(mediaPath) ?? defaults.get(extension) ?? mimeByExtension[extension];
    if (!mimeType?.startsWith('image/')) { issue(imageId, 'UNKNOWN_MEDIA', 'The embedded resource has no supported image content type.'); continue; }
    try {
      let base64 = mediaCache.get(mediaPath);
      if (base64 === undefined) {
        const data = await read(media, IMAGE_LIMIT);
        if (!data.length) { issue(imageId, 'EMPTY_MEDIA', 'The embedded image file is empty.'); continue; }
        base64 = toBase64(data);
        mediaCache.set(mediaPath, base64);
      }
      result.resources.set(imageId, { imageId, relationshipId, mediaPath, mimeType, base64 });
    } catch (error) {
      if (error instanceof WorkbookError && error.code === 'RESOURCE_LIMIT') throw error;
      issue(imageId, 'INVALID_MEDIA', 'The embedded image file cannot be read.');
    }
  }
  return result;
}
