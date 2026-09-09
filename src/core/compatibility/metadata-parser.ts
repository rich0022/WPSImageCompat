import JSZip from 'jszip';
import { attribute, nodes, parseXml } from '../../utils/xml';
import { WorkbookError } from '../../utils/errors';
import { MAX_WORKBOOK_BYTES } from '../../utils/limits';
import { throwIfCancelled } from '../../utils/cancellation';
import type { Metadata } from './types';

/** Read only workbook metadata. Never resolve or fetch external relationship targets. */
export async function parseCompatibilityMetadata(bytes: Uint8Array, signal?: AbortSignal): Promise<Metadata> {
  throwIfCancelled(signal);
  if (!bytes.length || bytes.length > MAX_WORKBOOK_BYTES) throw new WorkbookError('FILE_SIZE', 'Workbook size limit.');
  let zip: JSZip;
  try { zip = await JSZip.loadAsync(bytes); }
  catch { throw new WorkbookError('INVALID_XLSX', 'Unreadable XLSX archive.'); }
  throwIfCancelled(signal);
  if (Object.keys(zip.files).length > 20000) throw new WorkbookError('RESOURCE_LIMIT', 'Too many ZIP entries.');
  const part = zip.file('xl/workbook.xml');
  if (!part || (part.unsafeOriginalName && part.unsafeOriginalName !== 'xl/workbook.xml')) {
    throw new WorkbookError('WORKBOOK_METADATA_MISSING', 'Workbook metadata is missing or has an unsafe path.');
  }
  const xml = await new Promise<Uint8Array>((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    let size = 0;
    let stopped = false;
    // JSZip exposes internalStream at runtime but omits it from JSZipObject declarations.
    const stream = (part as typeof part & {
      internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array>;
    }).internalStream('uint8array');
    const fail = (error: unknown) => { stopped = true; chunks.length = 0; stream.pause(); signal?.removeEventListener('abort', cancel); reject(error); };
    const cancel = () => fail(new WorkbookError('OPERATION_CANCELLED', 'Cancelled.'));
    signal?.addEventListener('abort', cancel, { once: true });
    stream.on('data', chunk => {
      if (stopped) return;
      size += chunk.length;
      if (size > 4 * 1024 * 1024) { fail(new WorkbookError('RESOURCE_LIMIT', 'Workbook XML exceeds 4 MiB.')); return; }
      chunks.push(chunk);
    }).on('error', fail).on('end', () => {
      signal?.removeEventListener('abort', cancel);
      if (stopped) return;
      const data = new Uint8Array(size);
      let offset = 0;
      for (const chunk of chunks) { data.set(chunk, offset); offset += chunk.length; }
      resolve(data);
    }).resume();
  });
  throwIfCancelled(signal);
  const root = parseXml(xml, 'workbook');
  const properties = nodes(root.workbookPr);
  const date = properties.length ? properties[0]!['@_date1904'] : undefined;
  const propertyShapeValid = root.workbookPr === undefined || root.workbookPr === '' || properties.length === 1;
  const valid = propertyShapeValid && properties.length <= 1 && (date === undefined || typeof date === 'string' && ['0', '1', 'true', 'false'].includes(date));
  const groups = nodes(root.externalReferences);
  const references = groups.flatMap(group => nodes(group.externalReference));
  const ids = references.map(reference => attribute(reference, 'id'));
  // Empty malformed records must not be interpreted as absence of external dependencies.
  const malformed = (root.externalReferences !== undefined && root.externalReferences !== '' && !groups.length) || groups.length > 1 || groups.some(group => group.externalReference !== undefined &&
      nodes(group.externalReference).length !== (Array.isArray(group.externalReference) ? group.externalReference.length : 1)) ||
    ids.some(id => !id) || new Set(ids).size !== ids.length;
  return { dateSystem: valid ? (date === '1' || date === 'true' ? '1904' : '1900') : 'unknown',
    ...(valid ? {} : { dateReason: 'INVALID_DATE_SYSTEM' }),
    externalReferenceIds: ids.filter((id): id is string => !!id),
    ...(malformed ? { externalReason: 'INVALID_EXTERNAL_RECORDS' } : {}) };
}
