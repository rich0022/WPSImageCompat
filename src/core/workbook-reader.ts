import { WorkbookError } from '../utils/errors';

import { MAX_WORKBOOK_BYTES } from '../utils/limits';
const SLICE_BYTES = 1024 * 1024;
export interface WorkbookSlice { index: number; size: number; data: unknown }
export interface WorkbookFile {
  size: number;
  sliceCount: number;
  getSlice(index: number): Promise<WorkbookSlice>;
  close(): Promise<void>;
}
export type WorkbookFileSource = () => Promise<WorkbookFile>;

/** Testable byte assembly with guaranteed handle close after success or failure. */
export async function readCompressedFile(open: WorkbookFileSource): Promise<Uint8Array> {
  const file = await open();
  let failed = false;
  try {
    if (!Number.isSafeInteger(file.size) || file.size <= 0 || file.size > MAX_WORKBOOK_BYTES ||
        !Number.isSafeInteger(file.sliceCount) || file.sliceCount < 1 || file.sliceCount > file.size) {
      throw new WorkbookError('FILE_SIZE', 'Workbook is empty, invalid, or exceeds the 100 MiB limit.');
    }
    const bytes = new Uint8Array(file.size);
    let offset = 0;
    for (let index = 0; index < file.sliceCount; index++) {
      const slice = await file.getSlice(index);
      const data = slice.data;
      if (!(data instanceof Uint8Array) &&
          !(Array.isArray(data) && data.every(value => Number.isInteger(value) && value >= 0 && value <= 255))) {
        throw new WorkbookError('INVALID_SLICE', 'Excel returned invalid binary data.');
      }
      if (slice.index !== index || slice.size !== data.length || data.length === 0 || offset + data.length > bytes.length) {
        throw new WorkbookError('INVALID_SLICE', 'Excel returned an inconsistent workbook slice.');
      }
      bytes.set(data, offset);
      offset += data.length;
    }
    if (offset !== bytes.length) throw new WorkbookError('INCOMPLETE_FILE', 'Excel returned an incomplete workbook.');
    return bytes;
  } catch (error) {
    failed = true;
    throw error;
  } finally {
    try { await file.close(); }
    catch (error) { if (!failed) throw error; }
  }
}

function officeResult<T>(operation: (callback: (result: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) => {
    operation(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new WorkbookError('OFFICE_FILE', `Excel file operation failed (${result.error?.code ?? 'unknown'}).`));
    });
  });
}

export function canReadWorkbook(): boolean {
  return typeof Office !== 'undefined' &&
    Office.context.requirements.isSetSupported('CompressedFile', '1.1') &&
    typeof Office.context.document.getFileAsync === 'function';
}

/** Only acquires the current workbook binary; it knows nothing about WPS XML. */
export async function readWorkbook(): Promise<Uint8Array> {
  if (!canReadWorkbook()) {
    throw new WorkbookError('UNSUPPORTED_HOST', 'This Excel host cannot read a compressed workbook. Use an updated desktop Excel on Windows or macOS.');
  }
  return readCompressedFile(async () => {
    const file = await officeResult<Office.File>(callback =>
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: SLICE_BYTES }, callback));
    return {
      size: file.size, sliceCount: file.sliceCount,
      getSlice: index => officeResult<Office.Slice>(callback => file.getSliceAsync(index, callback)),
      close: async () => { await officeResult<void>(callback => file.closeAsync(callback)); },
    };
  });
}
