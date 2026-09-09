import { WorkbookError } from './errors';

export interface OfficeAsyncOptions<T> {
  timeoutMs?: number;
  signal?: AbortSignal;
  /** Only for acquired handles that arrive after timeout or cancellation. */
  onLateSuccess?: (value: T) => void;
}

/** Bound callback waits without retrying Office operations. Late file handles are released. */
export function officeAsync<T>(
  operation: (callback: (result: Office.AsyncResult<T>) => void) => void,
  options: OfficeAsyncOptions<T> = {},
): Promise<T> {
  return new Promise((resolve, reject) => {
    let settled = false;
    let callbackSeen = false;
    const clear = () => { clearTimeout(timer); options.signal?.removeEventListener('abort', cancel); };
    const cancel = () => {
      if (settled) return;
      settled = true;
      clear();
      reject(new WorkbookError('OPERATION_CANCELLED', 'Operation cancelled. Original formulas were not changed.'));
    };
    const timer = setTimeout(() => {
      settled = true;
      clear();
      reject(new WorkbookError('OFFICE_TIMEOUT', 'Excel did not respond within the allowed time. Close and reopen the task pane before retrying if this continues.'));
    }, options.timeoutMs ?? 30000);
    options.signal?.addEventListener('abort', cancel, { once: true });
    if (options.signal?.aborted) { cancel(); return; }
    try {
      operation(result => {
        if (callbackSeen) return;
        callbackSeen = true;
        clear();
        const success = result.status === Office.AsyncResultStatus.Succeeded;
        if (settled) {
          if (success) options.onLateSuccess?.(result.value);
          return;
        }
        settled = true;
        if (success) resolve(result.value);
        else reject(new WorkbookError('OFFICE_FILE', `Excel file operation failed (${result.error?.code ?? 'unknown'}).`));
      });
    } catch (error) {
      clear();
      if (!settled) { settled = true; reject(error); }
    }
  });
}
