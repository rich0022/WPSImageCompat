import { WorkbookError } from './errors';
export function throwIfCancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new WorkbookError('OPERATION_CANCELLED', 'Operation cancelled. Original formulas were not changed.');
}
