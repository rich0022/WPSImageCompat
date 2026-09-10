/**
 * The add-in must be possible to pause without touching the workbook.  Keep the
 * storage value deliberately small and tolerant of unavailable web storage.
 */
export function featureEnabledPreference(value: string | null | undefined): boolean {
  return value !== 'false';
}
