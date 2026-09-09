export interface ReleaseInfo { version: string; buildId: string }
export function parseRelease(value: unknown): ReleaseInfo | undefined {
  if (!value || typeof value !== 'object') return;
  const { version, buildId } = value as Record<string, unknown>;
  if (typeof version !== 'string' || !/^\d+\.\d+\.\d+(?:-[a-zA-Z0-9.-]+)?$/.test(version) || version.length > 60 ||
      typeof buildId !== 'string' || !/^[a-zA-Z0-9-]{1,80}$/.test(buildId)) return;
  return { version, buildId };
}
/** Checks same-origin metadata only. It never reloads the pane or touches a workbook. */
export function createUpdateChecker(current: ReleaseInfo, onChange: (release?: ReleaseInfo) => void,
  fetcher: typeof fetch = fetch, timeoutMs = 10000) {
  let inFlight = false;
  return async (): Promise<void> => {
    if (inFlight) return;
    inFlight = true;
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const response = await fetcher('/version.json', { cache: 'no-store', credentials: 'omit', signal: controller.signal });
      if (!response.ok) return;
      const release = parseRelease(await response.json());
      if (release) onChange(release.buildId === current.buildId ? undefined : release);
    } catch { /* Offline or an older deployment without metadata must not interrupt Excel. */ }
    finally { clearTimeout(timer); inFlight = false; }
  };
}
export function canApplyUpdate(busy: boolean, available?: ReleaseInfo): boolean { return !!available && !busy; }
