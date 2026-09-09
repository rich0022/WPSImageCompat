import './taskpane.css';
import { detectWorkbook } from '../core/workbook-detector';
import type { DetectionResult } from '../core/workbook-detector';
import { showWorkbookImages } from '../core/preview-controller';
import { canRenderImages, removePreviewImages } from '../core/image-renderer';
import type { PreviewResult } from '../core/image-renderer';
import { WorkbookError } from '../utils/errors';
import { createDiagnosticReport, readHostCapabilities } from '../core/diagnostics';
import type { HostCapabilities } from '../core/diagnostics';
import { createUpdateChecker, canApplyUpdate } from '../core/update-checker';
import type { ReleaseInfo } from '../core/update-checker';
import { languagePreference, resolveLocale, translate, errorText } from './i18n';
import type { LanguagePreference, MessageKey, MessageParams } from './i18n';

const element = <T extends HTMLElement>(id: string): T => {
  const result = document.getElementById(id);
  if (!result) throw new Error('Missing UI element: ' + id);
  return result as T;
};
const status = element('scan-status');
const hostStatus = element('host-status');
const languageSelect = element<HTMLSelectElement>('language');
const languageKey = 'wps-image-compat.language';
let preference: LanguagePreference = 'auto';
try { preference = languagePreference(localStorage.getItem(languageKey)); } catch { /* Restricted storage: use this session. */ }
let officeLanguage: string | undefined;
let locale = resolveLocale(preference, officeLanguage, navigator.language);
const t = (key: MessageKey, params?: MessageParams) => translate(locale, key, params);
let ready = false;
let shapesReady = false;
let busy = false;
let cancelAllowed = false;
let controller: AbortController | undefined;
let host: HostCapabilities | undefined;
let lastDetection: DetectionResult | undefined;
let lastPreview: PreviewResult | undefined;
let lastErrorCode: string | undefined;
let lastErrorMessage: string | undefined;
let lastOperation = 'initialize';
let availableUpdate: ReleaseInfo | undefined;
let statusText: () => string = () => '';
let hostText = () => t('connecting');
function setStatus(key: MessageKey, params?: MessageParams): void {
  statusText = () => t(key, params);
  status.textContent = statusText();
}
function updateControls(): void {
  element<HTMLButtonElement>('scan').disabled = !ready || busy;
  for (const id of ['show', 'refresh', 'remove']) element<HTMLButtonElement>(id).disabled = !shapesReady || busy;
  element<HTMLFieldSetElement>('preview-settings').disabled = !shapesReady || busy;
  element<HTMLButtonElement>('cancel').disabled = !busy || !cancelAllowed || !!controller?.signal.aborted;
  element<HTMLButtonElement>('diagnostics').disabled = !host || busy;
  element<HTMLButtonElement>('apply-update').disabled = !canApplyUpdate(busy, availableUpdate);
  element('update-panel').hidden = !availableUpdate;
  element('update-status').textContent = availableUpdate ?
    t('updateAvailable', { version: availableUpdate.version }) + (busy ? ' ' + t('updateBusy') : '') : '';
}
function appendIssue(location: string, code: string, original: string): void {
  const item = document.createElement('li');
  const description = errorText(locale, code, original);
  item.textContent = location + ': ' + description;
  // Preserve the exact diagnostic, including rollback details, behind an explicit disclosure.
  if (original && description !== original) {
    const details = document.createElement('details');
    const summary = document.createElement('summary');
    summary.textContent = t('details');
    const text = document.createElement('p');
    text.textContent = code + ': ' + original;
    details.append(summary, text);
    item.append(details);
  }
  element('results').append(item);
}
function renderResults(): void {
  element('results').replaceChildren();
  if (lastDetection) {
    const { scan, mappings, parsed } = lastDetection;
    const stateLabels = { found: 'foundStatus', missing: 'missingStatus', error: 'errorStatus', rendered: 'renderedStatus' } as const;
    for (const [index, cell] of scan.cells.slice(0, 100).entries()) {
      const item = document.createElement('li');
      const state = mappings?.[index]?.status;
      item.textContent = cell.worksheetName + '!' + cell.address + ' · ' + cell.imageId + ' · ' +
        t(state ? stateLabels[state] : 'uncheckedStatus');
      element('results').append(item);
    }
    for (const issue of (parsed?.issues ?? []).slice(0, 20)) appendIssue(issue.imageId ?? t('workbook'), issue.code, issue.message);
    if (lastDetection.resourceError) appendIssue(t('workbook'), lastDetection.resourceErrorCode ?? 'RESOURCE_READ_FAILED', lastDetection.resourceError);
  }
  for (const issue of (lastPreview?.issues ?? []).slice(0, 20)) {
    appendIssue(issue.location, issue.code ?? 'OPERATION_FAILED', issue.message);
  }
  if (lastErrorCode) appendIssue(t('workbook'), lastErrorCode, lastErrorMessage ?? '');
}
function applyLanguage(): void {
  locale = resolveLocale(preference, officeLanguage, navigator.language);
  document.documentElement.lang = locale;
  languageSelect.value = preference;
  for (const node of document.querySelectorAll<HTMLElement>('[data-i18n]')) {
    node.textContent = t(node.dataset.i18n as MessageKey);
  }
  for (const node of document.querySelectorAll<HTMLElement>('[data-i18n-label]')) {
    node.setAttribute('aria-label', t(node.dataset.i18nLabel as MessageKey));
  }
  element('app-version').textContent = t('currentVersion', { version: __APP_VERSION__ });
  status.textContent = statusText();
  hostStatus.textContent = hostText();
  renderResults();
  updateControls();
}
function clearResults(): void {
  lastDetection = undefined;
  lastPreview = undefined;
  lastErrorCode = undefined;
  lastErrorMessage = undefined;
  for (const id of ['image-count', 'sheet-count', 'parsed-count', 'missing-count', 'error-count']) element(id).textContent = '—';
  renderResults();
}
function displayDetection(detection: DetectionResult): void {
  lastDetection = detection;
  element('image-count').textContent = String(detection.scan.cells.length);
  element('sheet-count').textContent = String(detection.scan.worksheetCount);
  if (detection.mappings) {
    for (const [id, state] of [['parsed-count', 'found'], ['missing-count', 'missing'], ['error-count', 'error']]) {
      element(id!).textContent = String(detection.mappings.filter(item => item.status === state).length);
    }
  }
  renderResults();
}
function displayPreview(result: PreviewResult): void {
  lastPreview = result;
  statusText = () => t('previewComplete', { inserted: result.inserted, existing: result.existing, removed: result.removed, skipped: result.skipped }) +
    (result.issues.length ? ' ' + t('issueCount', { count: result.issues.length }) : '') + ' ' + t('preserved');
  status.textContent = statusText();
  renderResults();
}
async function runAction(operation: string, action: (signal: AbortSignal) => Promise<void>, cancellable = true): Promise<void> {
  if (!ready || busy) return;
  busy = true;
  cancelAllowed = cancellable;
  controller = new AbortController();
  lastOperation = operation;
  clearResults();
  updateControls();
  try { await action(controller.signal); }
  catch (error) {
    lastErrorCode = error instanceof WorkbookError ? error.code : 'OPERATION_FAILED';
    lastErrorMessage = error instanceof WorkbookError ? error.message : undefined;
    statusText = () => errorText(locale, lastErrorCode!, lastErrorMessage);
    status.textContent = statusText();
    renderResults();
  } finally { busy = false; cancelAllowed = false; controller = undefined; updateControls(); }
}
languageSelect.addEventListener('change', () => {
  preference = languagePreference(languageSelect.value);
  try { localStorage.setItem(languageKey, preference); } catch { /* Language still changes for this session. */ }
  applyLanguage();
});
applyLanguage();
const timeout = window.setTimeout(() => {
  if (!ready) { hostText = () => t('connectionTimeout'); hostStatus.textContent = hostText(); }
}, 15000);
if (typeof Office === 'undefined') {
  window.clearTimeout(timeout);
  hostText = () => t('officeMissing');
  hostStatus.textContent = hostText();
} else {
  Office.onReady().then(info => {
    window.clearTimeout(timeout);
    if (info.host !== Office.HostType.Excel) { hostText = () => t('excelOnly'); applyLanguage(); return; }
    officeLanguage = Office.context.displayLanguage;
    host = readHostCapabilities();
    if (!host.scan) { hostText = () => t('scanUnsupported'); applyLanguage(); return; }
    ready = true;
    shapesReady = canRenderImages();
    hostText = () => t('connected', { platform: String(info.platform) }) + (shapesReady ? '' : ' · ' + t('scanOnly'));
    applyLanguage();
  }).catch(() => {
    window.clearTimeout(timeout);
    hostText = () => t('initFailed');
    applyLanguage();
  });
}
element('scan').addEventListener('click', () => void runAction('scan', async signal => {
  setStatus('scanning');
  const detection = await detectWorkbook(({ worksheetName, scannedCells }) => {
    if (!signal.aborted) setStatus('scanProgress', { sheet: worksheetName, count: scannedCells });
  }, () => { if (!signal.aborted) setStatus('reading'); }, undefined, signal);
  displayDetection(detection);
  statusText = () => t('scanComplete', { count: detection.scan.scannedWorksheetCount }) +
    (detection.scan.cells.length > 100 ? ' ' + t('first100') : '') +
    (detection.resourceError ? ' ' + t('notChecked') + ' ' + errorText(locale, detection.resourceErrorCode ?? 'RESOURCE_READ_FAILED', detection.resourceError) : '') +
    (detection.parsed && !detection.parsed.hasCellImages ? ' ' + t('noParts') : '');
  status.textContent = statusText();
}));
for (const mode of ['show', 'refresh'] as const) {
  element(mode).addEventListener('click', () => void runAction(mode, async signal => {
    const result = await showWorkbookImages(mode, {
      keepAspectRatio: element<HTMLInputElement>('keep-aspect-ratio').checked,
      fitInsideCell: element<HTMLInputElement>('fit-inside-cell').checked,
    }, stage => { if (!signal.aborted) setStatus(stage); }, undefined, {
      signal, onRendering: () => { cancelAllowed = false; updateControls(); },
    });
    displayDetection(result.detection);
    displayPreview(result.preview);
  }));
}
element('remove').addEventListener('click', () => void runAction('remove', async () => {
  setStatus('removing');
  displayPreview(await removePreviewImages());
}, false));
element('cancel').addEventListener('click', () => {
  if (!busy || !cancelAllowed) return;
  controller?.abort();
  setStatus('cancelling');
  updateControls();
});
element('diagnostics').addEventListener('click', () => {
  if (!host || busy) return;
  const report = { ...createDiagnosticReport({ host, operation: lastOperation,
    detection: lastDetection, preview: lastPreview, errorCode: lastErrorCode }),
    appVersion: __APP_VERSION__, buildId: __BUILD_ID__, language: locale };
  const url = URL.createObjectURL(new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' }));
  const link = document.createElement('a');
  link.href = url;
  link.download = 'wps-image-compat-diagnostics.json';
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60000);
});
element('apply-update').addEventListener('click', () => {
  if (canApplyUpdate(busy, availableUpdate)) window.location.reload();
});
if (__CHECK_UPDATES__) {
  const check = createUpdateChecker({ version: __APP_VERSION__, buildId: __BUILD_ID__ }, release => {
    availableUpdate = release;
    updateControls();
  });
  void check();
  const checkVisible = () => { if (document.visibilityState === 'visible') void check(); };
  const interval = window.setInterval(checkVisible, 5 * 60 * 1000);
  window.addEventListener('focus', checkVisible);
  window.addEventListener('pagehide', () => {
    window.clearInterval(interval);
    window.removeEventListener('focus', checkVisible);
  }, { once: true });
}
