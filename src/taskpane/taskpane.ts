import './taskpane.css';
import { checkCompatibility } from '../core/compatibility/controller';
import type { CompatibilityReport } from '../core/compatibility/types';
import { renderCompatibility } from './compatibility-view';
import { navigateToCompatibilityCell } from '../core/compatibility/navigator';
import { convertWorkbook } from '../core/workbook-converter';
import type { WorkbookConversionResult } from '../core/workbook-converter';
import { detectWorkbook } from '../core/workbook-detector';
import type { DetectionResult } from '../core/workbook-detector';
import { showWorkbookImages } from '../core/preview-controller';
import { canRenderImages, removePreviewImages } from '../core/image-renderer';
import type { PreviewResult } from '../core/image-renderer';
import { recoverDamagedImageCells, recoverSelectedDamagedImageCell } from '../core/damaged-image-recovery';
import { scanExcelImages } from '../core/excel-image-scanner';
import type { ExcelImageScanResult } from '../core/excel-image-scanner';
import { inspectExcelEmbeddedImages } from '../core/excel-embedded-image-inspector';
import type { ExcelEmbeddedImageParseResult } from '../core/excel-drawing-parser';
import { canObserveWorksheetImageSelection, listManagedWorksheetImages, normalizeManagedWorksheetImageMetadata, toggleWorksheetImageZoom, watchManagedWorksheetImages } from '../core/worksheet-image-zoom';
import type { ManagedWorksheetImage } from '../core/worksheet-image-zoom';
import { WorkbookError } from '../utils/errors';
import { createDiagnosticReport, readHostCapabilities } from '../core/diagnostics';
import type { HostCapabilities } from '../core/diagnostics';
import { createUpdateChecker, canApplyUpdate } from '../core/update-checker';
import type { ReleaseInfo } from '../core/update-checker';
import { languagePreference, resolveLocale, translate, errorText } from './i18n';
import type { LanguagePreference, MessageKey, MessageParams } from './i18n';
import { featureEnabledPreference } from './feature-toggle';

const element = <T extends HTMLElement>(id: string): T => {
  const result = document.getElementById(id);
  if (!result) throw new Error('Missing UI element: ' + id);
  return result as T;
};
const status = element('scan-status');
const hostStatus = element('host-status');
const languageSelect = element<HTMLSelectElement>('language');
const autoToggleWorksheetImage = element<HTMLInputElement>('auto-toggle-worksheet-image');
const addinEnabled = element<HTMLInputElement>('addin-enabled');
const languageKey = 'wps-image-compat.language';
const featureEnabledKey = 'wps-image-compat.enabled';
let preference: LanguagePreference = 'auto';
try { preference = languagePreference(localStorage.getItem(languageKey)); } catch { /* Restricted storage: use this session. */ }
let featuresEnabled = true;
try { featuresEnabled = featureEnabledPreference(localStorage.getItem(featureEnabledKey)); } catch { /* Restricted storage: use this session. */ }
addinEnabled.checked = featuresEnabled;
let officeLanguage: string | undefined;
let locale = resolveLocale(preference, officeLanguage, navigator.language);
const t = (key: MessageKey, params?: MessageParams) => translate(locale, key, params);
let ready = false;
let shapesReady = false;
let busy = false;
let cancelAllowed = false;
let controller: AbortController | undefined;
let host: HostCapabilities | undefined;
let lastCompatibility: CompatibilityReport | undefined;
let lastConversion: WorkbookConversionResult | undefined;
let lastDetection: DetectionResult | undefined;
let lastExcelImages: ExcelImageScanResult | undefined;
let lastEmbeddedExcelImages: ExcelEmbeddedImageParseResult | undefined;
let lastPreview: PreviewResult | undefined;
let lastErrorCode: string | undefined;
let lastErrorMessage: string | undefined;
let lastOperation = 'initialize';
let availableUpdate: ReleaseInfo | undefined;
let statusText: () => string = () => '';
let hostText = () => t('connecting');
const worksheetImageToggles = new Set<string>();
let activatedWorksheetImage: ManagedWorksheetImage | undefined;
/** Watches worksheet selection so the optional automatic behavior and manual control use the same image. */
async function enableWorksheetImageInteraction(): Promise<void> {
  if (!featuresEnabled || !canObserveWorksheetImageSelection()) return;
  await watchManagedWorksheetImages(active => {
    const key = `${active.worksheetName}\u0000${active.shapeId}`;
    void (async () => {
      let toggling = false;
      try {
        if (!featuresEnabled) return;
        const image = (await listManagedWorksheetImages()).find(item =>
          item.worksheetName === active.worksheetName && item.shapeId === active.shapeId);
        if (!image) return;
        activatedWorksheetImage = image;
        updateControls();
        setStatus('worksheetImageSelected');
        if (!autoToggleWorksheetImage.checked || busy || worksheetImageToggles.has(key)) return;
        worksheetImageToggles.add(key);
        toggling = true;
        const zoomed = await toggleWorksheetImageZoom(image);
        activatedWorksheetImage = { ...image, zoomed };
        updateControls();
        setStatus(zoomed ? 'worksheetImageZoomed' : 'worksheetImageRestored');
      } catch {
        // An image may have been removed or a sheet protected after it was selected.
      } finally { if (toggling) worksheetImageToggles.delete(key); }
    })();
  }, deactivated => {
    if (activatedWorksheetImage?.worksheetName === deactivated.worksheetName && activatedWorksheetImage.shapeId === deactivated.shapeId) {
      activatedWorksheetImage = undefined;
      updateControls();
    }
  });
}
function setStatus(key: MessageKey, params?: MessageParams): void {
  statusText = () => t(key, params);
  status.textContent = statusText();
}
function updateControls(): void {
  const unavailable = !featuresEnabled || busy;
  addinEnabled.disabled = busy;
  element<HTMLButtonElement>('compatibility').disabled = !ready || unavailable;
  element<HTMLButtonElement>('compatibility-download').disabled = !featuresEnabled || !lastCompatibility || busy;
  element<HTMLButtonElement>('scan').disabled = !ready || unavailable;
  element<HTMLButtonElement>('inspect-excel-images').disabled = !ready || unavailable;
  for (const id of ['show', 'refresh', 'remove']) element<HTMLButtonElement>(id).disabled = !shapesReady || unavailable;
  element<HTMLButtonElement>('repair-damaged-images').disabled = !shapesReady || unavailable;
  element<HTMLButtonElement>('repair-selected-image').disabled = !shapesReady || unavailable;
  const toggleSelected = element<HTMLButtonElement>('toggle-selected-image');
  toggleSelected.disabled = !featuresEnabled || !activatedWorksheetImage || busy;
  toggleSelected.textContent = t(activatedWorksheetImage?.zoomed ? 'restoreSelectedImage' : 'zoomSelectedImage');
  element<HTMLFieldSetElement>('preview-settings').disabled = !shapesReady || unavailable;
  element<HTMLFieldSetElement>('conversion-settings').disabled = !ready || unavailable;
  const selectedConversion = element<HTMLInputElement>('convert-dispimg').checked ||
    element<HTMLInputElement>('freeze-unsupported').checked || element<HTMLInputElement>('freeze-external').checked;
  element<HTMLButtonElement>('convert').disabled = !ready || unavailable || !selectedConversion ||
    !element<HTMLInputElement>('conversion-confirm').checked;
  element<HTMLButtonElement>('cancel').disabled = !featuresEnabled || !busy || !cancelAllowed || !!controller?.signal.aborted;
  element<HTMLButtonElement>('diagnostics').disabled = !featuresEnabled || !host || busy;
  element<HTMLButtonElement>('apply-update').disabled = !featuresEnabled || !canApplyUpdate(busy, availableUpdate);
  element('addin-paused-hint').hidden = featuresEnabled;
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
  renderCompatibility(element('compatibility-results'), lastCompatibility, locale, location => {
    if (busy) return;
    setStatus('compatOpeningLocation', { sheet: location.worksheetName, address: location.address });
    void navigateToCompatibilityCell(location).then(() => {
      setStatus('compatLocationOpened', { sheet: location.worksheetName, address: location.address });
    }).catch(error => {
      lastErrorCode = error instanceof WorkbookError ? error.code : 'LOCATION_NOT_FOUND';
      lastErrorMessage = error instanceof Error ? error.message : undefined;
      statusText = () => errorText(locale, lastErrorCode!, lastErrorMessage);
      status.textContent = statusText();
      renderResults();
    });
  });
  element('results').replaceChildren();
  if (lastDetection) {
    const { scan, mappings, parsed } = lastDetection;
    const item = document.createElement('li');
    const found = mappings?.filter(mapping => mapping.status === 'found').length ?? 0;
    const unresolved = mappings ? mappings.length - found : scan.cells.length;
    item.textContent = t('imageSummary', { found, unresolved });
    element('results').append(item);
    for (const issue of (parsed?.issues ?? []).slice(0, 20)) appendIssue(issue.imageId ?? t('workbook'), issue.code, issue.message);
    if (lastDetection.resourceError) appendIssue(t('workbook'), lastDetection.resourceErrorCode ?? 'RESOURCE_READ_FAILED', lastDetection.resourceError);
  }
  if (lastExcelImages?.supported) {
    const item = document.createElement('li');
    item.textContent = t('excelImageSummary', { images: lastExcelImages.imageCount, sheets: lastExcelImages.worksheetCount });
    element('results').append(item);
  }
  if (lastEmbeddedExcelImages) {
    const sheets = new Set(lastEmbeddedExcelImages.images.map(image => image.worksheetName)).size;
    const item = document.createElement('li');
    item.textContent = t('excelEmbeddedSummary', { images: lastEmbeddedExcelImages.images.length, sheets });
    element('results').append(item);
    for (const issue of lastEmbeddedExcelImages.issues.slice(0, 20)) appendIssue(issue.location ?? t('workbook'), issue.code, issue.message);
  }
  for (const issue of (lastPreview?.issues ?? []).slice(0, 20)) {
    appendIssue(issue.location, issue.code ?? 'OPERATION_FAILED', issue.message);
  }
  if (lastConversion) {
    const item = document.createElement('li');
    item.textContent = t('conversionComplete', { images: lastConversion.convertedImages,
      formulas: lastConversion.clearedDispimgFormulas + lastConversion.frozenFormulas,
      skipped: lastConversion.skipped, broken: lastConversion.unresolvedBrokenReferences });
    element('results').append(item);
    for (const issue of lastConversion.issues.slice(0, 20)) appendIssue(issue.location, issue.code, issue.message);
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
  lastCompatibility = undefined;
  lastConversion = undefined;
  lastDetection = undefined;
  lastExcelImages = undefined;
  lastEmbeddedExcelImages = undefined;
  activatedWorksheetImage = undefined;
  lastPreview = undefined;
  lastErrorCode = undefined;
  lastErrorMessage = undefined;
  for (const id of ['image-count', 'sheet-count', 'excel-image-count', 'excel-embedded-count', 'parsed-count', 'missing-count', 'error-count']) element(id).textContent = '—';
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
function displayExcelImages(result: ExcelImageScanResult): void {
  lastExcelImages = result;
  element('excel-image-count').textContent = result.supported ? String(result.imageCount) : '—';
  renderResults();
}
function displayEmbeddedExcelImages(result: ExcelEmbeddedImageParseResult): void {
  lastEmbeddedExcelImages = result;
  element('excel-embedded-count').textContent = String(result.images.length);
  renderResults();
}
function displayPreview(result: PreviewResult): void {
  lastPreview = result;
  statusText = () => t('previewComplete', { inserted: result.inserted, existing: result.existing, removed: result.removed, skipped: result.skipped }) +
    (result.issues.length ? ' ' + t('issueCount', { count: result.issues.length }) : '') + ' ' + t('preserved');
  status.textContent = statusText();
  renderResults();
  void enableWorksheetImageInteraction().catch(() => { /* Direct worksheet interaction is optional. */ });
}
function displayConversion(result: WorkbookConversionResult): void {
  lastConversion = result;
  statusText = () => t('conversionComplete', { images: result.convertedImages,
    formulas: result.clearedDispimgFormulas + result.frozenFormulas, skipped: result.skipped,
    broken: result.unresolvedBrokenReferences });
  status.textContent = statusText();
  renderResults();
  void enableWorksheetImageInteraction().catch(() => { /* Direct worksheet interaction is optional. */ });
}
async function runAction(operation: string, action: (signal: AbortSignal) => Promise<void>, cancellable = true, preserveResults = false): Promise<void> {
  if (!featuresEnabled || !ready || busy) return;
  busy = true;
  cancelAllowed = cancellable;
  controller = new AbortController();
  lastOperation = operation;
  if (!preserveResults) clearResults();
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
addinEnabled.addEventListener('change', () => {
  featuresEnabled = addinEnabled.checked;
  try { localStorage.setItem(featureEnabledKey, String(featuresEnabled)); } catch { /* Preference still applies in this pane. */ }
  if (!featuresEnabled) activatedWorksheetImage = undefined;
  if (featuresEnabled && shapesReady) {
    void normalizeManagedWorksheetImageMetadata().then(enableWorksheetImageInteraction)
      .catch(() => { /* Legacy cleanup and image observation are optional. */ });
  }
  updateControls();
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
    if (featuresEnabled && shapesReady) void normalizeManagedWorksheetImageMetadata()
      .then(enableWorksheetImageInteraction).catch(() => { /* Legacy image cleanup is optional. */ });
    hostText = () => t('connected', { platform: String(info.platform) }) + (shapesReady ? '' : ' · ' + t('scanOnly'));
    applyLanguage();
  }).catch(() => {
    window.clearTimeout(timeout);
    hostText = () => t('initFailed');
    applyLanguage();
  });
}
element('compatibility').addEventListener('click', () => void runAction('compatibility', async signal => {
  setStatus('compatScanning', { count: 0 });
  lastCompatibility = await checkCompatibility(signal, count => setStatus('compatScanning', { count }));
  renderResults();
  setStatus('compatDone');
}));
element('compatibility-download').addEventListener('click', () => {
  if (!lastCompatibility || busy) return;
  const report = { ...lastCompatibility, appVersion: __APP_VERSION__, buildId: __BUILD_ID__, host,
    scope: { cells: 'Current formulas and host-typed formula errors; excludes constants, defined names and INDIRECT text targets.',
      metadata: 'xl/workbook.xml date system and externalReference IDs only; no external target validation.',
      excluded: ['connections', 'charts', 'macros', 'layout', 'date value interpretation'],
      timing: 'Live cells and workbook snapshot read sequentially; not an atomic snapshot.' } };
  const url = URL.createObjectURL(new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' }));
  const link = document.createElement('a');
  link.href = url; link.download = 'wps-excel-compatibility-report.json';
  document.body.append(link); link.click(); link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60000);
});
for (const id of ['convert-dispimg', 'freeze-unsupported', 'freeze-external', 'conversion-confirm']) {
  element<HTMLInputElement>(id).addEventListener('change', updateControls);
}
autoToggleWorksheetImage.addEventListener('change', updateControls);
element('convert').addEventListener('click', () => void runAction('convert', async () => {
  const result = await convertWorkbook({
    convertDispimg: element<HTMLInputElement>('convert-dispimg').checked,
    freezeUnsupported: element<HTMLInputElement>('freeze-unsupported').checked,
    freezeExternal: element<HTMLInputElement>('freeze-external').checked,
    keepAspectRatio: element<HTMLInputElement>('keep-aspect-ratio').checked,
    fitInsideCell: element<HTMLInputElement>('fit-inside-cell').checked,
  }, stage => {
    const message = { reading: 'conversionReading', checking: 'conversionChecking',
      convertingImages: 'conversionImages', convertingFormulas: 'conversionFormulas' } as const;
    setStatus(message[stage]);
  });
  displayConversion(result);
}, false));
element('scan').addEventListener('click', () => void runAction('scan', async signal => {
  setStatus('scanning');
  if (shapesReady) {
    try { await normalizeManagedWorksheetImageMetadata(); }
    catch { /* Legacy descriptions do not prevent a read-only scan. */ }
  }
  const detection = await detectWorkbook(({ worksheetName, scannedCells }) => {
    if (!signal.aborted) setStatus('scanProgress', { sheet: worksheetName, count: scannedCells });
  }, () => { if (!signal.aborted) setStatus('reading'); }, undefined, signal);
  displayDetection(detection);
  displayExcelImages(await scanExcelImages());
  statusText = () => t('scanComplete', { count: detection.scan.scannedWorksheetCount }) +
    (detection.scan.cells.length > 100 ? ' ' + t('first100') : '') +
    (detection.resourceError ? ' ' + t('notChecked') + ' ' + errorText(locale, detection.resourceErrorCode ?? 'RESOURCE_READ_FAILED', detection.resourceError) : '') +
    (detection.parsed && !detection.parsed.hasCellImages ? ' ' + t('noParts') : '');
  status.textContent = statusText();
}));
element('inspect-excel-images').addEventListener('click', () => void runAction('inspect-excel-images', async signal => {
  setStatus('reading');
  const inventory = await inspectExcelEmbeddedImages(signal);
  displayEmbeddedExcelImages(inventory);
  setStatus('excelPictureInventory', { images: inventory.images.length });
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
async function repairImageMetadata(selectedOnly: boolean): Promise<void> {
  await runAction(selectedOnly ? 'repair-selected-image' : 'repair-damaged-images', async () => {
  setStatus('reading');
  const recovered = await (selectedOnly ? recoverSelectedDamagedImageCell : recoverDamagedImageCells)({
    keepAspectRatio: element<HTMLInputElement>('keep-aspect-ratio').checked,
    fitInsideCell: element<HTMLInputElement>('fit-inside-cell').checked,
  });
  displayPreview(recovered.preview);
  statusText = () => selectedOnly && recovered.detected === 0 ? t('repairSelectedEmpty') :
    t('repairComplete', { detected: recovered.detected, restored: recovered.restored, cleared: recovered.cleared });
  status.textContent = statusText();
  }, false);
}
element('repair-damaged-images').addEventListener('click', () => void repairImageMetadata(false));
element('repair-selected-image').addEventListener('click', () => void repairImageMetadata(true));
element('toggle-selected-image').addEventListener('click', () => {
  const selected = activatedWorksheetImage;
  if (!selected) return;
  void runAction('worksheet-image-zoom', async () => {
    const image = (await listManagedWorksheetImages()).find(item =>
      item.worksheetName === selected.worksheetName && item.shapeId === selected.shapeId);
    if (!image) throw new WorkbookError('IMAGE_NOT_FOUND', 'The selected WPS Image Compat picture is no longer available.');
    const zoomed = await toggleWorksheetImageZoom(image);
    activatedWorksheetImage = { ...image, zoomed };
    setStatus(zoomed ? 'worksheetImageZoomed' : 'worksheetImageRestored');
  }, false, true);
});
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
