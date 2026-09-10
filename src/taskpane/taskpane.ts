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
import { collectViewableImages, createImageViewer } from './image-viewer';
import { canObserveWorksheetImageSelection, listManagedWorksheetImages, normalizeManagedWorksheetImageMetadata, toggleWorksheetImageZoom, watchManagedWorksheetImages } from '../core/worksheet-image-zoom';
import type { ManagedWorksheetImage } from '../core/worksheet-image-zoom';
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
let lastCompatibility: CompatibilityReport | undefined;
let lastConversion: WorkbookConversionResult | undefined;
let lastDetection: DetectionResult | undefined;
let lastPreview: PreviewResult | undefined;
let lastErrorCode: string | undefined;
let lastErrorMessage: string | undefined;
let lastOperation = 'initialize';
let availableUpdate: ReleaseInfo | undefined;
let statusText: () => string = () => '';
let hostText = () => t('connecting');
const imageViewer = createImageViewer();
let managedWorksheetImages: ManagedWorksheetImage[] = [];
let activatedWorksheetImage: ManagedWorksheetImage | undefined;
async function watchWorksheetImages(): Promise<void> {
  await watchManagedWorksheetImages(active => {
    void listManagedWorksheetImages().then(images => {
      activatedWorksheetImage = images.find(image => image.shapeId === active.shapeId && image.worksheetName === active.worksheetName);
      if (activatedWorksheetImage) setStatus('worksheetImageSelected');
      updateControls();
    }).catch(() => { /* The list fallback remains available. */ });
  });
}
function renderManagedWorksheetImages(): void {
  const panel = element('worksheet-image-panel'); const list = element('worksheet-image-list');
  panel.hidden = !managedWorksheetImages.length; list.replaceChildren();
  for (const image of managedWorksheetImages) {
    const item = document.createElement('li');
    const button = document.createElement('button'); button.className = 'location-button';
    button.textContent = image.zoomed ? t('worksheetImageRestore') : t('worksheetImageZoom');
    button.addEventListener('click', () => void runAction('worksheet-image-zoom', async () => {
      const zoomed = await toggleWorksheetImageZoom(image);
      managedWorksheetImages = await listManagedWorksheetImages();
      renderManagedWorksheetImages();
      setStatus(zoomed ? 'worksheetImageZoomed' : 'worksheetImageRestored');
    }, false));
    item.append(`${image.worksheetName}!${image.address} · ${image.kind === 'converted' ? t('worksheetImageConverted') : t('worksheetImagePreview')} · `, button);
    list.append(item);
  }
}
function setStatus(key: MessageKey, params?: MessageParams): void {
  statusText = () => t(key, params);
  status.textContent = statusText();
}
function updateControls(): void {
  element<HTMLButtonElement>('compatibility').disabled = !ready || busy;
  element<HTMLButtonElement>('compatibility-download').disabled = !lastCompatibility || busy;
  element<HTMLButtonElement>('scan').disabled = !ready || busy;
  for (const id of ['show', 'refresh', 'remove']) element<HTMLButtonElement>(id).disabled = !shapesReady || busy;
  element<HTMLButtonElement>('view-images').disabled = !imageViewer.hasImages() || busy;
  element<HTMLButtonElement>('manage-worksheet-images').disabled = !ready || busy;
  const selectedZoom = element<HTMLButtonElement>('zoom-selected-image');
  selectedZoom.disabled = !activatedWorksheetImage || busy;
  selectedZoom.textContent = t(activatedWorksheetImage?.zoomed ? 'restoreSelectedImage' : 'zoomSelectedImage');
  element<HTMLFieldSetElement>('preview-settings').disabled = !shapesReady || busy;
  element<HTMLFieldSetElement>('conversion-settings').disabled = !ready || busy;
  const selectedConversion = element<HTMLInputElement>('convert-dispimg').checked ||
    element<HTMLInputElement>('freeze-unsupported').checked || element<HTMLInputElement>('freeze-external').checked;
  element<HTMLButtonElement>('convert').disabled = !ready || busy || !selectedConversion ||
    !element<HTMLInputElement>('conversion-confirm').checked;
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
  imageViewer.setImages(collectViewableImages(lastDetection?.mappings), t('imageViewerOpen'));
  renderManagedWorksheetImages();
  renderResults();
  updateControls();
}
function clearResults(): void {
  lastCompatibility = undefined;
  lastConversion = undefined;
  lastDetection = undefined;
  lastPreview = undefined;
  lastErrorCode = undefined;
  lastErrorMessage = undefined;
  imageViewer.setImages([], t('imageViewerOpen'));
  managedWorksheetImages = [];
  activatedWorksheetImage = undefined;
  renderManagedWorksheetImages();
  for (const id of ['image-count', 'sheet-count', 'parsed-count', 'missing-count', 'error-count']) element(id).textContent = '—';
  renderResults();
}
function displayDetection(detection: DetectionResult): void {
  lastDetection = detection;
  imageViewer.setImages(collectViewableImages(detection.mappings), t('imageViewerOpen'));
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
  void watchWorksheetImages().catch(() => { /* Shape activation is optional; the list remains available. */ });
}
function displayConversion(result: WorkbookConversionResult): void {
  lastConversion = result;
  statusText = () => t('conversionComplete', { images: result.convertedImages,
    formulas: result.clearedDispimgFormulas + result.frozenFormulas, skipped: result.skipped,
    broken: result.unresolvedBrokenReferences });
  status.textContent = statusText();
  renderResults();
  void watchWorksheetImages().catch(() => { /* Shape activation is optional; the list remains available. */ });
}
async function runAction(operation: string, action: (signal: AbortSignal) => Promise<void>, cancellable = true, preserveResults = false): Promise<void> {
  if (!ready || busy) return;
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
    if (canObserveWorksheetImageSelection()) void normalizeManagedWorksheetImageMetadata()
      .then(watchWorksheetImages).catch(() => { /* optional capability */ });
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
element('view-images').addEventListener('click', () => {
  if (busy || !imageViewer.hasImages()) return;
  document.querySelector<HTMLButtonElement>('.image-viewer-thumbnail')?.click();
});
element('manage-worksheet-images').addEventListener('click', () => void runAction('worksheet-images', async () => {
  managedWorksheetImages = await listManagedWorksheetImages();
  renderManagedWorksheetImages();
  setStatus(managedWorksheetImages.length ? 'worksheetImagesReady' : 'worksheetImagesEmpty');
}, false));
element('zoom-selected-image').addEventListener('click', () => {
  const selected = activatedWorksheetImage;
  if (!selected) return;
  void runAction('worksheet-image-zoom', async () => {
  const image = (await listManagedWorksheetImages()).find(item => item.shapeId === selected.shapeId && item.worksheetName === selected.worksheetName);
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
