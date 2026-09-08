import './taskpane.css';
import { detectWorkbook } from '../core/workbook-detector';
import type { DetectionResult } from '../core/workbook-detector';
import { showWorkbookImages } from '../core/preview-controller';
import { canRenderImages, removePreviewImages } from '../core/image-renderer';
import type { PreviewResult } from '../core/image-renderer';
import { WorkbookError } from '../utils/errors';

const element = <T extends HTMLElement>(id: string): T => {
  const result = document.getElementById(id);
  if (!result) throw new Error(`Missing UI element: ${id}`);
  return result as T;
};
const status = element('scan-status');
const hostStatus = element('host-status');
let ready = false;
let shapesReady = false;
let busy = false;
function updateControls(): void {
  element<HTMLButtonElement>('scan').disabled = !ready || busy;
  for (const id of ['show', 'refresh', 'remove']) element<HTMLButtonElement>(id).disabled = !shapesReady || busy;
  element<HTMLFieldSetElement>('preview-settings').disabled = !shapesReady || busy;
}
function clearResults(): void {
  for (const id of ['image-count', 'sheet-count', 'parsed-count', 'missing-count', 'error-count']) element(id).textContent = '—';
  element('results').replaceChildren();
}
function displayDetection(detection: DetectionResult): void {
  const { scan, mappings, parsed } = detection;
  element('image-count').textContent = String(scan.cells.length);
  element('sheet-count').textContent = String(scan.worksheetCount);
  if (mappings) {
    for (const [id, state] of [['parsed-count', 'found'], ['missing-count', 'missing'], ['error-count', 'error']]) {
      element(id!).textContent = String(mappings.filter(item => item.status === state).length);
    }
  }
  const list = element('results');
  list.replaceChildren();
  for (const [index, cell] of scan.cells.slice(0, 100).entries()) {
    const item = document.createElement('li');
    item.textContent = `${cell.worksheetName}!${cell.address} · ${cell.imageId} · ${mappings?.[index]?.status ?? 'not checked'}`;
    list.append(item);
  }
  for (const issue of (parsed?.issues ?? []).slice(0, 20)) {
    const item = document.createElement('li');
    item.textContent = `${issue.imageId ?? 'Workbook'}: ${issue.message}`;
    list.append(item);
  }
}
function displayPreview(result: PreviewResult): void {
  status.textContent = `${result.inserted} inserted · ${result.existing} already shown · ${result.removed} removed · ${result.skipped} skipped.` +
    (result.issues.length ? ` ${result.issues.length} issues; see details below.` : '') + ' Original formulas were preserved.';
  for (const issue of result.issues.slice(0, 20)) {
    const item = document.createElement('li');
    item.textContent = `${issue.location}: ${issue.message}`;
    element('results').append(item);
  }
}
async function runAction(action: () => Promise<void>): Promise<void> {
  if (!ready || busy) return;
  busy = true;
  updateControls();
  try { await action(); }
  catch (error) {
    status.textContent = error instanceof WorkbookError ? error.message :
      'The operation did not complete. Some preview changes may have occurred. Check worksheet protection and retry. Original formulas were not changed.';
  } finally { busy = false; updateControls(); }
}
const timeout = window.setTimeout(() => {
  if (!ready) hostStatus.textContent = 'Excel has not connected. Open this add-in inside Excel and check your connection.';
}, 15000);
if (typeof Office === 'undefined') {
  window.clearTimeout(timeout);
  hostStatus.textContent = 'Office.js could not load. Check your network and reopen the add-in in Excel.';
} else {
  Office.onReady().then(info => {
    window.clearTimeout(timeout);
    if (info.host !== Office.HostType.Excel) {
      hostStatus.textContent = 'Open this add-in inside Microsoft Excel to scan a workbook.';
      return;
    }
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.4')) {
      hostStatus.textContent = 'This Excel version does not support ExcelApi 1.4. Please update Excel.';
      return;
    }
    ready = true;
    shapesReady = canRenderImages();
    hostStatus.textContent = `Connected to Excel · ${info.platform}` +
      (shapesReady ? '' : ' · Scanning only. Image previews require ExcelApi 1.10.');
    updateControls();
  }).catch(() => {
    window.clearTimeout(timeout);
    hostStatus.textContent = 'Could not initialize the Excel connection. Reopen the add-in.';
  });
}
element('scan').addEventListener('click', () => void runAction(async () => {
  clearResults();
  status.textContent = 'Scanning workbook…';
  const detection = await detectWorkbook(({ worksheetName, scannedCells }) => {
    status.textContent = `Scanning ${worksheetName} · ${scannedCells.toLocaleString()} cells checked…`;
  }, () => { status.textContent = 'Reading workbook and parsing WPS image resources…'; });
  displayDetection(detection);
  status.textContent = `Scan complete. ${detection.scan.scannedWorksheetCount} worksheets checked.` +
    (detection.scan.cells.length > 100 ? ' Showing the first 100 matches.' : '') +
    (detection.resourceError ? ` Resources not checked: ${detection.resourceError}` : '') +
    (detection.parsed && !detection.parsed.hasCellImages ? ' This workbook snapshot has no WPS image part. Excel may not have preserved it.' : '');
}));
for (const mode of ['show', 'refresh'] as const) {
  element(mode).addEventListener('click', () => void runAction(async () => {
    clearResults();
    const result = await showWorkbookImages(mode, {
      keepAspectRatio: element<HTMLInputElement>('keep-aspect-ratio').checked,
      fitInsideCell: element<HTMLInputElement>('fit-inside-cell').checked,
    }, text => { status.textContent = text; });
    displayDetection(result.detection);
    displayPreview(result.preview);
  }));
}
element('remove').addEventListener('click', () => void runAction(async () => {
  clearResults();
  status.textContent = 'Removing plugin preview images…';
  displayPreview(await removePreviewImages());
}));
