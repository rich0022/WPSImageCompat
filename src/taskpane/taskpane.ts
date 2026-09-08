import './taskpane.css';
import { detectWorkbook } from '../core/workbook-detector';

const element = <T extends HTMLElement>(id: string): T => {
  const result = document.getElementById(id);
  if (!result) throw new Error(`Missing UI element: ${id}`);
  return result as T;
};
const scanButton = element<HTMLButtonElement>('scan');
const status = element('scan-status');
const hostStatus = element('host-status');
let ready = false;

const timeout = window.setTimeout(() => {
  if (!ready) hostStatus.textContent = 'Excel has not connected. Open this add-in inside Excel and check your connection.';
}, 15000);

if (typeof Office === 'undefined') {
  window.clearTimeout(timeout);
  hostStatus.textContent = 'Office.js could not load. Check your network connection and reopen the add-in in Excel.';
} else {
  Office.onReady().then((info) => {
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
    hostStatus.textContent = `Connected to Excel · ${info.platform}`;
    scanButton.disabled = false;
  }).catch(() => {
    window.clearTimeout(timeout);
    hostStatus.textContent = 'Could not initialize the Excel connection. Reopen the add-in.';
  });
}

scanButton.addEventListener('click', async () => {
  if (!ready || scanButton.disabled) return;
  scanButton.disabled = true;
  element('image-count').textContent = '—';
  element('sheet-count').textContent = '—';
  element('results').replaceChildren();
  for (const id of ['parsed-count', 'missing-count', 'error-count']) element(id).textContent = '—';
  status.textContent = 'Scanning workbook…';
  try {
    const detection = await detectWorkbook(({ worksheetName, scannedCells }) => {
      status.textContent = `Scanning ${worksheetName} · ${scannedCells.toLocaleString()} cells checked…`;
    }, () => { status.textContent = 'Reading workbook and parsing WPS image resources…'; });
    const result = detection.scan;
    if (detection.mappings) {
      element('parsed-count').textContent = String(detection.mappings.filter(item => item.status === 'found').length);
      element('missing-count').textContent = String(detection.mappings.filter(item => item.status === 'missing').length);
      element('error-count').textContent = String(detection.mappings.filter(item => item.status === 'error').length);
    }
    element('image-count').textContent = String(result.cells.length);
    element('sheet-count').textContent = String(result.worksheetCount);
    const list = element('results');
    for (const [index, cell] of result.cells.slice(0, 100).entries()) {
      const item = document.createElement('li');
      item.textContent = `${cell.worksheetName}!${cell.address} · ${cell.imageId} · ${detection.mappings?.[index]?.status ?? 'not checked'}`;
      list.append(item);
    }
    status.textContent = `Scan complete. ${result.scannedWorksheetCount} worksheets checked.` +
      (result.cells.length > 100 ? ' Showing the first 100 matches.' : '') +
      (detection.resourceError ? ` Resources not checked: ${detection.resourceError}` : '') +
      (detection.parsed && !detection.parsed.hasCellImages ? ' This workbook snapshot has no WPS image part. Excel may not have preserved it.' : '') +
      (detection.parsed?.issues.length ? ` ${detection.parsed.issues.length} resource issues found.` : '');
    for (const issue of (detection.parsed?.issues ?? []).slice(0, 20)) {
      const item = document.createElement('li');
      item.textContent = `${issue.imageId ?? 'Workbook'}: ${issue.message}`;
      list.append(item);
    }
  } catch (error: unknown) {
    const code = typeof OfficeExtension !== 'undefined' && error instanceof OfficeExtension.Error ? ` (${error.code})` : '';
    status.textContent = `Scan failed${code}. No workbook data was changed. Please retry.`;
  } finally {
    scanButton.disabled = false;
  }
});
