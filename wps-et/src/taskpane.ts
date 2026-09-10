import '../ui/taskpane.css';
import { scanWpsEtWorkbook, selectWpsEtPicture, type WpsEtScanResult } from './wps-et-scanner';

declare global { interface Window { Application?: unknown; wps?: unknown } }
const byId = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;
const scan = byId<HTMLButtonElement>('scan');
const selectFirst = byId<HTMLButtonElement>('select-first');
const status = byId('status');
let result: WpsEtScanResult | undefined;
const host = () => window.Application ?? window.wps;
function showScan(scanResult: WpsEtScanResult): void {
  result = scanResult;
  byId('dispimg-count').textContent = String(scanResult.dispimgCells.length);
  byId('picture-count').textContent = String(scanResult.pictures.length);
  byId('sheet-count').textContent = String(scanResult.worksheetCount);
  selectFirst.disabled = scanResult.pictures.length === 0;
  const list = byId('results'); list.replaceChildren();
  for (const cell of scanResult.dispimgCells.slice(0, 8)) {
    const item = document.createElement('li'); item.textContent = `DISPIMG：${cell.worksheetName}!${cell.address}${cell.imageId ? ` · ${cell.imageId}` : ''}`; list.append(item);
  }
  for (const picture of scanResult.pictures.slice(0, 8)) {
    const item = document.createElement('li'); item.textContent = `图片：${picture.worksheetName}${picture.address ? `!${picture.address}` : ''}${picture.name ? ` · ${picture.name}` : ''}`; list.append(item);
  }
  for (const issue of scanResult.issues.slice(0, 8)) { const item = document.createElement('li'); item.textContent = `${issue.worksheetName ?? '工作簿'}：${issue.message}`; list.append(item); }
  status.textContent = `扫描完成：读取 ${scanResult.scannedCells} 个单元格。`;
}
scan.addEventListener('click', () => void (async () => {
  const application = host();
  if (!application) { status.textContent = '未检测到 WPS 加载项接口。请从 WPS 表格功能区打开此窗格。'; return; }
  scan.disabled = true; status.textContent = '正在扫描当前 WPS 工作簿…';
  try { showScan(await scanWpsEtWorkbook(application)); } catch (error) { status.textContent = error instanceof Error ? error.message : '扫描失败。'; }
  finally { scan.disabled = false; }
})());
selectFirst.addEventListener('click', () => void (async () => {
  const first = result?.pictures[0], application = host();
  if (!first || !application) return;
  try { await selectWpsEtPicture(application, first); status.textContent = `已选中 ${first.worksheetName}${first.address ? `!${first.address}` : ''} 的图片。`; }
  catch (error) { status.textContent = error instanceof Error ? error.message : '无法选中图片。'; }
})());
