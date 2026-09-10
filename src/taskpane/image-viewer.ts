import type { ImageMapping } from '../types/wps';

export interface ViewableImage {
  key: string;
  label: string;
  src: string;
}

/** Converts already-parsed local resources into safe task-pane image sources. */
export function collectViewableImages(mappings: readonly ImageMapping[] | undefined): ViewableImage[] {
  const seen = new Set<string>();
  const result: ViewableImage[] = [];
  for (const mapping of mappings ?? []) {
    const resource = mapping.resource;
    if (mapping.status !== 'found' || !resource?.base64 || !/^image\/(png|jpeg)$/i.test(resource.mimeType)) continue;
    const key = `${mapping.cell.worksheetName}\u0000${mapping.cell.address}\u0000${mapping.cell.imageId}`;
    if (seen.has(key)) continue;
    seen.add(key);
    result.push({ key, label: `${mapping.cell.worksheetName}!${mapping.cell.address} · ${mapping.cell.imageId}`,
      src: `data:${resource.mimeType};base64,${resource.base64}` });
  }
  return result;
}

export interface ImageViewer {
  setImages(images: readonly ViewableImage[], openLabel: string): void;
  hasImages(): boolean;
  close(): void;
}

/** Task-pane-only viewer: no workbook data is sent outside the add-in. */
export function createImageViewer(): ImageViewer {
  const panel = document.getElementById('image-viewer') as HTMLElement;
  const grid = document.getElementById('image-viewer-grid') as HTMLElement;
  const image = document.getElementById('image-viewer-image') as HTMLImageElement;
  const caption = document.getElementById('image-viewer-caption') as HTMLElement;
  const closeButton = document.getElementById('image-viewer-close') as HTMLButtonElement;
  let images: readonly ViewableImage[] = [];
  const close = () => { panel.hidden = true; image.removeAttribute('src'); image.classList.remove('image-viewer-zoomed'); };
  const open = (selected: ViewableImage) => {
    image.src = selected.src;
    image.alt = selected.label;
    caption.textContent = selected.label;
    image.classList.remove('image-viewer-zoomed');
    panel.hidden = false;
    closeButton.focus();
  };
  closeButton.addEventListener('click', close);
  panel.addEventListener('click', event => { if (event.target === panel) close(); });
  document.addEventListener('keydown', event => { if (event.key === 'Escape' && !panel.hidden) close(); });
  image.addEventListener('dblclick', () => image.classList.toggle('image-viewer-zoomed'));
  return {
    setImages(next, openLabel) {
      images = next;
      grid.replaceChildren();
      for (const candidate of images) {
        const button = document.createElement('button');
        button.type = 'button'; button.className = 'image-viewer-thumbnail';
        button.title = `${openLabel}: ${candidate.label}`;
        button.setAttribute('aria-label', button.title);
        const thumbnail = document.createElement('img');
        thumbnail.src = candidate.src; thumbnail.alt = candidate.label;
        button.append(thumbnail);
        button.addEventListener('dblclick', () => open(candidate));
        button.addEventListener('click', () => open(candidate));
        grid.append(button);
      }
      if (!images.length) close();
    },
    hasImages: () => images.length > 0,
    close,
  };
}
