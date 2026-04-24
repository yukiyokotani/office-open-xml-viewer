/**
 * Webview bootstrap script.
 *
 * This script runs inside the VSCode Webview iframe. It:
 * 1. Waits for the `ooxml-init` message from the extension host containing the file bytes.
 * 2. Instantiates the appropriate Viewer (XlsxViewer / DocxViewer / PptxViewer).
 * 3. Bridges selection events back to the extension host via acquireVsCodeApi().postMessage().
 */

// These globals are injected by the HTML template.
declare const __OOXML_FILE_TYPE__: 'xlsx' | 'docx' | 'pptx';
declare const __OOXML_WASM_URL__: string;

// VSCode API bridge — available only inside a Webview context.
declare function acquireVsCodeApi(): {
  postMessage(msg: unknown): void;
  getState(): unknown;
  setState(state: unknown): void;
};

import { XlsxViewer } from '@silurus/ooxml-xlsx';
import { DocxViewer } from '@silurus/ooxml-docx';
import { PptxViewer } from '@silurus/ooxml-pptx';
import type { CellRange } from '@silurus/ooxml-xlsx';

const vscodeApi = acquireVsCodeApi();
const fileType = __OOXML_FILE_TYPE__;

const statusEl = document.getElementById('status')!;
const navBar = document.getElementById('nav-bar')!;
const prevBtn = document.getElementById('prev-btn') as HTMLButtonElement;
const nextBtn = document.getElementById('next-btn') as HTMLButtonElement;
const pageInfo = document.getElementById('page-info')!;
const viewerContainer = document.getElementById('viewer-container')!;

function setStatus(msg: string): void {
  statusEl.style.display = '';
  statusEl.textContent = msg;
}

function hideStatus(): void {
  statusEl.style.display = 'none';
}

window.addEventListener('message', async (event: MessageEvent) => {
  const msg = event.data;
  if (msg.type !== 'ooxml-init') return;

  const buffer = new Uint8Array(msg.data as number[]).buffer;

  try {
    if (fileType === 'xlsx') {
      await initXlsx(buffer);
    } else if (fileType === 'docx') {
      await initDocx(buffer);
    } else if (fileType === 'pptx') {
      await initPptx(buffer);
    }
  } catch (err) {
    setStatus(`Error: ${err instanceof Error ? err.message : String(err)}`);
  }
});

// ── XLSX ─────────────────────────────────────────────────────────────────────

async function initXlsx(buffer: ArrayBuffer): Promise<void> {
  hideStatus();
  const container = document.createElement('div');
  container.style.cssText = 'width:100%;height:calc(100vh - 40px);';
  viewerContainer.appendChild(container);

  const viewer = new XlsxViewer(container, {
    onReady(sheetNames) {
      setStatus(`Loaded — ${sheetNames.length} sheet(s). Click a cell to select.`);
    },
    onError(err) {
      setStatus(`Error: ${err.message}`);
    },
    onSelectionChange(sel: CellRange | null) {
      if (!sel) return;
      vscodeApi.postMessage({ type: 'selection', fileType: 'xlsx', selection: sel });
    },
  });

  await viewer.load(buffer);
  hideStatus();

  // Ctrl+C / Cmd+C: copy selection through VSCode clipboard API
  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
      const sel = viewer.selection;
      if (!sel) return;
      // Let the viewer handle copy internally (it writes to navigator.clipboard),
      // then also notify the extension host.
      vscodeApi.postMessage({ type: 'copy-request', fileType: 'xlsx', selection: sel });
    }
  });
}

// ── DOCX ─────────────────────────────────────────────────────────────────────

async function initDocx(buffer: ArrayBuffer): Promise<void> {
  hideStatus();
  const canvas = document.createElement('canvas');
  viewerContainer.appendChild(canvas);

  const viewer = new DocxViewer(canvas, {
    width: Math.min(window.innerWidth - 32, 900),
    enableTextSelection: true,
  });

  await viewer.load(buffer);
  hideStatus();

  navBar.classList.add('visible');
  updateDocxNav(viewer);

  prevBtn.addEventListener('click', () => { viewer.prevPage(); updateDocxNav(viewer); });
  nextBtn.addEventListener('click', () => { viewer.nextPage(); updateDocxNav(viewer); });
}

function updateDocxNav(viewer: DocxViewer): void {
  pageInfo.textContent = `Page ${viewer.currentPage + 1} / ${viewer.pageCount}`;
  prevBtn.disabled = viewer.currentPage === 0;
  nextBtn.disabled = viewer.currentPage === viewer.pageCount - 1;
}

// ── PPTX ─────────────────────────────────────────────────────────────────────

async function initPptx(buffer: ArrayBuffer): Promise<void> {
  hideStatus();
  const container = document.createElement('div');
  viewerContainer.appendChild(container);

  const viewer = new PptxViewer(container, {
    width: Math.min(window.innerWidth - 32, 960),
    enableTextSelection: true,
    onSlideChange(index, total) {
      pageInfo.textContent = `Slide ${index + 1} / ${total}`;
      prevBtn.disabled = index === 0;
      nextBtn.disabled = index === total - 1;
    },
    onError(err) {
      setStatus(`Error: ${err.message}`);
    },
  });

  await viewer.load(buffer);
  hideStatus();
  navBar.classList.add('visible');

  prevBtn.addEventListener('click', () => viewer.prevSlide());
  nextBtn.addEventListener('click', () => viewer.nextSlide());
}
