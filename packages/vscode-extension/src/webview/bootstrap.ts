/**
 * Webview bootstrap script.
 *
 * Runs inside the VSCode Webview iframe. Receives the file bytes via the
 * `ooxml-init` message and instantiates the appropriate viewer:
 *   - docx / pptx: scroll-stacked render of every page / slide with a transparent
 *     text layer for native selection (PDF.js-style).
 *   - xlsx: XlsxViewer (sheet-based, no scroll stack)
 */

declare const __OOXML_FILE_TYPE__: 'docx' | 'xlsx' | 'pptx';

declare function acquireVsCodeApi(): {
  postMessage(msg: unknown): void;
  getState(): unknown;
  setState(state: unknown): void;
};

import { XlsxViewer, type CellRange } from '@silurus/ooxml-xlsx';
import { DocxDocument, type DocxTextRunInfo } from '@silurus/ooxml-docx';
import { PptxPresentation, type TextRunInfo } from '@silurus/ooxml-pptx';

const vscodeApi = acquireVsCodeApi();
const fileType = __OOXML_FILE_TYPE__;

const statusEl = document.getElementById('status')!;
const viewerContainer = document.getElementById('viewer-container')!;

function setStatus(msg: string): void {
  statusEl.style.display = '';
  statusEl.textContent = msg;
}

function hideStatus(): void {
  statusEl.style.display = 'none';
}

// Notify extension host that the webview script is ready to receive messages.
vscodeApi.postMessage({ type: 'webview-ready' });

window.addEventListener('message', async (event: MessageEvent) => {
  const msg = event.data;
  if (msg.type !== 'ooxml-init') return;

  const buffer = new Uint8Array(msg.data as number[]).buffer;

  try {
    if (fileType === 'docx') {
      await initDocx(buffer);
    } else if (fileType === 'xlsx') {
      await initXlsx(buffer);
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
  container.style.cssText = 'width:100%;height:100vh;';
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

  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
      const sel = viewer.selection;
      if (!sel) return;
      vscodeApi.postMessage({ type: 'copy-request', fileType: 'xlsx', selection: sel });
    }
  });
}

// ── DOCX (scroll view) ───────────────────────────────────────────────────────

function buildDocxTextLayer(layer: HTMLDivElement, runs: DocxTextRunInfo[]): void {
  layer.replaceChildren();
  for (const run of runs) {
    const span = document.createElement('span');
    span.textContent = run.text;
    span.style.cssText =
      `position:absolute;left:${run.x}px;top:${run.y}px;` +
      `font-size:${run.fontSize}px;line-height:${run.h}px;white-space:pre;color:transparent;cursor:text;pointer-events:all;`;
    layer.appendChild(span);
  }
}

async function initDocx(buffer: ArrayBuffer): Promise<void> {
  setStatus('Parsing…');
  const doc = await DocxDocument.load(buffer);
  setStatus(`Rendering ${doc.pageCount} page(s)…`);

  const stack = document.createElement('div');
  stack.className = 'page-stack';
  viewerContainer.appendChild(stack);

  const widthPx = Math.min(window.innerWidth - 64, 900);

  for (let i = 0; i < doc.pageCount; i++) {
    const wrapper = document.createElement('div');
    wrapper.className = 'page-wrapper';
    wrapper.style.maxWidth = `${widthPx}px`;

    const canvas = document.createElement('canvas');
    canvas.className = 'page-canvas';

    const textLayer = document.createElement('div');
    textLayer.className = 'text-layer';

    wrapper.append(canvas, textLayer);
    stack.appendChild(wrapper);

    const runs: DocxTextRunInfo[] = [];
    await doc.renderPage(canvas, i, { width: widthPx, onTextRun: (r) => runs.push(r) });
    buildDocxTextLayer(textLayer, runs);
  }

  hideStatus();
}

// ── PPTX (scroll view) ───────────────────────────────────────────────────────

function buildPptxTextLayer(
  layer: HTMLDivElement,
  runs: TextRunInfo[],
  cssWidth: number,
  cssHeight: number,
): void {
  layer.replaceChildren();
  layer.style.width = `${cssWidth}px`;
  layer.style.height = `${cssHeight}px`;

  const shapeMap = new Map<string, HTMLDivElement>();
  for (const run of runs) {
    const totalRot = run.rotation + (run.textBodyRotation ?? 0);
    const key = `${run.shapeX},${run.shapeY},${run.shapeW},${run.shapeH},${totalRot}`;
    let shape = shapeMap.get(key);
    if (!shape) {
      shape = document.createElement('div');
      shape.style.cssText =
        `position:absolute;left:${run.shapeX}px;top:${run.shapeY}px;` +
        `width:${run.shapeW}px;height:${run.shapeH}px;pointer-events:all;overflow:hidden;`;
      if (totalRot !== 0) {
        shape.style.transformOrigin = 'center center';
        shape.style.transform = `rotate(${totalRot}deg)`;
      }
      shapeMap.set(key, shape);
      layer.appendChild(shape);
    }
    const span = document.createElement('span');
    span.textContent = run.text;
    span.style.cssText =
      `position:absolute;left:${run.inShapeX}px;top:${run.inShapeY}px;` +
      `font-size:${run.fontSize}px;line-height:${run.h}px;white-space:pre;color:transparent;cursor:text;`;
    shape.appendChild(span);
  }
}

async function initPptx(buffer: ArrayBuffer): Promise<void> {
  setStatus('Parsing…');
  const pres = await PptxPresentation.load(buffer);
  setStatus(`Rendering ${pres.slideCount} slide(s)…`);

  const stack = document.createElement('div');
  stack.className = 'page-stack';
  viewerContainer.appendChild(stack);

  const widthPx = Math.min(window.innerWidth - 64, 960);
  const cssHeight = pres.slideWidth > 0
    ? Math.round((pres.slideHeight * widthPx) / pres.slideWidth)
    : Math.round((widthPx * 9) / 16);

  for (let i = 0; i < pres.slideCount; i++) {
    const wrapper = document.createElement('div');
    wrapper.className = 'page-wrapper';
    wrapper.style.maxWidth = `${widthPx}px`;

    const canvas = document.createElement('canvas');
    canvas.className = 'page-canvas';

    const textLayer = document.createElement('div');
    textLayer.className = 'text-layer';

    wrapper.append(canvas, textLayer);
    stack.appendChild(wrapper);

    const runs: TextRunInfo[] = [];
    await pres.renderSlide(canvas, i, { width: widthPx, onTextRun: (r) => runs.push(r) });
    buildPptxTextLayer(textLayer, runs, widthPx, cssHeight);
  }

  hideStatus();
}
