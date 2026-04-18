import type { Presentation, WorkerRequest, WorkerResponse } from './types';
import init, { parse_pptx } from './wasm/pptx_parser.js';
import { renderSlide } from './renderer';

let ready = false;
let internalCanvas: OffscreenCanvas | null = null;
let storedPresentation: Presentation | null = null;

async function initWasm(wasmUrl: string) {
  await init(wasmUrl);
  ready = true;
  const msg: WorkerResponse = { kind: 'ready' };
  self.postMessage(msg);
}

self.onmessage = (e: MessageEvent<WorkerRequest>) => {
  const req = e.data;

  if (req.kind === 'init') {
    initWasm(req.wasmUrl).catch((err) => {
      console.error('[pptx-worker] WASM init failed:', err);
    });
    return;
  }

  if (req.kind === 'parse') {
    if (!ready) {
      const msg: WorkerResponse = { kind: 'error', id: req.id, message: 'WASM not initialized' };
      self.postMessage(msg);
      return;
    }
    try {
      const jsonStr = parse_pptx(new Uint8Array(req.buffer));
      const presentation: Presentation = JSON.parse(jsonStr);
      storedPresentation = presentation;
      const msg: WorkerResponse = { kind: 'parsed', id: req.id, presentation };
      self.postMessage(msg);
    } catch (err) {
      const msg: WorkerResponse = {
        kind: 'error',
        id: req.id,
        message: err instanceof Error ? err.message : String(err),
      };
      self.postMessage(msg);
    }
    return;
  }

  if (req.kind === 'render') {
    if (!storedPresentation) {
      const msg: WorkerResponse = { kind: 'error', id: req.id, message: 'Presentation not loaded' };
      self.postMessage(msg);
      return;
    }
    const slide = storedPresentation.slides[req.slideIndex];
    if (!slide) {
      const msg: WorkerResponse = { kind: 'error', id: req.id, message: `Slide ${req.slideIndex} not found` };
      self.postMessage(msg);
      return;
    }
    if (!internalCanvas) {
      internalCanvas = new OffscreenCanvas(1, 1);
    }
    const canvas = internalCanvas;
    renderSlide(canvas, slide, storedPresentation.slideWidth, storedPresentation.slideHeight, {
      width: req.targetWidth,
      defaultTextColor: req.defaultTextColor,
      majorFont: req.majorFont,
      minorFont: req.minorFont,
      dpr: req.dpr,
    }).then(() => {
      const bitmap = canvas.transferToImageBitmap();
      const msg: WorkerResponse = { kind: 'bitmap', id: req.id, bitmap };
      (self as unknown as Worker).postMessage(msg, [bitmap]);
    }).catch((err) => {
      const msg: WorkerResponse = {
        kind: 'error',
        id: req.id,
        message: err instanceof Error ? err.message : String(err),
      };
      self.postMessage(msg);
    });
    return;
  }
};
