import type { Presentation, WorkerRequest, WorkerResponse } from './types';
import init, { parse_pptx } from './wasm/pptx_parser.js';
import { renderSlide } from './renderer';

let ready = false;
let offscreenCanvas: OffscreenCanvas | null = null;
let canvasDpr = 1;
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

  if (req.kind === 'transferCanvas') {
    offscreenCanvas = req.canvas;
    canvasDpr = req.dpr;
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
    if (!offscreenCanvas || !storedPresentation) {
      const msg: WorkerResponse = { kind: 'error', id: req.id, message: 'Canvas or presentation not ready' };
      self.postMessage(msg);
      return;
    }
    const slide = storedPresentation.slides[req.slideIndex];
    if (!slide) {
      const msg: WorkerResponse = { kind: 'error', id: req.id, message: `Slide ${req.slideIndex} not found` };
      self.postMessage(msg);
      return;
    }
    renderSlide(offscreenCanvas, slide, storedPresentation.slideWidth, storedPresentation.slideHeight, {
      width: req.targetWidth,
      defaultTextColor: req.defaultTextColor,
      dpr: canvasDpr,
    }).then(() => {
      const msg: WorkerResponse = { kind: 'rendered', id: req.id };
      self.postMessage(msg);
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
