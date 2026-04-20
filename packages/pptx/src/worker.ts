import type { Presentation, WorkerRequest, WorkerResponse } from './types';
import init, { parse_pptx, extract_media } from './wasm/pptx_parser.js';

let ready = false;
let currentBuffer: Uint8Array | null = null;

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
      const bytes = new Uint8Array(req.buffer);
      currentBuffer = bytes;
      const jsonStr = parse_pptx(bytes);
      const presentation: Presentation = JSON.parse(jsonStr);
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

  if (req.kind === 'extractMedia') {
    if (!currentBuffer) {
      const msg: WorkerResponse = { kind: 'error', id: req.id, message: 'No pptx loaded' };
      self.postMessage(msg);
      return;
    }
    try {
      const bytes = extract_media(currentBuffer, req.path);
      const copy = new Uint8Array(bytes).slice().buffer;
      const msg: WorkerResponse = { kind: 'mediaExtracted', id: req.id, bytes: copy };
      (self.postMessage as (message: unknown, transfer: Transferable[]) => void)(msg, [copy]);
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
};
