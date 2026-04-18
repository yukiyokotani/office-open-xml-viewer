import type { WorkerRequest, WorkerResponse } from './types';
import init, { parse_pptx } from './wasm/pptx_parser.js';
// Explicit URL import so Vite resolves it at build time — avoids invalid
// import.meta.url when the worker is inlined as a blob URL (e.g. Storybook static).
import wasmUrl from './wasm/pptx_parser_bg.wasm?url';

let ready = false;

async function initWasm() {
  await init(wasmUrl);
  ready = true;
  const msg: WorkerResponse = { kind: 'ready' };
  self.postMessage(msg);
}

self.onmessage = (e: MessageEvent<WorkerRequest>) => {
  const req = e.data;
  if (req.kind === 'parse') {
    if (!ready) {
      const msg: WorkerResponse = { kind: 'error', id: req.id, message: 'WASM not initialized' };
      self.postMessage(msg);
      return;
    }
    try {
      const jsonStr = parse_pptx(new Uint8Array(req.buffer));
      const presentation = JSON.parse(jsonStr);
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
  }
};

initWasm().catch((err) => {
  console.error('[pptx-worker] WASM init failed:', err);
});
