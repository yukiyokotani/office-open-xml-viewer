import init, { parse_docx } from './wasm/docx_parser.js';
import type { WorkerRequest, WorkerResponse } from './types';

let initPromise: Promise<void> | null = null;

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const req = e.data;

  if (req.type === 'init') {
    initPromise = init(req.wasmUrl);
    return;
  }

  try {
    await initPromise;
    if (req.type === 'parse') {
      const json = parse_docx(new Uint8Array(req.data));
      const document = JSON.parse(json);
      if (document.error) throw new Error(`Parse error: ${document.error}`);
      const res: WorkerResponse = { type: 'parsed', document };
      self.postMessage(res);
    }
  } catch (err) {
    const res: WorkerResponse = { type: 'error', message: String(err) };
    self.postMessage(res);
  }
};
