import init, { parse_docx } from './wasm/docx_parser.js';
import type { WorkerRequest, WorkerResponse } from './types';

let initPromise: Promise<unknown> | null = null;

function decodeDataUrl(url: string): ArrayBuffer | null {
  if (!url.startsWith('data:')) return null;
  const comma = url.indexOf(',');
  if (comma === -1) return null;
  const binary = atob(url.slice(comma + 1));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const req = e.data;

  if (req.type === 'init') {
    initPromise = init(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
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
