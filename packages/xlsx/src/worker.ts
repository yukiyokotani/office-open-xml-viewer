import init, { parse_xlsx, parse_sheet } from './wasm/xlsx_parser.js';
import type { WorkerRequest, WorkerResponse } from './types.js';

let initialized = false;

async function ensureInit(): Promise<void> {
  if (!initialized) {
    await init();
    initialized = true;
  }
}

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  try {
    await ensureInit();
    const req = e.data;

    if (req.type === 'parse') {
      const json = parse_xlsx(new Uint8Array(req.data));
      const workbook = JSON.parse(json);
      const res: WorkerResponse = { type: 'parsed', workbook };
      self.postMessage(res);
    } else if (req.type === 'parseSheet') {
      const json = parse_sheet(new Uint8Array(req.data), req.sheetIndex, req.sheetName);
      const worksheet = JSON.parse(json);
      const res: WorkerResponse = { type: 'parsedSheet', worksheet };
      self.postMessage(res);
    }
  } catch (err) {
    const res: WorkerResponse = { type: 'error', message: String(err) };
    self.postMessage(res);
  }
};
