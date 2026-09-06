import { expect, it, vi } from 'vitest';
import { attachXlsFontMeasurement, requestXlsFontMeasurement, XLS_FONT_REQUEST, XLS_FONT_RESULT } from './xls-font-worker.js';

class Port extends EventTarget {
  peer?: Port;
  readonly sent: unknown[] = [];
  postMessage(message: unknown) {
    this.sent.push(message);
    queueMicrotask(() => this.peer?.dispatchEvent(new MessageEvent('message', { data: structuredClone(message) })));
  }
}
function pair() {
  const main = new Port(); const worker = new Port();
  main.peer = worker; worker.peer = main;
  return { main, worker };
}
const font = { family: 'Arial', sizePoints: 11, bold: false, italic: false };

it('exchanges only metadata/width, not callbacks or conversion request identifiers', async () => {
  const { main, worker } = pair();
  const measure = vi.fn(() => 7);
  const detach = attachXlsFontMeasurement(main, measure);
  const request = requestXlsFontMeasurement(worker);
  expect(await request(font, new AbortController().signal)).toBe(7);
  expect(worker.sent).toEqual([{ type: XLS_FONT_REQUEST, font }]);
  expect(main.sent).toEqual([{ type: XLS_FONT_RESULT, width: 7 }]);
  expect(measure).toHaveBeenCalledOnce();
  await expect(request(font, new AbortController().signal)).rejects.toThrow();
  detach();
});

it('aborts host measurement and suppresses a late result after worker disposal', async () => {
  const { main, worker } = pair();
  let resolve: ((n: number) => void) | undefined;
  let signal: AbortSignal | undefined;
  const detach = attachXlsFontMeasurement(main, (_, s) => { signal = s; return new Promise<number>((r) => { resolve = r; }); });
  const controller = new AbortController();
  const pending = requestXlsFontMeasurement(worker)(font, controller.signal);
  const rejected = expect(pending).rejects.toThrow();
  await vi.waitFor(() => expect(signal).toBeDefined());
  detach(); controller.abort(); await rejected;
  expect(signal?.aborted).toBe(true);
  resolve?.(7); await Promise.resolve(); await Promise.resolve();
  expect(main.sent).toEqual([]);
});

it.each([0, 7.5, Infinity, 4097])('rejects malformed metric replies: %s', async (width) => {
  const { main, worker } = pair();
  const pending = requestXlsFontMeasurement(worker)(font, new AbortController().signal);
  main.postMessage({ type: XLS_FONT_RESULT, width });
  await expect(pending).rejects.toThrow();
});

it('refuses malformed fonts without exposing them to the callback and bounds duplicate requests', async () => {
  const { main, worker } = pair();
  const measure = vi.fn(() => 7);
  const detach = attachXlsFontMeasurement(main, measure);
  worker.postMessage({ type: XLS_FONT_REQUEST, font: { ...font, family: 'x'.repeat(256) } });
  worker.postMessage({ type: XLS_FONT_REQUEST, font });
  await vi.waitFor(() => expect(main.sent).toHaveLength(1));
  expect(main.sent).toEqual([{ type: XLS_FONT_RESULT, failed: true }]);
  expect(measure).not.toHaveBeenCalled(); detach();
});

it('passes unavailable fonts and reports callback errors without arbitrary error text', async () => {
  for (const fails of [false, true]) {
    const { main, worker } = pair();
    const detach = attachXlsFontMeasurement(main, () => { if (fails) throw new Error('private host diagnostic'); return undefined; });
    const pending = requestXlsFontMeasurement(worker)(font, new AbortController().signal);
    if (fails) await expect(pending).rejects.toThrow('XLS measurement failed');
    else expect(await pending).toBeUndefined();
    expect(JSON.stringify(main.sent)).not.toContain('private host diagnostic');
    detach();
  }
});
