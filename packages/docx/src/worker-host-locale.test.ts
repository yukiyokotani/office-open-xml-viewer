import { afterEach, beforeAll, expect, it, vi } from 'vitest';
import { DocxDocument } from './document.js';

/**
 * The main thread sends its host locale with every worker parse request
 * (document.ts `hostLocale: hostLocale()`), and the render worker adopts it
 * before any chart text or font demand is produced: a dedicated worker does
 * not see the page's emulated locale, so a request without the field leaves
 * its separators and short dates on the worker's own navigator.
 */

vi.mock('./wasm/docx_parser.js', () => ({
  default: async () => undefined, reinit: async () => undefined, DocxArchive: class { free(): void {} },
}));

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

// Main mode parses in a worker that paints nothing, so only worker mode (the
// render worker) needs the locale.
it('sends the main thread host locale with the worker-mode parse request', async () => {
  const posted: Record<string, unknown>[] = [];
  vi.stubGlobal('navigator', { language: 'fr-FR' });
  vi.stubGlobal('location', { href: 'http://localhost/' });
  vi.stubGlobal('OffscreenCanvas', class {});
  vi.stubGlobal('Worker', class {
    postMessage(message: Record<string, unknown>): void { posted.push(message); }
    addEventListener(): void {}
    removeEventListener(): void {}
    terminate(): void {}
  });
  const loading = DocxDocument.load(new ArrayBuffer(8), { mode: 'worker' }).catch(() => undefined);
  await vi.waitFor(() => expect(posted.some(message => message.type === 'parse')).toBe(true));
  expect(posted.find(message => message.type === 'parse')).toMatchObject({ hostLocale: 'fr-FR' });
  void loading;
});

// The render worker is imported once in beforeAll (its module graph is the
// expensive part, so it stays out of each test's timeout) with a fake `self`.
const worker: { onmessage: ((event: MessageEvent) => unknown) | null } = { onmessage: null };
// The worker's own module graph: adoption happens in it, so the painted-text
// and font-demand functions are imported from the same graph.
let core: typeof import('@silurus/ooxml-core');
let demandModule: typeof import('@silurus/ooxml-core/internal/font-scalar-demand');
beforeAll(async () => {
  vi.stubGlobal('self', { postMessage: () => undefined, set onmessage(handler: never) { worker.onmessage = handler; } });
  vi.resetModules();
  await import('./render-worker.js');
  core = await import('@silurus/ooxml-core');
  demandModule = await import('@silurus/ooxml-core/internal/font-scalar-demand');
}, 60_000);

it('render worker paints and demands the request locale\'s separators, and keeps its own without one', async () => {
  const paint = () => core.formatLocalizedExcelShortDate(45_659, false);
  const demand = () => (demandModule.chartPaintTexts({ chartType: 'bar', series: [] }) as string[]).join('');
  expect(paint()).toBe('1/2/2025');
  expect(demand()).not.toContain('\u202f');
  // A request without the field leaves the worker's own locale in place.
  await Promise.resolve(worker.onmessage?.({ data: { type: 'unsupported', id: 1 } } as MessageEvent)).catch(() => undefined);
  expect(paint()).toBe('1/2/2025');
  // The request's hostLocale is adopted before the request is handled.
  await Promise.resolve(worker.onmessage?.({ data: { type: 'unsupported', id: 2, hostLocale: 'fr-FR' } } as MessageEvent)).catch(() => undefined);
  expect(paint()).toBe('02/01/2025');
  expect(demand()).toContain('\u202f');
});
