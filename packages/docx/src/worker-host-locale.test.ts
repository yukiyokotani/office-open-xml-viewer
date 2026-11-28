import { afterEach, expect, it, vi } from 'vitest';

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
  const { DocxDocument } = await import('./document.js');
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

it('render worker adopts the request locale before handling it, and keeps its own without one', async () => {
  const posted: unknown[] = [];
  const fake = { onmessage: null as ((event: MessageEvent) => unknown) | null, postMessage: (message: unknown) => { posted.push(message); } };
  vi.stubGlobal('self', fake);
  vi.stubGlobal('navigator', { language: 'en-US' });
  vi.resetModules();
  await import('./render-worker.js');
  const core = await import('@silurus/ooxml-core');
  expect(core.hostLocale()).toBe('en-US');
  await Promise.resolve(fake.onmessage?.({ data: { type: 'unknown-request', id: 1 } } as MessageEvent)).catch(() => undefined);
  expect(core.hostLocale()).toBe('en-US');
  await Promise.resolve(fake.onmessage?.({ data: { type: 'unknown-request', id: 2, hostLocale: 'fr-FR' } } as MessageEvent)).catch(() => undefined);
  expect(core.hostLocale()).toBe('fr-FR');
});
