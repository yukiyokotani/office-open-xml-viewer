import { afterEach, expect, it, vi } from 'vitest';
import { DocxDocument } from './document.js';

/**
 * The main thread sends its host locale with every worker parse request
 * (document.ts `hostLocale: hostLocale()`); the render worker's adoption is
 * covered in render-worker-canvas-transfer.test.ts through a real parse and
 * render request: a dedicated worker does
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
