import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import { DocxDocument } from './document.js';
import { publishDocxLayout } from './document-layout-events.js';
import { installDom, makeEl, FakeDocxEngine } from './scroll-viewer-test-dom.js';
import type { HyperlinkTarget } from '@silurus/ooxml-core';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const PAGE = [{ widthPt: 595, heightPt: 842 }];

/**
 * IX-nav wiring — DocxViewer's internal-anchor click default.
 *
 * The single-canvas viewer builds its click callback in `_hyperlinkHandler()`,
 * exercised here via the fake-engine seam the render-error tests use (spy
 * `DocxDocument.load` to install a FakeDocxEngine). These pin the branch IX-nav
 * adds: an internal anchor resolves a bookmark → page via the document and jumps
 * there; unknown bookmark ⇒ no-op; a custom handler owns the click; external
 * unchanged.
 */
async function mountLoaded(opts: Record<string, unknown> = {}) {
  installDom();
  const canvas = makeEl('canvas');
  const engine = new FakeDocxEngine(6, PAGE);
  vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
  const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, opts);
  await v.load('x.docx');
  // The overlay's click callback (what a span click invokes).
  const handler = (
    v as unknown as { _hyperlinkHandler: () => (t: HyperlinkTarget) => void }
  )._hyperlinkHandler();
  return { v, engine, click: handler };
}

describe('DocxViewer — internal-anchor click default (IX-nav)', () => {
  it('resolves a bookmark name via the document and navigates to its page', async () => {
    const { v, engine, click } = await mountLoaded();
    engine.bookmarkPages.set('Heading2', 3);
    const goTo = vi.spyOn(v, 'goToPage');

    click({ kind: 'internal', ref: 'Heading2' });

    expect(engine.bookmarkCalls).toContain('Heading2');
    expect(goTo).toHaveBeenCalledWith(3);
    v.destroy();
  });

  it('is a safe no-op when the anchor names no known bookmark', async () => {
    const { v, engine, click } = await mountLoaded();
    const goTo = vi.spyOn(v, 'goToPage');
    click({ kind: 'internal', ref: 'DoesNotExist' });
    expect(engine.bookmarkCalls).toContain('DoesNotExist'); // tried
    expect(goTo).not.toHaveBeenCalled(); // but did not navigate
    v.destroy();
  });

  it('waits for the authoritative bookmark projection during progressive layout', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, PAGE);
    engine.setLayoutComplete(false);
    const doc = engine.asDoc();
    const v = DocxViewer.fromDocument(
      makeEl('canvas') as unknown as HTMLCanvasElement,
      doc,
    ) as DocxViewer;
    const click = (
      v as unknown as { _hyperlinkHandler: () => (t: HyperlinkTarget) => void }
    )._hyperlinkHandler();
    const goTo = vi.spyOn(v, 'goToPage');

    click({ kind: 'internal', ref: 'Later' });
    expect(engine.bookmarkCalls).not.toContain('Later');

    engine.bookmarkPages.set('Later', 2);
    engine.setPageCount(3);
    engine.setLayoutComplete(true);
    publishDocxLayout(doc, { pageCount: 3, exact: true, complete: true });
    await vi.waitFor(() => expect(goTo).toHaveBeenCalledWith(2));
    v.destroy();
  });

  it('a caller-supplied onHyperlinkClick fully owns the click (no bookmark resolution, no nav)', async () => {
    const onHyperlinkClick = vi.fn<(t: HyperlinkTarget) => void>();
    const { v, engine, click } = await mountLoaded({ onHyperlinkClick });
    engine.bookmarkPages.set('Heading2', 3);
    const goTo = vi.spyOn(v, 'goToPage');

    click({ kind: 'internal', ref: 'Heading2' });

    expect(onHyperlinkClick).toHaveBeenCalledWith({ kind: 'internal', ref: 'Heading2' });
    expect(engine.bookmarkCalls).not.toContain('Heading2'); // handler owns it
    expect(goTo).not.toHaveBeenCalled();
    v.destroy();
  });

  it('opens an external link in a new tab', async () => {
    const { v, click } = await mountLoaded();
    const winOpen = vi.fn();
    (globalThis as unknown as { window: { open: unknown } }).window.open = winOpen;
    click({ kind: 'external', url: 'https://example.com/' });
    expect(winOpen).toHaveBeenCalledWith('https://example.com/', '_blank', 'noopener,noreferrer');
    v.destroy();
  });
});
