import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxScrollViewer } from './scroll-viewer.js';
import { publishDocxLayout } from './document-layout-events.js';
import { installDom, makeContainer, makeBorrowedDocxScrollViewer, FakeDocxEngine, type FakeEl } from './scroll-viewer-test-dom.js';
import type { DocxTextRunInfo } from './renderer';
import type { HyperlinkTarget } from '@silurus/ooxml-core';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX-nav wiring — DocxScrollViewer's internal-hyperlink click default.
 *
 * IX1 built the clickable overlay and resolved the external branch, but the
 * INTERNAL `<w:anchor>` branch was a documented no-op (no bookmark → page map).
 * IX-nav supplies that map on `DocxDocument` (`getBookmarkPage`); these tests pin
 * that the scroll viewer's default click behaviour now (a) resolves a bookmark
 * name via the document and scrolls to its destination page, (b) no-ops safely
 * for an unknown bookmark, (c) fully delegates to a caller-supplied
 * `onHyperlinkClick`, and (d) still opens external links.
 *
 * The click is driven END TO END through the real overlay: a hyperlink text run
 * is fed to `onTextRun`, `buildDocxTextLayer` makes a clickable span, and the test
 * fires its `click` listener — exactly the path a user click takes.
 */

function linkRun(hyperlink: HyperlinkTarget): DocxTextRunInfo {
  return { text: 'go', x: 1, y: 2, w: 10, h: 12, fontSize: 12, font: '12px serif', hyperlink };
}

/** Mount a scroll viewer with one hyperlink run fed into the top page's overlay,
 *  and return the pieces + a helper to click the link span. */
async function setup(
  pageCount: number,
  hyperlink: HyperlinkTarget,
  opts: Record<string, unknown> = {},
) {
  installDom();
  const container = makeContainer(200, 400);
  const engine = new FakeDocxEngine(
    pageCount,
    Array.from({ length: pageCount }, () => ({ widthPt: 100, heightPt: 200 })),
  );
  engine.feedTextRuns = [linkRun(hyperlink)];
  const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
    document: engine.asDoc(),
    gap: 10,
    overscan: 1,
    paddingLeft: 0,
    paddingRight: 0,
    enableTextSelection: true,
    ...opts,
  });
  const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
  scrollHost.clientHeight = 400;
  scrollHost.clientWidth = 200;
  v.relayout();
  await Promise.resolve();
  await Promise.resolve();

  function clickLink(): void {
    // The docx overlay places spans directly under the slot's text-layer div.
    for (const slot of scrollHost.children) {
      const layer = slot.children.find((k) => k.tag === 'div') as FakeEl | undefined;
      const span = layer?.children[0] as FakeEl | undefined;
      if (span && '_listeners' in span) {
        (span as unknown as { dispatch: (t: string, e?: unknown) => void }).dispatch('click', {
          preventDefault() {},
        });
        return;
      }
    }
    throw new Error('no clickable link span found in any mounted slot');
  }

  return { container, engine, v, scrollHost, clickLink };
}

describe('DocxScrollViewer — internal-anchor click default (IX-nav)', () => {
  it('resolves a bookmark name via the document and scrolls to its destination page', async () => {
    const target: HyperlinkTarget = { kind: 'internal', ref: 'Heading2' };
    const { engine, v, scrollHost, clickLink } = await setup(6, target);
    engine.bookmarkPages.set('Heading2', 3); // bookmark lives on page index 3

    clickLink();

    expect(engine.bookmarkCalls).toContain('Heading2');
    // The click must land exactly where scrollToPage(3) lands.
    const afterClick = scrollHost.scrollTop;
    v.scrollToPage(3);
    expect(afterClick).toBeCloseTo(scrollHost.scrollTop, 3);
    expect(afterClick).toBeGreaterThan(0); // it actually moved to a lower page
    v.destroy();
  });

  it('is a safe no-op when the anchor names no known bookmark', async () => {
    const target: HyperlinkTarget = { kind: 'internal', ref: 'DoesNotExist' };
    const { engine, v, scrollHost, clickLink } = await setup(6, target);
    const before = scrollHost.scrollTop;
    clickLink();
    expect(engine.bookmarkCalls).toContain('DoesNotExist'); // it TRIED to resolve
    expect(scrollHost.scrollTop).toBe(before); // ...but did not move
    v.destroy();
  });

  it('waits for the authoritative bookmark projection during progressive layout', async () => {
    const target: HyperlinkTarget = { kind: 'internal', ref: 'Later' };
    const { engine, v, scrollHost, clickLink } = await setup(1, target);
    const doc = engine.asDoc();
    engine.setLayoutComplete(false);
    const before = scrollHost.scrollTop;

    clickLink();
    expect(engine.bookmarkCalls).not.toContain('Later');

    engine.bookmarkPages.set('Later', 2);
    engine.setPageCount(3);
    engine.setLayoutComplete(true);
    publishDocxLayout(doc, { pageCount: 3, exact: true, complete: true });
    await vi.waitFor(() => expect(engine.bookmarkCalls).toContain('Later'));
    expect(scrollHost.scrollTop).toBeGreaterThan(before);
    v.destroy();
  });

  it('delegates fully to a caller-supplied onHyperlinkClick and does NOT scroll itself', async () => {
    const target: HyperlinkTarget = { kind: 'internal', ref: 'Heading2' };
    const onHyperlinkClick = vi.fn<(t: HyperlinkTarget) => void>();
    const { engine, v, scrollHost, clickLink } = await setup(6, target, { onHyperlinkClick });
    engine.bookmarkPages.set('Heading2', 3);
    const before = scrollHost.scrollTop;

    clickLink();

    expect(onHyperlinkClick).toHaveBeenCalledTimes(1);
    expect(onHyperlinkClick).toHaveBeenCalledWith(target);
    // A custom handler OWNS the click — the viewer neither resolves nor scrolls.
    expect(engine.bookmarkCalls).not.toContain('Heading2');
    expect(scrollHost.scrollTop).toBe(before);
    v.destroy();
  });

  it('still opens an external link in a new tab (external branch unchanged)', async () => {
    const target: HyperlinkTarget = { kind: 'external', url: 'https://example.com/' };
    const { v, clickLink } = await setup(3, target);
    const winOpen = vi.fn();
    (globalThis as unknown as { window: { open: unknown } }).window.open = winOpen;
    clickLink();
    expect(winOpen).toHaveBeenCalledWith('https://example.com/', '_blank', 'noopener,noreferrer');
    v.destroy();
  });
});
