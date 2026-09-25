import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { XlsxWorkbook } from './workbook.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/** A minimal XlsxWorkbook stand-in covering the surface load() → buildTabs()
 *  touches, plus a `destroy` spy so the latch can assert the loser's workbook
 *  (its worker + WASM) is torn down. */
function fakeWorkbook() {
  const destroy = vi.fn();
  const wb = {
    sheetNames: ['Sheet1'],
    tabColors: {} as Record<number, string>,
    destroy,
    getWorksheet: vi.fn().mockResolvedValue(undefined),
  };
  return { wb: wb as unknown as XlsxWorkbook, destroy };
}

/**
 * Concurrent-load latch (composes with SC20's success-after-swap): if a caller
 * fires `load(A)` and, before it resolves, `load(B)`, both loads race the WASM
 * parse / worker init concurrently. Whichever resolves LAST must NOT win the swap
 * when it is the stale one — the loser's freshly-loaded workbook (never installed,
 * or installed then overwritten) must be destroyed, not leaked, and the winner's
 * workbook must stay live and untouched. A generation token (`_loadGen`) closes it.
 */
describe('XlsxViewer.load() — concurrent-load latch', () => {
  function build() {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    // Isolate the latch from the sheet-render path: showSheet needs a full
    // worksheet model, out of scope here. The engine-swap happens in load() BEFORE
    // showSheet, so a resolved no-op keeps this test on the leak.
    vi.spyOn(
      v as unknown as { showSheet: (i: number) => Promise<void> },
      'showSheet',
    ).mockResolvedValue(undefined);
    return { v };
  }

  function deferredLoad(wb: XlsxWorkbook): { resolve: () => void; promise: Promise<XlsxWorkbook> } {
    let resolve!: () => void;
    const promise = new Promise<XlsxWorkbook>((r) => {
      resolve = () => r(wb);
    });
    return { resolve, promise };
  }

  it('the later-started load winning first leaves the stale load a no-op (its workbook destroyed)', async () => {
    const { v } = build();
    const a = fakeWorkbook();
    const b = fakeWorkbook();
    const da = deferredLoad(a.wb);
    const db = deferredLoad(b.wb);
    vi.spyOn(XlsxWorkbook, 'load')
      .mockImplementationOnce(() => da.promise)
      .mockImplementationOnce(() => db.promise);

    const pa = v.load('a.xlsx'); // gen 1
    const pb = v.load('b.xlsx'); // gen 2 — supersedes A

    db.resolve();
    await pb;
    expect(b.destroy).not.toHaveBeenCalled();
    expect(a.destroy).not.toHaveBeenCalled();

    da.resolve();
    await pa;
    expect(a.destroy).toHaveBeenCalledTimes(1); // loser's workbook cleaned up
    expect(b.destroy).not.toHaveBeenCalled(); // winner untouched — still current

    v.destroy();
    expect(b.destroy).toHaveBeenCalledTimes(1);
    expect(a.destroy).toHaveBeenCalledTimes(1); // still exactly once
  });

  it('resolving in start order (A then B) behaves like today — B wins normally', async () => {
    const { v } = build();
    const a = fakeWorkbook();
    const b = fakeWorkbook();
    const da = deferredLoad(a.wb);
    const db = deferredLoad(b.wb);
    vi.spyOn(XlsxWorkbook, 'load')
      .mockImplementationOnce(() => da.promise)
      .mockImplementationOnce(() => db.promise);

    const pa = v.load('a.xlsx');
    const pb = v.load('b.xlsx');

    da.resolve();
    await pa;
    expect(a.destroy).toHaveBeenCalledTimes(1); // superseded loser cleaned up

    db.resolve();
    await pb;
    expect(b.destroy).not.toHaveBeenCalled();

    v.destroy();
    expect(b.destroy).toHaveBeenCalledTimes(1);
  });
});
