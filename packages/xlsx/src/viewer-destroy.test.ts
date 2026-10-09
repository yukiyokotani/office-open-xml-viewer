import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { XlsxWorkbook } from './workbook.js';
import { installDom, makeContainer, type FakeDocument, type FakeEl } from './viewer-destroy-test-dom.js';
import type { Worksheet } from './types.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * XlsxViewer builds its whole UI subtree (a wrapper div holding the canvas
 * area + sheet-tab bar) inside the caller's container, and injects a `<style>`
 * into `document.head`. destroy() must (1) remove that subtree so the container
 * returns to the empty state it had before construction, and (2) NOT leak a new
 * <style> per instance. These tests pin both, plus removal of the document-level
 * focus-scoped viewport keydown listener.
 */
describe('XlsxViewer.destroy() — subtree + listeners + style', () => {
  it('leaves outer framing to the caller-owned container', () => {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    const wrapper = container.children[0] as FakeEl;

    expect(wrapper.style.border).toBe('');
    v.destroy();
  });

  it('empties the container (removes the wrapper subtree)', () => {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    // Construction mounted exactly one wrapper subtree.
    expect(container.childNodes.length).toBe(1);
    v.destroy();
    expect(container.childNodes.length).toBe(0);
  });

  it('injects the viewer <style> once across 3 mount/unmount cycles (module-level, tagged)', () => {
    const doc = installDom() as FakeDocument;
    const container = makeContainer();
    for (let i = 0; i < 3; i++) {
      const v = new XlsxViewer(container as unknown as HTMLElement);
      v.destroy();
    }
    // Exactly one tagged stylesheet survives in <head> — not three (and not zero:
    // it is a class-constant sheet, kept after destroy for any live instances).
    const styles = doc.head.children.filter(
      (c: FakeEl) => c.tag === 'style' && c.hasAttribute('data-xlsx-viewer-styles'),
    );
    expect(styles.length).toBe(1);
    // And it is still present after the last destroy (destroy must NOT remove it).
    expect(doc.head.querySelector('style[data-xlsx-viewer-styles]')).not.toBeNull();
  });

  it('keeps a single tagged stylesheet even while several viewers are alive at once', () => {
    const doc = installDom() as FakeDocument;
    const a = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const b = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const c = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const count = () =>
      doc.head.children.filter(
        (e: FakeEl) => e.tag === 'style' && e.hasAttribute('data-xlsx-viewer-styles'),
      ).length;
    expect(count()).toBe(1);
    a.destroy();
    // b and c are still alive — the shared sheet must remain.
    expect(count()).toBe(1);
    expect(doc.head.querySelector('style[data-xlsx-viewer-styles]')).not.toBeNull();
    b.destroy();
    c.destroy();
  });

  it('never installs a document-level copy listener', () => {
    const doc = installDom() as FakeDocument;
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    // Copy is scoped to the focusable viewport, so unrelated document input and
    // other Viewer instances cannot race to overwrite the clipboard.
    expect(doc.listenerCount('keydown')).toBe(0);
    v.destroy();
    expect(doc.listenerCount('keydown')).toBe(0);
    // Dispatching a keydown after destroy must not throw (no live handler).
    expect(() => doc.dispatchEvent('keydown', { key: 'c', ctrlKey: true })).not.toThrow();
  });

  it('detaches every element, document, media and observer listener it installed', async () => {
    const doc = installDom();
    // Record every element the viewer creates, so listeners on nodes that are
    // no longer attached (tab buttons, detached outline gutters, value-list
    // items) are checked too, not only the live subtree.
    const created: FakeEl[] = [];
    const createElement = doc.createElement;
    doc.createElement = (tag: string) => {
      const element = createElement(tag);
      created.push(element);
      return element;
    };
    const observers: Array<{ observing: boolean }> = [];
    class RecordingObserver {
      observing = false;
      constructor(_callback: () => void) { observers.push(this); }
      observe(): void { this.observing = true; }
      disconnect(): void { this.observing = false; }
      unobserve(): void {}
    }
    const mediaListeners = new Set<unknown>();
    Object.assign(doc.defaultView, {
      ResizeObserver: RecordingObserver,
      MutationObserver: RecordingObserver,
      matchMedia: () => ({
        addEventListener: (_type: string, listener: unknown) => mediaListeners.add(listener),
        removeEventListener: (_type: string, listener: unknown) => mediaListeners.delete(listener),
      }),
    });
    const workbook = {
      mode: 'main',
      sheetCount: 2,
      sheetNames: ['Sheet1', 'Sheet2'],
      tabColors: [null, null],
      isHidden: () => false,
      getWorksheet: () => new Promise(() => {}),
      resolveValidationList: async () => ({ kind: 'values', values: ['A', 'B'] }),
      destroy: vi.fn(),
    } as unknown as XlsxWorkbook;
    const viewer = XlsxViewer.fromWorkbook(
      makeContainer() as unknown as HTMLElement,
      workbook,
      { onContextMenu: () => undefined },
    );
    const internals = viewer as unknown as {
      currentSheet: number;
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      selectionController: { select(cell: { row: number; col: number }): void };
      validation: { toggle(): void };
    };
    internals.currentSheet = 0;
    internals.canvasArea.clientWidth = 800;
    internals.canvasArea.clientHeight = 600;
    internals.currentWorksheet = {
      name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {},
      defaultColWidth: 8.43, defaultRowHeight: 15, freezeRows: 0, freezeCols: 0,
      mergeCells: [], conditionalFormats: [], images: [], charts: [],
      dataValidations: [{ validationType: 'list', sqref: 'A1', formula1: 'A1:A2' }],
    } as unknown as Worksheet;
    internals.selectionController.select({ row: 1, col: 1 });
    internals.validation.toggle();
    await vi.waitFor(() => expect(doc.listenerCount('pointerdown')).toBe(1));

    const listenerCount = () => created.reduce((total, element) =>
      total + [...element._listeners.values()].reduce((sum, list) => sum + list.length, 0), 0);
    const listenedTypes = (predicate: (element: FakeEl) => boolean) => new Set(
      created.filter(predicate).flatMap((element) =>
        [...element._listeners].filter(([, list]) => list.length > 0).map(([type]) => type)),
    );
    // Precondition: the viewport, gutters, tab strip, zoom control, overlay and
    // value-list items really did register listeners.
    expect(listenedTypes((element) => element.hasAttribute('data-xlsx-viewport-input')))
      .toEqual(new Set([
        'scroll', 'contextmenu', 'pointerdown', 'pointermove', 'pointerup',
        'pointercancel', 'lostpointercapture', 'wheel', 'pointerleave', 'focus', 'keydown',
      ]));
    expect(listenedTypes((element) => element.hasAttribute('data-xlsx-outline')))
      .toEqual(new Set(['pointerdown']));
    expect(listenedTypes((element) => element.tag === 'button').has('click')).toBe(true);
    expect(listenedTypes((element) => element.tag === 'input')).toEqual(new Set(['input']));
    expect(listenedTypes((element) => element.hasAttribute('data-xlsx-validation-item')))
      .toEqual(new Set(['pointerenter', 'pointerleave']));
    expect(listenedTypes((element) => element.hasAttribute('data-xlsx-validation-panel')))
      .toEqual(new Set(['wheel']));
    expect(mediaListeners.size).toBe(1);
    expect(observers.some((observer) => observer.observing)).toBe(true);
    expect(listenerCount()).toBeGreaterThan(0);

    viewer.destroy();

    expect(listenerCount()).toBe(0);
    expect(doc.listenerCount('pointerdown')).toBe(0);
    expect(doc.listenerCount('keydown')).toBe(0);
    expect(mediaListeners.size).toBe(0);
    expect(observers.filter((observer) => observer.observing)).toEqual([]);
    expect(workbook.destroy).not.toHaveBeenCalled();
  });

  it('is safe to call destroy() twice', () => {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    v.destroy();
    expect(() => v.destroy()).not.toThrow();
    expect(container.childNodes.length).toBe(0);
  });

  it('permanently rejects a new load after destroy without acquiring a workbook', async () => {
    installDom();
    const viewer = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const load = vi.spyOn(XlsxWorkbook, 'load');
    viewer.destroy();

    const closed = 'XlsxViewer is destroyed';
    await expect(viewer.load(new ArrayBuffer(0))).rejects.toThrow(closed);
    expect(load).not.toHaveBeenCalled();
  });

  it('borrows a workbook through fromWorkbook() and leaves its lifecycle with the caller', async () => {
    installDom();
    const destroy = vi.fn();
    const workbook = {
      mode: 'main',
      sheetCount: 1,
      sheetNames: ['Sheet1'],
      tabColors: {} as Record<number, string>,
      isHidden: () => false,
      getWorksheet: () => new Promise(() => {}),
      destroy,
    } as unknown as XlsxWorkbook;
    const viewer = XlsxViewer.fromWorkbook(
      makeContainer() as unknown as HTMLElement,
      workbook,
    );

    expect(viewer.sheetCount).toBe(1);
    await expect((viewer as XlsxViewer).load(new ArrayBuffer(0))).rejects.toThrow(/fromWorkbook/);
    viewer.destroy();
    expect(destroy).not.toHaveBeenCalled();
  });
});
