import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const mocks = vi.hoisted(() => ({
  docx: [] as Array<Record<string, any>>,
  pptx: [] as Array<Record<string, any>>,
  xlsx: [] as Array<Record<string, any>>,
  xlsxSheet: [] as Array<Record<string, any>>,
  deferDocx: false,
  deferDocxLayout: false,
  deferPptxLayout: false,
  rejectXlsx: false,
  math: { loadMathJax: vi.fn(), mathMLToSvg: vi.fn() },
  threeD: { render: vi.fn() },
  regionMap: { render: vi.fn() },
  chartEx: { render: vi.fn() },
}));

vi.mock('@silurus/ooxml-docx', () => {
  class DocxScrollViewer {
    pageCount = 4;
    destroyed = false;
    readonly relayout = vi.fn();
    layoutComplete = true;
    resolveLoad: (() => void) | null = null;
    rejectLoad: ((error: Error) => void) | null = null;
    resolveLayout: (() => void) | null = null;
    readonly events: string[] = [];
    readonly setScaleCalls: number[] = [];
    constructor(
      public readonly host: HTMLElement,
      public readonly opts: Record<string, any>,
    ) {
      mocks.docx.push(this as unknown as Record<string, any>);
    }
    load(): Promise<void> {
      this.events.push('load');
      if (this.opts.progressiveLayout && mocks.deferDocxLayout) this.layoutComplete = false;
      if (!mocks.deferDocx) return Promise.resolve();
      return new Promise<void>((resolve, reject) => {
        this.resolveLoad = resolve;
        this.rejectLoad = reject;
      });
    }
    waitUntilLayoutComplete(): Promise<void> {
      if (this.layoutComplete) return Promise.resolve();
      return new Promise<void>((resolve) => {
        this.resolveLayout = () => {
          this.pageCount = 7;
          this.layoutComplete = true;
          resolve();
        };
      });
    }
    setScale(scale: number): void {
      this.events.push('setScale');
      this.setScaleCalls.push(scale);
    }
    getScale(): number { return this.setScaleCalls.at(-1) ?? 1; }
    destroy(): void { this.destroyed = true; }
  }
  return {
    DocxScrollViewer,
    DocxDocument: class {
      static load(): Promise<{ destroy(): void }> {
        return Promise.resolve({ destroy() {} });
      }
    },
  };
});

vi.mock('@silurus/ooxml-pptx', () => {
  class PptxScrollViewer {
    slideCount = 6;
    availableSlideCount = 6;
    destroyed = false;
    readonly relayout = vi.fn();
    readonly events: string[] = [];
    readonly setScaleCalls: number[] = [];
    layoutComplete = true;
    resolveLayout: (() => void) | null = null;
    constructor(
      public readonly host: HTMLElement,
      public readonly opts: Record<string, any>,
    ) {
      mocks.pptx.push(this as unknown as Record<string, any>);
    }
    load(): Promise<void> {
      this.events.push('load');
      if (this.opts.progressiveLayout && mocks.deferPptxLayout) {
        this.layoutComplete = false;
        this.availableSlideCount = 2;
      }
      return Promise.resolve();
    }
    waitUntilLayoutComplete(): Promise<void> {
      if (this.layoutComplete) return Promise.resolve();
      return new Promise<void>((resolve) => {
        this.resolveLayout = () => {
          this.availableSlideCount = this.slideCount;
          this.layoutComplete = true;
          resolve();
        };
      });
    }
    setScale(scale: number): void {
      this.events.push('setScale');
      this.setScaleCalls.push(scale);
    }
    getScale(): number { return this.setScaleCalls.at(-1) ?? 1; }
    destroy(): void { this.destroyed = true; }
  }
  return {
    PptxScrollViewer,
    PptxPresentation: class {
      static load(): Promise<{ destroy(): void }> {
        return Promise.resolve({ destroy() {} });
      }
    },
  };
});

vi.mock('@silurus/ooxml-xlsx', () => ({
  XlsxViewer: class {
    destroyed = false;
    readonly loadCalls: unknown[][] = [];
    constructor(
      public readonly host: HTMLElement,
      public readonly opts: Record<string, any>,
    ) {
      mocks.xlsx.push(this as unknown as Record<string, any>);
    }
    load(...args: unknown[]): Promise<void> {
      this.loadCalls.push(args);
      return mocks.rejectXlsx
        ? Promise.reject(new Error('xlsx parse failed'))
        : Promise.resolve();
    }
    destroy(): void { this.destroyed = true; }
  },
  XlsxSheetViewer: class {
    destroyed = false;
    readonly loadCalls: unknown[][] = [];
    constructor(
      public readonly canvas: HTMLCanvasElement,
      public readonly opts: Record<string, any>,
    ) {
      mocks.xlsxSheet.push(this as unknown as Record<string, any>);
    }
    load(...args: unknown[]): Promise<void> {
      this.loadCalls.push(args);
      return mocks.rejectXlsx
        ? Promise.reject(new Error('delimited text parse failed'))
        : Promise.resolve();
    }
    destroy(): void { this.destroyed = true; }
  },
}));

vi.mock('../../../src/math', () => ({ math: mocks.math }));
vi.mock('../../../src/three-d', () => ({ threeD: mocks.threeD }));
vi.mock('../../../src/region-map', () => ({ regionMap: mocks.regionMap }));
vi.mock('../../../src/chart-ex', () => ({ chartEx: mocks.chartEx }));

import { disposeRenderedFile, renderFile } from './try';

class FakeElement {
  className = '';
  textContent = '';
  title = '';
  type = '';
  disabled = false;
  clientWidth = 960;
  readonly style: Record<string, string> = {};
  readonly children: FakeElement[] = [];
  readonly attributes = new Map<string, string>();
  readonly listeners = new Map<string, Array<() => void>>();
  private html = '';

  get innerHTML(): string {
    return this.html;
  }
  set innerHTML(value: string) {
    this.html = value;
    if (value === '') this.children.length = 0;
  }
  append(...children: FakeElement[]): void {
    this.children.push(...children);
  }
  appendChild(child: FakeElement): FakeElement {
    this.children.push(child);
    return child;
  }
  setAttribute(name: string, value: string): void {
    this.attributes.set(name, value);
  }
  addEventListener(type: string, listener: () => void): void {
    const listeners = this.listeners.get(type) ?? [];
    listeners.push(listener);
    this.listeners.set(type, listeners);
  }
}

const file = (name: string): File => ({
  name,
  arrayBuffer: () => Promise.resolve(new ArrayBuffer(8)),
}) as unknown as File;

const stage = (): HTMLElement => new FakeElement() as unknown as HTMLElement;

beforeEach(() => {
  vi.stubGlobal('document', {
    createElement: () => new FakeElement(),
  });
  mocks.deferDocx = false;
  mocks.deferDocxLayout = false;
  mocks.deferPptxLayout = false;
  mocks.rejectXlsx = false;
});

afterEach(() => {
  disposeRenderedFile();
  mocks.docx.length = 0;
  mocks.pptx.length = 0;
  mocks.xlsx.length = 0;
  mocks.xlsxSheet.length = 0;
  vi.unstubAllGlobals();
  vi.clearAllMocks();
});

describe('Try Yours ScrollViewer integration', () => {
  it('lets DOCX fit the preview width and mounts every selectable page for native Find', async () => {
    const hostStage = stage();
    const result = await renderFile(hostStage, file('sample.docx'));
    const viewer = mocks.docx[0];

    expect(result.units).toBe(4);
    expect('width' in viewer.opts).toBe(false);
    expect('onScaleChange' in viewer.opts).toBe(false);
    expect(viewer.opts.enableTextSelection).toBe(true);
    expect(viewer.opts.enableZoom).toBe(true);
    expect(viewer.opts.zoomMin).toBe(0.5);
    expect(viewer.opts.pageShadow).toBe(false);
    expect(viewer.opts.mode).toBe('worker');
    expect(viewer.opts.progressiveLayout).toBe(true);
    expect(viewer.opts.comments).toBe(false);
    expect(viewer.opts.math).toBe(mocks.math);
    expect(viewer.opts.threeD).toBe(mocks.threeD);
    expect(viewer.opts.regionMap).toBe(mocks.regionMap);
    expect(viewer.opts.chartEx).toBe(mocks.chartEx);
    expect(viewer.setScaleCalls).toEqual([]);
    expect(viewer.events[0]).toBe('load');
    expect(viewer.opts.overscan).toBe(4);
    expect(viewer.relayout).toHaveBeenCalledTimes(1);
    const renderedStage = hostStage as unknown as FakeElement;
    expect(renderedStage.children).toHaveLength(1);
    expect(renderedStage.children[0].className).toBe('lv-scroll-viewer');
    expect(renderedStage.children[0].children).toHaveLength(0);
  });

  it('returns opening DOCX pages early and mounts the full native-Find surface after layout converges', async () => {
    mocks.deferDocxLayout = true;
    const progress = vi.fn();
    const result = await renderFile(stage(), file('large.docx'), progress);
    const viewer = mocks.docx[0];

    expect(result.units).toBe(4);
    expect(result.finalUnits).toBeDefined();
    expect(viewer.opts.progressiveLayout).toBe(true);
    expect(viewer.opts.overscan).toBe(0);
    expect(viewer.relayout).not.toHaveBeenCalled();

    viewer.pageCount = 6;
    viewer.opts.onLayoutPartial({ availableUnits: 6, exact: false });
    expect(viewer.opts.overscan).toBe(6);
    expect(viewer.relayout).toHaveBeenCalledTimes(1);
    expect(progress).toHaveBeenLastCalledWith({ format: 'docx', units: 6, unitLabel: 'page' });

    viewer.resolveLayout();
    await expect(result.finalUnits).resolves.toBe(7);
    expect(viewer.opts.overscan).toBe(7);
    expect(viewer.relayout).toHaveBeenCalledTimes(2);
  });

  it('lets PPTX fit the preview width while keeping selection, media, and native Find', async () => {
    const result = await renderFile(stage(), file('sample.pptx'));
    const viewer = mocks.pptx[0];

    expect(result.units).toBe(6);
    expect('width' in viewer.opts).toBe(false);
    expect('onScaleChange' in viewer.opts).toBe(false);
    expect(viewer.opts.enableTextSelection).toBe(true);
    expect(viewer.opts.enableMediaPlayback).toBe(true);
    expect(viewer.opts.mediaOverscan).toBe(1);
    expect(viewer.opts.zoomMin).toBe(0.5);
    expect(viewer.opts.pageShadow).toBe(false);
    expect(viewer.opts.mode).toBe('worker');
    expect(viewer.opts.progressiveLayout).toBe(true);
    expect(viewer.opts.comments).toBe(false);
    expect(viewer.opts.threeD).toBe(mocks.threeD);
    expect(viewer.opts.regionMap).toBe(mocks.regionMap);
    expect(viewer.opts.chartEx).toBe(mocks.chartEx);
    expect(viewer.setScaleCalls).toEqual([]);
    expect(viewer.events[0]).toBe('load');
    expect(viewer.opts.overscan).toBe(6);
    expect(viewer.relayout).toHaveBeenCalledTimes(1);
  });

  it('reports PPTX progressive completion without changing its final slide count', async () => {
    mocks.deferPptxLayout = true;
    const progress = vi.fn();
    const result = await renderFile(stage(), file('progressive.pptx'), progress);
    const viewer = mocks.pptx[0];

    expect(result.units).toBe(2);
    expect(result.finalUnits).toBeDefined();
    expect(viewer.opts.overscan).toBe(0);
    expect(viewer.relayout).not.toHaveBeenCalled();

    viewer.availableSlideCount = 4;
    viewer.opts.onLayoutPartial({ availableUnits: 4, totalUnits: 6, exact: false });
    expect(viewer.opts.overscan).toBe(4);
    expect(viewer.relayout).toHaveBeenCalledTimes(1);
    expect(progress).toHaveBeenLastCalledWith({ format: 'pptx', units: 4, unitLabel: 'slide' });

    viewer.resolveLayout();
    await expect(result.finalUnits).resolves.toBe(6);
    expect(viewer.opts.overscan).toBe(6);
    expect(viewer.relayout).toHaveBeenCalledTimes(2);
  });

  it.each([
    { ext: 'docx', instances: mocks.docx },
    { ext: 'pptx', instances: mocks.pptx },
  ])('does not override the width-derived $ext scale', async ({
    ext,
    instances,
  }) => {
    await renderFile(stage(), file(`narrow.${ext}`));
    expect(instances[0].opts.zoomMin).toBe(0.5);
    expect(instances[0].setScaleCalls).toEqual([]);
  });

  it('destroys an XLSX viewer when load fails', async () => {
    mocks.rejectXlsx = true;
    await expect(renderFile(stage(), file('broken.xlsx'))).rejects.toThrow('xlsx parse failed');
    expect(mocks.xlsx[0].destroyed).toBe(true);
  });

  it('enables equations and every optional chart renderer for XLSX', async () => {
    await renderFile(stage(), file('advanced.xlsx'));
    const viewer = mocks.xlsx[0];

    expect(viewer.opts.mode).toBe('main');
    expect(viewer.opts.comments).toBe(false);
    expect(viewer.opts.math).toBeDefined();
    expect(viewer.opts.threeD).toBe(mocks.threeD);
    expect(viewer.opts.regionMap).toBe(mocks.regionMap);
    expect(viewer.opts.chartEx).toBe(mocks.chartEx);
    expect(viewer.opts.useGoogleFonts).toBe(true);
  });

  it.each(['csv', 'tsv'] as const)(
    'opens a selected %s file in the sheet-only viewer',
    async (format) => {
      const hostStage = stage();
      const result = await renderFile(hostStage, file(`table.${format}`));
      const viewer = mocks.xlsxSheet[0];

      expect(result).toEqual({ format: 'xlsx', units: 0, unitLabel: 'sheet' });
      expect(mocks.xlsx).toHaveLength(0);
      expect(mocks.xlsxSheet).toHaveLength(1);
      expect(viewer.loadCalls).toHaveLength(1);
      expect(viewer.loadCalls[0][0]).toBeInstanceOf(ArrayBuffer);
      expect(viewer.loadCalls[0][1]).toEqual({ format });
      expect((hostStage as unknown as FakeElement).children[0].className).toBe('lv-xlsx');
    },
  );

  it('destroys a CSV sheet viewer when parsing fails', async () => {
    mocks.rejectXlsx = true;
    await expect(renderFile(stage(), file('broken.csv'))).rejects.toThrow(
      'delimited text parse failed',
    );
    expect(mocks.xlsxSheet[0].destroyed).toBe(true);
  });

  it.each(['pdf', 'dat'])(
    'destroys the active viewer when a later .%s selection is unsupported',
    async (extension) => {
      await renderFile(stage(), file('valid.docx'));
      const active = mocks.docx[0];

      await expect(renderFile(stage(), file(`unsupported.${extension}`))).rejects.toThrow(
        'Unsupported file',
      );
      expect(active.destroyed).toBe(true);
      expect(mocks.xlsxSheet).toHaveLength(0);
    },
  );

  it('destroys a superseded viewer and prevents it from becoming active', async () => {
    mocks.deferDocx = true;
    const sharedStage = stage();
    const firstPromise = renderFile(sharedStage, file('first.docx'));
    await Promise.resolve();
    await Promise.resolve();
    const first = mocks.docx[0];

    const secondPromise = renderFile(sharedStage, file('second.docx'));
    await Promise.resolve();
    await Promise.resolve();
    const second = mocks.docx[1];

    second.resolveLoad();
    await secondPromise;
    first.resolveLoad();
    await expect(firstPromise).rejects.toMatchObject({ name: 'AbortError' });

    expect(first.destroyed).toBe(true);
    expect(second.destroyed).toBe(false);
    disposeRenderedFile();
    expect(second.destroyed).toBe(true);
  });
});
