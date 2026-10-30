import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ModelSource } from '@silurus/ooxml-core';
import { PptxPresentation } from './presentation.js';
import { PptxScrollViewer } from './scroll-viewer.js';
import { PptxViewer } from './viewer.js';
import { FakePptxEngine, installDom, makeContainer, makeEl, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const source = { target: 'pptx', claim: () => false, beginLoad: vi.fn() } as unknown as ModelSource<'pptx'>;

function slideViewer(options: ConstructorParameters<typeof PptxViewer>[1]) {
  installDom();
  return new PptxViewer(makeEl('canvas') as unknown as HTMLCanvasElement, options);
}

function scrollViewer(options: ConstructorParameters<typeof PptxScrollViewer>[1]) {
  installDom();
  const container = makeContainer(200, 400);
  const viewer = new PptxScrollViewer(container as unknown as HTMLElement, { gap: 10, ...options });
  const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
  scrollHost.clientHeight = 400;
  scrollHost.clientWidth = 200;
  return viewer;
}

// The viewers own the load call: a model source they drop would silently send
// the input down the OOXML path instead.
describe.each([
  { name: 'PptxViewer', create: slideViewer },
  { name: 'PptxScrollViewer', create: scrollViewer },
])('$name load options', ({ create }) => {
  it('forwards legacyPpt to PptxPresentation.load only when configured', async () => {
    const load = vi.spyOn(PptxPresentation, 'load')
      .mockImplementation(async () => new FakePptxEngine(1, 9144000, 6858000).asPres());
    const odp = { ...source } as typeof source;
    for (const [options, legacyPpt, expectedOdp] of [
      [{ legacyPpt: source }, source, undefined],
      [{ odp }, undefined, odp],
      [{ legacyPpt: source, odp }, source, odp],
      [{}, undefined, undefined],
    ] as const) {
      const viewer = create(options);
      await viewer.load(new ArrayBuffer(1));
      expect(load.mock.lastCall?.[1]?.legacyPpt).toEqual(legacyPpt);
      expect(load.mock.lastCall?.[1]?.odp).toEqual(expectedOdp);
      viewer.destroy();
    }
  });
});

// Each surface takes only a source for its own format, under its own key.
describe('PPTX source option typing', () => {
  it('rejects another format\'s source and the removed modelSources option', () => {
    const docx = { target: 'docx', claim: () => false, beginLoad: vi.fn() } as unknown as ModelSource<'docx'>;
    const xlsx = { target: 'xlsx', claim: () => false, beginLoad: vi.fn() } as unknown as ModelSource<'xlsx'>;
    // @ts-expect-error A DOCX source cannot configure a PPTX surface.
    const pptxViewerDocx: ConstructorParameters<typeof PptxViewer>[1] = { legacyPpt: docx };
    // @ts-expect-error An XLSX source cannot configure a PPTX surface.
    const pptxViewerXlsx: ConstructorParameters<typeof PptxViewer>[1] = { legacyPpt: xlsx };
    // @ts-expect-error The generic list option was replaced by `legacyPpt`.
    const pptxViewerList: ConstructorParameters<typeof PptxViewer>[1] = { modelSources: [] };
    // @ts-expect-error A DOCX source cannot configure a PPTX surface.
    const pptxScrollViewerDocx: ConstructorParameters<typeof PptxScrollViewer>[1] = { legacyPpt: docx };
    // @ts-expect-error An XLSX source cannot configure a PPTX surface.
    const pptxScrollViewerXlsx: ConstructorParameters<typeof PptxScrollViewer>[1] = { legacyPpt: xlsx };
    // @ts-expect-error The generic list option was replaced by `legacyPpt`.
    const pptxScrollViewerList: ConstructorParameters<typeof PptxScrollViewer>[1] = { modelSources: [] };
    // @ts-expect-error A DOCX source cannot load PPTX input.
    const loadDocx: NonNullable<Parameters<typeof PptxPresentation.load>[1]> = { legacyPpt: docx };
    // @ts-expect-error An XLSX source cannot load PPTX input.
    const loadXlsx: NonNullable<Parameters<typeof PptxPresentation.load>[1]> = { legacyPpt: xlsx };
    expect([pptxViewerDocx, pptxViewerXlsx, pptxViewerList, pptxScrollViewerDocx, pptxScrollViewerXlsx, pptxScrollViewerList, loadDocx, loadXlsx]).toHaveLength(8);
  });
});
