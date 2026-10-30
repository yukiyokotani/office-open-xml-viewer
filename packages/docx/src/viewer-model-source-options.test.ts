import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ModelSource } from '@silurus/ooxml-core';
import { DocxDocument } from './document.js';
import { DocxScrollViewer } from './scroll-viewer.js';
import { DocxViewer } from './viewer.js';
import { FakeDocxEngine, installDom, makeContainer, makeEl, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SIZE = [{ widthPt: 100, heightPt: 200 }];
const source = { target: 'docx', claim: () => false, beginLoad: vi.fn() } as unknown as ModelSource<'docx'>;

function docxViewer(options: ConstructorParameters<typeof DocxViewer>[1]) {
  installDom();
  return new DocxViewer(makeEl('canvas') as unknown as HTMLCanvasElement, options);
}

function scrollViewer(options: ConstructorParameters<typeof DocxScrollViewer>[1]) {
  installDom();
  const container = makeContainer(200, 400);
  const viewer = new DocxScrollViewer(container as unknown as HTMLElement, { gap: 10, ...options });
  const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
  scrollHost.clientHeight = 400;
  scrollHost.clientWidth = 200;
  return viewer;
}

// The viewers own the load call, so a model source or an explicit `false` view
// that they drop never reaches DocxDocument.load: the document would open on the
// OOXML path, or the source's markup default would override the caller.
describe.each([
  // DocxViewer leaves the per-render view unset so the document fills in its
  // active view; the scroll viewer adopts the loaded view and passes it.
  { name: 'DocxViewer', create: docxViewer, renderedView: undefined },
  { name: 'DocxScrollViewer', create: scrollViewer, renderedView: true },
])('$name load options', ({ create, renderedView }) => {
  it('forwards legacyDoc and an explicit showTrackedChanges: false', async () => {
    const engine = new FakeDocxEngine(1, SIZE);
    const load = vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    const viewer = create({ legacyDoc: source, showTrackedChanges: false });
    await viewer.load(new ArrayBuffer(1));
    expect(load).toHaveBeenCalledExactlyOnceWith(
      expect.any(ArrayBuffer),
      expect.objectContaining({ legacyDoc: source, showTrackedChanges: false }),
    );
    viewer.destroy();
  });

  it('forwards odt beside legacyDoc and applies the same source view precedence to it', async () => {
    const engine = new FakeDocxEngine(1, SIZE);
    engine.layoutView = { showTrackedChanges: true, currentDate: 0 };
    const load = vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    const odt = { ...source } as typeof source;
    const viewer = create({ legacyDoc: source, odt });
    await viewer.load(new ArrayBuffer(1));
    const options = load.mock.calls[0]![1]!;
    expect(options.legacyDoc).toBe(source);
    expect(options.odt).toBe(odt);
    expect(options).not.toHaveProperty('showTrackedChanges');
    viewer.destroy();

    const only = create({ odt });
    await only.load(new ArrayBuffer(1));
    expect(load.mock.lastCall![1]).toMatchObject({ odt });
    expect(load.mock.lastCall![1]).not.toHaveProperty('legacyDoc');
    await vi.waitFor(() => expect(engine.renderCalls.length).toBeGreaterThan(0));
    expect(new Set(engine.renderCalls.map((call) => call.showTrackedChanges))).toEqual(new Set([renderedView]));
    only.destroy();
  });

  it('leaves an unchosen view to the document and renders the view it loaded with', async () => {
    const engine = new FakeDocxEngine(1, SIZE);
    // What a model source's markup view default leaves on the loaded document.
    engine.layoutView = { showTrackedChanges: true, currentDate: 0 };
    const load = vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    const viewer = create({ legacyDoc: source });
    await viewer.load(new ArrayBuffer(1));
    const options = load.mock.calls[0]![1]!;
    expect(options).not.toHaveProperty('showTrackedChanges');
    expect(options.legacyDoc).toBe(source);
    await vi.waitFor(() => expect(engine.renderCalls.length).toBeGreaterThan(0));
    expect(new Set(engine.renderCalls.map((call) => call.showTrackedChanges))).toEqual(new Set([renderedView]));
    viewer.destroy();
  });

  it('preserves an explicit false setter call before load over the source default', async () => {
    const engine = new FakeDocxEngine(1, SIZE);
    engine.layoutView = { showTrackedChanges: true, currentDate: 0 };
    const load = vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    const viewer = create({ legacyDoc: source });
    await viewer.setShowTrackedChanges(false);
    await viewer.load(new ArrayBuffer(1));
    expect(load.mock.calls[0]![1]).toHaveProperty('showTrackedChanges', false);
    viewer.destroy();
  });
});

// Each surface takes only a source for its own format, under its own key.
describe('DOCX source option typing', () => {
  it('rejects another format\'s source and the removed modelSources option', () => {
    const xlsx = { target: 'xlsx', claim: () => false, beginLoad: vi.fn() } as unknown as ModelSource<'xlsx'>;
    const pptx = { target: 'pptx', claim: () => false, beginLoad: vi.fn() } as unknown as ModelSource<'pptx'>;
    // @ts-expect-error An XLSX source cannot configure a DOCX surface.
    const docxViewerXlsx: ConstructorParameters<typeof DocxViewer>[1] = { legacyDoc: xlsx };
    // @ts-expect-error A PPTX source cannot configure a DOCX surface.
    const docxViewerPptx: ConstructorParameters<typeof DocxViewer>[1] = { legacyDoc: pptx };
    // @ts-expect-error The generic list option was replaced by `legacyDoc`.
    const docxViewerList: ConstructorParameters<typeof DocxViewer>[1] = { modelSources: [] };
    // @ts-expect-error An XLSX source cannot configure a DOCX surface.
    const docxScrollViewerXlsx: ConstructorParameters<typeof DocxScrollViewer>[1] = { legacyDoc: xlsx };
    // @ts-expect-error A PPTX source cannot configure a DOCX surface.
    const docxScrollViewerPptx: ConstructorParameters<typeof DocxScrollViewer>[1] = { legacyDoc: pptx };
    // @ts-expect-error The generic list option was replaced by `legacyDoc`.
    const docxScrollViewerList: ConstructorParameters<typeof DocxScrollViewer>[1] = { modelSources: [] };
    // @ts-expect-error An XLSX source cannot load DOCX input.
    const loadXlsx: NonNullable<Parameters<typeof DocxDocument.load>[1]> = { legacyDoc: xlsx };
    // @ts-expect-error A PPTX source cannot load DOCX input.
    const loadPptx: NonNullable<Parameters<typeof DocxDocument.load>[1]> = { legacyDoc: pptx };
    expect([docxViewerXlsx, docxViewerPptx, docxViewerList, docxScrollViewerXlsx, docxScrollViewerPptx, docxScrollViewerList, loadXlsx, loadPptx]).toHaveLength(8);
  });
});
