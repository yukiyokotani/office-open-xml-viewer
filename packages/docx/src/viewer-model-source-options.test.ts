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
  it('forwards modelSources and an explicit showTrackedChanges: false', async () => {
    const engine = new FakeDocxEngine(1, SIZE);
    const load = vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    const viewer = create({ modelSources: [source], showTrackedChanges: false });
    await viewer.load(new ArrayBuffer(1));
    expect(load).toHaveBeenCalledExactlyOnceWith(
      expect.any(ArrayBuffer),
      expect.objectContaining({ modelSources: [source], showTrackedChanges: false }),
    );
    viewer.destroy();
  });

  it('leaves an unchosen view to the document and renders the view it loaded with', async () => {
    const engine = new FakeDocxEngine(1, SIZE);
    // What a model source's markup view default leaves on the loaded document.
    engine.layoutView = { showTrackedChanges: true, currentDate: 0 };
    const load = vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    const viewer = create({});
    await viewer.load(new ArrayBuffer(1));
    const options = load.mock.calls[0]![1]!;
    expect(options).not.toHaveProperty('showTrackedChanges');
    expect(options).not.toHaveProperty('modelSources');
    await vi.waitFor(() => expect(engine.renderCalls.length).toBeGreaterThan(0));
    expect(new Set(engine.renderCalls.map((call) => call.showTrackedChanges))).toEqual(new Set([renderedView]));
    viewer.destroy();
  });
});
