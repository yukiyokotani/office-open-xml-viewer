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
  it('forwards modelSources to PptxPresentation.load only when configured', async () => {
    const load = vi.spyOn(PptxPresentation, 'load')
      .mockImplementation(async () => new FakePptxEngine(1, 9144000, 6858000).asPres());
    for (const [options, expected] of [[{ modelSources: [source] }, [source]], [{}, undefined]] as const) {
      const viewer = create(options);
      await viewer.load(new ArrayBuffer(1));
      expect(load.mock.lastCall?.[1]?.modelSources).toEqual(expected);
      viewer.destroy();
    }
  });
});
