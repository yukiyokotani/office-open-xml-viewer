import { afterEach, describe, expect, it, vi } from 'vitest';
import { PptxPresentation } from './presentation';
import {
  installDom,
  makeEl,
  FakePptxEngine,
  type FakeEl,
} from './scroll-viewer-test-dom';
import { PptxViewer, type PptxShapeClickEvent } from './viewer';
import type { ShapeElement } from './types';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SLIDE_WIDTH = 9_144_000;
const SLIDE_HEIGHT = 6_858_000;

async function mountLoaded(onShapeClick: (event: PptxShapeClickEvent) => void) {
  installDom();
  const canvas = makeEl('canvas');
  canvas.clientWidth = 960;
  canvas.clientHeight = 720;
  const engine = new FakePptxEngine(1, SLIDE_WIDTH, SLIDE_HEIGHT);
  engine.shapeHit = {
    slideIndex: 0,
    shapeId: '7',
    shape: { type: 'shape', id: '7' } as unknown as ShapeElement,
    point: { x: SLIDE_WIDTH / 2, y: SLIDE_HEIGHT / 2 },
  };
  vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
  const viewer = new PptxViewer(
    canvas as unknown as HTMLCanvasElement,
    { onShapeClick },
  );
  await viewer.load('x.pptx');
  return {
    viewer,
    engine,
    canvas,
    wrapper: canvas.parentElement as FakeEl,
  };
}

describe('PptxViewer shape clicks', () => {
  it('converts wrapper clicks to slide EMU and emits the topmost shape hit', async () => {
    const onShapeClick = vi.fn<(event: PptxShapeClickEvent) => void>();
    const { viewer, engine, wrapper } = await mountLoaded(onShapeClick);
    const nativeEvent = {
      clientX: 480,
      clientY: 360,
      defaultPrevented: false,
    };

    wrapper.dispatch('click', nativeEvent);

    expect(engine.hitTestCalls).toEqual([{
      slide: 0,
      point: { x: SLIDE_WIDTH / 2, y: SLIDE_HEIGHT / 2 },
      opts: { tolerance: (6 / 960) * SLIDE_WIDTH },
    }]);
    expect(onShapeClick).toHaveBeenCalledWith({
      ...engine.shapeHit,
      nativeEvent,
    });
    viewer.destroy();
  });

  it('ignores prevented hyperlink clicks and empty-space hits', async () => {
    const onShapeClick = vi.fn<(event: PptxShapeClickEvent) => void>();
    const { viewer, engine, wrapper } = await mountLoaded(onShapeClick);

    wrapper.dispatch('click', {
      clientX: 100,
      clientY: 100,
      defaultPrevented: true,
    });
    engine.shapeHit = null;
    wrapper.dispatch('click', {
      clientX: 100,
      clientY: 100,
      defaultPrevented: false,
    });

    expect(engine.hitTestCalls).toHaveLength(1);
    expect(onShapeClick).not.toHaveBeenCalled();
    viewer.destroy();
  });

  it('removes its stable click listener on destroy', async () => {
    const onShapeClick = vi.fn<(event: PptxShapeClickEvent) => void>();
    const { viewer, engine, wrapper } = await mountLoaded(onShapeClick);

    viewer.destroy();
    wrapper.dispatch('click', {
      clientX: 480,
      clientY: 360,
      defaultPrevented: false,
    });

    expect(engine.hitTestCalls).toEqual([]);
    expect(onShapeClick).not.toHaveBeenCalled();
  });

  it('rejects shape click wiring in worker mode', () => {
    installDom();
    const canvas = makeEl('canvas');
    expect(() =>
      new PptxViewer(
        canvas as unknown as HTMLCanvasElement,
        {
          mode: 'worker',
          onShapeClick: () => {},
        },
      ),
    ).toThrow(/onShapeClick.*mode.*main/i);
  });
});
