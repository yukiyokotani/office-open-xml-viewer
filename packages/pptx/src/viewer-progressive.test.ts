import { afterEach, describe, expect, it, vi } from 'vitest';
import { PptxPresentation } from './presentation.js';
import { PptxViewer } from './viewer.js';
import { FakePptxEngine, installDom, makeEl, type FakeEl } from './scroll-viewer-test-dom.js';

const WIDTH = 9144000;
const HEIGHT = 5143500;

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function mountCanvas(): FakeEl {
  installDom();
  const parent = makeEl('div');
  const canvas = makeEl('canvas');
  parent.appendChild(canvas);
  return canvas;
}

describe('PptxViewer progressive layout', () => {
  it('forwards the symmetric lifecycle options and exposes presentation state', async () => {
    const canvas = mountCanvas();
    const engine = new FakePptxEngine(3, WIDTH, HEIGHT);
    const onLayoutProgress = vi.fn();
    const onLayoutPartial = vi.fn();
    const onLayoutComplete = vi.fn();
    const load = vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
    const viewer = new PptxViewer(canvas as unknown as HTMLCanvasElement, {
      progressiveLayout: true,
      onLayoutProgress,
      onLayoutPartial,
      onLayoutComplete,
    });

    await viewer.load('deck.pptx');

    expect(load).toHaveBeenCalledWith('deck.pptx', expect.objectContaining({
      progressiveLayout: true,
      onLayoutProgress,
      onLayoutPartial,
      onLayoutComplete,
    }));
    expect(viewer.availableSlideCount).toBe(3);
    expect(viewer.layoutComplete).toBe(true);
    await expect(viewer.waitUntilLayoutComplete()).resolves.toBeUndefined();
    viewer.destroy();
  });

  it('shows a loading overlay while navigation waits and reports completion changes', async () => {
    const canvas = mountCanvas();
    const engine = new FakePptxEngine(3, WIDTH, HEIGHT);
    engine.setLayoutProgress(1, false);
    const onSlideChange = vi.fn();
    const viewer = PptxViewer.fromPresentation(
      canvas as unknown as HTMLCanvasElement,
      engine.asPres(),
      { onSlideChange },
    ) as PptxViewer;
    const wrapper = canvas.parentElement as FakeEl;
    const loading = wrapper.children.find((child) =>
      child.children.some((grandchild) =>
        grandchild.className === 'ooxml-pptx-progress-circle'))!;

    expect(loading.children.some((child) => child.tag === 'progress')).toBe(false);
    expect(loading.children[0]?.style['border-radius']).toBe('50%');

    const navigation = viewer.goToSlide(2);
    expect(loading.style.display).toBe('flex');
    expect(viewer.slideCount).toBe(3);
    expect(viewer.availableSlideCount).toBe(1);
    expect(viewer.layoutComplete).toBe(false);

    engine.setLayoutProgress(3, true);
    await navigation;
    expect(loading.style.display).toBe('none');
    expect(onSlideChange).toHaveBeenLastCalledWith(2, 3, true);
    viewer.destroy();
  });

  it('waits for slide facts before applying hidden-slide dimming', async () => {
    const canvas = mountCanvas();
    const engine = new FakePptxEngine(3, WIDTH, HEIGHT);
    engine.hiddenSlides.add(2);
    engine.setLayoutProgress(1, false);
    const renderSlide = vi.spyOn(engine, 'renderSlide');
    const viewer = PptxViewer.fromPresentation(
      canvas as unknown as HTMLCanvasElement,
      engine.asPres(),
      { hiddenSlideMode: 'dim' },
    ) as PptxViewer;

    const navigation = viewer.goToSlide(2);
    expect(renderSlide).not.toHaveBeenCalled();
    engine.setLayoutProgress(3, true);
    await navigation;

    expect(renderSlide).toHaveBeenLastCalledWith(
      expect.anything(),
      2,
      expect.objectContaining({ dim: { color: '#ffffff', opacity: 0.6 } }),
    );
    viewer.destroy();
  });

  it('waits through an unavailable hidden slide before skip navigation chooses the next visible one', async () => {
    const canvas = mountCanvas();
    const engine = new FakePptxEngine(3, WIDTH, HEIGHT);
    engine.hiddenSlides.add(1);
    engine.setLayoutProgress(1, false);
    const viewer = PptxViewer.fromPresentation(
      canvas as unknown as HTMLCanvasElement,
      engine.asPres(),
      { hiddenSlideMode: 'skip' },
    ) as PptxViewer;

    const navigation = viewer.nextSlide();
    engine.setLayoutProgress(2, false);
    await Promise.resolve();
    expect(viewer.slideIndex).toBe(0);

    engine.setLayoutProgress(3, true);
    await navigation;
    expect(viewer.slideIndex).toBe(2);
    viewer.destroy();
  });

  it('settles superseded navigation immediately instead of waiting for another layout publication', async () => {
    const canvas = mountCanvas();
    const engine = new FakePptxEngine(3, WIDTH, HEIGHT);
    engine.setLayoutProgress(1, false);
    const viewer = PptxViewer.fromPresentation(
      canvas as unknown as HTMLCanvasElement,
      engine.asPres(),
    ) as PptxViewer;

    let firstSettled = false;
    const first = viewer.goToSlide(2).then(() => { firstSettled = true; });
    await viewer.goToSlide(0);
    await first;

    expect(firstSettled).toBe(true);
    expect(viewer.slideIndex).toBe(0);
    viewer.destroy();
  });

  it('keeps a concurrent progressive failure on the canceled navigation Promise channel', async () => {
    const canvas = mountCanvas();
    const engine = new FakePptxEngine(3, WIDTH, HEIGHT);
    engine.setLayoutProgress(1, false);
    const onError = vi.fn();
    const viewer = PptxViewer.fromPresentation(
      canvas as unknown as HTMLCanvasElement,
      engine.asPres(),
      { onError },
    ) as PptxViewer;
    const failure = new Error('layout failed during navigation cancellation');

    const navigation = viewer.goToSlide(2);
    void viewer.goToSlide(0);
    engine.setLayoutProgress(1, false, failure);

    await expect(navigation).rejects.toBe(failure);
    expect(onError).not.toHaveBeenCalled();
    viewer.destroy();
  });

  it('does not lose a layout failure when clearFind supersedes a progressive search', async () => {
    const canvas = mountCanvas();
    const engine = new FakePptxEngine(3, WIDTH, HEIGHT);
    engine.setLayoutProgress(1, false);
    const onError = vi.fn();
    const viewer = PptxViewer.fromPresentation(
      canvas as unknown as HTMLCanvasElement,
      engine.asPres(),
      { onError },
    ) as PptxViewer;
    const failure = new Error('layout failed during find cancellation');

    const find = viewer.findText('needle');
    viewer.clearFind();
    engine.setLayoutProgress(1, false, failure);

    await expect(find).rejects.toBe(failure);
    expect(onError).not.toHaveBeenCalled();
    viewer.destroy();
  });
});
