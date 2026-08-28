import { afterEach, describe, expect, it, vi } from 'vitest';
import { PptxPresentation, type LoadOptions } from './presentation.js';
import { publishPptxLayout } from './presentation-layout-events.js';
import { PptxViewer } from './viewer.js';
import { PptxScrollViewer } from './scroll-viewer.js';
import { FakePptxEngine, installDom, makeContainer, makeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

type Surface = PptxViewer | PptxScrollViewer;

async function loadedSurface(
  kind: 'viewer' | 'scroll',
  callbacks: Pick<LoadOptions, 'onLayoutComplete'> & { onError?: (error: Error) => void },
): Promise<{ engine: FakePptxEngine; surface: Surface; loadOptions: LoadOptions }> {
  installDom();
  const engine = new FakePptxEngine(2, 9144000, 5143500);
  engine.setLayoutProgress(1, false);
  const load = vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
  const surface = kind === 'viewer'
    ? new PptxViewer(makeEl('canvas') as unknown as HTMLCanvasElement, callbacks)
    : new PptxScrollViewer(makeContainer(200, 50) as unknown as HTMLElement, {
        ...callbacks,
        overscan: 0,
      });
  await surface.load('progressive.pptx');
  return { engine, surface, loadOptions: load.mock.calls[0]![1] as LoadOptions };
}

describe.each(['viewer', 'scroll'] as const)('PPTX %s progressive failure routing', (kind) => {
  it('routes an actively awaited failure only through the returned Promise', async () => {
    const onError = vi.fn();
    const { engine, surface } = await loadedSurface(kind, { onError });
    const error = new Error('layout failed');
    const waiting = surface.waitUntilLayoutComplete();
    engine.setLayoutProgress(1, false, error);

    await expect(waiting).rejects.toBe(error);
    expect(onError).not.toHaveBeenCalled();
    surface.destroy();
  });

  it('routes a progressive search failure only through the find Promise', async () => {
    const onError = vi.fn();
    const { engine, surface } = await loadedSurface(kind, { onError });
    const error = new Error('layout failed during find');
    vi.spyOn(engine, 'collectSlideRuns').mockImplementation(async (slide) => {
      if (slide >= engine.availableSlideCount) await engine.waitUntilLayoutComplete();
      return [];
    });
    const searching = surface.findText('needle');
    engine.setLayoutProgress(1, false, error);

    await expect(searching).rejects.toBe(error);
    expect(onError).not.toHaveBeenCalled();
    surface.destroy();
  });

  it('routes an unawaited failure once through onError', async () => {
    const onError = vi.fn();
    const { engine, surface } = await loadedSurface(kind, { onError });
    const error = new Error('layout failed');
    engine.setLayoutProgress(1, false, error);

    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError).toHaveBeenCalledWith(error);
    surface.destroy();
  });

  it('leaves an explicitly configured completion callback as the sole owner', async () => {
    const onError = vi.fn();
    const onLayoutComplete = vi.fn();
    const { engine, surface, loadOptions } = await loadedSurface(
      kind,
      { onError, onLayoutComplete },
    );
    const error = new Error('layout failed');
    engine.setLayoutProgress(1, false, error);
    loadOptions.onLayoutComplete?.(error);

    expect(onLayoutComplete).toHaveBeenCalledTimes(1);
    expect(onLayoutComplete).toHaveBeenCalledWith(error);
    expect(onError).not.toHaveBeenCalled();
    surface.destroy();
  });
});
