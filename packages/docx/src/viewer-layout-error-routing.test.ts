import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxDocument, type LoadOptions } from './document.js';
import { publishDocxLayout } from './document-layout-events.js';
import { DocxViewer } from './viewer.js';
import { DocxScrollViewer } from './scroll-viewer.js';
import { FakeDocxEngine, installDom, makeContainer, makeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

type Surface = DocxViewer | DocxScrollViewer;

async function loadedSurface(
  kind: 'viewer' | 'scroll',
  callbacks: Pick<LoadOptions, 'onLayoutComplete'> & { onError?: (error: Error) => void },
): Promise<{ engine: FakeDocxEngine; surface: Surface; loadOptions: LoadOptions }> {
  installDom();
  const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
  engine.setLayoutComplete(false);
  const load = vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
  const surface = kind === 'viewer'
    ? new DocxViewer(makeEl('canvas') as unknown as HTMLCanvasElement, callbacks)
    : new DocxScrollViewer(makeContainer() as unknown as HTMLElement, callbacks);
  await surface.load('progressive.docx');
  return { engine, surface, loadOptions: load.mock.calls[0]![1] as LoadOptions };
}

describe.each(['viewer', 'scroll'] as const)('DOCX %s progressive failure routing', (kind) => {
  it('routes an actively awaited failure only through the returned Promise', async () => {
    const onError = vi.fn();
    const { engine, surface } = await loadedSurface(kind, { onError });
    const error = new Error('layout failed');
    const waiting = surface.waitUntilLayoutComplete();
    engine.setLayoutFailure(error);
    publishDocxLayout(engine.asDoc(), {
      pageCount: 1, exact: false, complete: false, error,
    });

    await expect(waiting).rejects.toBe(error);
    expect(onError).not.toHaveBeenCalled();
    surface.destroy();
  });

  it('routes an unawaited failure once through onError', async () => {
    const onError = vi.fn();
    const { engine, surface } = await loadedSurface(kind, { onError });
    const error = new Error('layout failed');
    engine.setLayoutFailure(error);
    publishDocxLayout(engine.asDoc(), {
      pageCount: 1, exact: false, complete: false, error,
    });

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
    engine.setLayoutFailure(error);
    loadOptions.onLayoutComplete?.(error);
    publishDocxLayout(engine.asDoc(), {
      pageCount: 1, exact: false, complete: false, error,
    });

    expect(onLayoutComplete).toHaveBeenCalledTimes(1);
    expect(onLayoutComplete).toHaveBeenCalledWith(error);
    expect(onError).not.toHaveBeenCalled();
    surface.destroy();
  });
});
