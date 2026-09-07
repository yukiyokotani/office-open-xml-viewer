import { describe, expect, it, vi } from 'vitest';
import { TerminalResourceOwner } from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import type { PptxPresentation } from './presentation.js';
import { settleLegacyPptLoad } from './legacy-ppt-load.js';

const sourceOptions = {
  ppt: { source: {
    protocol: 'ooxml-legacy-ppt-source/v1' as const,
    builtin: 'ppt' as const,
    wasmUrl: 'https://example.test/direct.wasm',
  } },
};

describe('settleLegacyPptLoad', () => {
  function ownedPresentation() {
    let retained: () => void = () => undefined;
    const presentation = {
      _retainLegacyPptSignalCleanup(cleanup: () => void) { retained = cleanup; },
      destroy() {
        const cleanup = retained;
        retained = () => undefined;
        cleanup();
      },
    } as unknown as PptxPresentation;
    return presentation;
  }

  it('cleans direct-source signal wiring when loading fails', async () => {
    const failure = new Error('load failed');
    const cleanup = vi.fn();
    await expect(settleLegacyPptLoad(Promise.reject(failure), {
      options: sourceOptions,
      cleanup,
    })).rejects.toBe(failure);
    expect(cleanup).toHaveBeenCalledOnce();
  });

  it('transfers successful direct-source cleanup without releasing it early', async () => {
    const cleanup = vi.fn();
    const retain = vi.fn();
    const presentation = {
      _retainLegacyPptSignalCleanup: retain,
    } as unknown as PptxPresentation;
    await expect(settleLegacyPptLoad(Promise.resolve(presentation), {
      options: sourceOptions,
      cleanup,
    })).resolves.toBe(presentation);
    expect(cleanup).not.toHaveBeenCalled();
    expect(retain).toHaveBeenCalledWith(cleanup);
  });

  it('keeps byte-converter cleanup scoped to promise settlement', async () => {
    const cleanup = vi.fn();
    const presentation = {} as PptxPresentation;
    await settleLegacyPptLoad(Promise.resolve(presentation), {
      options: { ppt: { converter: { convert: vi.fn() } } },
      cleanup,
    });
    expect(cleanup).toHaveBeenCalledOnce();
  });

  it('releases transferred cleanup for both swapped and stale candidates', async () => {
    const owner = new TerminalResourceOwner<PptxPresentation>('test');
    const firstCleanup = vi.fn();
    const staleCleanup = vi.fn();
    const winnerCleanup = vi.fn();
    await owner.replace(() => settleLegacyPptLoad(Promise.resolve(ownedPresentation()), {
      options: sourceOptions, cleanup: firstCleanup,
    }));

    let resolveStale!: (presentation: PptxPresentation) => void;
    const stale = owner.replace(() => settleLegacyPptLoad(
      new Promise<PptxPresentation>((resolve) => { resolveStale = resolve; }),
      { options: sourceOptions, cleanup: staleCleanup },
    ));
    await owner.replace(() => settleLegacyPptLoad(Promise.resolve(ownedPresentation()), {
      options: sourceOptions, cleanup: winnerCleanup,
    }));
    expect(firstCleanup).toHaveBeenCalledOnce();
    resolveStale(ownedPresentation());
    await stale;
    expect(staleCleanup).toHaveBeenCalledOnce();
    expect(winnerCleanup).not.toHaveBeenCalled();

    owner.close();
    expect(winnerCleanup).toHaveBeenCalledOnce();
  });
});
