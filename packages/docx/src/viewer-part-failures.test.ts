import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import { DocxScrollViewer } from './scroll-viewer.js';
import { installDom, makeEl, FakeDocxEngine } from './scroll-viewer-test-dom.js';

/** #1974 — viewer delivery of DOCX part failures and option forwarding. */
afterEach(() => {
  vi.restoreAllMocks();
});

const FAILURE = { kind: 'document', message: 'word/document.xml: bad' } as const;

function brokenEngine(): FakeDocxEngine {
  const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
  engine._partFailureLog.add(FAILURE);
  return engine;
}

describe('DocxViewer part failures (#1974)', () => {
  it('replays the failure at engine commit and forwards errorPlaceholder', async () => {
    installDom();
    const parent = makeEl('div');
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const engine = brokenEngine();
    const seen: unknown[] = [];
    const errorPlaceholder = () => ({ title: 'Unavailable' });
    const viewer = DocxViewer.fromDocument(canvas as unknown as HTMLCanvasElement, engine.asDoc(), {
      onPartFailure: (failure) => seen.push(failure),
      errorPlaceholder,
    });
    expect(seen).toEqual([FAILURE]);
    expect(viewer.partFailures).toEqual([FAILURE]);
    await viewer.goToPage(0);
    expect(engine.renderCalls.at(-1)?.errorPlaceholder).toBe(errorPlaceholder);
    viewer.destroy();
    expect(viewer.partFailures).toEqual([]);
    expect(seen).toHaveLength(1);
  });

  it('isolates a throwing callback from onError', () => {
    installDom();
    const canvas = makeEl('canvas');
    makeEl('div').appendChild(canvas);
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const onError = vi.fn();
    const viewer = DocxViewer.fromDocument(canvas as unknown as HTMLCanvasElement, brokenEngine().asDoc(), {
      onPartFailure: () => { throw new Error('observer failed'); },
      onError,
    });
    expect(onError).not.toHaveBeenCalled();
    expect(consoleError).toHaveBeenCalled();
    viewer.destroy();
  });

  it('rejects an invalid static option at construction', () => {
    installDom();
    const canvas = makeEl('canvas');
    makeEl('div').appendChild(canvas);
    expect(() => new DocxViewer(canvas as unknown as HTMLCanvasElement, {
      errorPlaceholder: { title: 5 } as never,
    })).toThrow(TypeError);
  });
});

describe('DocxScrollViewer part failures (#1974)', () => {
  it('replays the failure for a borrowed document and empties on destroy', () => {
    installDom();
    const container = makeEl('div');
    const seen: unknown[] = [];
    const viewer = DocxScrollViewer.fromDocument(container as unknown as HTMLElement, brokenEngine().asDoc(), {
      onPartFailure: (failure) => seen.push(failure),
    });
    expect(seen).toEqual([FAILURE]);
    expect(viewer.partFailures).toEqual([FAILURE]);
    viewer.destroy();
    expect(viewer.partFailures).toEqual([]);
  });
});
