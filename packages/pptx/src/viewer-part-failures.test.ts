import { describe, expect, it } from 'vitest';
import { PptxScrollViewer } from './scroll-viewer.js';
import { installDom, makeEl, FakePptxEngine } from './scroll-viewer-test-dom.js';

/** #1974 — viewer delivery of PPTX slide failures (replay + live). */
const failure = (slideIndex: number) => ({
  kind: 'slide', slideIndex, slideNumber: slideIndex + 1, message: `ppt/slides/slide${slideIndex + 1}.xml: bad`,
}) as const;

describe('PptxScrollViewer part failures (#1974)', () => {
  it('replays known failures, delivers later ones once, and stops after destroy', () => {
    installDom();
    const engine = new FakePptxEngine(3, 9_144_000, 6_858_000);
    engine._partFailureLog.add(failure(0));
    const seen: number[] = [];
    const viewer = PptxScrollViewer.fromPresentation(makeEl('div') as unknown as HTMLElement, engine.asPres(), {
      onPartFailure: (f) => seen.push(f.slideIndex),
    });
    expect(seen).toEqual([0]);
    engine._partFailureLog.add(failure(2));
    expect(seen).toEqual([0, 2]);
    expect(viewer.partFailures.map((f) => f.slideIndex)).toEqual([0, 2]);
    viewer.destroy();
    engine._partFailureLog.add(failure(1));
    expect(seen).toEqual([0, 2]);
    expect(viewer.partFailures).toEqual([]);
  });
});
