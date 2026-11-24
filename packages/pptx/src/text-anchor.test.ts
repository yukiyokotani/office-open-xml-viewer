import { describe, expect, it } from 'vitest';
import { layoutParagraph } from './renderer.js';
import { createPptxTextAnchor, resolvePptxTextAnchor, sourcePptxTextRun } from './text-anchor.js';
import type { PptxTextRunInfo } from './renderer.js';
import type { Paragraph, Slide, ShapeElement } from './types.js';

/** Source-backed runs with independently specified UTF-16 boundaries. */
function runs(): PptxTextRunInfo[] {
  return [
    { text: 'A𝄞', start: 0, end: 3 },
    { text: 'éZ', start: 3, end: 6 },
    { text: 'Next', start: 7, end: 11 },
  ].map(({ text, start, end }) => ({
    text, inShapeX: 0, inShapeY: 0, w: 30, h: 12, fontSize: 12, font: '12px serif',
    shapeX: 0, shapeY: 0, shapeW: 100, shapeH: 50, rotation: 0,
    source: {
      body: { slideIndex: 0, slidePart: 'ppt/slides/slide1.xml', shapeId: '2' },
      context: { text: 'A𝄞éZ\nNext' }, start, end,
    },
  }));
}

describe('durable PPTX text ranges', () => {
  it('rejects cross-body, missing, overlapping, mismatched and incomplete mappings', () => {
    const original = runs();
    const created = createPptxTextAnchor(original, [{ runIndex: 0, start: 1, end: 3 }]);
    expect(created.isSuccess).toBe(true);
    if (!created.isSuccess) return;
    for (const body of [
      { ...original[0].source!.body, shapeId: '3' },
      { ...original[0].source!.body, slideIndex: 1 },
      { ...original[0].source!.body, cell: { row: 0, column: 1 } },
    ]) {
      const other = { ...original[0], source: { ...original[0].source!, body } };
      expect(createPptxTextAnchor([original[0], other], [
        { runIndex: 0, start: 1, end: 3 }, { runIndex: 1, start: 0, end: 1 },
      ])).toEqual({ isSuccess: false, reason: 'cross-body' });
      expect(resolvePptxTextAnchor(created.anchor, [other])).toEqual({ isSuccess: false, reason: 'source-not-found' });
    }
    expect(resolvePptxTextAnchor(created.anchor, [...original, original[0]])).toEqual({ isSuccess: false, reason: 'ambiguous' });
    expect(resolvePptxTextAnchor({ ...created.anchor, exact: 'xx' }, original)).toEqual({ isSuccess: false, reason: 'text-mismatch' });
    expect(createPptxTextAnchor(original, [{ runIndex: 0, start: 2, end: 3 }])).toEqual({ isSuccess: false, reason: 'invalid-range' });
    expect(createPptxTextAnchor(original, [{ runIndex: 1, start: 1, end: 2 }])).toEqual({ isSuccess: false, reason: 'invalid-range' });
    expect(createPptxTextAnchor([{ ...original[0], source: undefined }], [{ runIndex: 0, start: 0, end: 1 }])).toEqual({ isSuccess: false, reason: 'unsupported' });
    expect(createPptxTextAnchor(original, [{ runIndex: 0, start: 0, end: 3 }, { runIndex: 2, start: 0, end: 4 }])).toEqual({ isSuccess: false, reason: 'incomplete' });
    expect(resolvePptxTextAnchor({ ...created.anchor, prefix: '' }, original)).toEqual({ isSuccess: false, reason: 'invalid-range' });
  });

  it('survives storage, worker cloning and rewrapping across paragraphs', () => {
    const original = runs();
    const created = createPptxTextAnchor(original, [
      { runIndex: 0, start: 1, end: 3 }, { runIndex: 1, start: 0, end: 3 }, { runIndex: 2, start: 0, end: 2 },
    ]);
    expect(created.isSuccess).toBe(true);
    if (!created.isSuccess) return;
    expect(created.anchor).toMatchObject({ start: 1, end: 9, exact: '𝄞éZ\nNe' });
    const rewrapped = structuredClone([
      { ...original[0], text: 'A𝄞éZ', source: { ...original[0].source!, end: 6 } }, original[2],
    ]);
    expect(resolvePptxTextAnchor(JSON.parse(JSON.stringify(created.anchor)), rewrapped)).toEqual({
      isSuccess: true, slices: [{ runIndex: 0, start: 1, end: 6 }, { runIndex: 1, start: 0, end: 2 }],
    });
  });

  it('allows native nonpainted gaps but rejects missing painted spaces and glyphs', () => {
    for (const gap of [' ', '\n', '\t']) {
      const text = `a${gap}b`;
      const full = { ...runs()[0], text, source: { ...runs()[0].source!, context: { text }, start: 0, end: 3 } };
      const created = createPptxTextAnchor([full], [{ runIndex: 0, start: 0, end: 3 }]);
      expect(created.isSuccess).toBe(true);
      if (!created.isSuccess) continue;
      const painted = [
        { ...full, text: 'a', source: { ...full.source, context: { text, nonpainted: [{ start: 1, end: 2 }] }, end: 1 } },
        { ...full, text: 'b', source: { ...full.source, start: 2 } },
      ];
      expect(resolvePptxTextAnchor(created.anchor, painted)).toEqual({
        isSuccess: true, slices: [{ runIndex: 0, start: 0, end: 1 }, { runIndex: 1, start: 0, end: 1 }],
      });
      expect(resolvePptxTextAnchor(created.anchor, [painted[0]])).toEqual({ isSuccess: false, reason: 'incomplete' });
      if (gap === ' ') {
        expect(resolvePptxTextAnchor(created.anchor, painted.map((run) => ({ ...run, source: { ...run.source, context: { text } } })))).toEqual({ isSuccess: false, reason: 'incomplete' });
      }
    }
  });

  it('retains native Unicode offsets through wrapping and formatting seams', () => {
    const paragraph = { marL: 0, marR: 0, indent: 0,
      runs: [{ type: 'text', text: 'A𝄞' }, { type: 'text', text: 'éZ', bold: true }],
    } as Paragraph;
    const ctx = { measureText: (text: string) => ({ width: [...text].length * 10 }) } as CanvasRenderingContext2D;
    expect(layoutParagraph(ctx, paragraph, 30, 12, '#000', 1 / 9525, 0)
      .flatMap((line) => line.segments).filter((segment) => segment.text)
      .map(({ text, sourceStart, sourceEnd }) => ({ text, sourceStart, sourceEnd }))).toEqual([
      { text: 'A𝄞', sourceStart: 0, sourceEnd: 3 }, { text: 'éZ', sourceStart: 3, sourceEnd: 6 },
    ]);
  });

  it('attaches only unique slide-owned source bodies, sharing context per render', () => {
    const element = { type: 'shape', id: '2', textBody: { paragraphs: [{ runs: [{ type: 'text', text: 'AB' }] }] } } as ShapeElement;
    const slide = { index: 0, partName: 'ppt/slides/slide1.xml', elements: [element], elementSources: [{ origin: 'slide' }] } as Slide;
    const run = { ...runs()[0], text: 'A', origin: 'slide' as const, elementIndex: 0, sourceStart: 0, sourceEnd: 1 };
    const cache = new Map();
    const first = sourcePptxTextRun(slide, element, run, cache);
    const second = sourcePptxTextRun(slide, element, { ...run, text: 'B', sourceStart: 1, sourceEnd: 2 }, cache);
    expect(first).toMatchObject({ body: { slideIndex: 0, shapeId: '2' }, start: 0, end: 1 });
    expect(second?.context).toBe(first?.context);
    expect(sourcePptxTextRun({ ...slide, elements: [element, element], elementSources: [{ origin: 'slide' }, { origin: 'slide' }] }, element, run, new Map())).toBeUndefined();
    expect(sourcePptxTextRun(slide, element, { ...run, origin: 'master' }, new Map())).toBeUndefined();
  });
});
