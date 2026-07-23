import { describe, it, expect } from 'vitest';
import { paintHighlight } from './renderer.js';
// layoutParagraph is module-internal (not re-exported from index.ts), exported
// only so this regression can drive the real segmentation path.
import { layoutParagraph } from './renderer.js';
import type { Paragraph } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// --- A mock 2D context that records fillRect / fillText / fillStyle in order.
// layoutParagraph and paintHighlight only need measureText, font, fillStyle,
// fillRect and fillText, so a tiny stub suffices (same approach as
// tabular-text.test.ts). Every glyph advances 10px so widths are predictable.
function mockCtx() {
  const ops: Array<
    | { op: 'rect'; x: number; y: number; w: number; h: number; style: string }
    | { op: 'text'; t: string; x: number; style: string }
  > = [];
  let fillStyle = '';
  const ctx = {
    font: '',
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(v: string) {
      fillStyle = v;
    },
    measureText: (s: string) => ({ width: s.length * 10 }),
    fillRect: (x: number, y: number, w: number, h: number) =>
      ops.push({ op: 'rect', x, y, w, h, style: fillStyle }),
    fillText: (t: string, x: number) => ops.push({ op: 'text', t, x, style: fillStyle }),
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, ops };
}

// --- Minimal but type-complete fixtures -------------------------------------
function run(text: string, over: Partial<TextRunData> = {}): TextRunData {
  return {
    type: 'text',
    text,
    bold: null,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize: 20,
    color: '000000',
    fontFamily: 'Arial',
    ...over,
  };
}

function para(runs: TextRunData[]): Paragraph {
  return {
    alignment: 'l',
    marL: 0,
    marR: 0,
    indent: 0,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl: 0,
    bullet: { type: 'none' },
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs,
  } as Paragraph;
}

describe('layoutParagraph — highlight is part of segment identity (sameMeta)', () => {
  // ECMA-376 §21.1.2.3.4. Adjacent runs that are identical except for the
  // highlight colour must NOT be coalesced, or one run's marker would bleed
  // over the other. sameMeta() includes the resolved highlight, so the two
  // runs stay as two segments each carrying its own colour.
  it('splits adjacent runs whose highlight differs', () => {
    const { ctx } = mockCtx();
    const lines = layoutParagraph(
      ctx,
      para([run('AAA', { highlight: 'FFFF00' }), run('BBB', { highlight: '00FF00' })]),
      10_000, // wide enough that everything fits on one line
      20,
      '#000000',
      1,
      0,
    );
    expect(lines).toHaveLength(1);
    const segs = lines[0].segments;
    expect(segs.map((s) => s.text)).toEqual(['AAA', 'BBB']);
    // Resolved to rgba() by hexToRgba; the two markers keep distinct colours.
    expect(segs[0].highlight).toBe('rgba(255,255,0,1)');
    expect(segs[1].highlight).toBe('rgba(0,255,0,1)');
  });

  it('coalesces adjacent runs that share the same highlight', () => {
    const { ctx } = mockCtx();
    const lines = layoutParagraph(
      ctx,
      para([run('AAA', { highlight: 'FFFF00' }), run('BBB', { highlight: 'FFFF00' })]),
      10_000,
      20,
      '#000000',
      1,
      0,
    );
    const segs = lines[0].segments;
    expect(segs).toHaveLength(1);
    expect(segs[0].text).toBe('AAABBB');
    expect(segs[0].highlight).toBe('rgba(255,255,0,1)');
  });

  it('splits a highlighted run from an unhighlighted neighbour', () => {
    const { ctx } = mockCtx();
    const lines = layoutParagraph(
      ctx,
      para([run('AAA', { highlight: 'FFFF00' }), run('BBB')]),
      10_000,
      20,
      '#000000',
      1,
      0,
    );
    const segs = lines[0].segments;
    expect(segs).toHaveLength(2);
    expect(segs[0].highlight).toBe('rgba(255,255,0,1)');
    expect(segs[1].highlight).toBeUndefined();
  });
});

describe('paintHighlight — paints the marker box behind the glyphs', () => {
  // The marker rectangle must be filled BEFORE the glyphs and must restore the
  // glyph colour so the following fillText draws the text, not the marker.
  it('fills one rect with the highlight colour then restores the glyph colour', () => {
    const { ctx, ops } = mockCtx();
    paintHighlight(ctx, 5, 100, 30, 20, 'rgba(255,255,0,1)', 'rgba(0,0,0,1)');

    expect(ops).toHaveLength(1);
    const rect = ops[0];
    expect(rect.op).toBe('rect');
    if (rect.op === 'rect') {
      expect(rect.style).toBe('rgba(255,255,0,1)'); // box drawn in highlight colour
      expect(rect.x).toBe(5);
      expect(rect.w).toBe(30);
      // Vertical band from highlightBox: top = baseline − 0.85·em, height = 1.1·em.
      expect(rect.y).toBeCloseTo(100 - 20 * 0.85, 6);
      expect(rect.h).toBeCloseTo(20 * 1.1, 6);
    }
    // fillStyle is left on the glyph colour for the caller's fillText.
    expect(ctx.fillStyle).toBe('rgba(0,0,0,1)');
  });

  it('draws the box under a glyph painted afterwards (background then glyph)', () => {
    const { ctx, ops } = mockCtx();
    // Simulate the renderer's call order: paint the box, then the glyph.
    paintHighlight(ctx, 0, 100, 30, 20, 'rgba(255,255,0,1)', 'rgba(0,0,0,1)');
    ctx.fillText('AAA', 0, 100);

    expect(ops.map((o) => o.op)).toEqual(['rect', 'text']);
    const [box, glyph] = ops;
    expect(box.style).toBe('rgba(255,255,0,1)');
    expect(glyph.style).toBe('rgba(0,0,0,1)'); // glyph on top, in text colour
  });
});
