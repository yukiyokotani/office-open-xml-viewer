import { afterEach, describe, expect, it, vi } from 'vitest';
import { rasterizeMathSvg } from '@silurus/ooxml-core';
import type { HyperlinkTarget } from '@silurus/ooxml-core';
import { layoutParagraph, prepareSlideMath, renderTextBody, shapeTextRotation, type PptxTextRunInfo } from './renderer.js';
import { buildPptxTextLayer } from './text-layer.js';
import type { Paragraph, TextBody, TextRunData } from './types.js';

vi.mock('@silurus/ooxml-core', async (load) => ({
  ...await load<typeof import('@silurus/ooxml-core')>(),
  rasterizeMathSvg: vi.fn(async () => ({ source: {} })),
  tintMathRaster: () => ({}),
}));

// ECMA-376 §20.1.10.83 wordArtVert / wordArtVertRtl through the PowerPoint
// renderer: segments become stacked glyphs (core layoutStackedText) and each
// glyph class paints as the issue #1626 controls show.
const SCALE = 1 / 12700; // 1 pt → 1 px

interface Stroke { points: { x: number; y: number }[]; width: number; color: unknown }

interface Call { text: string; x: number; y: number; rot: number; tx: number; ty: number; align: string }

function mockCtx(): { ctx: CanvasRenderingContext2D; calls: Call[]; images: { x: number; y: number; w: number; h: number; rot: number }[]; strokes: Stroke[] } {
  let font = '24px serif';
  let textAlign: CanvasTextAlign = 'left';
  let textBaseline: CanvasTextBaseline = 'alphabetic';
  let rot = 0;
  let tx = 0;
  let ty = 0;
  const stack: [number, number, number][] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '24');
  const calls: Call[] = [];
  const images: { x: number; y: number; w: number; h: number; rot: number }[] = [];
  const strokes: Stroke[] = [];
  let points: Stroke['points'] = [];
  const point = (x: number, y: number) => points.push({
    x: tx + x * Math.cos(rot) - y * Math.sin(rot),
    y: ty + x * Math.sin(rot) + y * Math.cos(rot),
  });
  const ctx = {
    canvas: { style: {} },
    get font() { return font; }, set font(v: string) { font = v; },
    drawImage(_image: unknown, x: number, y: number, w: number, h: number) { images.push({ x, y, w, h, rot }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, letterSpacing: '0px', direction: 'ltr',
    get textAlign() { return textAlign; }, set textAlign(v: CanvasTextAlign) { textAlign = v; },
    get textBaseline() { return textBaseline; }, set textBaseline(v: CanvasTextBaseline) { textBaseline = v; },
    measureText: (s: string) => ({
      width: [...s].length * px() * 0.6,
      fontBoundingBoxAscent: px() * 0.9, fontBoundingBoxDescent: px() * 0.2,
      actualBoundingBoxAscent: px() * 0.7, actualBoundingBoxDescent: 0,
    }) as TextMetrics,
    fillText: (text: string, x: number, y: number) => calls.push({ text, x, y, rot, tx, ty, align: textAlign }),
    save: () => { stack.push([rot, tx, ty]); },
    restore: () => { const s = stack.pop(); if (s) [rot, tx, ty] = s; },
    translate: (x: number, y: number) => { tx += x; ty += y; },
    rotate: (a: number) => { rot += a; },
    scale: () => {}, beginPath: () => { points = []; }, moveTo: point, lineTo: point,
    stroke: () => strokes.push({ points: [...points], width: ctx.lineWidth, color: ctx.strokeStyle }),
    fill: () => {}, clip: () => {}, rect: () => {}, fillRect: () => {}, setLineDash: () => {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls, images, strokes };
}

function body(text: string, vert: string, fontFamilyEa?: string, hyperlink?: string): TextBody {
  const run = {
    type: 'text', text, bold: null, italic: null, underline: false, strikethrough: false,
    fontSize: 24, color: '000000', fontFamily: 'Arial', fontFamilyEa, hyperlink,
  } as TextRunData;
  const para = {
    alignment: 'l', marL: 0, marR: 0, indent: 0, spaceBefore: null, spaceAfter: null, spaceLine: null,
    lvl: 0, bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops: [], eaLnBrk: true, runs: [run],
  } as Paragraph;
  return {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: 24, defaultBold: null, defaultItalic: null,
    lIns: 91440, rIns: 91440, tIns: 45720, bIns: 45720, wrap: 'square', vert, autoFit: 'none',
  };
}

interface FakeEl {
  textContent: string; innerHTML: string; title: string; dataset: Record<string, string>;
  style: Record<string, string>; children: FakeEl[]; listeners: ((e: { preventDefault(): void }) => void)[];
  appendChild(c: FakeEl): void; setAttribute(): void; addEventListener(t: string, f: (e: { preventDefault(): void }) => void): void; click(): void;
}
function fakeEl(): FakeEl {
  const el: FakeEl = {
    textContent: '', innerHTML: '', title: '', dataset: {}, style: {}, children: [], listeners: [],
    appendChild(c) { el.children.push(c); },
    setAttribute() {},
    addEventListener(t, f) { if (t === 'click') el.listeners.push(f); },
    click() { for (const f of el.listeners) f({ preventDefault() {} }); },
  };
  return el;
}
afterEach(() => { vi.unstubAllGlobals(); vi.restoreAllMocks(); });

// Arial 24 pt: cell 7/6 × 1.1172 em × 24 = 31.28, descent 0.2119 em.
const CELL = (7 / 6) * (2288 / 2048) * 24;

describe('pptx stacked vertical text (wordArtVert / wordArtVertRtl)', () => {
  it('stacks upright Latin glyphs on the column axis, one cell apart', () => {
    const { ctx, calls } = mockCtx();
    renderTextBody(ctx, body('AB', 'wordArtVert'), 0, 0, 60, 470, SCALE);
    const [a, b] = calls;
    expect(a).toMatchObject({ text: 'A', align: 'center', rot: 0 });
    expect(a.x).toBeCloseTo(7.2 + CELL / 2, 6);
    expect(a.y).toBeCloseTo(3.6 + CELL - (434 / 2048) * 24, 6);
    expect(b.y - a.y).toBeCloseTo(CELL, 6);
    // wordArtVertRtl starts at the right edge.
    const r = mockCtx();
    renderTextBody(r.ctx, body('AB', 'wordArtVertRtl'), 0, 0, 60, 470, SCALE);
    expect(r.calls[0].x).toBeCloseTo(60 - 7.2 - CELL / 2, 6);
  });

  it('uses the font-table cell for an exceptional drawing face and its Latin column face', () => {
    const { ctx, calls } = mockCtx();
    const b = body('日日', 'wordArtVert', 'Yu Gothic');
    renderTextBody(ctx, b, 0, 0, 100, 470, SCALE);
    const cell = (7 / 6) * (2636 / 2048) * 24; // OS/2 usWin box, not a fitted pitch.
    expect(calls[1].y - calls[0].y).toBeCloseTo(cell, 6);
    const latin = mockCtx();
    const mixed = body('日日', 'wordArtVert', 'MS Mincho');
    (mixed.paragraphs[0].runs[0] as TextRunData).fontFamily = 'Yu Gothic';
    renderTextBody(latin.ctx, mixed, 0, 0, 100, 470, SCALE);
    expect(latin.calls[0].x).toBeCloseTo(7.2 + cell / 2, 6);
  });

  it('turns ASCII brackets 90° and draws CJK brackets with the face vertical glyph', () => {
    const { ctx, calls } = mockCtx();
    renderTextBody(ctx, body('(A「', 'wordArtVert', 'Yu Gothic'), 0, 0, 60, 470, SCALE);
    const paren = calls.find((c) => c.text === '(');
    expect(paren?.rot).toBeCloseTo(Math.PI / 2, 6);
    // 「 is drawn upright as its vertical presentation form ﹁ (U+FE41).
    const bracket = calls.find((c) => c.text === '﹁');
    expect(bracket?.rot).toBe(0);
    expect(calls.some((c) => c.text === '「')).toBe(false);
  });

  // #1827 PowerPoint controls: vrt2-only ∥ (Yu Gothic) stands unturned in a
  // 1 em cell, left-aligned with the （ form; MS Gothic's vert ∥ is turned.
  it('draws a vrt2-only ∥ unturned in the vertical-form frame and a vert ∥ turned', () => {
    const face = (text: string, family: string) => {
      const b = body(text, 'wordArtVert', family);
      (b.paragraphs[0].runs[0] as TextRunData).fontFamily = family;
      return b;
    };
    const { ctx, calls } = mockCtx();
    const runs: PptxTextRunInfo[] = [];
    renderTextBody(ctx, face('∥（', 'Yu Gothic'), 0, 0, 60, 470, SCALE,
      null, 0, false, false, '#000000', undefined, undefined, (run) => runs.push(run));
    const bar = calls.find((c) => c.text === '∥')!;
    const form = calls.find((c) => c.text === '︵')!;
    expect(bar).toMatchObject({ rot: 0, align: 'left' });
    expect(bar.x).toBeCloseTo(form.x, 6);
    expect(form.y - bar.y).toBeCloseTo(24, 6);
    expect(runs[0].h).toBeCloseTo(24, 6);
    const ms = mockCtx();
    renderTextBody(ms.ctx, face('∥', 'MS Gothic'), 0, 0, 60, 470, SCALE);
    expect(ms.calls[0].rot).toBeCloseTo(Math.PI / 2, 6);
  });

  it.each([false, true])('uses bold-specific vert coverage for the drawing face (italic=%s)', (italic) => {
    const advances = [false, true].map((bold) => {
      const { ctx } = mockCtx();
      const b = body('ーA', 'wordArtVert', 'Microsoft JhengHei');
      (b.paragraphs[0].runs[0] as TextRunData).bold = bold;
      (b.paragraphs[0].runs[0] as TextRunData).italic = italic;
      const runs: PptxTextRunInfo[] = [];
      renderTextBody(ctx, b, 0, 0, 100, 470, SCALE,
        null, 0, false, false, '#000000', undefined, undefined, (run) => runs.push(run));
      return runs[0].h;
    });
    // Regular has a vert glyph (one em); bold rotates the horizontal glyph.
    expect(advances).toEqual([24, 24 * 0.6]);
  });

  it('stacks a grapheme cluster in one cell', () => {
    const { ctx, calls } = mockCtx();
    renderTextBody(ctx, body('e\u0301B', 'wordArtVert'), 0, 0, 60, 470, SCALE);
    expect(calls.map((c) => c.text)).toEqual(['e\u0301', 'B']);
    expect(calls[1].y - calls[0].y).toBeCloseTo(CELL, 6);
  });

  it('hands each glyph its run hyperlink so the text layer installs link handlers', () => {
    const { ctx } = mockCtx();
    const runs: PptxTextRunInfo[] = [];
    renderTextBody(ctx, body('AB', 'wordArtVert', undefined, 'https://example.com/'), 0, 0, 60, 470, SCALE,
      null, 0, false, false, '#000000', undefined, undefined, (r) => runs.push(r));
    const target: HyperlinkTarget = { kind: 'external', url: 'https://example.com/' };
    expect(runs.map((r) => r.hyperlink)).toEqual([target, target]);
    vi.stubGlobal('document', { createElement: () => fakeEl() });
    const layer = fakeEl();
    const onClick = vi.fn<(t: HyperlinkTarget) => void>();
    buildPptxTextLayer(layer as unknown as HTMLDivElement, runs, 960, 540, onClick);
    const spans = layer.children.flatMap((group) => group.children);
    expect(spans).toHaveLength(2);
    spans[0].click();
    expect(onClick).toHaveBeenCalledWith(target);
  });

  it('keeps a stacked body upright under flipH and turns it 180° under flipV', () => {
    expect(shapeTextRotation('wordArtVert', 0, true, false)).toBe(0);
    expect(shapeTextRotation('wordArtVertRtl', 0, false, true)).toBe(180);
    expect(shapeTextRotation('horz', 0, false, true)).toBe(0);
  });
});

// A grapheme extender (variation selector, combining mark) follows its base's
// font slot, so no cluster straddles two segments. Arial latin, Yu Gothic ea.
describe('font slots keep grapheme clusters whole', () => {
  const IVS = '\u{E0100}'; // ideographic variation selector (Latin slot by itself)
  it('stacked: a CJK base and its selector share one Yu Gothic cell', () => {
    const { ctx, calls } = mockCtx();
    renderTextBody(ctx, body(`葛${IVS}B`, 'wordArtVert', 'Yu Gothic'), 0, 0, 60, 470, SCALE);
    expect(calls.map((c) => c.text)).toEqual([`葛${IVS}`, 'B']);
  });

  it('stacked: a bracket with a selector keeps its glyphs (no lone-character substitution)', () => {
    const { ctx, calls } = mockCtx();
    renderTextBody(ctx, body('「\uFE0F', 'wordArtVert', 'Yu Gothic'), 0, 0, 60, 470, SCALE);
    expect(calls.map((c) => c.text)).toEqual(['「\uFE0F']);
  });

  it('horizontal: the cluster stays in the base segment and font', () => {
    const { ctx } = mockCtx();
    const para = body(`葛${IVS}B e\u0301`, 'horz', 'Yu Gothic').paragraphs[0];
    const [line] = layoutParagraph(ctx, para, 10_000, 24, '#000', SCALE, 0);
    const segs = line.segments.filter((g) => g.text);
    expect(segs.map((g) => g.text)).toEqual([`葛${IVS}`, 'B e\u0301']);
    expect(segs[0].font).toContain('Yu Gothic');
    expect(segs[1].font).not.toContain('Yu Gothic');
  });
});

// A cluster whose extender opens the next authored run: clusters are
// segmented over the paragraph text, and the carried extender takes the base
// run's formatting (the second run here is bold).
describe('grapheme clusters across run seams', () => {
  const SEAMS: [string, string][] = [['葛', '\u{E0100}'], ['「', '\uFE0F'], ['가', '\u11A8']];
  const split = (vert: string, base: string, ext: string): TextBody => {
    const b = body(`${base}${ext}B`, vert, 'Yu Gothic');
    const r = b.paragraphs[0].runs[0] as TextRunData;
    b.paragraphs[0].runs = [{ ...r, text: base }, { ...r, text: `${ext}B`, bold: true }];
    return b;
  };
  for (const [base, ext] of SEAMS) {
    it(`stacked: ${JSON.stringify(base + ext)} split across runs stays one cell`, () => {
      for (const vert of ['wordArtVert', 'wordArtVertRtl']) {
        const { ctx, calls } = mockCtx();
        renderTextBody(ctx, split(vert, base, ext), 0, 0, 100, 470, SCALE);
        expect(calls.map((c) => c.text)).toEqual([base + ext, 'B']);
      }
    });
    it(`horizontal: ${JSON.stringify(base + ext)} split across runs stays in the base segment`, () => {
      const { ctx } = mockCtx();
      const [line] = layoutParagraph(ctx, split('horz', base, ext).paragraphs[0], 10_000, 24, '#000', SCALE, 0);
      const segs = line.segments.filter((g) => g.text);
      expect(segs.map((g) => g.text)).toEqual([base + ext, 'B']);
      expect(segs[0].font).toContain('Yu Gothic');
      expect(segs[0].font).not.toContain('bold');
      expect(segs[1].font).toContain('bold');
    });
  }
});

// One base followed by 80,000 runs that hold only a combining mark: cluster
// segmentation must stay linear in the paragraph length (a rescan of the
// growing cluster per run took about 19 s here; main lays the runs out in
// about 0.35 s).
describe('cluster segmentation cost', () => {
  it('handles 80,000 extender-only runs in linear time, horizontal and stacked', () => {
    const b = body('a', 'wordArtVert', 'Yu Gothic');
    const r = b.paragraphs[0].runs[0] as TextRunData;
    b.paragraphs[0].runs = [{ ...r, text: 'a' }, ...Array.from({ length: 80_000 }, () => ({ ...r, text: '\u0301' })), { ...r, text: 'B' }];
    const start = performance.now();
    const lines = layoutParagraph(mockCtx().ctx, b.paragraphs[0], 1e9, 24, '#000', SCALE, 0);
    expect(lines).toHaveLength(1);
    const [line] = lines;
    const segs = line.segments.filter((g) => g.text);
    expect(segs.map((g) => g.text).join('')).toBe(`a${'\u0301'.repeat(80_000)}B`);
    const { ctx, calls } = mockCtx();
    renderTextBody(ctx, b, 0, 0, 100, 470, SCALE);
    expect(calls.map((c) => c.text.length)).toEqual([80_001, 1]);
    expect(performance.now() - start).toBeLessThan(5_000);
  }, 60_000);
});

// Compare cached and unavailable OMML geometry with the horizontal host rule.
it.each(['cached', 'no engine', 'conversion', 'rasterization'] as const)('%s preserves horizontal equation geometry in both stacked modes', async (failure) => {
  const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
  const error = vi.spyOn(console, 'error').mockImplementation(() => {});
  if (failure === 'rasterization') vi.mocked(rasterizeMathSvg).mockRejectedValue(new Error('rasterization failed'));
  try {
    for (const display of [false, true]) {
      const outcomes = [];
      let horizontalImage: { w: number; h: number } | undefined;
      for (const vert of ['horz', 'wordArtVert', 'wordArtVertRtl']) {
        warn.mockClear(); error.mockClear();
        const equation = { type: 'math' as const, nodes: [{ kind: 'run' as const, text: 'x', style: 'italic' as const }], display, fontSize: 18 };
        const engine = {
          loadMathJax: async () => {},
          mathMLToSvg: async () => {
            if (failure === 'conversion') throw new Error('conversion failed');
            return { svg: '<svg/>', widthEm: 3, ascentEm: 1.5, descentEm: .5 };
          },
        };
        const b = body('A', vert);
        const text = b.paragraphs[0].runs[0] as TextRunData;
        b.paragraphs[0].runs = [text, equation, { ...text, text: 'B' }];
        if (failure !== 'no engine') {
          await prepareSlideMath({ elements: [{ type: 'shape', textBody: b }] } as unknown as import('./types.js').Slide, engine);
        }
        const { ctx, calls, images } = mockCtx();
        renderTextBody(ctx, b, 0, 0, 300, 300, SCALE);
        if (failure === 'cached') {
          expect(images).toHaveLength(1);
          const image = images[0];
          expect(image).toMatchObject({ w: 54, h: 36, rot: 0 });
          if (vert === 'horz') {
            horizontalImage = image;
            if (!display) {
              expect(image.x - calls[0].x).toBeCloseTo(14.4);
              expect(calls[1].x - image.x).toBeCloseTo(image.w);
              expect(calls[1].y).toBeCloseTo(calls[0].y);
            } else {
              expect(image.y).toBeGreaterThan(calls[0].y);
              expect(calls[1].y).toBeGreaterThan(image.y + image.h);
            }
          } else {
            expect(image.w).toBe(horizontalImage!.w);
            expect(image.h).toBe(horizontalImage!.h);
            const sign = vert === 'wordArtVert' ? 1 : -1;
            expect(calls[1].x - calls[0].x).toBeCloseTo(display ? sign * (CELL + image.w) : 0);
            expect(calls[1].y - calls[0].y).toBeCloseTo(display ? 0 : CELL + image.h);
            expect(image.x + image.w / 2 - calls[0].x).toBeCloseTo(display ? sign * (CELL + image.w) / 2 : 0);
            expect(image.y).toBeCloseTo(3.6 + (display ? 0 : CELL));
          }
        } else if (failure === 'no engine') {
          // Explicit engine opt-out preserves the existing empty geometry.
          // Horizontal reserves a blank display line, but no inline advance.
          // Its stacked counterpart uses the mark cell for that blank column.
          const control = mockCtx();
          const blank = { ...b, paragraphs: [{ ...b.paragraphs[0], runs: [text,
            ...(display ? [{ type: 'break' as const }, { type: 'break' as const }] : []),
            { ...text, text: 'B' }] }] };
          renderTextBody(control.ctx, blank, 0, 0, 300, 300, SCALE);
          if (vert !== 'horz') {
            expect(calls).toEqual(control.calls);
            expect(calls[1].x - calls[0].x).toBeCloseTo(display
              ? (vert === 'wordArtVert' ? 2 * CELL : -2 * CELL) : 0);
            expect(calls[1].y - calls[0].y).toBeCloseTo(display ? 0 : CELL);
          } else {
            const lines = layoutParagraph(ctx, b.paragraphs[0], 300, 24, '#000', SCALE, 0);
            expect(lines).toHaveLength(display ? 3 : 1);
            expect(calls[1].x - calls[0].x).toBeCloseTo(display ? 0 : 14.4);
            // Mock font box = 1.1 em; horizontal default line leading = -1.2 px.
            expect(calls[1].y - calls[0].y).toBeCloseTo(display ? 2 * (24 * 1.1 - 1.2) : 0);
          }
        }
        if (failure === 'conversion' || failure === 'rasterization') {
          expect(images).toHaveLength(0);
          expect(calls.map(c => c.text)).toEqual(['A', 'Equation unavailable', 'B']);
          expect(calls[2].x !== calls[0].x || calls[2].y !== calls[0].y).toBe(true);
        }
        const outcome = { text: calls.map((c) => c.text).join(''), images: images.length,
          warnings: [...warn.mock.calls], errors: [...error.mock.calls] };
        expect(outcome).toEqual({ text: failure === 'conversion' || failure === 'rasterization' ? 'AEquation unavailableB' : 'AB', images: failure === 'cached' ? 1 : 0, warnings: [], errors: [] });
        outcomes.push(outcome);
      }
      expect(outcomes.slice(1)).toEqual([outcomes[0], outcomes[0]]);
    }
  } finally {
    vi.mocked(rasterizeMathSvg).mockResolvedValue({ source: {} as CanvasImageSource, widthPx: 3, heightPx: 2 });
  }
});


// Office-export controls: both column progressions keep underlines on the left;
// strike strokes follow upright glyphs, and markers consume full stacked cells.
it.each(['wordArtVert', 'wordArtVertRtl'])('paints glyph-sized horizontal strikes in %s', (vert) => {
  const b = body('A', vert);
  const r = b.paragraphs[0].runs[0] as TextRunData;
  Object.assign(r, { strikethrough: true, strikeDouble: true });
  b.paragraphs[0].runs.push({ ...r, text: 'B', fontSize: 32 });
  const { ctx, calls, strokes } = mockCtx();
  const measure = ctx.measureText.bind(ctx);
  ctx.measureText = text => text === ' ' ? { ...measure(text), width: 8 } : measure(text);
  renderTextBody(ctx, b, 0, 0, 200, 200, SCALE);
  expect(strokes).toHaveLength(3); // two glyphs and the invisible paragraph mark
  expect(strokes[2].points[1].x - strokes[2].points[0].x).toBeCloseTo(8);
  expect(strokes[2].points[0].y - strokes[1].points[0].y).toBeCloseTo(CELL * 32 / 24, 0);
  for (const [i, stroke] of strokes.slice(0, 2).entries()) {
    expect(stroke.points).toHaveLength(4);
    for (const pair of [stroke.points.slice(0, 2), stroke.points.slice(2)]) {
      expect(pair[0].y).toBeCloseTo(pair[1].y);
      expect(pair[1].x - pair[0].x).toBeCloseTo((i === 0 ? 24 : 32) * .6);
      expect((pair[0].x + pair[1].x) / 2).toBeCloseTo(calls[i].x);
      expect(pair[0].y).toBeLessThan(calls[i].y);
      expect(pair[0].y).toBeGreaterThan(calls[i].y - (i === 0 ? 24 : 32) * .5);
    }
  }
});

it('keeps strikes in a rotated glyph frame and a vertical presentation glyph frame', () => {
  const b = body('(A「', 'wordArtVert', 'Yu Gothic');
  Object.assign(b.paragraphs[0].runs[0], { strikethrough: true });
  const { ctx, strokes } = mockCtx();
  renderTextBody(ctx, b, 0, 0, 200, 200, SCALE);
  expect(strokes).toHaveLength(4);
  expect(strokes[0].points[0].x).toBeCloseTo(strokes[0].points[1].x);
  expect(strokes[0].points[1].y - strokes[0].points[0].y).toBeCloseTo(24 * .6);
  for (const stroke of strokes.slice(1)) expect(stroke.points[0].y).toBeCloseTo(stroke.points[1].y);
});

it('honors an explicit unstruck paragraph insertion style', () => {
  const b = body('A', 'wordArtVert');
  const run = b.paragraphs[0].runs[0] as TextRunData;
  // Parser JSON carries run properties only: no `type` discriminant here.
  const { type: _type, ...properties } = run;
  b.paragraphs[0].endRunProperties = { ...properties, text: '' };
  run.strikethrough = true;
  const { ctx, strokes } = mockCtx();
  renderTextBody(ctx, b, 0, 0, 200, 200, SCALE);
  expect(strokes).toHaveLength(1);
});

it.each(['wordArtVert', 'wordArtVertRtl'])('uses one underline frame per mixed-size wrapped %s column', (vert) => {
  for (const underlineStyle of ['sng', 'dbl', 'wavy']) {
    const b = body('A', vert);
    const r = b.paragraphs[0].runs[0] as TextRunData;
    Object.assign(r, { underline: true, underlineStyle, underlineColor: 'FF0000' });
    b.paragraphs[0].runs.push({ ...r, text: 'B', fontSize: 32 }, { ...r, text: 'CD', fontSize: 18 });
    const { ctx, calls, strokes } = mockCtx();
    // First column A/B/C, second column D.
    renderTextBody(ctx, b, 0, 0, 200, 110, SCALE);
    expect(calls[0].x).toBe(calls[1].x);
    expect(calls[1].x).toBe(calls[2].x);
    expect(calls[3].x).not.toBe(calls[0].x);
    expect(strokes).toHaveLength(2);
    expect(strokes.map(s => s.width)).toEqual([32 * .05, 1]);
    expect(strokes.every(s => s.color === 'rgba(255,0,0,1)')).toBe(true);
    for (const [i, stroke] of strokes.entries()) {
      const axis = calls[i === 0 ? 0 : 3].x;
      expect(stroke.points.every(p => p.x < axis)).toBe(true);
    }
    // Include the stacked cell's extra leading in the rotated baseline.
    // A raw (ascent-descent)/2 baseline is too far right by half that leading.
    if (underlineStyle === 'sng') {
      const columnLeft = calls[0].x - CELL * (32 / 24) / 2;
      expect(strokes[0].points[0].x).toBeCloseTo(columnLeft + (434 / 2048) * 32 - 2);
    }
  }
});

it('places underline from the drawing face descent and full cell rather than a font-specific offset', () => {
  const b = body('A', 'wordArtVert');
  Object.assign(b.paragraphs[0].runs[0], { fontFamily: 'Yu Gothic', underline: true });
  const { ctx, strokes } = mockCtx();
  renderTextBody(ctx, b, 0, 0, 200, 200, SCALE);
  // Yu Gothic's usWin descent is 619/2048 em. A 24px single underline
  // uses the same 1.2px weight and 2.5px snapped offset as horizontal text.
  expect(strokes).toHaveLength(1);
  expect(strokes[0].points[0].x).toBeCloseTo(7.2 + (619 / 2048) * 24 - 2.5);
});

it.each(['wordArtVert', 'wordArtVertRtl'])('reserves stacked character and number prefixes only at paragraph start in %s', (vert) => {
  for (const bullet of [
    { type: 'char' as const, char: 'R', fontFamily: 'Arial', color: null, sizePct: null },
    { type: 'autoNum' as const, numType: 'arabicPeriod', startAt: 12, color: null },
  ]) {
    const b = body('ABCDEFGHIJKLMNO', vert);
    const para = b.paragraphs[0];
    para.bullet = bullet;
    b.paragraphs.push({ ...para, runs: [{ ...(para.runs[0] as TextRunData), text: 'X' }] });
    const { ctx, calls } = mockCtx();
    renderTextBody(ctx, b, 0, 0, 300, 250, SCALE);
    const prefix = bullet.type === 'char' ? ['R'] : ['1', '2', '.'];
    const firstColumn = calls.filter(c => c.x === calls.find(c => c.text === 'A')!.x);
    expect(firstColumn.slice(0, prefix.length).map(c => c.text)).toEqual(prefix);
    const a = calls.find(c => c.text === 'A')!;
    expect(a.y - firstColumn[0].y).toBeCloseTo(CELL * prefix.length);
    expect(firstColumn.every(c => c.rot === 0 && c.align === 'center')).toBe(true);
    expect(calls.filter(c => c.text === prefix[0])).toHaveLength(2);
    if (bullet.type === 'autoNum') expect(calls.map(c => c.text).join('')).toContain('13.X');
    const continuation = calls.find(c => /^[B-O]$/.test(c.text) && c.x !== a.x)!;
    expect(continuation.y).toBeCloseTo(firstColumn[0].y);
    const indented = mockCtx();
    Object.assign(para, { marL: 36 * 12700, indent: -24 * 12700 });
    renderTextBody(indented.ctx, b, 0, 0, 300, 250, SCALE);
    expect(indented.calls.find(c => c.text === prefix[0])!.y - firstColumn[0].y).toBeCloseTo(12);
    expect(indented.calls.find(c => c.text === 'A')!.y - a.y).toBeCloseTo(12);
  }
});

it('keeps picture markers per paragraph with authored size and source aspect', () => {
  const b = body('A', 'wordArtVertRtl');
  const para = b.paragraphs[0];
  Object.assign(para, { bullet: { type: 'blip', imagePath: 'marker.png', mimeType: 'image/png', sizePct: 50 } });
  const second = { ...para, runs: [{ ...(para.runs[0] as TextRunData), text: 'B' }] };
  Object.assign(second, { bullet: { type: 'blip', imagePath: 'marker.png', mimeType: 'image/png', sizePct: 50, sizePts: 10 } });
  b.paragraphs.push(second);
  const { ctx, images } = mockCtx();
  renderTextBody(ctx, b, 0, 0, 200, 200, SCALE, null, 0, false, false, '#000000', undefined, {
    themeMajorFont: null, themeMinorFont: null, dpr: 1,
    pictureBulletImages: new Map([['marker.png', { width: 36, height: 24 } as ImageBitmap]]),
  });
  expect(images.map(({ w, h, rot }) => ({ w, h, rot }))).toEqual([
    { w: 18, h: 12, rot: 0 }, { w: 15, h: 10, rot: 0 },
  ]);
});

it.each(['horz', 'wordArtVert', 'wordArtVertRtl'])('replaces a complete inline equation branch in %s', async (vert) => {
  const b = body('', vert);
  b.paragraphs[0].runs = ['x', 'y'].map(text => ({ type: 'math',
    nodes: [{ kind: 'run', text, style: 'roman' }], display: false, fontSize: 18 }));
  b.paragraphs[0].mathFallbacks = [{ start: 0, end: 2, text: 'authored preview' }];
  await prepareSlideMath({ elements: [{ type: 'shape', textBody: b }] } as unknown as import('./types.js').Slide,
    { loadMathJax: async () => {}, mathMLToSvg: async () => { throw new Error('conversion'); } });
  const { ctx, calls } = mockCtx();
  renderTextBody(ctx, b, 0, 0, 200, 200, SCALE);
  expect(calls.map(c => c.text)).toEqual(['authored preview']);
});
