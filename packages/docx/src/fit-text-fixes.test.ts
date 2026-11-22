import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  buildSegments,
  layoutLines,
  type LayoutSeg,
  type LayoutTextSeg,
} from './line-layout.js';
import { renderDocumentToCanvas } from './renderer.js';
import type { DocxTextRunInfo } from './renderer.js';
import { buildDocxTextLayer } from './text-layer.js';
import { buildDocxHighlightLayer, type DocxHighlightMatch } from './find-highlight-layer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ─────────────────────────────────────────────────────────────────────────────
// ECMA-376 §17.3.2.14 `<w:fitText>` — review-fix contracts (PR #918 follow-up).
//
// Each block pins one adjudicated finding of the fitText adversarial review:
//   1. a SINGLE-character region still occupies exactly `w:val` (the glyph sits
//      at the region start at its natural width; the cell pads to the target —
//      glyph stretching is deferred until compression ground truth exists),
//   2. a TAB inside a fitText run splits the region at the tab boundary (a tab
//      advance is not a glyph advance): each text fragment is its own region,
//   3. fitted geometry reaches the selection/find overlays: `onTextRun` reports
//      the drawn per-glyph letter-spacing, the text layer applies it to the
//      selection span, and the find-highlight layer offsets slice extents by it,
//   4. small-caps case transforms that CHANGE the code-point count (ß → SS)
//      count the EMITTED text, keeping the region advance pinned to `w:val`,
//   7. a fit region inside a justified (§17.18.44) line contributes NO internal
//      gaps to the paragraph-justify distribution, so the line still reaches
//      the right margin and the region's internal pitch stays the region's own.
// (Finding numbers follow the review; 5/6/8 are Rust-side and pinned in
// styles.rs / parser.rs.)
// ─────────────────────────────────────────────────────────────────────────────

const ENV = { pageIndex: 0, totalPages: 1 };

function modelRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun & { type: 'text' } {
  return {
    type: 'text',
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 12, color: null, fontFamily: 'serif', isLink: false,
    background: null, vertAlign: null, hyperlink: null, allCaps: false,
    smallCaps: false, doubleStrikethrough: false, ...extra,
  } as DocxTextRun & { type: 'text' };
}

/** Linear ctx: every code point advances exactly the font px (12pt → 12px). */
function makeLinearCtx(): CanvasRenderingContext2D {
  let font = '12px serif';
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText(text: string) {
      const px = Number(/([\d.]+)px/.exec(font)?.[1] ?? 12);
      return {
        width: [...text].length * px,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function textSegments(segs: LayoutSeg[]): LayoutTextSeg[] {
  return segs.filter((seg): seg is LayoutTextSeg => 'text' in seg);
}

describe('§17.3.2.14 single-character region (finding 1)', () => {
  it('a one-glyph region occupies exactly w:val (advance = target, glyph at region start)', () => {
    // One run "氏" (1 cp, 12 px natural), val=2400 twips = 120 pt. §17.3.2.14
    // displays the contents "in the width specified": the region cell is 120 px
    // even though no inter-character gap exists. The glyph is NOT stretched
    // (compression/stretch ground truth pending); the cell pads after it.
    const lines = layoutLines(
      makeLinearCtx(),
      buildSegments([modelRun('氏', { fitTextVal: 2400 })], ENV),
      1000, 0, 1,
    );
    const segs = textSegments(lines[0].segments);
    expect(segs).toHaveLength(1);
    expect(segs[0].measuredWidth).toBeCloseTo(120, 9);
  });

  it('a following run starts at the single-glyph region target edge', () => {
    const lines = layoutLines(
      makeLinearCtx(),
      buildSegments([
        modelRun('氏', { fitTextVal: 2400 }),
        modelRun('あと'),
      ], ENV),
      1000, 0, 1,
    );
    const segs = textSegments(lines[0].segments);
    expect(segs).toHaveLength(2);
    // Pen space: seg 0 advance is the full 120 px cell, so 'あと' begins at 120.
    expect(segs[0].measuredWidth).toBeCloseTo(120, 9);
  });
});

describe('§17.3.2.14 tab inside a fitText run (finding 2)', () => {
  it('a tab splits the region: each text fragment is its own full region', () => {
    // "AB\tCD" with val=2400: a tab advance is resolved against tab stops, not
    // glyph advances, so it cannot participate in a fit region. The region is
    // split at the tab boundary; each fragment fits `w:val` independently
    // (adjudicated contract — Word's behaviour for this degenerate input is
    // unobserved). Neither fragment's charCount may include the tab.
    const lines = layoutLines(
      makeLinearCtx(),
      buildSegments([modelRun('AB\tCD', { fitTextVal: 2400, fitTextId: 5 })], ENV),
      1000, 0, 1,
    );
    const segs = textSegments(lines[0].segments);
    expect(segs.map((s) => s.text)).toEqual(['AB', 'CD']);
    // Each fragment is a self-consistent region: natural 24 + gaps = 120.
    expect(segs[0].measuredWidth).toBeCloseTo(120, 9);
    expect(segs[1].measuredWidth).toBeCloseTo(120, 9);
    // Fragments are SEPARATE regions (no cross-tab boundary gap linking them).
    expect(segs[0].fitTextRegionIndex).not.toBe(segs[1].fitTextRegionIndex);
  });
});

describe('§17.3.2.33 small caps changing the code-point count (finding 4)', () => {
  it('ß → SS: the region charCount follows the EMITTED text and the advance stays w:val', () => {
    // "aß" small-caps emits "ASS" (3 cps) — one more code point than the source
    // run text (2 cps). The per-gap divisor must count the emitted glyphs or the
    // folded advance overshoots the target.
    const lines = layoutLines(
      makeLinearCtx(),
      buildSegments([modelRun('aß', { fitTextVal: 2400, smallCaps: true })], ENV),
      1000, 0, 1,
    );
    const segs = textSegments(lines[0].segments);
    expect(segs.map((s) => s.text).join('')).toBe('ASS');
    expect(segs.reduce((sum, s) => sum + s.measuredWidth, 0)).toBeCloseTo(120, 9);
  });
});

// ── Paint-level contracts (recording canvas, model → renderDocumentToCanvas) ──

const FONT_PX = 20;

interface FillCall {
  text: string;
  x: number;
  y: number;
  letterSpacing: string;
  translateX: number;
  scaleX: number;
  scaleY: number;
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[] } {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fills: FillCall[] = [];
  let scaleX = 1;
  let scaleY = 1;
  let translateX = 0;
  const stack: { scaleX: number; scaleY: number; translateX: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    fontKerning: 'auto',
    measureText: (s: string) => {
      const p = px();
      const w = [...s].length * p;
      return {
        width: w,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() { stack.push({ scaleX, scaleY, translateX }); },
    restore() { const s = stack.pop(); if (s) { scaleX = s.scaleX; scaleY = s.scaleY; translateX = s.translateX; } },
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {},
    fillRect() {}, strokeRect() {}, clip() {}, rect() {}, setLineDash() {},
    drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    scale(sx: number, sy?: number) { scaleX *= sx; scaleY *= sy ?? sx; },
    translate(tx: number) { translateX += tx; },
    rotate() {},
    fillText(text: string, x: number, y: number) {
      fills.push({ text, x, y, letterSpacing, translateX, scaleX, scaleY });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, fills };
}

function paintRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'NotInMetrics', isLink: false,
    background: null, vertAlign: null, hyperlink: null, ...extra,
  };
}

type ParaRun = DocParagraph['runs'][number];

function para(runs: DocxTextRun[], alignment = 'left'): BodyElement {
  const p: DocParagraph = {
    alignment, indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: runs.map((r) => ({ type: 'text', ...r }) as ParaRun),
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as DocParagraph;
  return { type: 'paragraph', ...p } as BodyElement;
}

function section(extra: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 590, pageHeight: 600, marginTop: 0, marginRight: 0, marginBottom: 0,
    marginLeft: 0, headerDistance: 0, footerDistance: 0, titlePage: false,
    evenAndOddHeaders: false, ...extra,
  } as SectionProps;
}

function doc(body: BodyElement[], sec: SectionProps): DocxDocumentModel {
  return {
    section: sec, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function render(
  body: BodyElement[],
  sec: SectionProps = section(),
  onTextRun?: (r: DocxTextRunInfo) => void,
): Promise<FillCall[]> {
  const { canvas, fills } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc(body, sec), canvas, 0, {
    dpr: 1, width: sec.pageWidth, ...(onTextRun ? { onTextRun } : {}),
  });
  return fills;
}

describe('§17.18.44 justify around a fit region (finding 7)', () => {
  it('paragraph justify excludes region-internal gaps and the line reaches the right margin', async () => {
    // Page 590 wide, margins 0, jc=both. Line 1: fit region "あい" (val=1600 =
    // 80 pt cell, natural 40, per-gap 40) + 25 of the following CJK run's glyphs
    // (25 × 20 = 500) → natural line 580, slack 10. The fit region's INTERNAL
    // boundary must NOT receive a share of the paragraph slack (its pitch is
    // fixed by §17.3.2.14); the slack goes to the legal inter-CJK gaps so the
    // line's last glyph lands on the right margin (590).
    const fills = await render(
      [para([
        paintRun('あい', { fitTextVal: 1600 }),
        paintRun('い'.repeat(40)),
      ], 'both')],
    );
    const fit = fills.find((f) => f.text === 'あい');
    expect(fit, 'fit region painted as one contextual fillText').toBeDefined();
    // The region's internal pitch stays the region per-gap — 40 px exactly, no
    // justify pollution on top.
    expect(parseFloat((fit as FillCall).letterSpacing)).toBeCloseTo(40, 6);

    const line1Tail = fills.filter(
      (operation) => operation.text.includes('い') && operation.y === (fit as FillCall).y,
    );
    expect(line1Tail.length, 'line-1 tail of the following run painted').toBeGreaterThan(0);
    // The retained representation may be one contextual operation or several
    // positioned slices. Union their transformed ink extents instead of
    // requiring either representation; the final painted edge is invariant.
    const rightEdge = Math.max(...line1Tail.map((operation) => {
      const codePoints = [...operation.text].length;
      const internalSpacing = Math.max(0, codePoints - 1) * parseFloat(operation.letterSpacing);
      return operation.translateX
        + operation.scaleX * (operation.x + codePoints * FONT_PX + internalSpacing);
    }));
    expect(rightEdge).toBeCloseTo(590, 4);
  });
});

describe('fitted geometry reaches onTextRun (finding 3)', () => {
  it('onTextRun reports the drawn per-glyph letter-spacing of a fit segment', async () => {
    const infos: DocxTextRunInfo[] = [];
    await render(
      [para([paintRun('あい', { fitTextVal: 1600 }), paintRun('かき')])],
      section(),
      (r) => infos.push(r),
    );
    const fit = infos.find((r) => r.text === 'あい');
    const plain = infos.find((r) => r.text === 'かき');
    expect(fit).toBeDefined();
    expect(plain).toBeDefined();
    // Contract: DocxTextRunInfo.letterSpacingPx = the uniform per-code-point
    // pitch the run was DRAWN with (the §17.3.2.14 region per-gap here), so the
    // selection / find overlays can reproduce the fitted glyph geometry.
    expect((fit as DocxTextRunInfo & { letterSpacingPx?: number }).letterSpacingPx)
      .toBeCloseTo(40, 6);
    expect((plain as DocxTextRunInfo & { letterSpacingPx?: number }).letterSpacingPx ?? 0)
      .toBeCloseTo(0, 6);
    // The reported width stays the fitted advance (cell width for the whole
    // region: natural 40 + 1 gap × 40 = 80).
    expect((fit as DocxTextRunInfo).w).toBeCloseTo(80, 6);
  });
});

// ── Overlay contracts (fake DOM, node env) ──

interface FakeEl {
  tag: string;
  textContent: string;
  innerHTML: string;
  style: Record<string, string> & { cssText: string };
  children: FakeEl[];
  appendChild(c: FakeEl): void;
  addEventListener(): void;
  title?: string;
}

function makeEl(tag: string): FakeEl {
  const style: Record<string, string> = {};
  const el: FakeEl = {
    tag,
    textContent: '',
    innerHTML: '',
    children: [],
    style: new Proxy(style as Record<string, string> & { cssText: string }, {
      set(target, prop: string, value: string) {
        if (prop === 'cssText') {
          for (const decl of value.split(';')) {
            const idx = decl.indexOf(':');
            if (idx > 0) target[decl.slice(0, idx).trim()] = decl.slice(idx + 1).trim();
          }
          target.cssText = value;
        } else {
          target[prop] = value;
        }
        return true;
      },
    }),
    appendChild(c: FakeEl) { this.children.push(c); },
    addEventListener() {},
  };
  return el;
}

afterEach(() => vi.unstubAllGlobals());

function runInfo(partial: Record<string, unknown>): DocxTextRunInfo {
  return {
    text: 'X', x: 0, y: 0, w: 10, h: 12, fontSize: 12, font: '12px serif',
    ...partial,
  } as DocxTextRunInfo;
}

describe('selection span uses the drawn letter-spacing (finding 3)', () => {
  it('a run with letterSpacingPx renders its span with that CSS letter-spacing', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    buildDocxTextLayer(
      layer as unknown as HTMLDivElement,
      [runInfo({ text: '氏名', letterSpacingPx: 9.6 })],
      700, 900,
    );
    expect(layer.children).toHaveLength(1);
    // In units of the layer width (700px), so it scales with the page.
    expect(layer.children[0].style['letter-spacing']).toBe(`${(9.6 / 700) * 100}cqw`);
  });

  it('a run without letterSpacingPx keeps the letter-spacing reset (0)', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    buildDocxTextLayer(
      layer as unknown as HTMLDivElement,
      [runInfo({ text: 'plain' })],
      700, 900,
    );
    expect(layer.children[0].style['letter-spacing']).toBe('0');
  });
});

describe('find-highlight slices use the fitted advance (finding 3)', () => {
  // Monospace measurer: each char 12 px.
  const measureForFont = () => (s: string) => [...s].length * 12;

  it('a partial match inside a fitted run is offset and widened by the per-glyph pitch', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    // Drawn geometry with letterSpacingPx=10: glyph i starts at 12·i + 10·i.
    // Slice [1,3) ("BC"): left = 12 + 10 = 22; right = measure("ABC") + 2·10
    // = 36 + 20 = 56 → width = 34.
    const runs = [runInfo({ text: 'ABCD', x: 100, y: 50, w: 78, h: 16, letterSpacingPx: 10 })];
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 1, end: 3 }], active: false },
    ];
    buildDocxHighlightLayer(
      layer as unknown as HTMLDivElement,
      runs, matches, 700, 900, measureForFont,
    );
    expect(layer.children).toHaveLength(1);
    const box = layer.children[0];
    expect(box.style.left).toBe(`${((100 + 22) / 700) * 100}%`);
    expect(box.style.width).toBe(`${(34 / 700) * 100}%`);
  });

  it('a full-run match spans exactly the fitted width (no trailing gap)', () => {
    vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
    const layer = makeEl('div');
    // Full slice [0,4): width = natural 48 + (4−1)·10 = 78 — the region-end
    // fitted advance (no pitch after the last glyph).
    const runs = [runInfo({ text: 'ABCD', x: 100, y: 50, w: 78, h: 16, letterSpacingPx: 10 })];
    const matches: DocxHighlightMatch[] = [
      { slices: [{ runIndex: 0, start: 0, end: 4 }], active: false },
    ];
    buildDocxHighlightLayer(
      layer as unknown as HTMLDivElement,
      runs, matches, 700, 900, measureForFont,
    );
    const box = layer.children[0];
    expect(box.style.left).toBe(`${(100 / 700) * 100}%`);
    expect(box.style.width).toBe(`${(78 / 700) * 100}%`);
  });
});
