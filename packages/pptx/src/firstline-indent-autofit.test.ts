import { describe, it, expect } from 'vitest';
import { renderTextBody, naturalWidthExceedsBbox } from './renderer.js';
import type { TextBody, Paragraph, Bullet } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// ECMA-376 §21.1.2.2.7 (a:pPr @indent): the first-line indent is applied to the
// FIRST line only. The SAME amount must be used by all three code paths that
// consume it — the spAutoFit overflow MEASUREMENT (`naturalWidthExceedsBbox`),
// the wrap budget (`layoutParagraph`/`lineMaxW`), and the DRAW-side first-line
// offset (`textXOffset`) — otherwise they disagree:
//
//  1. A BULLETED paragraph's first line is NOT narrowed by `indent` (the gutter
//     is handled by the bullet/textX geometry). `naturalWidthExceedsBbox` used
//     to subtract `Math.max(0, indentPx)` UNCONDITIONALLY — never checking
//     `hasBullet` — so it measured a narrower box for a bulleted paragraph than
//     the wrap/draw passes actually use. (End-to-end this happens to leave the
//     rendered line count unchanged — `doWrap` only alters output when some
//     paragraph genuinely wraps, which trips the consistent measurement anyway
//     — but the measurement's *contract* was wrong, so the three expressions
//     could drift apart in the future. The fix makes them one shared helper.)
//  2. A NEGATIVE non-bullet indent extends the first line left of marL. The
//     measurement, wrap and draw paths must all preserve this signed value;
//     clamping every path consistently is still incorrect DrawingML geometry.

const SCALE = 1 / 12700; // emuToPx(emu) = emu * scale ⇒ 1pt → 1px
const emu = (px: number) => Math.round(px * 12700); // px → EMU at this scale
const RC = { themeMajorFont: null, themeMinorFont: null, dpr: 1 };

function mockCtx() {
  const texts: Array<{ text: string; x: number; y: number }> = [];
  let fillStyle = ''; let font = ''; let direction: CanvasDirection = 'ltr';
  const ctx = {
    get fillStyle() { return fillStyle; }, set fillStyle(v: string) { fillStyle = v; },
    get font() { return font; }, set font(v: string) { font = v; },
    get direction() { return direction; }, set direction(v: CanvasDirection) { direction = v; },
    // 10px advance per code point (font size ignored) → predictable line widths.
    measureText: (s: string) => ({
      width: [...s].length * 10, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
    }),
    fillText: (t: string, x: number, y: number) => texts.push({ text: t, x, y }),
    fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    measureTextWidth: undefined,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, texts };
}

function run(text: string, letterSpacing?: number): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', fontFamily: 'Arial',
    ...(letterSpacing != null ? { letterSpacing } : {}),
  };
}

const charBullet: Bullet = { type: 'char', char: '•', color: null, sizePct: null, fontFamily: null };
const noBullet: Bullet = { type: 'none' };

function makePara(opts: {
  runs: TextRunData[];
  indent?: number; // EMU
  bullet?: Bullet;
  alignment?: Paragraph['alignment'];
}): Paragraph {
  return {
    alignment: opts.alignment ?? 'l',
    marL: 0, marR: 0, indent: opts.indent ?? 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: opts.bullet ?? noBullet, defFontSize: null, defColor: null, defBold: null,
    defItalic: null, defFontFamily: null, tabStops: [], eaLnBrk: true, runs: opts.runs,
  } as Paragraph;
}

function makeBody(para: Paragraph, over: Partial<TextBody> = {}): TextBody {
  return {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: 20,
    defaultBold: null, defaultItalic: null,
    lIns: 91440, rIns: 91440, tIns: 45720, bIns: 45720, // 7.2px insets each
    wrap: 'square', vert: 'horz', autoFit: 'none', ...over,
  } as TextBody;
}

describe('pptx first-line indent: the measurement path ignores indent for a bulleted paragraph (§21.1.2.2.7)', () => {
  // Box 200px wide, insets 7.2px each ⇒ full text room ≈ 185.6px.
  // 18 glyphs (180px) fit the FULL width but NOT (185.6 − 50) = 135.6.
  const TEXT18 = 'あ'.repeat(18);
  const INDENT_50 = emu(50);
  const measure = (para: Paragraph) => {
    const { ctx } = mockCtx();
    return naturalWidthExceedsBbox(ctx, makeBody(para), 200, 7.2, 7.2, SCALE, RC);
  };

  it('returns FALSE for a bulleted paragraph whose text fits the box without the indent', () => {
    // A bullet does not consume first-line width, so the measurement must NOT
    // subtract the indent: 180px ≤ 185.6px ⇒ does not overflow.
    expect(measure(makePara({ runs: [run(TEXT18)], indent: INDENT_50, bullet: charBullet }))).toBe(false);
  });

  it('control: returns TRUE for the SAME positive indent on a NON-bullet paragraph', () => {
    // A non-bullet positive indent legitimately eats the first line's width:
    // 180px > 135.6px ⇒ overflows (so the bullet — not a blanket "never
    // overflows" — is the discriminator).
    expect(measure(makePara({ runs: [run(TEXT18)], indent: INDENT_50, bullet: noBullet }))).toBe(true);
  });
});

describe('pptx spAutoFit measurement includes DrawingML rPr@spc (§21.1.2.3.9)', () => {
  const measure = (textRun: TextRunData) => {
    const { ctx } = mockCtx();
    return naturalWidthExceedsBbox(
      ctx,
      makeBody(makePara({ runs: [textRun] })),
      200,
      7.2,
      7.2,
      SCALE,
      RC,
    );
  };

  it('counts positive and negative spacing in the natural line width', () => {
    expect(measure(run('A'.repeat(16)))).toBe(false);
    expect(measure(run('A'.repeat(16), 2))).toBe(true);
    expect(measure(run('A'.repeat(19)))).toBe(true);
    expect(measure(run('A'.repeat(19), -1))).toBe(false);
  });
});

describe('pptx first-line indent: the draw path matches the wrap path for a negative indent (§21.1.2.2.7)', () => {
  it('draws a negative non-bullet first-line indent left of marL', () => {
    // Short single-line text, wrap disabled so only the draw offset is in play.
    const zero = mockCtx();
    renderTextBody(zero.ctx, makeBody(makePara({ runs: [run('AAAA')], indent: 0 }), { wrap: 'none' }), 0, 0, 200, 200, SCALE);
    const neg = mockCtx();
    renderTextBody(neg.ctx, makeBody(makePara({ runs: [run('AAAA')], indent: emu(-50) }), { wrap: 'none' }), 0, 0, 200, 200, SCALE);

    expect(zero.texts.length).toBeGreaterThan(0);
    expect(neg.texts.length).toBeGreaterThan(0);
    const firstX = (t: Array<{ x: number }>) => Math.min(...t.map((c) => c.x));
    expect(firstX(neg.texts) - firstX(zero.texts)).toBeCloseTo(-50, 5);
  });

  it('extends only the first wrap line and keeps continuation lines at marL', () => {
    const para = makePara({ runs: [run('あ'.repeat(40))], indent: emu(-50) });
    para.marL = emu(50);
    const { ctx, texts } = mockCtx();
    renderTextBody(ctx, makeBody(para, { lIns: 0, rIns: 0 }), 0, 0, 200, 200, SCALE);
    const rows = [...new Set(texts.map(t => t.y))].map(y => texts.filter(t => t.y === y));
    expect(rows.map(row => row.map(t => t.text).join('').length)).toEqual([20, 15, 5]);
    expect(rows.map(row => Math.min(...row.map(t => t.x)))).toEqual([0, 50, 50]);
    for (const row of rows) expect(Math.max(...row.map(t => t.x + t.text.length * 10))).toBeLessThanOrEqual(200);
  });

  it.each(['l', 'ctr', 'r'] as const)('uses the same expanded first-line region for %s alignment', alignment => {
    const para = makePara({ runs: [run('AAAA')], indent: emu(-50), alignment });
    para.marL = emu(50);
    const { ctx, texts } = mockCtx();
    renderTextBody(ctx, makeBody(para, { lIns: 0, rIns: 0, wrap: 'none' }), 0, 0, 200, 200, SCALE);
    expect(Math.min(...texts.map(t => t.x))).toBeCloseTo({ l: 0, ctr: 80, r: 160 }[alignment], 5);
  });

  it('keeps tab stops anchored to the text body while the first line hangs left', () => {
    const para = makePara({ runs: [run('A\tB')], indent: emu(-50) });
    para.marL = emu(50);
    para.tabStops = [{ pos: emu(100), algn: 'l' }];
    const { ctx, texts } = mockCtx();
    renderTextBody(ctx, makeBody(para, { lIns: 0, rIns: 0 }), 0, 0, 200, 200, SCALE);
    expect(texts.find(t => t.text === 'A')?.x).toBeCloseTo(0, 5);
    expect(texts.find(t => t.text === 'B')?.x).toBeCloseTo(100, 5);
  });

  it('does not request wrapping when text fits the expanded first-line region', () => {
    const para = makePara({ runs: [run('あ'.repeat(18))], indent: emu(-50) });
    para.marL = emu(50);
    const { ctx } = mockCtx();
    expect(naturalWidthExceedsBbox(ctx, makeBody(para), 200, 7.2, 7.2, SCALE, RC)).toBe(false);
    para.indent = 0;
    expect(naturalWidthExceedsBbox(ctx, makeBody(para), 200, 7.2, 7.2, SCALE, RC)).toBe(true);
  });

  it('still applies a POSITIVE non-bullet first-line indent at draw time (regression guard)', () => {
    const zero = mockCtx();
    renderTextBody(zero.ctx, makeBody(makePara({ runs: [run('AAAA')], indent: 0 }), { wrap: 'none' }), 0, 0, 200, 200, SCALE);
    const pos = mockCtx();
    renderTextBody(pos.ctx, makeBody(makePara({ runs: [run('AAAA')], indent: emu(50) }), { wrap: 'none' }), 0, 0, 200, 200, SCALE);
    const firstX = (t: Array<{ x: number }>) => Math.min(...t.map((c) => c.x));
    // A positive indent still shifts the first line right by ~50px.
    expect(firstX(pos.texts) - firstX(zero.texts)).toBeCloseTo(50, 1);
  });
});
