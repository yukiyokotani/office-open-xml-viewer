import { describe, it, expect } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import { segAdvanceWidth, type CondensedAdvanceUnits, type LayoutTextSeg } from './line-layout.js';
import { condensedCharacterAdjustments } from './line-breaker/advance.js';
import { resolvedBidiLevels } from './character-spacing.js';
import { LineMeasurementAdapter } from './line-breaker/measurement-adapter.js';
import { paintLayoutPage } from './paint/canvas-page.js';
import { planVerticalRunWithCapability } from './vertical-text.js';
import type { VerticalGlyphMeasurementService } from './layout/measurement-capabilities.js';
import { makeRecordingCanvas } from './test-support/recording-canvas.test-support.js';
import type { LineLayout, TextPlacement } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxTextRun, SectionProps } from './types';

// #1975 / WORD_CONDENSED_SPACING_ADVANCE_FLOOR. ECMA-376 §17.3.2.35 removes
// condensed pitch "after each character"; Word floors every character's
// advance at zero instead of moving the pen backwards. A Word export of
// "AB" + "Mi"/"iM" (w:spacing -10 pt, Arial 20 pt) places "CD" at AB + (M - 10)
// in both orders, and an over-condensed run collapses its underline and
// shading to nothing.

/** Stub glyph advance: 'i', 'י' and 'ي' are 0.2 em, U+0640 is 0.1 em, every
 * other character 0.5 em. */
function makeCtx(contextual = false, longContext = false, kerning = false): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].reduce((sum, ch, index, all) => sum + p * (
          ch === '\u200D' ? 0 : ch === '\u0640' ? 0.1 : 'iיي'.includes(ch) ? 0.2 : 0.5
        // Optionally a letter after U+0640 takes a 0.2 em wider joined form.
        ) + (contextual && index > 0 && all[index - 1] === '\u0640' ? p * 0.2 : 0)
        // Optionally each 'l' widens by 0.01 em per preceding 'l' (context
        // longer than any bounded window).
        + (longContext && ch === 'l' ? p * 0.01 * all.slice(0, index).filter((c) => c === 'l').length : 0)
        // Optionally the pair "AV" kerns by -0.2 em.
        + (kerning && ch === 'A' && all[index + 1] === 'V' ? -p * 0.2 : 0), 0),
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

const section: SectionProps = {
  pageWidth: 400, pageHeight: 400,
  marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
  headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
};

function textRun(text: string, extra: Partial<DocxTextRun> & { charSpacing?: number } = {}): DocParagraph['runs'][number] {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null, ...extra,
  };
  return { type: 'text', ...run } as DocParagraph['runs'][number];
}

function para(runs: DocParagraph['runs'], extra: Partial<DocParagraph> = {}): BodyElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs, defaultFontSize: 20, defaultFontFamily: 'NotInMetrics', ...extra,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function layoutOf(body: BodyElement[], contextual = false, longContext = false) {
  return layoutBodyModel(body, section, makeCtx(contextual, longContext));
}

function lines(body: BodyElement[], contextual = false): LineLayout[] {
  const node = layoutOf(body, contextual).pages[0]!.layers.body[0]!;
  if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
  return [...node.lines];
}

function texts(line: LineLayout): TextPlacement[] {
  return line.placements.filter((p): p is TextPlacement => p.kind === 'text');
}

const byText = (line: LineLayout, text: string): TextPlacement =>
  texts(line).find((p) => p.text === text)!;

describe('#1975 condensed w:spacing larger than the glyph advance', () => {
  it('floors each over-condensed character at zero advance instead of rejecting', () => {
    // M = 10 pt, spacing -30 pt: each M advances 0, so CD starts where AB ends.
    const [line] = lines([para([textRun('AB'), textRun('MM', { charSpacing: -30 }), textRun('CD')])]);
    const run = byText(line!, 'MM');
    expect(run.advancePt).toBe(0);
    expect(run.bounds.widthPt).toBe(0);
    expect(run.clusters.map((c) => [c.offset.xPt, c.advancePt])).toEqual([[0, 0], [0, 0]]);
    expect(byText(line!, 'CD').bounds.xPt).toBe(byText(line!, 'AB').bounds.xPt + 20);
    // Each floored glyph paints from its own retained origin.
    expect(run.paintOps.map((op) => [op.text, op.offset.xPt])).toEqual([['M', 0], ['M', 0]]);
  });

  it.each(['Mi', 'iM'])('floors per character, not per run (%s)', (text) => {
    // M = 10 → 4, i = 4 → -2 floored to 0: the run advances 4 in both orders.
    // A signed or per-run sum would give 2.
    const [line] = lines([para([textRun('AB'), textRun(text, { charSpacing: -6 }), textRun('CD')])]);
    const run = byText(line!, text);
    expect(run.advancePt).toBeCloseTo(4, 9);
    expect(byText(line!, 'CD').bounds.xPt).toBeCloseTo(byText(line!, 'AB').bounds.xPt + 24, 9);
    const offsets = run.clusters.map((c) => [c.offset.xPt, c.advancePt]);
    expect(offsets).toEqual(text === 'Mi' ? [[0, 4], [4, 0]] : [[0, 0], [0, 4]]);
    // The unfloored M keeps native letterSpacing; the floored i is separate
    // and carries no internal pitch.
    expect(run.paintOps.map((op) => [op.text, op.offset.xPt, op.letterSpacingPt]))
      .toEqual(text === 'Mi' ? [['M', 0, -6], ['i', 4, 0]] : [['i', 0, 0], ['M', 0, -6]]);
  });

  it('floors a condensed space run', () => {
    const [line] = lines([para([textRun('Lorem'), textRun(' ', { charSpacing: -30 }), textRun('ipsum')])]);
    expect(byText(line!, ' ').advancePt).toBe(0);
    expect(byText(line!, 'ipsum').bounds.xPt).toBe(byText(line!, 'Lorem').bounds.xPt + 50);
  });

  it('aligns right and centre with the floored width', () => {
    const runs = () => [textRun('AB'), textRun('MMMM', { charSpacing: -30 }), textRun('CD')];
    const [right] = lines([para(runs(), { alignment: 'right' })]);
    expect(byText(right!, 'CD').bounds.xPt + 20).toBeCloseTo(380, 9);
    expect(byText(right!, 'AB').bounds.xPt).toBeCloseTo(340, 9);
    const [center] = lines([para(runs(), { alignment: 'center' })]);
    expect(byText(center!, 'AB').bounds.xPt).toBeCloseTo(180, 9);
  });

  it('breaks lines with the floored width, never borrowing room from negative advances', () => {
    // 360 pt band. "Q" + 20 over-condensed M (0 pt) + 36 'x' (360 pt): a signed
    // run (-400 pt) would pull every x onto line 1; floored, they overflow it.
    const source = `Q${'M'.repeat(20)}${'x'.repeat(36)}`;
    const all = lines([para([
      textRun('Q'), textRun('M'.repeat(20), { charSpacing: -30 }), textRun('x'.repeat(36)),
    ])]);
    expect(all.length).toBe(2);
    expect(all.flatMap(texts).map((p) => p.text).join('')).toBe(source);
    const firstWidth = texts(all[0]!).reduce((sum, p) => sum + p.advancePt, 0);
    expect(firstWidth).toBeCloseTo(360, 9);
    expect(all.flatMap(texts).every((p) => p.advancePt >= 0)).toBe(true);
  });

  it('justifies the line holding a floored run to the band edge', () => {
    // Word lengths leave slack on every full line.
    const words = 'aa bbb cc ddd ee fff gg hhh ii jjj kk lll mm nnn oo ppp qq rrr ss ttt ';
    const all = lines([para([
      textRun(words), textRun('MMMM', { charSpacing: -30 }), textRun(words),
    ], { alignment: 'both' })]);
    const line = all.find((candidate) => texts(candidate).some((p) => p.text === 'MMMM'))!;
    expect(line).not.toBe(all.at(-1));
    const placed = texts(line);
    expect(placed.every((p) => p.bounds.widthPt >= 0)).toBe(true);
    // The justified line's visible text still ends on the band edge, as
    // uncondensed lines do; its line-end space (10 pt here) hangs past it
    // (WORD_LINE_END_SPACE_HANGS).
    const last = placed.at(-1)!;
    expect(last.text.endsWith(' ')).toBe(true);
    expect(last.bounds.xPt + last.advancePt).toBeCloseTo(390, 6);
  });

  it('draws no underline under a fully floored run', () => {
    // AB [20, 40), MMMM floored at 40, CD [40, 60): only CD's underline remains.
    const [line] = lines([para([
      textRun('AB'), textRun('MMMM', { charSpacing: -30, underline: true }),
      textRun('CD', { underline: true }),
    ])]);
    const spans = line!.placements.flatMap((p) => p.kind === 'text' ? p.decorations ?? [] : [])
      .filter((d) => d.kind === 'underline')
      .map((d) => [(d as { from: { xPt: number } }).from.xPt, (d as { to: { xPt: number } }).to.xPt]);
    expect(spans).toEqual([[40, 60]]);
  });
});

describe('#1975 one contextual authority for fitting and retained geometry', () => {
  it('attributes a pair adjustment as retained cluster geometry does', () => {
    // Retained clusters are contextual prefix differences, and punctuation
    // compression is bounded against them, so the floor uses the same
    // attribution: M 10 → 2, A 10 → 2, V 10 kerned -4 → -2 floored.
    const node = layoutBodyModel(
      [para([textRun('MAV', { charSpacing: -8 }), textRun('B')])],
      section,
      makeCtx(false, false, true),
    ).pages[0]!.layers.body[0]!;
    if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
    const run = byText(node.lines[0]!, 'MAV');
    expect(run.advancePt).toBeCloseTo(4, 9);
    expect(run.paintOps.map((op) => [op.text, op.offset.xPt])).toEqual([['MA', 0], ['V', 4]]);
    expect(byText(node.lines[0]!, 'B').bounds.xPt).toBeCloseTo(run.bounds.xPt + 4, 9);
  });

  it('places following text at the retained pen for long-context glyphs', () => {
    // 'l' is 10 pt + 0.2 pt per preceding 'l'; at -10.3 the first two floor
    // and later ones advance by an amount that depends on unbounded context.
    const layout = layoutOf([para([
      textRun('A'), textRun('l'.repeat(20), { charSpacing: -10.3 }), textRun('B'),
    ])], false, true);
    const node = layout.pages[0]!.layers.body[0]!;
    if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
    const line = node.lines[0]!;
    const run = byText(line, 'l'.repeat(20));
    const retainedEnd = run.clusters.reduce(
      (end, cluster) => Math.max(end, cluster.offset.xPt + cluster.advancePt), 0,
    );
    expect(run.advancePt).toBeCloseTo(retainedEnd, 9);
    expect(byText(line, 'B').bounds.xPt).toBeCloseTo(run.bounds.xPt + retainedEnd, 9);
  });
});

/** UAX#9 L2 order of `text` drawn alone by Canvas in `direction`. */
function canvasVisualOrder(text: string, direction: 'ltr' | 'rtl'): string[] {
  const levels = resolvedBidiLevels(text, direction);
  // The tested strings are BMP-only, so code points index the levels.
  let units = [...text].map((unit, index) => ({ unit, level: levels[index]! }));
  const highest = Math.max(0, ...units.map((entry) => entry.level));
  const lowestOdd = Math.min(...units.map((entry) => entry.level).filter((level) => level % 2 === 1));
  for (let level = highest; level >= lowestOdd; level -= 1) {
    const next: typeof units = [];
    let cursor = 0;
    while (cursor < units.length) {
      if (units[cursor]!.level < level) {
        next.push(units[cursor]!);
        cursor += 1;
        continue;
      }
      let end = cursor;
      while (end < units.length && units[end]!.level >= level) end += 1;
      next.push(...units.slice(cursor, end).reverse());
      cursor = end;
    }
    units = next;
  }
  return units.map((entry) => entry.unit);
}

describe('#1975 condensed floor in right-to-left text', () => {
  it('places floored operations visually from the right edge', () => {
    // מ 10 → 4, י 4 → -2 floored, א 10 → 4: logical origins 0, 4, 4; width 8.
    const [line] = lines([para(
      [textRun('אב', { rtl: true }), textRun('מיא', { rtl: true, charSpacing: -6 }), textRun('גד', { rtl: true })],
      { bidi: true } as Partial<DocParagraph>,
    )]);
    const run = byText(line!, 'מיא');
    expect(run.advancePt).toBeCloseTo(8, 9);
    expect(run.paintOps.map((op) => [op.text, op.direction, op.offset.xPt]))
      .toEqual([['מ', 'rtl', 4], ['י', 'rtl', 0], ['א', 'rtl', 0]]);
    // Following text continues leftwards from the floored pen.
    expect(byText(line!, 'גד').bounds.xPt + 20).toBeCloseTo(run.bounds.xPt, 9);
  });

  it('mirrors a floored vertical piece without a split floored cluster', () => {
    // بي floors (ي 4 - 6) and stays one cursive cell; ! and ب do not floor:
    // logical origins 0, 4, 8 over width 12 paint at 8, 4, 0.
    const node = layoutBodyModel(
      [para([textRun('بي!ب', { rtl: true, charSpacing: -6 })], { bidi: true } as Partial<DocParagraph>)],
      { ...section, textDirection: 'tbRl' } as SectionProps,
      makeCtx(),
    ).pages[0]!.layers.body[0]!;
    if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
    expect(byText(node.lines[0]!, 'بي!ب').paintOps.map((op) => [op.text, op.offset.xPt]))
      .toEqual([['بي', 8], ['!', 4], ['ب', 0]]);
  });

  it.each([
    ['horizontal', undefined,
      [['אב', 'rtl', 16], ['A', 'ltr', 8], ['i', 'ltr', 12], ['B', 'ltr', 12], ['גד', 'rtl', 0]]],
    // A floored vertical piece paints one cell per grapheme.
    ['vertical', 'tbRl',
      [['א', 'rtl', 20], ['ב', 'rtl', 16], ['A', 'ltr', 8], ['i', 'ltr', 12], ['B', 'ltr', 12],
        ['ג', 'rtl', 4], ['ד', 'rtl', 0]]],
  ] as const)('keeps embedded left-to-right order in a floored right-to-left run (%s)', (_name, textDirection, expected) => {
    // אב, A, i (floored), B, גד at logical 0, 8, 12, 12, 16 over width 24.
    // UAX#9 (base RTL): AiB is level 2 and keeps its order, so visually
    // גד | A i B | אב, as one contextual Canvas operation would draw it; each
    // operation draws in its level's direction.
    const node = layoutBodyModel(
      [para([textRun('אבAiBגד', { rtl: true, charSpacing: -6 })], { bidi: true } as Partial<DocParagraph>)],
      textDirection ? { ...section, textDirection } as SectionProps : section,
      makeCtx(),
    ).pages[0]!.layers.body[0]!;
    if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
    expect(byText(node.lines[0]!, 'אבAiBגד').paintOps.map((op) => [op.text, op.direction, op.offset.xPt]))
      .toEqual(expected);
  });

  it.each([
    ['horizontal', undefined, [['אב', 26], ['A\u00ADB', 8], ['i', 26], ['גד', 0]]],
    ['vertical', 'tbRl',
      [['א', 30], ['ב', 26], ['A', 8], ['\u00AD', 12], ['B', 22], ['i', 26], ['ג', 4], ['ד', 0]]],
  ] as const)('does not let a removed control break an embedded run (%s)', (_name, textDirection, expected) => {
    // U+00AD (stub width 10, no pitch) is removed by UAX#9 X9 and takes its
    // neighbours' level: A, U+00AD, B and the floored i stay one LTR run
    // (8..26), visually between גד and אב.
    const text = 'אבA\u00ADBiגד';
    // U+00AD is removed by UAX#9 X9; A, B and the floored i keep their order.
    const node = layoutBodyModel(
      [para([textRun(text, { rtl: true, charSpacing: -6 })], { bidi: true } as Partial<DocParagraph>)],
      textDirection ? { ...section, textDirection } as SectionProps : section,
      makeCtx(),
    ).pages[0]!.layers.body[0]!;
    if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
    const ops = byText(node.lines[0]!, text).paintOps;
    expect(ops.map((op) => [op.text, op.offset.xPt])).toEqual(expected);
  });

  it('keeps a join control inside its vertical cursive group', () => {
    // U+200D is removed by UAX#9 X9; it must not split ب‍ت's contextual cell.
    const node = layoutBodyModel(
      [para([textRun('ب\u200Dت', { rtl: true, charSpacing: -12 })], { bidi: true } as Partial<DocParagraph>)],
      { ...section, textDirection: 'tbRl' } as SectionProps,
      makeCtx(),
    ).pages[0]!.layers.body[0]!;
    if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
    expect(byText(node.lines[0]!, 'ب\u200Dت').paintOps.map((op) => op.text)).toEqual(['ب\u200Dت']);
  });

  it.each([
    ['RLO', 'horizontal', undefined, '\u202E'],
    ['RLO', 'vertical', 'tbRl', '\u202E'],
    ['LRO', 'horizontal', undefined, '\u202D'],
    ['LRO', 'vertical', 'tbRl', '\u202D'],
    // At the UAX#9 depth limit the RLO reaches level 125 and the LRE after
    // it overflows, so the override still orders ABiC.
    ['RLO', 'horizontal at the depth limit', undefined, `${'\u202B'.repeat(61)}\u202E\u202A`],
    ['RLO', 'vertical at the depth limit', 'tbRl', `${'\u202B'.repeat(61)}\u202E\u202A`],
  ] as const)('keeps a %s override\'s order across split operations (%s)', (control, _name, textDirection, mark) => {
    // ABiC (i floors) inside RLO reads C i B A; inside LRO, A B i C.
    const text = `אב${mark}ABiC\u202Cגד`;
    const node = layoutBodyModel(
      [para([textRun(text, { rtl: true, charSpacing: -6 })], { bidi: true } as Partial<DocParagraph>)],
      textDirection ? { ...section, textDirection } as SectionProps : section,
      makeCtx(),
    ).pages[0]!.layers.body[0]!;
    if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
    // Visual order: operations by placed origin, each resolved by Canvas as
    // its own UAX#9 paragraph in its paint direction.
    const visual = byText(node.lines[0]!, text).paintOps
      .map((op, index) => ({ op, index }))
      .sort((left, right) => left.op.offset.xPt - right.op.offset.xPt || left.index - right.index)
      .flatMap(({ op }) => canvasVisualOrder(op.text, op.direction).filter((unit) => /[ABC]/u.test(unit)))
      .join('');
    expect(visual).toBe(control === 'RLO' ? 'CBA' : 'ABC');
  });

  it('draws a split operation in its own level direction', () => {
    // "A-" resolves to level 2 inside the RTL text; drawn alone with RTL
    // direction Canvas would make the trailing "-" level 1 and reorder it.
    const [line] = lines([para(
      [textRun('אבA-iBגד', { rtl: true, charSpacing: -6 })],
      { bidi: true } as Partial<DocParagraph>,
    )]);
    const ops = byText(line!, 'אבA-iBגד').paintOps;
    expect(ops.find((op) => op.text === 'A-')?.direction).toBe('ltr');
    expect(ops.find((op) => op.text === 'אב')?.direction).toBe('rtl');
  });

  it('places a floored mixed word as Canvas orders it', () => {
    // ab (level 2), ג, י (floored), ד (level 1): visually ד/י | ג | ab.
    const [line] = lines([para([textRun('abגיד', { charSpacing: -6 })])]);
    expect(byText(line!, 'abגיד').paintOps.map((op) => [op.text, op.direction, op.offset.xPt]))
      .toEqual([['ab', 'ltr', 8], ['ג', 'rtl', 4], ['י', 'rtl', 0], ['ד', 'rtl', 0]]);
  });

  it('keeps a floored cursive run as one contextual operation', () => {
    // Splitting would break the joined forms; Canvas does not track inside
    // cursive scripts, so only the advance is floored.
    const [line] = lines([para(
      [textRun('بيت', { rtl: true, charSpacing: -6 }), textRun('من', { rtl: true })],
      { bidi: true } as Partial<DocParagraph>,
    )]);
    const run = byText(line!, 'بيت');
    expect(run.paintOps.map((op) => op.text)).toEqual(['بيت']);
    // ب 10 → 4, ي 4 → 0, ت 10 → 4.
    expect(run.advancePt).toBeCloseTo(8, 9);
    expect(run.clusters.map((c) => c.advancePt)).toEqual([4, 0, 4]);
  });

  it('splits floored non-cursive clusters beside cursive text in one segment', () => {
    // A w:cs run keeps Latin, Arabic and join controls in one segment. The
    // ZWJ is pitch-bearing, so 'M\u200D' (10 - 2 x 6) floors too.
    for (const [text, expected] of [
      // Latin and Arabic take different bidi levels, so they split too.
      ['MiMب', [['M', -6], ['i', 0], ['M', -6], ['ب', -6]]],
      ['MiM\u200D', [['M', -6], ['i', 0], ['M\u200D', 0]]],
    ] as const) {
      const [line] = lines([para([textRun(text, { charSpacing: -6, cs: true } as Partial<DocxTextRun>)])]);
      expect(byText(line!, text).paintOps.map((op) => [op.text, op.letterSpacingPt])).toEqual(expected);
    }
  });

  it('paints a floored non-cursive cluster without internal pitch', async () => {
    const layout = layoutOf([para(
      [textRun('מיא', { rtl: true, charSpacing: -6 })],
      { bidi: true } as Partial<DocParagraph>,
    )]);
    const node = layout.pages[0]!.layers.body[0]!;
    if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
    const ops = byText(node.lines[0]!, 'מיא').paintOps;
    expect(ops.map((op) => [op.text, op.letterSpacingPt])).toEqual([['מ', -6], ['י', 0], ['א', -6]]);
    expect(ops.find((op) => op.text === 'י')?.naturalAdvancePt).toBe(4);
    const { canvas } = makeRecordingCanvas();
    await expect(paintLayoutPage(layout, 0, canvas, { dpr: 1, scale: 1 })).resolves.not.toThrow();
  });

  it('does not open a gap for trailing whitespace', () => {
    const offsets = (text: string) => {
      const [line] = lines([para(
        [textRun(text, { rtl: true, charSpacing: -6 })],
        { bidi: true } as Partial<DocParagraph>,
      )]);
      const run = byText(line!, text);
      return run.paintOps.filter((op) => op.text.trim() !== '')
        .map((op) => run.origin.xPt + op.offset.xPt);
    };
    // Visible glyphs keep their relative placement whatever whitespace follows.
    const relative = (xs: number[]) => xs.map((x) => x - xs[0]!);
    const bare = relative(offsets('מיא'));
    expect(relative(offsets('מיא '))).toEqual(bare);
    expect(relative(offsets('מיא  '))).toEqual(bare);
  });

  it.each([
    ['explicit w:rtl', 'מיא', { rtl: true }],
    ['bidi-resolved direction', 'מיא', {}],
    ['trailing whitespace', 'מיא ', { rtl: true }],
    ['two trailing spaces', 'מיא  ', { rtl: true }],
  ] as const)('keeps right-to-left order in a floored vertical sideways piece (%s)', (_name, text, extra) => {
    // tbRl: the same visual placement as horizontally (מ 4, י 0, א 0 from the
    // piece's start), not the logical origins 0, 4, 4. The paint origin
    // already moves past trailing whitespace, so it adds no shift here.
    const node = layoutBodyModel(
      [para([textRun(text, { ...extra, charSpacing: -6 })], { bidi: true } as Partial<DocParagraph>)],
      { ...section, textDirection: 'tbRl' } as SectionProps,
      makeCtx(),
    ).pages[0]!.layers.body[0]!;
    if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
    const run = byText(node.lines[0]!, text);
    expect(run.paintOps.filter((op) => op.text.trim() !== '')
      .map((op) => [op.text, op.direction, op.glyphOrientation, op.offset.xPt]))
      .toEqual([['מ', 'rtl', 'sideways', 4], ['י', 'rtl', 'sideways', 0], ['א', 'rtl', 'sideways', 0]]);
  });
});

describe('#1975 condensed floor and kashida justification', () => {
  const kashidaParagraph = (charSpacing: number) => lines([para(
    [textRun('بنت سلام كتب '.repeat(12), { rtl: true, charSpacing })],
    { bidi: true, alignment: 'lowKashida' } as Partial<DocParagraph>,
  )], true);
  const tatweels = (all: LineLayout[]) => all.flatMap(texts)
    .flatMap((p) => p.paintOps).filter((op) => op.text.includes('\u0640')).length;

  it('inserts no elongation that would itself reach the floor', () => {
    // U+0640 is 2 pt in the stub: at -6 pt pitch an inserted tatweel floors,
    // while the wider joined letter after it would still gain width.
    expect(tatweels(kashidaParagraph(-6))).toBe(0);
  });

  it('still elongates when the tatweel keeps a positive advance', () => {
    expect(tatweels(kashidaParagraph(-1))).toBeGreaterThan(0);
  });
});

describe('segAdvanceWidth condensed floor', () => {
  const seg = (text: string, charSpacing: number, extra: Partial<LayoutTextSeg> = {}): LayoutTextSeg => ({
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'X', vertAlign: null, measuredWidth: 0, charSpacing, ...extra,
  } as LayoutTextSeg);
  const natural = (clusters: Array<[string, number]>): CondensedAdvanceUnits => ({
    kind: 'natural-clusters',
    clusters: clusters.map(([text, advancePx]) => ({ text, advancePx })),
  });

  it('sums floored cluster advances', () => {
    expect(segAdvanceWidth(seg('Mi', -6), 14, undefined, 1, natural([['M', 10], ['i', 4]])))
      .toBeCloseTo(4, 12);
  });

  it('keeps a fully floored run exactly zero under fractional spacing and w:w', () => {
    const units = natural([['M', 10], ['M', 10], ['M', 10]]);
    expect(segAdvanceWidth(seg('MMM', -30.1), 30, undefined, 1, units)).toBe(0);
    expect(segAdvanceWidth(seg('MMM', -7.3, { charScale: 0.7 }), 30, undefined, 1, units)).toBe(0);
  });

  it('floors punctuation compression together with the run pitch', () => {
    const compressed = seg('。', -30, { punctuationCompressions: [{ end: 1, adjustmentPt: -5 }] });
    expect(segAdvanceWidth(compressed, 10, undefined, 1, natural([['。', 10]]))).toBe(0);
    expect(segAdvanceWidth(compressed, 10, undefined, 1)).toBe(0);
  });

  it('measures vertical runs from their planned cells', () => {
    // Planned cells already carry pitch, character adjustments and floor.
    const vertical = seg('ーー', -30, { verticalRun: true });
    const cells: CondensedAdvanceUnits = {
      kind: 'vertical-cells',
      cells: [{ text: 'ー', advancePx: 0, floored: true }, { text: 'ー', advancePx: 0, floored: true }],
    };
    expect(segAdvanceWidth(vertical, 40, undefined, 1, cells)).toBe(0);
  });

  it('is unchanged when no character reaches the floor', () => {
    const units = natural([['M', 10], ['i', 4]]);
    expect(segAdvanceWidth(seg('Mi', -2), 14, undefined, 1, units)).toBe(10);
    expect(segAdvanceWidth(seg('Mi', 3), 14, undefined, 1)).toBe(20);
  });
});

describe('condensed character adjustments', () => {
  it('assigns each compression to the unit holding its end, in linear time', () => {
    const text = '甲、'.repeat(20_000);
    const segment = {
      ...textRun(text, { charSpacing: -1 }), measuredWidth: 0,
      punctuationCompressions: Array.from({ length: 20_000 }, (_, index) => ({
        end: index * 2 + 2, adjustmentPt: -5,
      })),
    } as unknown as LayoutTextSeg;
    const adjustment = condensedCharacterAdjustments(segment, undefined, 1);
    expect(adjustment(0, 1)).toBe(0);
    expect(adjustment(1, 2)).toBe(-5);
    expect(adjustment(0, text.length)).toBe(-100_000);
  });
});

describe('condensed advance units work', () => {
  function countingContext(): { ctx: CanvasRenderingContext2D; measured: () => number } {
    let units = 0;
    const ctx = makeCtx();
    const measure = ctx.measureText.bind(ctx);
    ctx.measureText = (text: string) => {
      units += text.length;
      return measure(text);
    };
    return { ctx, measured: () => units };
  }

  it.each([100, 200, 400])('is linear in the run length (%s characters)', (length) => {
    const { ctx, measured } = countingContext();
    const adapter = new LineMeasurementAdapter(ctx, 1, () => '20px serif');
    const segment = { ...textRun('M'.repeat(length), { charSpacing: -30 }), measuredWidth: 0 } as unknown as LayoutTextSeg;
    let reserved = 0;
    const reserve = (units: number) => { reserved += units; };
    adapter.condensedAdvanceUnits(segment, segment.text, 0, undefined, reserve);
    const whole = measured();
    expect(whole).toBeLessThanOrEqual(12 * length);
    expect(reserved).toBeGreaterThanOrEqual(whole);
    // Prefix probes reuse the whole-segment clusters.
    for (let end = 1; end <= length; end += 1) {
      adapter.condensedAdvanceUnits(segment, segment.text.slice(0, end), 0, undefined, reserve);
    }
    expect(measured()).toBe(whole);
  });

  it('builds and charges the adjustment table once per segment across prefix probes', () => {
    const ctx = makeCtx();
    const text = 'あいう'.repeat(2_000);
    const segment = {
      ...textRun(text, { charSpacing: -30 }), measuredWidth: 0, verticalRun: true,
    } as unknown as LayoutTextSeg;
    let tableCharges = 0;
    const first = condensedCharacterAdjustments(segment, undefined, 1, (units) => { tableCharges += units; });
    for (let end = 1; end <= 200; end += 1) {
      expect(condensedCharacterAdjustments(segment, undefined, 1, (units) => { tableCharges += units; }))
        .toBe(first);
    }
    expect(tableCharges).toBe(text.length);
    // A probe of the same segment reuses the memoised table.
    const adapter = new LineMeasurementAdapter(ctx, 1, () => '20px serif', {
      planRun: (input: Parameters<VerticalGlyphMeasurementService['planRun']>[0]) => planVerticalRunWithCapability(
        ctx, input.text, input.fontSizePt, input.letterSpacingPt, input.charScale,
        input.growTrRotateInk, () => false, input.writingMode, input.characterGridDeltaPt,
        input.characterAdjustmentPt,
      ),
    } as unknown as VerticalGlyphMeasurementService);
    let reserved = 0;
    for (let end = 1; end <= 50; end += 1) {
      adapter.condensedAdvanceUnits(segment, text.slice(0, end), 0, undefined, (units) => { reserved += units; });
    }
    expect(reserved).toBeLessThanOrEqual(50 * 51 / 2 * 19);
  });

  it('charges vertical cell planning for every measured unit', () => {
    const { ctx, measured } = countingContext();
    const adapter = new LineMeasurementAdapter(ctx, 1, () => '20px serif', {
      planRun: (input: Parameters<VerticalGlyphMeasurementService['planRun']>[0]) => planVerticalRunWithCapability(
        ctx, input.text, input.fontSizePt, input.letterSpacingPt, input.charScale,
        input.growTrRotateInk, () => false, input.writingMode, input.characterGridDeltaPt,
        input.characterAdjustmentPt,
      ),
    } as unknown as VerticalGlyphMeasurementService);
    const text = 'ffiffiあいう'.repeat(25);
    const segment = {
      ...textRun(text, { charSpacing: -30 }), measuredWidth: 0, verticalRun: true,
    } as unknown as LayoutTextSeg;
    let reserved = 0;
    adapter.condensedAdvanceUnits(segment, text, 0, undefined, (units) => { reserved += units; });
    expect(measured()).toBeGreaterThan(0);
    expect(measured()).toBeLessThanOrEqual(reserved);
  });
});
