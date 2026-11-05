import { createLineOrderResolver } from './bidi-line.js';
import { physicalDecimalAlignmentPoint } from './line-breaker/decimal-anchor.js';
import type { LayoutSeg } from './line-breaker/model.js';
import { describe, it, expect } from 'vitest';
import {
  layoutBidiTabStops,
  type BidiTabItem,
} from './line-layout.js';
import { nextTabStopRtl } from './layout/text.js';
import { computeLineVisualOrder } from './bidi-line.js';
import { renderDocumentToCanvas } from './renderer.js';
import type { DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// ECMA-376 §17.3.1.6 (base RTL) + §17.3.1.37 (tabs) + §17.15.1.25 (default tab)
// + §17.18.84 (ST_TabJc start/end logical edges) — a bidi paragraph's tab stops
// are anchored at the LEADING (right) text edge and its cells reorder visually.
// Issue #820: the Arabic TOC's page numbers and dot/underscore leaders (and the
// footer's tab-aligned fields) landed on the wrong visual side (or wrapped to a
// new line) because tabs were resolved in LTR pen coordinates. These tests pin
// the mirrored resolution independently of any font metrics.

describe('nextTabStopRtl (§17.3.1.37 / §17.15.1.25, leading = right edge)', () => {
  const stops = [
    { pos: 100, alignment: 'left' as const, leader: 'none' as const },
    { pos: 300, alignment: 'right' as const, leader: 'underscore' as const },
  ];
  it('advances leftward to the nearest custom stop past the pen', () => {
    // Pen 50 from the right edge → next stop further left is pos 100.
    expect(nextTabStopRtl(50, stops, 36)?.pos).toBe(100);
    // Pen 100 (on the first stop) → advances to 300.
    expect(nextTabStopRtl(100, stops, 36)?.pos).toBe(300);
  });
  it('falls onto the §17.15.1.25 automatic grid AFTER all custom stops', () => {
    // Past the last custom stop (300); the grid is anchored at the leading edge
    // with interval 36, so the next multiple past 300 is 324.
    const s = nextTabStopRtl(300, stops, 36);
    expect(s?.pos).toBe(324);
    expect(s?.alignment).toBe('left');
    expect(s?.leader).toBeUndefined();
  });
  it('returns null when no interval and the pen is past every custom stop', () => {
    expect(nextTabStopRtl(400, stops, 0)).toBeNull();
  });
});

describe('computeLineVisualOrder treats a tab as a segment separator (UAX#9 S)', () => {
  it('reorders tab-delimited cells in mirrored order under an RTL base', () => {
    // Logical order: [chapNum][space][title] TAB [pageNum] (all rtl-marked).
    const segs = [
      { text: '1.1', rtl: true }, { text: ' ', rtl: true }, { text: 'TITLE', rtl: true },
      { isTab: true },
      { text: '4', rtl: true },
    ];
    const { order } = computeLineVisualOrder(segs as unknown[], true);
    // Visual L→R must be: pageNum, TAB, title, space, chapNum — the page number
    // ends up on the visual LEFT and the chapter number on the visual RIGHT, with
    // the tab (leader region) between the cells. Without the S classification the
    // whole line reversed as one run and the page number landed mid-line.
    expect(order.map((i) => (('isTab' in segs[i]) ? 'TAB' : (segs[i] as { text: string }).text)))
      .toEqual(['4', 'TAB', 'TITLE', ' ', '1.1']);
  });
});

describe('layoutBidiTabStops (§17.3.1.37 mirror — margin-anchored reading frame)', () => {
  const avail = 400; // margin-to-margin width px (no-indent case: startPen 0, leftLimit 400)

  /** Reading-frame LEFT edge (px from the right text margin) of each segment
   *  after the walk: startPen + cumulative widths through index i. */
  const readingEdges = (
    items: BidiTabItem[],
    res: { width: number }[],
    startPen: number,
  ): number[] => {
    let pen = startPen;
    return items.map((_, i) => (pen += res[i].width));
  };

  it('places an end (right) tab so its page number trails at the mirrored stop', () => {
    // TOC row logical order: chapNum(20) space(4) title(80) TAB pageNum(8).
    // The tab is the style right/underscore leader stop at pos 300 from the
    // right text margin. The page number is trailing-aligned: its trailing
    // (reading-left) edge sits ON the stop.
    const items: BidiTabItem[] = [
      { isTab: false, width: 20 }, { isTab: false, width: 4 }, { isTab: false, width: 80 },
      { isTab: true, width: 0 },
      { isTab: false, width: 8 },
    ];
    const stops = [{ pos: 300, alignment: 'right' as const, leader: 'underscore' as const }];
    const res = layoutBidiTabStops(items, stops, 0, avail, 36);
    const edge = readingEdges(items, res, 0);
    // The page number's trailing (reading-left) edge is on the stop: the tab
    // advances the pen to stop − fw (edge[3] = 292), and the page number's own
    // width carries it onto the stop (edge[4] = 300).
    expect(edge[3]).toBeCloseTo(292, 6);
    expect(edge[4]).toBeCloseTo(300, 6);
    // Chapter number (first logical, leading) starts at the RIGHT text margin.
    expect(edge[0]).toBeCloseTo(20, 6);
    // The tab carries the underscore leader and fills the visible gap.
    expect(res[3].leader).toBe('underscore');
    expect(res[3].width).toBeGreaterThan(0);
  });

  it('honors strict logical start/end aliases and the num leading role', () => {
    const items: BidiTabItem[] = [
      { isTab: false, width: 40 },
      { isTab: true, width: 0 },
      { isTab: false, width: 30 },
    ];
    const edgeAfterTab = (alignment: 'left' | 'start' | 'right' | 'end' | 'num') => {
      const result = layoutBidiTabStops(
        items,
        [{ pos: 150, alignment, leader: 'none' }],
        0,
        avail,
        1000,
      );
      return readingEdges(items, result, 0)[1];
    };

    expect(edgeAfterTab('start')).toBeCloseTo(edgeAfterTab('left'), 6);
    expect(edgeAfterTab('num')).toBeCloseTo(edgeAfterTab('left'), 6);
    expect(edgeAfterTab('end')).toBeCloseTo(edgeAfterTab('right'), 6);
    expect(edgeAfterTab('start')).toBeCloseTo(150, 6);
    expect(edgeAfterTab('end')).toBeCloseTo(120, 6);
  });

  it('pins a page number to the left text margin when the stop is past it', () => {
    // A right/leader stop at pos 420 (past the 400px left margin) would place
    // the page number past the margin; it pins so its far (left) edge is ON the
    // margin: reading left edge = 400.
    const items: BidiTabItem[] = [
      { isTab: false, width: 20 }, { isTab: false, width: 80 },
      { isTab: true, width: 0 },
      { isTab: false, width: 8 },
    ];
    const stops = [{ pos: 420, alignment: 'right' as const, leader: 'underscore' as const }];
    const res = layoutBidiTabStops(items, stops, 0, avail, 36);
    const edge = readingEdges(items, res, 0);
    expect(edge[3]).toBeCloseTo(avail, 6); // page number's far (left) edge on the margin
  });

  it('flips physical left (leading) so its content ends at the mirrored stop', () => {
    // Single leading (left) tab at pos 150 from the right margin. Following
    // content (width 30) has its LEADING (right) edge there: reading span
    // [150, 180] (= visual [220, 250]).
    const items: BidiTabItem[] = [
      { isTab: false, width: 40 }, // leading content
      { isTab: true, width: 0 },
      { isTab: false, width: 30 }, // follows the tab
    ];
    const stops = [{ pos: 150, alignment: 'left' as const, leader: 'none' as const }];
    const res = layoutBidiTabStops(items, stops, 0, avail, 1000 /* no auto grid */);
    const edge = readingEdges(items, res, 0);
    expect(edge[1]).toBeCloseTo(150, 6); // pen after the tab = the stop
    expect(edge[2]).toBeCloseTo(180, 6); // content extends 30 further left
    expect(res[1].leader ?? 'none').toBe('none');
  });

  it('centers content around a mirrored center stop', () => {
    const items: BidiTabItem[] = [
      { isTab: false, width: 40 },
      { isTab: true, width: 0 },
      { isTab: false, width: 20 },
    ];
    const stops = [{ pos: 200, alignment: 'center' as const, leader: 'none' as const }];
    const res = layoutBidiTabStops(items, stops, 0, avail, 1000);
    const edge = readingEdges(items, res, 0);
    // Content spans reading [190, 210]: centered on the stop (200).
    expect(edge[1]).toBeCloseTo(190, 6);
    expect(edge[2]).toBeCloseTo(210, 6);
  });

  it('aligns the first following period at a decimal stop across segment boundaries', () => {
    const items: BidiTabItem[] = [
      { isTab: false, width: 40 },
      { isTab: true, width: 0 },
      { isTab: false, width: 10 },
      { isTab: false, width: 30, decimalOffset: 5 },
      { isTab: false, width: 10 },
    ];
    const res = layoutBidiTabStops(
      items,
      [{ pos: 200, alignment: 'decimal', leader: 'none' }],
      0,
      avail,
      1000,
    );
    const edge = readingEdges(items, res, 0);
    // The logical prefix is 10 + 5px, but the LTR numeric cell is embedded in
    // a mirrored reading frame. Its physical period-left edge is therefore
    // total(50) - prefix(15) = 35px from the cell's physical right edge.
    expect(edge[1]).toBeCloseTo(165, 6);
    expect(edge[1] + 35).toBeCloseTo(200, 6);
  });

  it('aligns a no-decimal LTR cell by its physical right edge', () => {
    const items: BidiTabItem[] = [
      { isTab: false, width: 40 },
      { isTab: true, width: 0 },
      { isTab: false, width: 50 },
    ];
    const res = layoutBidiTabStops(
      items,
      [{ pos: 200, alignment: 'decimal', leader: 'none' }],
      0,
      avail,
      1000,
    );
    const edge = readingEdges(items, res, 0);
    // The mirrored reading frame starts at the physical right edge. Word puts
    // that edge on the stop when no decimal exists, so the tab itself advances
    // to 200 and the cell continues leftward from there.
    expect(edge[1]).toBeCloseTo(200, 6);
    expect(edge[2]).toBeCloseTo(250, 6);
  });

  it('assigns the Nth tab the Nth-reachable stop IN READING ORDER (leader cell)', () => {
    // sample-28 TOC2 shape: [chapNum][TAB][title][TAB][pageNum] with an early
    // left stop and the right/underscore leader stop. The FIRST logical tab must
    // take the early left stop and the SECOND (the one before the page number)
    // the leader stop — resolving against the VISUAL sequence reverses the
    // assignment and paints the leader between the title and the chapter number
    // (the #830 follow-up bug).
    const items: BidiTabItem[] = [
      { isTab: false, width: 20 },  // chapter number
      { isTab: true, width: 0 },    // → left@50
      { isTab: false, width: 100 }, // title
      { isTab: true, width: 0 },    // → right/underscore@380
      { isTab: false, width: 8 },   // page number
    ];
    const stops = [
      { pos: 50, alignment: 'left' as const, leader: 'none' as const },
      { pos: 380, alignment: 'right' as const, leader: 'underscore' as const },
    ];
    const res = layoutBidiTabStops(items, stops, 0, avail, 1000);
    expect(res[1].leader ?? 'none').toBe('none'); // first tab: plain left stop
    expect(res[3].leader).toBe('underscore');     // second tab: the leader
    const edge = readingEdges(items, res, 0);
    expect(edge[1]).toBeCloseTo(50, 6);  // title's leading (right) edge on 50
    expect(edge[3]).toBeCloseTo(372, 6); // page number trails at 380 − 8
  });

  it('anchors stops at the TEXT MARGIN, not the indented paragraph edge', () => {
    // A 36px leading indent (startPen 36, sample-28 TOC2's w:ind left=720): the
    // chapter number begins at the indent, but the tab's stop at pos 100 still
    // measures from the MARGIN — the following title's leading edge lands at
    // reading 100, NOT 136 (verified against the Word PDF: TOC2 titles align at
    // margin − 50.85pt despite the 36pt indent).
    const items: BidiTabItem[] = [
      { isTab: false, width: 20 },
      { isTab: true, width: 0 },
      { isTab: false, width: 50 },
    ];
    const stops = [{ pos: 100, alignment: 'left' as const, leader: 'none' as const }];
    const res = layoutBidiTabStops(items, stops, 36, avail, 1000);
    const edge = readingEdges(items, res, 36);
    expect(edge[0]).toBeCloseTo(56, 6);  // chapter: indent 36 + width 20
    expect(edge[1]).toBeCloseTo(100, 6); // stop is margin-anchored
  });

  it('is a no-op shape for a line with no tabs (widths unchanged)', () => {
    const items: BidiTabItem[] = [{ isTab: false, width: 10 }, { isTab: false, width: 20 }];
    const res = layoutBidiTabStops(items, [], 0, avail, 36);
    expect(res.map((r) => r.width)).toEqual([10, 20]);
  });
});

// ── End-to-end: render a synthetic bidi TOC row + footer row through
// renderDocumentToCanvas with a recording canvas (fixed-width glyphs), so the
// full layout→reorder→draw path is exercised (not just the pure helper). The
// mock glyph is `fontSize` px wide; positions are therefore exact and font-free.
interface FillCall { text: string; x: number; y: number; }
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[]; leaderXs: number[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const fills: FillCall[] = [];
  const leaderXs: number[] = [];
  let dir: CanvasDirection = 'ltr';
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    letterSpacing: '0px',
    get direction() { return dir; }, set direction(v: CanvasDirection) { dir = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, rect() {}, clip() {}, scale() {},
    translate() {}, setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; }, drawImage() {},
    fillText(text: string, x: number, y: number) {
      if (text === '_' || text === '.' || text === '·' || text === '-') leaderXs.push(x);
      else fills.push({ text, x, y });
    },
    strokeText() {}, fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, globalAlpha: 1,
    lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, fills, leaderXs };
}

function bidiPara(runs: unknown[], tabStops: unknown[], opts: Partial<DocParagraph> = {}): DocParagraph {
  return {
    alignment: 'left', bidi: true,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
    tabStops, runs,
    defaultFontSize: 10, defaultFontFamily: 'Arial', widowControl: false,
    ...opts,
  } as unknown as DocParagraph;
}
function txt(text: string, rtl = false) {
  return {
    type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Arial', fontFamilyEastAsia: 'Arial',
    isLink: false, background: null, vertAlign: null, hyperlink: null, rtl: rtl || undefined,
  };
}
function docOf(paras: DocParagraph[], width = 400): DocxDocumentModel {
  return {
    section: {
      pageWidth: width, pageHeight: 400, marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    settings: { defaultTabStop: 36 },
    body: paras.map((p) => ({ type: 'paragraph', ...p })),
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { Arial: 'swiss' },
  } as unknown as DocxDocumentModel;
}

describe('bidi TOC / footer rows render on one line, mirrored (issue #820)', () => {
  it('projects fresh cell membership and advances through reusable canonical order', () => {
    const segments = [{ text: 'AB' }, { isTab: true }, { text: '12.50', rtl: true },
      { text: ' ', rtl: true }, { text: 'אבגד', rtl: true }] as unknown as LayoutSeg[];
    const input = { segments, point: { segmentIndex: 2, charOffset: 2 },
      cellStart: 2, cellEnd: 5, advances: [20, 0, 50, 10, 40],
      selectedPrefixAdvance: 20, model: 'line-level' as const, resolveOrder: createLineOrderResolver() };
    // The RTL word and space precede the numeric slice physically. A new
    // acquired word width and then a shorter cell must not reuse an old offset.
    expect(physicalDecimalAlignmentPoint(input)).toBe(70);
    input.advances[4] = 80;
    expect(physicalDecimalAlignmentPoint(input)).toBe(110);
    expect(physicalDecimalAlignmentPoint({ ...input, cellEnd: 4 })).toBe(30);
  });

  it('aligns a decimal point by its visual cell position before following RTL words', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    const row = bidiPara([txt('AB'), txt('\t'), txt('12.50 אבגד הוזח', true)],
      [{ pos: 216, alignment: 'decimal', leader: 'none' }]);
    await renderDocumentToCanvas(docOf([row]), canvas, 0, { dpr: 1, width: 400 });
    const number = fills.find(f => f.text.startsWith('12.50'))!;
    // Decimal stops are measured from the right text margin. Two 10pt digits
    // precede the physical period, regardless of words reordered to its left.
    expect(number.x + 20).toBeCloseTo(400 - 216, 6);
  });

  it('retains the complete-cell anchor when the decimal lies on a later line', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    const numberText = '123456789012.50';
    const row = bidiPara([txt('AB'), txt('\t'), txt(`אבגד הוזח ${numberText} תאגת`, true)],
      [{ pos: 180, alignment: 'decimal', leader: 'none' }]);
    await renderDocumentToCanvas(docOf([row], 200), canvas, 0, { dpr: 1, width: 200 });
    const firstWord = fills.find(f => f.text === 'הוזח')!;
    // The 300pt queued cell has 100pt before the numeric slice in reading
    // order, and 30pt after its decimal. Its gap limit is 180-130-20=30pt.
    // The first line's 100pt cell therefore starts at 50, with its visible RTL
    // word 10pt past the retained space. The period itself is on the next line.
    expect(firstWord.x).toBeCloseTo(60, 6);
    expect(fills.find(f => f.text === numberText)!.y).toBeGreaterThan(firstWord.y);
  });

  it('preserves the numeric-end fallback when the actual cell has no separator', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    const row = bidiPara([txt('AB'), txt('\t'), txt('12 אבגד הוזח', true)],
      [{ pos: 216, alignment: 'decimal', leader: 'none' }]);
    await renderDocumentToCanvas(docOf([row]), canvas, 0, { dpr: 1, width: 400 });
    // The physical-period correction does not broaden the old numeric-end
    // fallback for a number followed by RTL prose (an unmeasured boundary).
    expect(fills.find(f => f.text === '12')!.x + 20).toBeCloseTo(284, 6);
  });

  it('preserves a nonterminal decimal cell while incremental context is unsupported', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    const row = bidiPara([txt('AB'), txt('\t'), txt('12.50 אבגד הוזח', true), txt('\t'), txt('9', true)],
      [{ pos: 216, alignment: 'decimal', leader: 'none' }, { pos: 350, alignment: 'left', leader: 'none' }]);
    await renderDocumentToCanvas(docOf([row]), canvas, 0, { dpr: 1, width: 400 });
    // Preserve the prior two-stop allocation: the new terminal-cell model
    // must not silently reinterpret a cell used by the incremental tab walk.
    expect(fills.find(f => f.text.startsWith('12.50'))!.x + 20).toBeCloseTo(284, 6);
    expect(fills.find(f => f.text === '9')!.x).toBeCloseTo(40, 6);
  });

  it('preserves a selected decimal paint slice containing a directional island', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    const mixed = 'Aאב.50';
    const row = bidiPara([txt('AB'), txt('\t'), { ...txt(mixed, true), fontFamilyCs: 'Arial' }],
      [{ pos: 216, alignment: 'decimal', leader: 'none' }]);
    await renderDocumentToCanvas(docOf([row]), canvas, 0, { dpr: 1, width: 400 });
    // The selected source slice contains both local parities. Keep its
    // previous allocation: the following two-digit island starts at 154.
    // A source-order prefix is not a physical caret inside the mixed slice.
    expect(fills.find(f => f.text === '50')!.x).toBeCloseTo(154, 6);
  });

  it('retains a leading tab leader when a second tab fixes its reading-order gap', async () => {
    const { canvas, fills, leaderXs } = makeRecordingCanvas();
    const row = bidiPara(
      [txt('AB', true), txt('\t', true), txt('TITLE', true), txt('\t', true), txt('9', true)],
      [
        { pos: 50, alignment: 'left', leader: 'underscore' },
        { pos: 180, alignment: 'left', leader: 'underscore' },
      ],
    );
    await renderDocumentToCanvas(docOf([row]), canvas, 0, { dpr: 1, width: 400 });
    const pageNum = fills.find((f) => f.text === '9')!;
    const title = fills.find((f) => f.text === 'TITLE')!;
    expect(pageNum.x).toBeCloseTo(210, 0);
    expect(title.x).toBeCloseTo(300, 0);
    // Both authored leaders survive fitting: the second leading gap is fixed
    // before bidi projection and still paints between the adjacent cells.
    expect(leaderXs.some((x) => x > pageNum.x + 10 && x < title.x)).toBe(true);
    expect(leaderXs.some((x) => x > title.x + 50 && x < 380)).toBe(true);
  });

  it('draws the TOC page number at the left, chapter at the right, with a leader', async () => {
    const { canvas, fills, leaderXs } = makeRecordingCanvas();
    // pageWidth 400, no margins ⇒ scale 1, content [0,400]. Right/underscore
    // stop at 380pt (near the leading/right edge distance), so the page number
    // trails at 400-380 = 20.
    const row = bidiPara(
      [txt('AB', true), txt(' ', true), txt('TITLE', true), txt('\t', true), txt('9', true)],
      [{ pos: 380, alignment: 'right', leader: 'underscore' }],
    );
    await renderDocumentToCanvas(docOf([row]), canvas, 0, { dpr: 1, width: 400 });

    const pageNum = fills.find((f) => f.text === '9');
    const chapter = fills.find((f) => f.text.trim() === 'AB');
    expect(pageNum, 'page number drawn').toBeDefined();
    expect(chapter, 'chapter number drawn').toBeDefined();
    // Page number on the visual LEFT (near x=20), chapter number on the visual
    // RIGHT (its 2 glyphs = 20px end at the right edge 400 ⇒ starts near 380).
    expect(pageNum!.x).toBeCloseTo(20, 0);
    expect(chapter!.x).toBeGreaterThan(pageNum!.x + 100);
    // A continuous underscore leader fills the gap between them.
    expect(leaderXs.length).toBeGreaterThan(3);
    expect(Math.min(...leaderXs)).toBeGreaterThan(pageNum!.x);
    expect(Math.max(...leaderXs)).toBeLessThan(chapter!.x);
    // Everything on ONE line (no wrap): the page number and chapter share a row.
    // A single-line paragraph draws each token exactly once.
    expect(fills.filter((f) => f.text === '9')).toHaveLength(1);
  });

  it('right-aligns a footer field row (Page N tab of M) to the leading edge', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    // Footer with a trailing right tab at 380: the "of" cell right-aligns near
    // the right edge. Verify the last token ends at/near the right margin.
    const row = bidiPara(
      [txt('P', true), txt('\t', true), txt('N', true)],
      [{ pos: 380, alignment: 'right', leader: 'none' }],
    );
    await renderDocumentToCanvas(docOf([row]), canvas, 0, { dpr: 1, width: 400 });
    const p = fills.find((f) => f.text === 'P');
    const n = fills.find((f) => f.text === 'N');
    expect(p).toBeDefined();
    expect(n).toBeDefined();
    // "P" is the leading (logical-first) token → visual RIGHT edge; "N" trails at
    // the mirrored stop (400-380 = 20) on the visual LEFT.
    expect(n!.x).toBeCloseTo(20, 0);
    expect(p!.x).toBeGreaterThan(n!.x);
  });

  it('paints the leader in the tab CELL GAP between page number and title (#830 follow-up)', async () => {
    const { canvas, fills, leaderXs } = makeRecordingCanvas();
    // sample-28 TOC2 shape: [chapNum][TAB→left@50][title][TAB→underscore@380][pageNum].
    // Correct visual layout (Word PDF): [pageNum][LEADER][title][blank][chapNum].
    // The regression drew [pageNum][blank][title][LEADER][chapNum] — the leader
    // migrated to the wrong cell gap because the tab→stop assignment was made
    // against the visual sequence instead of reading order.
    const row = bidiPara(
      [txt('AB', true), txt('\t', true), txt('TITLE', true), txt('\t', true), txt('9', true)],
      [
        { pos: 50, alignment: 'left', leader: 'none' },
        { pos: 380, alignment: 'right', leader: 'underscore' },
      ],
    );
    await renderDocumentToCanvas(docOf([row]), canvas, 0, { dpr: 1, width: 400 });
    const pageNum = fills.find((f) => f.text === '9');
    const title = fills.find((f) => f.text === 'TITLE');
    const chapter = fills.find((f) => f.text.trim() === 'AB');
    expect(pageNum).toBeDefined();
    expect(title).toBeDefined();
    expect(chapter).toBeDefined();
    // Page number trails at the mirrored leader stop: 400 − 380 = 20.
    expect(pageNum!.x).toBeCloseTo(20, 0);
    // Title's leading (right) edge on the mirrored left stop: 400 − 50 = 350
    // (5 glyphs × 10px ⇒ left edge 300).
    expect(title!.x).toBeCloseTo(300, 0);
    // The underscore leader lies ENTIRELY between the page number and the title
    // (the visual span of the leader tab's cell gap) — not right of the title.
    expect(leaderXs.length).toBeGreaterThan(3);
    expect(Math.min(...leaderXs)).toBeGreaterThan(pageNum!.x);
    expect(Math.max(...leaderXs)).toBeLessThan(title!.x);
  });

  it('anchors a bidi tab stop at the text margin under a leading indent', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    // Paragraph with logical-left indent 36 (physical RIGHT under bidi): the
    // chapter starts at the indented edge (400−36−20 ⇒ x=344), but the tab's
    // stop at pos 100 measures from the MARGIN, so the title's right edge lands
    // at 400−100 = 300 (left edge 250) — not 300−36.
    const row = bidiPara(
      [txt('AB', true), txt('\t', true), txt('TITLE', true)],
      [{ pos: 100, alignment: 'left', leader: 'none' }],
      { indentLeft: 36 },
    );
    await renderDocumentToCanvas(docOf([row]), canvas, 0, { dpr: 1, width: 400 });
    const chapter = fills.find((f) => f.text.trim() === 'AB');
    const title = fills.find((f) => f.text === 'TITLE');
    expect(chapter).toBeDefined();
    expect(title).toBeDefined();
    expect(chapter!.x).toBeCloseTo(344, 0);
    expect(title!.x).toBeCloseTo(250, 0);
  });
});
