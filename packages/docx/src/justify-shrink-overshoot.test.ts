import { describe, it, expect } from 'vitest';
import { createCanvasFontRoute, type CanvasFontRoute } from '@silurus/ooxml-core';
import { layoutLines, type LayoutSeg, type LayoutTextSeg } from './line-layout.js';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ECMA-376 §17.18.44 defines justification, not a general word-fit tolerance.
// Word-produced Calibri and Arial controls at their natural-width boundary
// wrap for left, center, and right alignment. Earlier justified controls also
// require natural fit. Keep the same fit decision across these paragraph paths.

const FONT_PX = 12; // linear stub: each code point advances FONT_PX at scale 1

/** Linear recording canvas: measureText advances FONT_PX per code point (space
 *  included), so a token's trailing-space width is exactly one FONT_PX. This makes
 *  the fit arithmetic explicit — see TEXT4/TEXT5 below. */
function makeLinearCanvas(): HTMLCanvasElement {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  const px = () => Number(/([\d.]+)px/.exec(font)?.[1] ?? FONT_PX);
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    fontKerning: 'auto',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0, style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return canvas as unknown as HTMLCanvasElement;
}

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'serif', isLink: false, background: null,
    vertAlign: null, hyperlink: null, ...extra,
  };
}

type DocRun = DocParagraph['runs'][number];

function para(runs: DocRun[], alignment: DocParagraph['alignment']): BodyElement {
  const p: DocParagraph = {
    alignment,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs,
    defaultFontSize: FONT_PX, defaultFontFamily: 'serif',
    widowControl: false,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function textPara(text: string, alignment: DocParagraph['alignment']): BodyElement {
  return para([{ type: 'text', ...textRun(text) } as DocRun], alignment);
}

function section(pageWidth: number): SectionProps {
  return {
    pageWidth, pageHeight: 400,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    docGridCharSpace: undefined,
  } as SectionProps;
}

function doc(el: BodyElement, sec: SectionProps): DocxDocumentModel {
  return {
    section: sec, body: [el],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

/** Render one paragraph and return its visual lines, top-to-bottom: each line is
 *  the concatenated text of the onTextRun segments sharing a baseline y. */
async function renderLines(el: BodyElement, pageWidth: number): Promise<string[]> {
  const runs: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(doc(el, section(pageWidth)), makeLinearCanvas(), 0, {
    dpr: 1,
    width: pageWidth, // scale = 1 px/pt
    onTextRun: (r) => { if (r.text && r.text.trim()) runs.push(r); },
  });
  const byY = new Map<number, DocxTextRunInfo[]>();
  for (const r of runs) {
    const key = Math.round(r.y);
    let arr = byY.get(key);
    if (!arr) { arr = []; byY.set(key, arr); }
    arr.push(r);
  }
  return [...byY.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, rs]) => rs.slice().sort((p, q) => p.x - q.x).map((r) => r.text).join(''));
}

const tokens = (line: string): number => (line.match(/AAAA/g) ?? []).length;

const route = (
  familyList: string,
  scope: CanvasFontRoute['scope'] = 'native',
): CanvasFontRoute => createCanvasFontRoute(familyList, scope);

function fitSegment(
  text: string,
  fontRoute: CanvasFontRoute,
  joinPrev = false,
): LayoutTextSeg {
  return {
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: FONT_PX,
    color: null,
    fontFamily: 'Georgia',
    fontRoute,
    vertAlign: null,
    measuredWidth: 0,
    ...(joinPrev ? { joinPrev: true } : {}),
  };
}

function gluedFitLines(
  leadRoute: CanvasFontRoute,
  followerRoute: CanvasFontRoute,
): string[] {
  const georgiaRoute = route('Georgia');
  const segments: LayoutSeg[] = [
    fitSegment('X ', georgiaRoute),
    fitSegment('AA', leadRoute),
    fitSegment('B', followerRoute, true),
    fitSegment('TAIL', georgiaRoute),
  ];
  const ctx = makeLinearCanvas().getContext('2d') as CanvasRenderingContext2D;
  return layoutLines(
    ctx,
    segments,
    59.9,
    0,
    1,
    [],
    undefined,
    {},
    undefined,
    undefined,
    undefined,
    36,
    59.9,
    false,
    true,
    false,
  ).map((line) => line.segments
    .filter((segment): segment is LayoutTextSeg => 'text' in segment)
    .map((segment) => segment.text)
    .join(''));
}

// Fit arithmetic (linear stub): each "AAAA " token advances 5·12 = 60px, the
// bare word 48px, its trailing space 12px. Testing the 4th word on a line
// already holding three tokens: currentWidth = 180, wForFit = 48 ⇒ natural end
// 228. Column = 225 ⇒ overflow 3px, so the fourth word wraps.
const TEXT4 = 'AAAA AAAA AAAA AAAA';      // marginal word is the paragraph-FINAL word
const TEXT5 = 'AAAA AAAA AAAA AAAA AAAA'; // marginal word is followed by more content
const COLUMN = 225;

describe('§17.18.44 — natural word fit across alignment modes', () => {
  it('counts a candidate trailing space when the prospective line will justify', async () => {
    const lines = await renderLines(textPara('AAAA AAAA BBBB', 'both'), 108);

    // A separator following the last admitted word still participates in the
    // wrap decision for a line that will justify. The first word plus the
    // second visible word is exactly 108px; counting the second token's 12px
    // separator therefore moves it to the next line.
    expect(tokens(lines[0])).toBe(1);
  });

  it('wraps the marginal word on a line that will justify', async () => {
    // Word PDF ground truth for #698's narrow justified column still shows a
    // natural-fit break (3 tokens). This line will justify, so it gets no
    // drawable-space budget.
    const lines = await renderLines(textPara(TEXT5, 'both'), COLUMN);
    expect(lines.length).toBe(2);
    expect(tokens(lines[0])).toBe(3);
  });

  it('wraps a marginal Latin word at natural width for ordinary alignments', async () => {
    // Word-controlled Calibri and Arial tables, each with left/center/right
    // alignment, wrap even when the natural line is under 1pt too wide.
    for (const alignment of ['left', 'center', 'right'] as const) {
      const lines = await renderLines(textPara(TEXT5, alignment), COLUMN);
      expect(lines.length, alignment).toBe(2);
      expect(tokens(lines[0]), alignment).toBe(3);
    }
  });

  it("wraps a `both` paragraph's marginal final word at natural width", async () => {
    const lines = await renderLines(textPara(TEXT4, 'both'), COLUMN);
    expect(lines.length).toBe(2);
    expect(tokens(lines[0])).toBe(3);
  });

  it('wraps a marginal word before a manual <w:br/> (§17.3.3.1)', async () => {
    const el = para(
      [
        { type: 'text', ...textRun(TEXT4) } as DocRun,
        { type: 'break', breakType: 'line' } as DocRun,
        { type: 'text', ...textRun('BBBB') } as DocRun,
      ],
      'both',
    );
    const lines = await renderLines(el, COLUMN);
    expect(lines.length).toBe(3);
    expect(tokens(lines[0])).toBe(3);
    expect(lines[2]).toContain('BBBB');
  });

  it('wraps the marginal word on a `distribute` last line', async () => {
    // Word PDF ground truth for #698 still requires the natural-fit break here:
    // §17.18.44 `distribute` stretches the final line, yielding 3 tokens across
    // 2 lines. `distribute` stretches even the logical last line, so the
    // drawable-space budget is suppressed.
    const lines = await renderLines(textPara(TEXT4, 'distribute'), COLUMN);
    expect(lines.length).toBe(2);
    expect(tokens(lines[0])).toBe(3);
  });

  it('does not force a wrap when the justified content genuinely fits at natural width', async () => {
    const lines = await renderLines(textPara(TEXT4, 'both'), 240);
    expect(lines.length).toBe(1);
  });

  it('uses the same measured width for Georgia and a generic face after authored w:w', async () => {
    const el = para(
      [{
        type: 'text',
        ...textRun('AAAA AAAA AAAA AAAA', { fontFamily: 'Georgia', charScale: 0.5 }),
      } as DocRun],
      'both',
    );

    // The synthetic Canvas measures both names identically. Authored w:w may
    // scale the advances, but the family name alone cannot add fit tolerance.
    const generic = para(
      [{
        type: 'text',
        ...textRun('AAAA AAAA AAAA AAAA', { fontFamily: 'serif', charScale: 0.5 }),
      } as DocRun],
      'both',
    );
    expect(tokens((await renderLines(el, 89.2))[0])).toBe(2);
    expect(tokens((await renderLines(generic, 89.2))[0])).toBe(2);
  });

  it('wraps a marginal glued group at natural width for one or mixed routes', () => {
    const georgiaRoute = route('Georgia');
    const otherResolvedRoute = route('Registered Georgia Alias', 'registered');
    // Natural end = 60px against 59.9px. The group is atomic and starts on the
    // next line, independent of name and route splitting.
    expect(gluedFitLines(georgiaRoute, georgiaRoute)[0]).toBe('X ');
    expect(gluedFitLines(georgiaRoute, otherResolvedRoute)[0]).toBe('X ');
    expect(gluedFitLines(otherResolvedRoute, georgiaRoute)[0]).toBe('X ');
  });
});
