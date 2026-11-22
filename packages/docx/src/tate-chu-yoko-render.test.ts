import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ECMA-376 §17.3.2.10 eastAsianLayout w:vert (horizontal-in-vertical / 縦中横).
//
// Regression: sample-26 ("９月２９日") renders the date-digit run "２９" with
// `<w:w w:val="67"/>` + `<w:eastAsianLayout w:vert="1" w:vertCompress="1"/>`.
// PR #813 (w:w) folded the 67% scale into the MEASURED advance, so the "２９"
// cell shrank to 67% along the vertical column — but the two full-width digits
// were painted full-width upright, overrunning the compressed cell and
// OVERLAPPING the following "日" (the reported "9 と 日 が重なる").
//
// Word draws "２９" as 縦中横: one cell of the vertical line with the two digits
// laid out horizontally side by side (sample-26.pdf: exactly one 12 pt cell).
// The fix makes a 縦中横 seg advance ONE em along the column and draws its whole
// text as a single upright fillText, so the following "日" abuts the next cell
// with no overlap.
//
// Because a vertical page is laid out in a SWAPPED LOGICAL frame (horizontal
// layout, then a +90° page rotation), `onTextRun` reports each cell with `y`
// as the ALONG-column position (cells stack DOWN the column via increasing y)
// and `h` as the along-column advance; `x`/`w` are the cross-column offset and
// width. The regression is about the along-column advance, so we assert on y/h.

const FONT_PX = 20; // px advance per full-width CJK char in the stub (scale 1)

/** Recording 2D context. `measureText` gives every code point FONT_PX width (so
 *  "２９" measures 2·FONT_PX naturally); `fillText` records text + x + y. The
 *  canvas is deliberately metric-free of contextual shaping — the 縦中横 draw is a
 *  single fillText whose position we assert, and the along-column advance is
 *  pinned to one em by the layout regardless of natural width. */
function makeRecordingCanvas(fontBoxHeightFactor = 1): {
  canvas: HTMLCanvasElement;
  fillTextCalls: { text: string; x: number; y: number }[];
  scaleCalls: { x: number; y: number }[];
} {
  let font = `${FONT_PX}px serif`;
  let textAlign = 'start';
  let textBaseline = 'alphabetic';
  let letterSpacing = '0px';
  const fillTextCalls: { text: string; x: number; y: number }[] = [];
  const scaleCalls: { x: number; y: number }[] = [];
  const ctx = {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    get textAlign() {
      return textAlign;
    },
    set textAlign(v: string) {
      textAlign = v;
    },
    get textBaseline() {
      return textBaseline;
    },
    set textBaseline(v: string) {
      textBaseline = v;
    },
    get letterSpacing() {
      return letterSpacing;
    },
    set letterSpacing(v: string) {
      letterSpacing = v;
    },
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    globalAlpha: 1,
    save() {},
    restore() {},
    translate() {},
    rotate() {},
    scale(x: number, y: number) { scaleCalls.push({ x, y }); },
    beginPath() {},
    closePath() {},
    moveTo() {},
    lineTo() {},
    rect() {},
    fill() {},
    stroke() {},
    clip() {},
    fillRect() {},
    strokeRect() {},
    clearRect() {},
    setTransform() {},
    resetTransform() {},
    measureText(s: string) {
      // Parse the current font px so effSize-scaled runs still measure per-cp.
      const m = /(\d+(?:\.\d+)?)px/.exec(font);
      const px = m ? parseFloat(m[1]) : FONT_PX;
      const per = px; // 1 em per code point (full-width CJK)
      return {
        width: [...s].length * per,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
        fontBoundingBoxAscent: px * 0.8 * fontBoxHeightFactor,
        fontBoundingBoxDescent: px * 0.2 * fontBoxHeightFactor,
      } as TextMetrics;
    },
    fillText(text: string, x: number, y: number) {
      fillTextCalls.push({ text, x, y });
    },
    strokeText() {},
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls, scaleCalls };
}

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: FONT_PX,
    color: null,
    // A face that is NOT in the metrics table so line height is the fallback —
    // keeps the stub deterministic (mirrors the sibling render tests).
    fontFamily: 'NotInMetrics',
    isLink: false,
    background: null,
    vertAlign: null,
    hyperlink: null,
    ...extra,
  };
}

type DocRun = DocParagraph['runs'][number];

function para(runs: DocxTextRun[]): BodyElement {
  const p: DocParagraph = {
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: runs.map((r) => ({ type: 'text', ...r }) as DocRun),
    defaultFontSize: FONT_PX,
    defaultFontFamily: 'NotInMetrics',
    widowControl: false,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

/** A vertical (tbRl) section. `pageHeight`/`pageWidth` are swapped internally by
 *  the renderer; the along-column advance is the logical x reported by
 *  onTextRun. */
function verticalSection(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 400,
    pageHeight: 600,
    marginTop: 0,
    marginRight: 0,
    marginBottom: 0,
    marginLeft: 0,
    headerDistance: 0,
    footerDistance: 0,
    titlePage: false,
    evenAndOddHeaders: false,
    textDirection: 'tbRl',
    ...overrides,
  } as SectionProps;
}

function horizontalSection(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 600,
    pageHeight: 400,
    marginTop: 0,
    marginRight: 0,
    marginBottom: 0,
    marginLeft: 0,
    headerDistance: 0,
    footerDistance: 0,
    titlePage: false,
    evenAndOddHeaders: false,
    ...overrides,
  } as SectionProps;
}

function doc(body: BodyElement[], sec: SectionProps): DocxDocumentModel {
  return {
    section: sec,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function render(
  body: BodyElement[],
  sec: SectionProps,
  fontBoxHeightFactor = 1,
): Promise<{
  runs: DocxTextRunInfo[];
  fillTextCalls: { text: string; x: number; y: number }[];
  scaleCalls: { x: number; y: number }[];
}> {
  const { canvas, fillTextCalls, scaleCalls } = makeRecordingCanvas(fontBoxHeightFactor);
  const runs: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(doc(body, sec), canvas, 0, {
    dpr: 1,
    width: sec.pageWidth,
    onTextRun: (r) => runs.push(r),
  });
  return { runs, fillTextCalls, scaleCalls };
}

// The sample-26 date runs: ９ / 月 / ２９(vert+w:w=67) / 日. Full-width digits so
// buildSegments keeps "２９" as ONE east-Asian segment.
const NINE = '９';
const GETSU = '月';
const TCY = '２９';
const NICHI = '日';

describe('§17.3.2.10 縦中横 (horizontal-in-vertical) — sample-26 "9月29日" regression', () => {
  it('advances the 縦中横 "２９" run exactly ONE cell along the vertical column', async () => {
    const { runs } = await render(
      [
        para([
          textRun(NINE),
          textRun(GETSU),
          textRun(TCY, {
            charScale: 0.67,
            eastAsianVert: true,
            eastAsianVertCompress: true,
          }),
          textRun(NICHI),
        ]),
      ],
      verticalSection(),
    );

    const nine = runs.find((r) => r.text === NINE);
    const tcy = runs.find((r) => r.text === TCY);
    expect(nine, '"９" reported').toBeDefined();
    expect(tcy, '"２９" reported').toBeDefined();

    // The 縦中横 cell's ALONG-column advance (the run's inline extent `w`; `h`
    // is the line's extent across the column) is ONE em — NOT 2·FONT_PX (full
    // width) and NOT 2·FONT_PX·0.67 (the #813 regression). One cell = FONT_PX
    // at scale 1.
    expect(tcy!.w).toBeCloseTo(FONT_PX, 6);
    // A normal single CJK cell ("９") is also one em — so "２９" occupies the same
    // along-column extent as a single character, exactly as sample-26.pdf shows.
    expect(tcy!.w).toBeCloseTo(nine!.w, 6);
  });

  it('draws "２９" as ONE upright fillText and the following "日" does not overlap', async () => {
    const { runs, fillTextCalls } = await render(
      [
        para([
          textRun(NINE),
          textRun(GETSU),
          textRun(TCY, { charScale: 0.67, eastAsianVert: true, eastAsianVertCompress: true }),
          textRun(NICHI),
        ]),
      ],
      verticalSection(),
    );

    // The whole run is painted by exactly ONE fillText carrying "２９" (縦中横),
    // not two isolated upright digit draws.
    const tcyCalls = fillTextCalls.filter((c) => c.text === TCY);
    expect(tcyCalls.length, 'one fillText for the 縦中横 run').toBe(1);

    const getsu = runs.find((r) => r.text === GETSU)!;
    const tcy = runs.find((r) => r.text === TCY)!;
    const nichi = runs.find((r) => r.text === NICHI)!;

    // Cells abut DOWN the column (along-column axis = `y`): each starts where the
    // previous ended (start + advance `w`). No overlap between "月" → "２９" →
    // "日" (the regression symptom: "２９" overran its cell and "日" was pulled
    // back over it).
    expect(tcy.y).toBeCloseTo(getsu.y + getsu.w, 6);
    expect(nichi.y).toBeCloseTo(tcy.y + tcy.w, 6);
    // "日" starts exactly ONE cell after "２９", which is one cell after "月":
    // "月" → "日" spans two 1-em cells ("２９" occupies exactly one of them).
    expect(nichi.y - getsu.y).toBeCloseTo(2 * FONT_PX, 6);
  });

  it('retains vertCompress from acquisition so tall font metrics fit the one-em cell', async () => {
    const { runs, scaleCalls } = await render(
      [
        para([
          textRun(TCY, { charScale: 0.67, eastAsianVert: true, eastAsianVertCompress: true }),
          textRun(NICHI),
        ]),
      ],
      verticalSection(),
      1.25,
    );

    const tcy = runs.find((run) => run.text === TCY)!;
    const following = runs.find((run) => run.text === NICHI)!;
    expect(following.y - tcy.y).toBeCloseTo(FONT_PX, 6);
    expect(scaleCalls).toContainEqual({ x: 0.67, y: 0.8 });
  });

  it('flags the 縦中横 run info with eastAsianVert so the overlays clamp it (#836)', async () => {
    const { runs } = await render(
      [
        para([
          textRun(NINE),
          textRun(TCY, { charScale: 0.67, eastAsianVert: true, eastAsianVertCompress: true }),
          textRun(NICHI),
        ]),
      ],
      verticalSection(),
    );
    const tcy = runs.find((r) => r.text === TCY)!;
    const nine = runs.find((r) => r.text === NINE)!;
    // Only the 縦中横 run carries the flag; ordinary vertical cells do not.
    expect(tcy.eastAsianVert).toBe(true);
    expect(nine.eastAsianVert).toBeUndefined();
    // The reported width IS the drawn one-em cell (what the overlay clamps to),
    // not the natural 2·FONT_PX. On a vertical page onTextRun's `w` is the
    // cross-column cell width = one em.
    expect(tcy.w).toBeCloseTo(FONT_PX, 6);
  });

  it('does NOT flag eastAsianVert on a horizontal page (the run is not 縦中横 there)', async () => {
    const { runs } = await render(
      [
        para([
          textRun(TCY, { charScale: 0.67, eastAsianVert: true, eastAsianVertCompress: true }),
        ]),
      ],
      horizontalSection(),
    );
    const tcy = runs.find((r) => r.text === TCY)!;
    expect(tcy.eastAsianVert).toBeUndefined();
  });

  it('leaves a horizontal page byte-identical (縦中横 flag inert without tbRl)', async () => {
    // The SAME runs on a HORIZONTAL page: eastAsianVert is meaningful only in
    // vertical text (§17.3.2.10), so buildSegments never sets the flag and the
    // "２９" run keeps its w:w=67 advance (2·FONT_PX·0.67), not one em.
    const { runs } = await render(
      [
        para([
          textRun(NINE),
          textRun(GETSU),
          textRun(TCY, { charScale: 0.67, eastAsianVert: true, eastAsianVertCompress: true }),
          textRun(NICHI),
        ]),
      ],
      horizontalSection(),
    );
    const tcy = runs.find((r) => r.text === TCY)!;
    // Horizontal: advance = natural (2·FONT_PX) × 0.67 = 26.8 — the w:w path,
    // unchanged by the 縦中横 code (which is gated on the vertical flag).
    expect(tcy.w).toBeCloseTo(2 * FONT_PX * 0.67, 6);
  });
});
