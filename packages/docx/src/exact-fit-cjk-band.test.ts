import { describe, expect, it } from 'vitest';
import { fitCJKPrefix, fitsLineWidth } from './line-breaker/fit-search.js';
import type { MeasurementTextContext } from './layout/measurement-capabilities.js';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocxDocumentModel, SectionProps } from './types';

// Issue #2102: Word keeps N CJK characters on a body band exactly N em wide.
// The band is authored twip geometry carried as binary64 pt (an A4 page with
// 1134-twip margins and a 7238-twip right indent leaves
// 595.3 - 56.7 - 56.7 - 361.9 = 119.99999999999994pt), so an exact fit must
// not wrap on round-off. One twip narrower still wraps.

const CJK = '吾輩は猫である名前はまだ無いどこで生れたかとんと見当がつかぬ何でも';

function recordingCanvas() {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const calls: { text: string; x: number; y: number; width: number }[] = [];
  const target: Record<string | symbol, unknown> = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const p = px();
      return {
        width: [...text].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    fillText(text: string, x: number, y: number) {
      calls.push({ text, x, y, width: [...text].length * px() });
    },
    createLinearGradient: () => ({ addColorStop() {} }),
  };
  const ctx = new Proxy(target, {
    get: (object, key) => (key in object ? object[key] : () => {}),
    set: (object, key, value) => { object[key] = value; return true; },
  });
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

const SECTION: SectionProps = {
  pageWidth: 11906 / 20, pageHeight: 16838 / 20,
  marginTop: 1134 / 20, marginRight: 1134 / 20, marginBottom: 1134 / 20, marginLeft: 1134 / 20,
  headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
};

async function firstLine(sizePt: number, n: number, jc: string, narrowerTwips: number) {
  const rightIndentTwips = 9638 - Math.round(n * sizePt * 20) + narrowerTwips;
  const paragraph = {
    type: 'paragraph', alignment: jc,
    indentLeft: 0, indentRight: rightIndentTwips / 20, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text: CJK.slice(0, n + 6), bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: sizePt, color: null, fontFamily: 'MS Mincho', isLink: false,
      background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: sizePt, defaultFontFamily: 'MS Mincho', widowControl: false,
  } as unknown as BodyElement;
  const { canvas, calls } = recordingCanvas();
  await renderDocumentToCanvas({
    section: SECTION, body: [paragraph],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel, canvas, 0, { dpr: 1, width: SECTION.pageWidth });
  const top = Math.min(...calls.map((call) => call.y));
  const line = calls.filter((call) => call.y === top);
  return {
    count: [...line.map((call) => call.text).join('')].length,
    endPt: Math.max(...line.map((call) => call.x + call.width)) - SECTION.marginLeft,
  };
}

describe('exact-fit CJK body band (#2102)', () => {
  const cases = [12, 10.5].flatMap((size) => [10, 15, 20].flatMap((n) =>
    ['both', 'left', 'distribute'].map((jc) => [size, n, jc] as const)));

  it.each(cases)('%spt N=%i jc=%s keeps N characters on an exactly N-em band', async (size, n, jc) => {
    const { count, endPt } = await firstLine(size, n, jc, 0);
    expect(count).toBe(n);
    // Paint agrees with break time: the line ends on the band edge.
    expect(endPt).toBeCloseTo(n * size, 9);
  });

  it.each(cases)('%spt N=%i jc=%s wraps one character earlier on a band one twip narrower', async (size, n, jc) => {
    expect((await firstLine(size, n, jc, 1)).count).toBe(n - 1);
  });

  it('fitCJKPrefix admits an exact fit whose band carries binary64 round-off', () => {
    const band = 11906 / 20 - 1134 / 20 - 1134 / 20 - 7238 / 20;
    expect(band).toBeLessThan(120);
    const ctx = { measureText: (text: string) => ({ width: [...text].length * 12 }) } as unknown as MeasurementTextContext;
    expect(fitCJKPrefix(ctx, CJK.slice(0, 16), band)).toBe(CJK.slice(0, 10));
    expect(fitCJKPrefix(ctx, CJK.slice(0, 16), 119.95)).toBe(CJK.slice(0, 9));
    expect(fitsLineWidth(120, band)).toBe(true);
    expect(fitsLineWidth(120.05, 120)).toBe(false);
  });
});
