import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
  NumberingInfo,
} from './types';

// docx VRT covers no list paragraphs (every visual sample is numPr=0), so the
// bullet-marker draw path has no regression net. These end-to-end renderer tests
// pin two §17.9.x bullet behaviors that have silently regressed before:
//
//  1. Symbol/Wingdings glyph bullets (§17.9.6 rPr + §17.3.2.26 rFonts) are stored
//     as the FONT's own PUA code point (Symbol U+F0B7 = "•", Wingdings U+F0A7 =
//     "▪"). The renderer must normalize them to the Unicode equivalent so a
//     fallback face shows the real glyph, never raw tofu (the PUA char).
//  2. Picture bullets (§17.9.9 lvlPicBulletId → §17.9.20 numPicBullet) draw the
//     marker as an IMAGE at the hanging-indent anchor, sized from the bullet
//     drawing's extent and — when the extent is absent — the resolved marker font
//     size (the S-9 unification: one default, no magic "9pt").

const FONT_CLASSES: Record<string, string> = {
  'Times New Roman': 'roman',
};

interface DrawImageCall { bmp: unknown; x: number; y: number; w: number; h: number; }

/** Recording 2D context. Records fillText WITH the active font (markers are drawn
 *  via fillText, not reported through onTextRun) AND drawImage args (picture
 *  bullets are drawn via drawImage). Glyph advance = charCount × fontPx. */
function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: { text: string; x: number; font: string }[];
  drawImageCalls: DrawImageCall[];
} {
  let font = '16px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '16');
  const fillTextCalls: { text: string; x: number; font: string }[] = [];
  const drawImageCalls: DrawImageCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
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
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage(bmp: unknown, x: number, y: number, w: number, h: number) {
      drawImageCalls.push({ bmp, x, y, w, h });
    },
    fillText(text: string, x: number, _y: number) { fillTextCalls.push({ text, x, font }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls, drawImageCalls };
}

function run(text: string, fontFamily = 'Times New Roman'): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 16, color: null, fontFamily, fontFamilyEastAsia: fontFamily,
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as DocxTextRun;
}

function bulletDoc(num: NumberingInfo): DocxDocumentModel {
  const p: DocParagraph = {
    alignment: 'left',
    indentLeft: 36, indentRight: 0, indentFirst: -18, // hanging indent (marker in the margin)
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: num, tabStops: [],
    runs: [{ type: 'text', ...run('body') } as DocParagraph['runs'][number]],
    defaultFontSize: 16, defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as unknown as DocParagraph;
  return {
    section: {
      pageWidth: 400, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'paragraph', ...p } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: FONT_CLASSES,
  } as unknown as DocxDocumentModel;
}

async function render(
  num: NumberingInfo,
  fetchImage?: (path: string, mime: string) => Promise<Blob>,
) {
  const { canvas, fillTextCalls, drawImageCalls } = makeRecordingCanvas();
  const runs: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(bulletDoc(num), canvas, 0, {
    dpr: 1,
    width: 400, // scale = 1 (px per pt)
    onTextRun: (r) => runs.push(r),
    fetchImage,
  });
  return { runs, fillTextCalls, drawImageCalls };
}

// A 1×1 transparent GIF — the smallest valid raster the createImageBitmap stub
// can stand in for. Used as the picture-bullet media part bytes.
const ONE_BY_ONE_GIF = new Uint8Array([
  0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00, 0x80, 0x00, 0x00,
  0xff, 0xff, 0xff, 0x00, 0x00, 0x00, 0x21, 0xf9, 0x04, 0x01, 0x00, 0x00, 0x00,
  0x00, 0x2c, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0x02, 0x02,
  0x44, 0x01, 0x00, 0x3b,
]);

describe('symbol/wingdings glyph bullet markers (§17.9.6 rPr + §17.3.2.26 rFonts)', () => {
  it('draws a Symbol bullet (U+F0B7) as "•", never the raw PUA char', async () => {
    const num: NumberingInfo = {
      numId: 1, level: 0, format: 'bullet', text: '',
      indentLeft: 36, tab: 18, suff: 'tab', fontFamily: 'Symbol',
    };
    const { fillTextCalls } = await render(num);
    const bullet = fillTextCalls.find((c) => c.text === '•');
    expect(bullet, 'Symbol U+F0B7 normalized to the Unicode bullet').toBeDefined();
    // The raw private-use code point must NOT reach the canvas (it renders as tofu
    // in any fallback face).
    expect(fillTextCalls.some((c) => c.text === '')).toBe(false);
  });

  it('draws a Wingdings bullet (U+F0A7) as "▪", never the raw PUA char', async () => {
    const num: NumberingInfo = {
      numId: 2, level: 0, format: 'bullet', text: '',
      indentLeft: 36, tab: 18, suff: 'tab', fontFamily: 'Wingdings',
    };
    const { fillTextCalls } = await render(num);
    const sq = fillTextCalls.find((c) => c.text === '▪');
    expect(sq, 'Wingdings U+F0A7 normalized to the Unicode square').toBeDefined();
    expect(fillTextCalls.some((c) => c.text === '')).toBe(false);
  });
});

describe('picture bullet marker (§17.9.9 lvlPicBulletId → §17.9.20 numPicBullet)', () => {
  beforeEach(() => {
    // createImageBitmap doesn't exist in node; stub the GIF's inspected 1×1
    // decoded grid so retained-surface safety checks observe browser behavior.
    // (size is irrelevant — the draw box is driven by the §17.9.20 pt size, not
    // the bitmap's intrinsic pixels).
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_src: unknown) => ({ width: 1, height: 1, close: () => {} }) as unknown as ImageBitmap),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  // Marker text is empty for a picture bullet (lvlText is typically ""); the
  // marker IS the image. No explicit width/height ⇒ the S-9 fallback = resolved
  // marker font size (16pt here). suff:'space' so the body offset is measured off
  // the marker width (no tab-stop dependency).
  const picNum = (): NumberingInfo => ({
    numId: 3, level: 0, format: 'bullet', text: '',
    indentLeft: 36, tab: 18, suff: 'space',
    picBulletImagePath: 'word/media/image1.gif',
    picBulletMimeType: 'image/gif',
    // picBulletWidthPt / picBulletHeightPt intentionally omitted (extent absent).
  });

  async function renderPic(num: NumberingInfo = picNum()) {
    const fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([ONE_BY_ONE_GIF], { type: mime }),
    );
    const out = await render(num, fetchImage);
    return { ...out, fetchImage };
  }

  it('(a) draws the bullet image via drawImage at the hanging-indent anchor', async () => {
    const { drawImageCalls, fetchImage } = await renderPic();
    expect(fetchImage).toHaveBeenCalledWith('word/media/image1.gif', 'image/gif');
    expect(drawImageCalls).toHaveLength(1);
    // LTR marker sits at lineLeft + indFirst = (contentX 0 + indentLeft 36) +
    // indentFirst(-18) = 18px (the hanging margin = line start − markerW band).
    expect(drawImageCalls[0].x).toBeCloseTo(18, 5);
  });

  it('(b) sizes the bullet to the resolved marker font size when the extent is absent (S-9: no magic 9pt)', async () => {
    const { drawImageCalls } = await renderPic();
    // 16pt font, scale = 1 px/pt ⇒ 16px box. The pre-S-9 draw default was a magic
    // 9px; this asserts the unified font-size fallback, NOT 9.
    expect(drawImageCalls[0].w).toBeCloseTo(16, 5);
    expect(drawImageCalls[0].h).toBeCloseTo(16, 5);
    expect(drawImageCalls[0].w).not.toBeCloseTo(9, 1);
  });

  it('(b2) honors an explicit picBullet extent over the font-size fallback', async () => {
    const num = { ...picNum(), picBulletWidthPt: 24, picBulletHeightPt: 12 };
    const { drawImageCalls } = await renderPic(num);
    expect(drawImageCalls[0].w).toBeCloseTo(24, 5);
    expect(drawImageCalls[0].h).toBeCloseTo(12, 5);
  });

  it('(c) advances the body text past the marker (markerW = bullet width)', async () => {
    const { runs } = await renderPic();
    const body = runs.find((r) => r.text === 'body');
    expect(body).toBeDefined();
    // suff:'space': body x = firstLineX (18) + markerW (16) + one space (16px) = 50.
    // The key assertion is that the body is pushed RIGHT of the marker band (not
    // overlapping the bullet at x=18), i.e. ≥ marker end (18 + 16 = 34).
    expect(body!.x).toBeGreaterThanOrEqual(34);
    expect(body!.x).toBeCloseTo(50, 5);
  });
});
