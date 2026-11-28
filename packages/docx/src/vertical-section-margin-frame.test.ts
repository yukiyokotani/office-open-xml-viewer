import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { transformRect } from './layout/coordinate-space.js';
import { composeAffine } from './layout/affine.js';
import { lineFrameTransform } from './layout/line-frame.js';
import type { DocumentLayout, LayoutPage, LayoutRect, ParagraphLayout } from './layout/types.js';
import { renderDocumentToCanvas } from './renderer.js';
import { textRunsForPage } from './text-run-projection.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// ECMA-376 §17.6.11 w:pgMar names physical page edges and §17.6.20 /
// §17.18.93 (Part 4 §14.11.7) choose the line progression: tbRl lines flow
// right to left, tbLrV (≡ lrV) lines flow left to right. The first line of
// the body therefore sits against the physical right margin in tbRl and
// against the physical left margin in tbLrV, and every line starts at the
// physical top margin. Word 16 (Windows) exports of A4 sections with distinct
// margins on all four edges confirm both placements, with btLr on the tbRl
// frame. The margins here are distinct on every edge so a mirrored or
// rotated mapping cannot pass.

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => {
      const size = px();
      return {
        width: [...text].length * size,
        fontBoundingBoxAscent: size * 0.8, fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8, actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, letterSpacing: '0px',
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  } as unknown as CanvasRenderingContext2D;
}

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeCtx(); }
};

function para(text: string): BodyElement {
  return {
    type: 'paragraph',
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
      vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph & BodyElement;
}

const PAGE = { width: 600, height: 800 };
const MARGINS = { top: 60, right: 90, bottom: 80, left: 40 };

function sectionDocument(textDirection: string, texts = ['AB', 'CD']): DocxDocumentModel {
  const section = {
    pageWidth: PAGE.width, pageHeight: PAGE.height,
    marginTop: MARGINS.top, marginRight: MARGINS.right,
    marginBottom: MARGINS.bottom, marginLeft: MARGINS.left,
    headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
    textDirection,
  } as SectionProps;
  return {
    section, body: texts.map(para),
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

function layoutOf(document: DocxDocumentModel): DocumentLayout {
  return layoutDocument(
    document,
    createLayoutServices(document, { measureContext: makeCtx() }),
    { currentDateMs: 0 },
  );
}

function layoutPage(textDirection: string): LayoutPage {
  const section = {
    pageWidth: PAGE.width, pageHeight: PAGE.height,
    marginTop: MARGINS.top, marginRight: MARGINS.right,
    marginBottom: MARGINS.bottom, marginLeft: MARGINS.left,
    headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
    textDirection,
  } as SectionProps;
  const document = {
    section, body: [para('AB'), para('CD')],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
  const layout = layoutDocument(
    document,
    createLayoutServices(document, { measureContext: makeCtx() }),
    { currentDateMs: 0 },
  );
  return layout.pages[0]!;
}

/** Physical rectangles of the body paragraphs' first lines, through the
 * page's own paint matrix. */
function physicalParagraphs(page: LayoutPage): LayoutRect[] {
  const matrix = page.sectionRegions[0]!.coordinateSpace.logicalToPhysical;
  return page.layers.body
    .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
    .map((node) => transformRect(matrix, node.lines[0]!.bounds));
}

describe('vertical section margins on the physical page (§17.6.11, §17.6.20)', () => {
  // tbRlV shares the tbLrV page frame (`word-section-tbrlv-tblrv-page-frame`).
  it.each(['tbLrV', 'tbRlV'])('starts a %s body at the physical left and top margins and advances rightward', (textDirection) => {
    const page = layoutPage(textDirection);
    const [first, second] = physicalParagraphs(page);

    expect(page.section.geometry).toMatchObject({
      pageWidth: PAGE.height, pageHeight: PAGE.width,
      marginTop: MARGINS.left, marginBottom: MARGINS.right,
      marginLeft: MARGINS.top, marginRight: MARGINS.bottom,
    });
    expect(first!.xPt).toBeCloseTo(MARGINS.left, 6);
    expect(first!.yPt).toBeCloseTo(MARGINS.top, 6);
    expect(second!.xPt).toBeCloseTo(first!.xPt + first!.widthPt, 6);
  });

  it.each(['tbRl', 'btLr'])(
    'starts a %s body at the physical right and top margins and advances leftward',
    (textDirection) => {
      const [first, second] = physicalParagraphs(layoutPage(textDirection));

      expect(first!.xPt + first!.widthPt).toBeCloseTo(PAGE.width - MARGINS.right, 6);
      expect(first!.yPt).toBeCloseTo(MARGINS.top, 6);
      expect(second!.xPt + second!.widthPt).toBeCloseTo(first!.xPt, 6);
    },
  );
});

// `word-section-tblrv-line-strip`: inside its strip a tbLrV line has the
// arrangement of a tbRl line; only the strip order differs.
describe('tbLrV line strips (§17.6.20)', () => {
  const texts = ['AB\u3042\u3044', 'CD\u3046'];

  it('projects every text run of a tbLrV strip as the matching tbRl strip', () => {
    const runsOf = (textDirection: string) => {
      const layout = layoutOf(sectionDocument(textDirection, texts));
      const page = layout.pages[0]!;
      const strips = page.layers.body
        .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
        .map((node) => transformRect(
          page.sectionRegions[0]!.coordinateSpace.logicalToPhysical,
          node.lines[0]!.bounds,
        ));
      return { strips, runs: textRunsForPage(layout, 0, { scale: 1 }) };
    };
    const lr = runsOf('tbLrV');
    const rl = runsOf('tbRl');

    expect(lr.runs.map((run) => run.text)).toEqual(rl.runs.map((run) => run.text));
    expect(lr.runs.length).toBeGreaterThan(0);
    lr.runs.forEach((run, index) => {
      const other = rl.runs[index]!;
      const strip = lr.strips.find((rect) => run.x > rect.xPt && run.x <= rect.xPt + rect.widthPt + 1e-9)!;
      const otherStrip = rl.strips.find((rect) => other.x > rect.xPt && other.x <= rect.xPt + rect.widthPt + 1e-9)!;
      expect(run.transform).toBe(other.transform);
      expect(run.x - strip.xPt).toBeCloseTo(other.x - otherStrip.xPt, 9);
      expect(run.y).toBeCloseTo(other.y, 9);
      expect([run.w, run.h]).toEqual([other.w, other.h]);
    });
  });

  it('paints tbLrV glyphs through a turn, never a reflection', async () => {
    const determinants: number[] = [];
    let matrix = [1, 0, 0, 1, 0, 0];
    const stack: number[][] = [];
    const multiply = (n: number[]) => {
      const [a, b, c, d, e, f] = matrix;
      matrix = [
        a! * n[0]! + c! * n[1]!, b! * n[0]! + d! * n[1]!,
        a! * n[2]! + c! * n[3]!, b! * n[2]! + d! * n[3]!,
        a! * n[4]! + c! * n[5]! + e!, b! * n[4]! + d! * n[5]! + f!,
      ];
    };
    const base = makeCtx() as unknown as Record<string, unknown>;
    const ctx = Object.assign(base, {
      save() { stack.push([...matrix]); },
      restore() { matrix = stack.pop()!; },
      translate(x: number, y: number) { multiply([1, 0, 0, 1, x, y]); },
      scale(x: number, y: number) { multiply([x, 0, 0, y, 0, 0]); },
      rotate(angle: number) {
        multiply([Math.cos(angle), Math.sin(angle), -Math.sin(angle), Math.cos(angle), 0, 0]);
      },
      transform(a: number, b: number, c: number, d: number, e: number, f: number) {
        multiply([a, b, c, d, e, f]);
      },
      setTransform(a: number, b: number, c: number, d: number, e: number, f: number) {
        matrix = [a, b, c, d, e, f];
      },
      fillText() { determinants.push(matrix[0]! * matrix[3]! - matrix[1]! * matrix[2]!); },
      clip() {}, rect() {}, closePath() {}, fill() {}, strokeRect() {}, clearRect() {},
      setLineDash() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
      createLinearGradient() { return { addColorStop() {} }; },
      globalAlpha: 1, textBaseline: 'alphabetic',
    });
    const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx } as unknown as HTMLCanvasElement;

    await renderDocumentToCanvas(sectionDocument('tbLrV', texts), canvas, 0, { dpr: 1, width: PAGE.width });

    expect(determinants.length).toBeGreaterThan(0);
    for (const determinant of determinants) expect(determinant).toBeGreaterThan(0);
  });
});

// `word-vertical-inline-picture-upright-extent`: an inline picture stays
// upright at its wp:extent; its height runs along the line.
describe('inline pictures in vertical sections (§20.4.2.7)', () => {
  it.each(['tbRl', 'tbLrV'])('keeps a 60 x 40pt picture upright in a %s strip', (textDirection) => {
    const document = sectionDocument(textDirection, ['AB']);
    (document.body[0] as DocParagraph).runs.push({
      type: 'image', imagePath: 'word/media/p.png', mimeType: 'image/png',
      widthPt: 60, heightPt: 40, anchor: false,
    } as unknown as DocParagraph['runs'][number]);
    const page = layoutOf(document).pages[0]!;
    const paragraph = page.layers.body.find((node): node is ParagraphLayout => node.kind === 'paragraph')!;
    const line = paragraph.lines[0]!;
    const resource = line.placements.find((placement) => placement.kind === 'resource')!;
    const lineFrame = lineFrameTransform(paragraph, line);
    const region = page.sectionRegions[0]!.coordinateSpace.logicalToPhysical;
    const toPage = lineFrame ? composeAffine(region, lineFrame) : region;
    const picture = transformRect(toPage, resource.bounds);
    const strip = transformRect(toPage, line.bounds);

    expect([picture.widthPt, picture.heightPt]).toEqual([60, 40]);
    expect(strip.widthPt).toBeGreaterThanOrEqual(60);
    expect(resource.kind === 'resource' && resource.orientation).toBe('upright-physical');
  });
});

// A hand-built anchored picture (no parser anchor acquisition) is placed in
// the section-logical frame itself. The vertical-lr frame is a reflection,
// so its local inverse must be the transpose, not a quarter turn.
describe('section-logical anchored pictures in vertical sections', () => {
  const anchoredOrientation = (textDirection: string) => {
    const document = sectionDocument(textDirection, ['AB']);
    (document.body[0] as DocParagraph).runs.push({
      type: 'image', imagePath: 'word/media/p.png', mimeType: 'image/png',
      widthPt: 60, heightPt: 40, anchor: true, anchorXPt: 10, anchorYPt: 20,
      anchorXFromMargin: true, anchorYFromPara: false,
    } as unknown as DocParagraph['runs'][number]);
    const page = layoutOf(document).pages[0]!;
    const paragraph = page.layers.body.find((node): node is ParagraphLayout => node.kind === 'paragraph')!;
    const command = paragraph.drawings.flatMap((drawing) => drawing.commands)
      .find((entry) => entry.kind === 'resource');
    return command?.kind === 'resource' ? command.orientation : undefined;
  };

  it('cancels each vertical frame with its own inverse', () => {
    expect(anchoredOrientation('tbRl')).toBe('upright-physical');
    expect(anchoredOrientation('tbLrV')).toBe('upright-physical-transposed');
  });

  it('paints a transposed picture upright through the vertical-lr frame', async () => {
    const { paintRetainedResource } = await import('./paint/canvas-resource.js');
    let matrix = [0, 1, 1, 0, 0, 0]; // the vertical-lr section frame
    const multiply = (n: number[]) => {
      const [a, b, c, d, e, f] = matrix;
      matrix = [
        a! * n[0]! + c! * n[1]!, b! * n[0]! + d! * n[1]!,
        a! * n[2]! + c! * n[3]!, b! * n[2]! + d! * n[3]!,
        a! * n[4]! + c! * n[5]! + e!, b! * n[4]! + d! * n[5]! + f!,
      ];
    };
    const stack: number[][] = [];
    let painted: { matrix: number[]; rect: LayoutRect } | undefined;
    const ctx = {
      save() { stack.push([...matrix]); },
      restore() { matrix = stack.pop()!; },
      translate(x: number, y: number) { multiply([1, 0, 0, 1, x, y]); },
      rotate(angle: number) {
        multiply([Math.cos(angle), Math.sin(angle), -Math.sin(angle), Math.cos(angle), 0, 0]);
      },
      transform(a: number, b: number, c: number, d: number, e: number, f: number) {
        multiply([a, b, c, d, e, f]);
      },
    };
    paintRetainedResource('key', 'image', { xPt: 100, yPt: 200, widthPt: 40, heightPt: 60 },
      'upright-physical-transposed', {
        ctx,
        resources: { paint: (_key: string, _kind: string, rect: LayoutRect) => { painted = { matrix: [...matrix], rect }; } },
      } as never);

    // The picture is drawn with an identity linear part at its physical size:
    // the 40 x 60 logical box is 60 wide and 40 tall on the page.
    expect(painted!.matrix.slice(0, 4)).toEqual([1, 0, 0, 1]);
    expect([painted!.rect.widthPt, painted!.rect.heightPt]).toEqual([60, 40]);
  });
});

// An inline drawing is line content: in a tbLrV line it paints and projects
// through the line's block-reflected frame, like the text beside it.
describe('inline drawings in tbLrV line strips', () => {
  const shapeDocument = (textDirection: string) => {
    const document = sectionDocument(textDirection, ['AB']);
    (document.body[0] as DocParagraph).runs.push({
      type: 'shape', inline: true, widthPt: 30, heightPt: 20,
      anchorXPt: 0, anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: false,
      zOrder: 0, stroke: '#000000', strokeWidth: 1, fill: null,
      textBlocks: [{ text: 'Q', fontSizePt: 10, fontFamily: 'NotInMetrics', alignment: 'left' }],
      subpaths: [[
        { cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'lineTo', x: 30, y: 10 },
        { cmd: 'lineTo', x: 0, y: 20 }, { cmd: 'close' },
      ]],
    } as unknown as DocParagraph['runs'][number]);
    return document;
  };

  it.each(['tbRl', 'tbLrV'])('paints a %s inline shape through a turn, as its text', async (textDirection) => {
    const determinants: Array<{ op: string; determinant: number }> = [];
    let matrix = [1, 0, 0, 1, 0, 0];
    const stack: number[][] = [];
    const multiply = (n: number[]) => {
      const [a, b, c, d, e, f] = matrix;
      matrix = [
        a! * n[0]! + c! * n[1]!, b! * n[0]! + d! * n[1]!,
        a! * n[2]! + c! * n[3]!, b! * n[2]! + d! * n[3]!,
        a! * n[4]! + c! * n[5]! + e!, b! * n[4]! + d! * n[5]! + f!,
      ];
    };
    const record = (op: string) => () => {
      determinants.push({ op, determinant: matrix[0]! * matrix[3]! - matrix[1]! * matrix[2]! });
    };
    const ctx = Object.assign(makeCtx() as unknown as Record<string, unknown>, {
      save() { stack.push([...matrix]); },
      restore() { matrix = stack.pop()!; },
      translate(x: number, y: number) { multiply([1, 0, 0, 1, x, y]); },
      scale(x: number, y: number) { multiply([x, 0, 0, y, 0, 0]); },
      rotate(angle: number) {
        multiply([Math.cos(angle), Math.sin(angle), -Math.sin(angle), Math.cos(angle), 0, 0]);
      },
      transform(a: number, b: number, c: number, d: number, e: number, f: number) {
        multiply([a, b, c, d, e, f]);
      },
      setTransform(a: number, b: number, c: number, d: number, e: number, f: number) {
        matrix = [a, b, c, d, e, f];
      },
      fillText: record('text'),
      stroke: record('stroke'),
      fill: record('fill'),
      clip() {}, rect() {}, closePath() {}, strokeRect() {}, clearRect() {},
      setLineDash() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {}, ellipse() {},
      createLinearGradient() { return { addColorStop() {} }; },
      globalAlpha: 1, textBaseline: 'alphabetic', lineCap: 'butt', lineJoin: 'miter',
    });
    const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx } as unknown as HTMLCanvasElement;

    await renderDocumentToCanvas(shapeDocument(textDirection), canvas, 0, { dpr: 1, width: PAGE.width });

    const strokes = determinants.filter((entry) => entry.op === 'stroke');
    const texts = determinants.filter((entry) => entry.op === 'text');
    expect(strokes.length).toBeGreaterThan(0);
    expect(texts.length).toBeGreaterThan(0);
    for (const entry of [...strokes, ...texts]) expect(entry.determinant).toBeGreaterThan(0);
  });

  it('projects the text a tbLrV inline shape owns through its line frame', () => {
    const shapeRun = (textDirection: string) => textRunsForPage(layoutOf(shapeDocument(textDirection)), 0, { scale: 1 })
      .find((run) => run.text === 'Q');
    const rl = shapeRun('tbRl');
    const lr = shapeRun('tbLrV');
    expect(rl).toBeDefined();
    // The owned text is turned exactly as in tbRl, not reflected.
    expect(lr?.transform).toBe(rl!.transform);
  });
});
