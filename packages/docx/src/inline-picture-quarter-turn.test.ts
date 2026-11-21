import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type { BodyElement, DocxDocumentModel, SectionProps } from './types';

// Compatibility rule `word-inline-picture-quarter-turn-footprint`: Word lays
// out an inline picture turned a quarter turn by its a:xfrm rot in the turned
// footprint (wp:extent cy wide, cx tall), not in the authored wp:extent box.

const FONT = 'Synthetic Untabled Serif';

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => ({
      width: [...text].length * px(),
      fontBoundingBoxAscent: px() * 0.8, fontBoundingBoxDescent: px() * 0.2,
      actualBoundingBoxAscent: px() * 0.8, actualBoundingBoxDescent: px() * 0.2,
    }) as TextMetrics,
  } as unknown as CanvasRenderingContext2D;
}

const text = (value: string) => ({
  type: 'text', text: value, bold: false, italic: false, underline: false, strikethrough: false,
  fontSize: 10, color: null, fontFamily: FONT, fontFamilyEastAsia: '',
  isLink: false, background: null, vertAlign: null, hyperlink: null,
});

const paragraph = (runs: unknown[]): BodyElement => ({
  type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
  spaceBefore: 0, spaceAfter: 0, lineSpacing: { value: 1, rule: 'auto', explicit: true },
  numbering: null, tabStops: [], runs, defaultFontSize: 10, defaultFontFamily: FONT,
  widowControl: false,
} as unknown as BodyElement);

/** A 100 x 40pt wp:extent picture (effect extent l=1 t=2 r=3 b=4) followed by
 * text on its line, then a second paragraph. */
function layoutPicture(rotation: number) {
  const picture = {
    type: 'image', imagePath: 'word/media/picture.png', mimeType: 'image/png',
    widthPt: 100, heightPt: 40, anchor: false, rotation,
    inlineEffectExtent: [1, 2, 3, 4],
  };
  const model = {
    section: {
      pageWidth: 200, pageHeight: 1000,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [paragraph([picture, text('a')]), paragraph([text('Next')])],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [FONT]: 'roman' },
  } as unknown as DocxDocumentModel;
  const services = createLayoutServices(model, {
    localMetrics: testFontSnapshot([{ family: FONT }]), measureContext: measureContext(),
  });
  const [first, second] = layoutDocument(model, services).pages[0]!.layers.body;
  if (first?.kind !== 'paragraph' || second?.kind !== 'paragraph') throw new Error('expected two paragraphs');
  const line = first.lines[0]!;
  const image = line.placements.find((placement) => placement.kind === 'resource');
  const tail = line.placements.find((placement) => placement.kind === 'text');
  if (image?.kind !== 'resource' || !tail) throw new Error('expected the picture and its tail text');
  return { image: image.bounds, tailX: tail.bounds.xPt, nextTop: second.lines[0]!.bounds.yPt };
}

describe('inline picture turned by a quarter turn (word-inline-picture-quarter-turn-footprint)', () => {
  it.each([90, 270, -90])('lays out a %s-degree picture in its turned footprint', (rotation) => {
    const { image, tailX, nextTop } = layoutPicture(rotation);
    // Turned footprint: 40 + l + r wide, 100 + t + b tall. The authored
    // 100 x 40 object paints centred in it, so after the turn its ink covers
    // x 11..51 and y 12..112 exactly.
    expect(tailX).toBeCloseTo(10 + 1 + 40 + 3, 6);
    expect(image.widthPt).toBeCloseTo(100, 6);
    expect(image.heightPt).toBeCloseTo(40, 6);
    expect(image.xPt + image.widthPt / 2).toBeCloseTo(11 + 20, 6);
    expect(image.yPt + image.heightPt / 2).toBeCloseTo(12 + 50, 6);
    expect(nextTop).toBeGreaterThanOrEqual(10 + 2 + 100 + 4 - 1e-6);
  });

  it.each([0, 180, 30])('keeps the authored wp:extent box for a %s-degree picture', (rotation) => {
    const { image, tailX, nextTop } = layoutPicture(rotation);
    expect(tailX).toBeCloseTo(10 + 1 + 100 + 3, 6);
    expect(image.xPt).toBeCloseTo(11, 6);
    expect(image.yPt + image.heightPt / 2).toBeCloseTo(12 + 20, 6);
    expect(nextTop).toBeLessThan(10 + 2 + 100 + 4);
  });
});
