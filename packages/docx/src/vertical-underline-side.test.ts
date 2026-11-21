import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { ParagraphLayout, TabPlacement, TextPlacement } from './layout/types.js';
import { transformPoint, transformRect } from './layout/coordinate-space.js';
import { composeAffine } from './layout/affine.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// `word-vertical-underline-over-side`: an upright vertical line carries its
// underline on the over side of the glyph column (logical top in the
// section frame); horizontal and all-rotated btLr lines keep it below.

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

type Run = Readonly<{ text: string; fontSize: number; underline?: boolean; strikethrough?: boolean }>;

function layoutLine(runs: readonly Run[], textDirection?: string) {
  const paragraph = {
    type: 'paragraph',
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: runs.map((run) => ({
      type: 'text', text: run.text, bold: false, italic: false,
      underline: run.underline === true, strikethrough: run.strikethrough === true,
      fontSize: run.fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
      vertAlign: null, hyperlink: null,
    })),
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph & BodyElement;
  const section = {
    pageWidth: 600, pageHeight: 800,
    marginTop: 60, marginRight: 90, marginBottom: 80, marginLeft: 40,
    headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
    ...(textDirection ? { textDirection } : {}),
  } as SectionProps;
  const document = {
    section, body: [paragraph],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
  const page = layoutDocument(
    document,
    createLayoutServices(document, { measureContext: makeCtx() }),
    { currentDateMs: 0 },
  ).pages[0]!;
  const node = page.layers.body.find((entry): entry is ParagraphLayout => entry.kind === 'paragraph')!;
  const line = node.lines[0]!;
  // The paint matrix of this line: its section frame, composed with any
  // retained line frame the paragraph carries.
  const region = page.sectionRegions[0]!.coordinateSpace.logicalToPhysical;
  const lineFrame = (node as { lineFrame?: string }).lineFrame === 'block-reflected'
    ? { a: 1, b: 0, c: 0, d: -1, e: 0, f: 2 * line.bounds.yPt + line.bounds.heightPt }
    : null;
  const toPage = lineFrame ? composeAffine(region, lineFrame) : region;
  return { line, toPage };
}

function underlined(textDirection?: string, strikethrough = false): TextPlacement {
  const { line } = layoutLine([{ text: 'Ab\u3042', fontSize: 10, underline: !strikethrough, strikethrough }], textDirection);
  return line.placements.find((entry): entry is TextPlacement => entry.kind === 'text')!;
}

/** Physical x of a run's underline minus the physical centre of its font box. */
function physicalUnderlineOffset(textDirection: string): number {
  const { line, toPage } = layoutLine([{ text: 'Ab\u3042', fontSize: 10, underline: true }], textDirection);
  const placement = line.placements.find((entry): entry is TextPlacement => entry.kind === 'text')!;
  const decoration = placement.decorations.find((entry) => entry.kind === 'underline')!;
  const box = transformRect(toPage, placement.highlightBounds!);
  return transformPoint(toPage, decoration.from).xPt - (box.xPt + box.widthPt / 2);
}

/** Signed offset of the decoration from the run's font-box centre; negative
 * is the logical top (the over side once the frame turns). */
function offsetFromCentre(placement: TextPlacement, kind: 'underline' | 'strikethrough'): number {
  const decoration = placement.decorations.find((entry) => entry.kind === kind)!;
  const box = placement.highlightBounds!;
  return decoration.from.yPt - (box.yPt + box.heightPt / 2);
}

describe('underline side in vertical sections (§17.3.2.40)', () => {
  it.each(['tbRl'])('mirrors a %s underline across its font box in the section frame', (direction) => {
    const horizontal = offsetFromCentre(underlined(), 'underline');
    expect(horizontal).toBeGreaterThan(0);
    expect(offsetFromCentre(underlined(direction), 'underline')).toBeCloseTo(-horizontal, 9);
  });

  it('keeps the all-rotated btLr underline and every strikethrough in place', () => {
    const horizontal = offsetFromCentre(underlined(), 'underline');
    expect(offsetFromCentre(underlined('btLr'), 'underline')).toBeCloseTo(horizontal, 9);
    expect(offsetFromCentre(underlined('tbRl', true), 'strikethrough'))
      .toBeCloseTo(offsetFromCentre(underlined(undefined, true), 'strikethrough'), 9);
  });

  it.each(['tbRl', 'tbLrV'])('paints a %s underline on the physical right of its glyph column', (direction) => {
    const horizontal = offsetFromCentre(underlined(), 'underline');
    expect(physicalUnderlineOffset(direction)).toBeCloseTo(horizontal, 9);
  });

  it('centres an underlined tab on its own font box, not a larger neighbouring run', () => {
    const tabOffset = (textDirection?: string) => {
      const { line } = layoutLine([
        { text: 'A', fontSize: 30 },
        { text: '\t', fontSize: 10, underline: true },
      ], textDirection);
      const tab = line.placements.find((entry): entry is TabPlacement => entry.kind === 'tab')!;
      const decoration = tab.decorations!.find((entry) => entry.kind === 'underline')!;
      // The 10pt tab's font box (ascent 8, descent 2) centres 3pt above its
      // baseline.
      return decoration.from.yPt - (line.baselinePt - 3);
    };
    expect(tabOffset()).toBeGreaterThan(0);
    expect(tabOffset('tbRl')).toBeCloseTo(-tabOffset(), 9);
  });
});
