/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import type { DocxDocumentModel } from './types.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { storeZip } from './conformance/generate.js';
import {
  logicalToPhysicalMatrix,
  transformRect,
  writingModeFromTextDirection,
} from './layout/coordinate-space.js';
import { logicalSectionGeometry, uprightPhysicalPageFrames } from './layout/context.js';
import type { LayoutRect } from './layout/types.js';
import type { SectionGeom } from './types.js';

// ECMA-376 §20.4.3.4/§20.4.3.5 `relativeFrom="margin"` with §17.6.11 pgMar,
// under `word-vertical-section-physical-drawing-layer`: a tbRl section
// resolves positionH/positionV on the upright physical page, so the margin
// reference is the physical margin box. Word places a margin-left/top aligned
// shape of an A4 tbRl page with 54 pt left/right and 72 pt top/bottom margins
// at (54, 72), and a margin posOffset of 144.75 pt at y = 216.75. Resolving the
// section-logical margin box unrotated swaps the two axes' margins whenever
// they differ. Every supported vertical direction must recover the physical
// margins with the exact inverse of its logical section mapping, whatever its
// paint matrix (tbLrV paints through the vertical-lr reflection). Expected
// values are physical page points.

const EMU = 12700;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PAGE_W = 11906 / 20;
const PAGE_H = 16838 / 20;

interface Margins { readonly left: number; readonly right: number; readonly top: number; readonly bottom: number }
/** The Word-measured A4 tbRl margins. */
const MEASURED: Margins = { left: 54, right: 54, top: 72, bottom: 72 };
/** Four distinct margins so any edge swap or reflection is visible. */
const ASYMMETRIC: Margins = { left: 40, right: 90, top: 60, bottom: 80 };
const VERTICAL_DIRECTIONS = ['tbRl', 'tbRlV', 'tbLrV', 'btLr'] as const;
const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);

interface Shape {
  readonly h: string;
  readonly v: string;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly wrap: string;
}

const align = (axis: 'H' | 'V', value: string) =>
  `<wp:position${axis} relativeFrom="margin"><wp:align>${value}</wp:align></wp:position${axis}>`;
const offset = (axis: 'H' | 'V', pt: number) =>
  `<wp:position${axis} relativeFrom="margin"><wp:posOffset>${Math.round(pt * EMU)}</wp:posOffset></wp:position${axis}>`;

function shapeRun(shape: Shape, index: number): string {
  const cx = Math.round(shape.widthPt * EMU);
  const cy = Math.round(shape.heightPt * EMU);
  return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" `
    + `relativeHeight="${251659264 + index}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">`
    + `<wp:simplePos x="0" y="0"/>${shape.h}${shape.v}<wp:extent cx="${cx}" cy="${cy}"/>`
    + `<wp:effectExtent l="0" t="0" r="0" b="0"/>${shape.wrap}<wp:docPr id="${index + 1}" name="s${index + 1}"/>`
    + `<a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr/><wps:spPr>`
    + `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp>'
    + '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
}

function docx(
  textBefore: number,
  shapes: readonly Shape[],
  direction = 'tbRl',
  margins: Margins = MEASURED,
): Uint8Array {
  const tw = (pt: number) => Math.round(pt * 20);
  const paragraphs = Array.from({ length: textBefore }, () =>
    '<w:p><w:r><w:t>ABCDEFGHIJKLMNOPQRST</w:t></w:r></w:p>');
  paragraphs.push(`<w:p>${shapes.map(shapeRun).join('')}</w:p>`);
  const document = `<w:document xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}">`
    + `<w:body>${paragraphs.join('')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>`
    + `<w:pgMar w:top="${tw(margins.top)}" w:right="${tw(margins.right)}" w:bottom="${tw(margins.bottom)}" `
    + `w:left="${tw(margins.left)}" w:header="851" w:footer="992" w:gutter="0"/>`
    + `<w:textDirection w:val="${direction}"/><w:docGrid w:type="lines" w:linePitch="360"/>`
    + '</w:sectPr></w:body></w:document>';
  return storeZip(new Map([
    ['[Content_Types].xml', xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '</Types>')],
    ['_rels/.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/document.xml', xml(document)],
  ]));
}

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'normal',
    measureText: (text: string) => {
      const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * px,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

interface PageOne {
  /** Logical (section-frame) boxes of the anchored drawings, in source order. */
  readonly drawings: LayoutRect[];
  /** Logical inline extents of the text placements of each text paragraph. */
  readonly textSpans: Array<Array<readonly [number, number]>>;
}

function layOutPageOne(bytes: Uint8Array): PageOne {
  const archive = new DocxArchive(bytes);
  let model: DocxDocumentModel;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
  const layout = layoutDocument(
    model,
    createLayoutServices(model, { measureContext: measureContext() }),
    { currentDateMs: 0 },
  );
  const drawings: LayoutRect[] = [];
  const textSpans: Array<Array<readonly [number, number]>> = [];
  for (const node of layout.pages[0]!.layers.body) {
    if (node.kind !== 'paragraph') continue;
    for (const drawing of node.drawings) drawings.push(drawing.flowBounds);
    const spans = node.lines.flatMap((line) => line.placements
      .filter((placement) => placement.kind === 'text')
      .map((placement) => [placement.bounds.xPt, placement.bounds.xPt + placement.bounds.widthPt] as const));
    if (spans.length) textSpans.push(spans);
  }
  return { drawings, textSpans };
}

/** Physical page boxes of the anchored drawings on page 1. Retained body
 * geometry of a vertical page is section-logical; the page paint applies the
 * direction's writing-mode matrix. */
function physicalDrawings(bytes: Uint8Array, direction = 'tbRl'): LayoutRect[] {
  const toPhysical = logicalToPhysicalMatrix(
    writingModeFromTextDirection(direction),
    { widthPt: PAGE_W, heightPt: PAGE_H },
  );
  return layOutPageOne(bytes).drawings.map((box) => transformRect(toPhysical, box));
}

const close = (actual: number | undefined, expected: number) => expect(actual).toBeCloseTo(expected, 1);

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('vertical-section margin-relative drawings resolve on the physical margins', () => {
  it.each([
    ['wrapNone', '<wp:wrapNone/>'],
    ['wrapSquare', '<wp:wrapSquare wrapText="bothSides"/>'],
  ])('places %s margin align and offset shapes in the physical margin box', (_name, wrap) => {
    const [aligned, offsetShape] = physicalDrawings(docx(0, [
      { h: align('H', 'left'), v: align('V', 'top'), widthPt: 400, heightPt: 100, wrap },
      { h: align('H', 'left'), v: offset('V', 144.75), widthPt: 300, heightPt: 40, wrap: '<wp:wrapNone/>' },
    ]));
    close(aligned?.xPt, 54);
    close(aligned?.yPt, 72);
    close(aligned?.widthPt, 400);
    close(aligned?.heightPt, 100);
    close(offsetShape?.xPt, 54);
    close(offsetShape?.yPt, 72 + 144.75);
  });

  it('aligns to the far physical margins (right/bottom)', () => {
    const [shape] = physicalDrawings(docx(0, [
      { h: align('H', 'right'), v: align('V', 'bottom'), widthPt: 100, heightPt: 50, wrap: '<wp:wrapNone/>' },
    ]));
    close(shape?.xPt, PAGE_W - 54 - 100);
    close(shape?.yPt, PAGE_H - 72 - 50);
  });

  it.each(VERTICAL_DIRECTIONS)('recovers asymmetric physical margins in a %s section', (direction) => {
    const [leading, trailing] = physicalDrawings(docx(0, [
      { h: align('H', 'left'), v: align('V', 'top'), widthPt: 100, heightPt: 50, wrap: '<wp:wrapNone/>' },
      { h: align('H', 'right'), v: align('V', 'bottom'), widthPt: 100, heightPt: 50, wrap: '<wp:wrapSquare wrapText="bothSides"/>' },
    ], direction, ASYMMETRIC), direction);
    close(leading?.xPt, ASYMMETRIC.left);
    close(leading?.yPt, ASYMMETRIC.top);
    close(trailing?.xPt, PAGE_W - ASYMMETRIC.right - 100);
    close(trailing?.yPt, PAGE_H - ASYMMETRIC.bottom - 50);
  });

  it.each([null, 'bottomToTop'] as const)(
    'inverts the logical section frame exactly (section flow %s)',
    (sectionFlow) => {
      // A bottomToTop section flow is only produced from a private wire, so it
      // is pinned at the frame boundary rather than through parsed XML.
      const physical: SectionGeom = {
        pageWidth: PAGE_W, pageHeight: PAGE_H,
        marginLeft: 40, marginRight: 90, marginTop: 60, marginBottom: 80,
        headerDistance: 0, footerDistance: 0,
      };
      const logical = logicalSectionGeometry(physical, sectionFlow);
      const frames = uprightPhysicalPageFrames({
        page: { xPt: 0, yPt: 0, widthPt: logical.pageWidth, heightPt: logical.pageHeight },
        margin: {
          xPt: logical.marginLeft,
          yPt: logical.marginTop,
          widthPt: logical.pageWidth - logical.marginLeft - logical.marginRight,
          heightPt: logical.pageHeight - logical.marginTop - logical.marginBottom,
        },
      }, sectionFlow);
      expect(frames.page).toEqual({ xPt: 0, yPt: 0, widthPt: PAGE_W, heightPt: PAGE_H });
      close(frames.margin?.xPt, 40);
      close(frames.margin?.yPt, 60);
      close(frames.margin?.widthPt, PAGE_W - 40 - 90);
      close(frames.margin?.heightPt, PAGE_H - 60 - 80);
    },
  );

  it.each([
    ['tbRl', 'right'],
    ['tbLrV', 'left'],
  ] as const)('excludes earlier %s text with the physical box of a page-start prescanned float', (direction, edge) => {
    // A later paragraph's margin-relative wrapping float is registered at page
    // start (`word-page-level-float-prescan`) through its own frame resolution,
    // so the earlier lines must already wrap around its physical position. The
    // float sits at the physical edge where this direction's lines start and
    // is wide enough to span the six earlier lines.
    const bytes = docx(6, [
      { h: align('H', edge), v: offset('V', 20), widthPt: 200, heightPt: 60, wrap: '<wp:wrapSquare wrapText="bothSides"/>' },
    ], direction, ASYMMETRIC);
    const [shape] = physicalDrawings(bytes, direction);
    close(shape?.xPt, edge === 'left' ? ASYMMETRIC.left : PAGE_W - ASYMMETRIC.right - 200);
    close(shape?.yPt, ASYMMETRIC.top + 20);
    // Physical y [80, 140] is the logical inline band the earlier lines skip.
    const top = ASYMMETRIC.top + 20;
    const bottom = top + 60;
    const { textSpans } = layOutPageOne(bytes);
    expect(textSpans).toHaveLength(6);
    for (const spans of textSpans) {
      for (const [start, end] of spans) expect(end <= top + 1e-6 || start >= bottom - 1e-6).toBe(true);
      expect(spans.some(([start]) => Math.abs(start - bottom) < 0.1)).toBe(true);
    }
  });
});
