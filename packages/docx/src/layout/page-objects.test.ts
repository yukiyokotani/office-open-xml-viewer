/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { storeZip } from '../conformance/generate.js';
import { layoutDocument } from '../document-layout.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import { layoutFingerprint } from './invariants.js';
import { normalizeLayoutOptions } from './options.js';
import { layoutDocumentProgressively, type ProgressiveLayoutPreview } from './progressive.js';
import { createLayoutServices } from '../layout-runtime.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import type { DocxDocumentModel } from '../types.js';
import type { DocumentLayout, LayoutPage } from './types.js';

// Producer page objects (ODF text:anchor-page-number, Part 3 §19.758): a
// drawing anchored to a physical page is painted on that page, wraps that
// page's body text, and extends the document to its page. DOCX never
// produces them. Lines are exact 24 pt lines measured by a linear context, so
// a US Letter page with 1 in margins holds 27 lines.

const EMU = 12700;
const PNG = Uint8Array.from(Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC',
  'base64',
));
const encoder = new TextEncoder();

type Box = Readonly<{ x: number; y: number; w: number; h: number; id: number; alignBottom?: boolean }>;
const DEFAULT_BOX: Box = { x: 72, y: 200, w: 200, h: 120, id: 1 };

function anchorXml(box: Box = DEFAULT_BOX): string {
  // Default: 200 x 120 pt at page (72, 200), square wrap on both sides.
  const { x, y, w, h, id, alignBottom } = box;
  return '<w:r><w:drawing>'
    + '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
    + `relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">`
    + '<wp:simplePos x="0" y="0"/>'
    + `<wp:positionH relativeFrom="page"><wp:posOffset>${x * EMU}</wp:posOffset></wp:positionH>`
    + (alignBottom
      ? '<wp:positionV relativeFrom="page"><wp:align>bottom</wp:align></wp:positionV>'
      : `<wp:positionV relativeFrom="page"><wp:posOffset>${y * EMU}</wp:posOffset></wp:positionV>`)
    + `<wp:extent cx="${w * EMU}" cy="${h * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
    + `<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="${id}" name="Anchor ${id}"/>`
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="a.png"/><pic:cNvPicPr/></pic:nvPicPr>`
    + '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * EMU}" cy="${h * EMU}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
    + '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
}

/** Paragraph 0 holds the anchor (and paragraph `bodyAnchorAt`, when given,
 * a second body page anchor); the others are one labelled line each. */
function docx(
  paragraphCount: number,
  box: Box = DEFAULT_BOX,
  bodyAnchorAt?: number,
  oddSectionAfter?: number,
): Uint8Array {
  const body = Array.from({ length: paragraphCount }, (_, index) => {
    const label = `P${String(index).padStart(3, '0')} words words words words words words words words`;
    const content = index === 0 ? anchorXml(box)
      : `<w:r><w:t>${label}</w:t></w:r>`
        + (index === bodyAnchorAt ? anchorXml({ x: 300, y: 500, w: 150, h: 100, id: 2 }) : '');
    // A section ending here; the next one starts on an odd page.
    const sectPr = index === oddSectionAfter
      ? '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
      : '';
    return '<w:p><w:pPr><w:widowControl w:val="0"/>'
      + '<w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/>' + sectPr + '</w:pPr>'
      + content + '</w:p>';
  }).join('');
  const document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    + 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>'
    + body
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
    + '<w:cols w:space="720"/>'
    + (oddSectionAfter !== undefined ? '<w:type w:val="oddPage"/>' : '')
    + '</w:sectPr></w:body></w:document>';
  return storeZip(new Map([
    ['[Content_Types].xml', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '</Types>')],
    ['_rels/.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      + '</Relationships>')],
    ['word/_rels/document.xml.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/a.png"/>'
      + '</Relationships>')],
    ['word/document.xml', encoder.encode(document)],
    ['word/media/a.png', PNG],
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
        width: [...text].length * px * 0.5,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

/** The parsed document; with `page`, its anchor paragraph moves out of the
 * body into a page object on that page (an empty paragraph keeps its slot). */
function model(
  paragraphCount: number,
  page?: number,
  box?: Box,
  bodyAnchorAt?: number,
  oddSectionAfter?: number,
): DocxDocumentModel {
  const archive = new DocxArchive(docx(paragraphCount, box, bodyAnchorAt, oddSectionAfter));
  let raw: DocxDocumentModel & { __pageObjects?: unknown };
  try {
    raw = JSON.parse(new TextDecoder().decode(archive.parse()));
  } finally {
    archive.free();
  }
  if (page !== undefined) {
    const [host] = raw.body.splice(0, 1, { ...raw.body[0]!, runs: [] } as typeof raw.body[number]);
    raw.__pageObjects = [{ page, body: [host] }];
  }
  return normalizeInternalDocumentModel(raw).document;
}

function layout(doc: DocxDocumentModel): DocumentLayout {
  return layoutDocument(doc, createLayoutServices(doc, { measureContext: measureContext() }), {
    currentDateMs: 0,
  });
}

/** Pages (1-based) holding a painted drawing, from any layer. */
function drawingPages(result: DocumentLayout): number[] {
  return result.pages.flatMap((page) => [...page.layers.body, ...page.layers.header]
    .some((node) => node.kind === 'paragraph' && node.drawings.length > 0) ? [page.pageIndex + 1] : []);
}

/** Left edge of each body line of a page, by its top. */
function lineStarts(result: DocumentLayout, pageNumber: number): Map<number, number> {
  const starts = new Map<number, number>();
  for (const node of result.pages[pageNumber - 1]!.layers.body) {
    if (node.kind !== 'paragraph') continue;
    for (const line of node.lines) {
      const text = line.placements.find((placement) => placement.kind === 'text');
      if (text) starts.set(Math.round(line.bounds.yPt), Math.round(text.bounds.xPt));
    }
  }
  return starts;
}

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('producer page objects', () => {
  it('paint on their page and wrap that page body only', () => {
    // 60 lines: pages 1-3. The anchor belongs to page 2.
    const result = layout(model(60, 2));
    expect(result.pages.length).toBe(3);
    expect(drawingPages(result)).toEqual([2]);
    // Page 2's lines beside the drawing (page y 200-320) start right of it
    // (x 272); page 1 and 3 lines start at the margin.
    const page2 = lineStarts(result, 2);
    const beside = [...page2].filter(([top]) => top >= 200 && top < 320);
    expect(beside.length).toBeGreaterThan(0);
    expect(beside.every(([, x]) => x >= 272)).toBe(true);
    expect([...page2].filter(([top]) => top < 176).every(([, x]) => x === 72)).toBe(true);
    expect([...lineStarts(result, 1).values()].every((x) => x === 72)).toBe(true);
  });

  it('extend the document to a page past the content', () => {
    const result = layout(model(30, 5));
    expect(result.pages.length).toBe(5);
    expect(drawingPages(result)).toEqual([5]);
  });

  it('converge: progressive previews keep only pages the final layout keeps', async () => {
    // The anchor depends on its page index alone, so every pass registers it
    // identically and published pages never change.
    const doc = model(200, 4);
    const source = layoutSourceStore(doc);
    const previews: ProgressiveLayoutPreview[] = [];
    const final = await layoutDocumentProgressively(
      source.bodyLayoutInput,
      createLayoutServices(source, { measureContext: measureContext() }),
      normalizeLayoutOptions(undefined, 0),
      { onPreview: (preview) => { previews.push(preview); } },
    );
    expect(drawingPages(final as DocumentLayout)).toEqual([4]);
    expect(previews.length).toBeGreaterThan(1);
    const fingerprint = (page: LayoutPage) => layoutFingerprint({ pages: [page], diagnostics: [] });
    for (const preview of previews) {
      preview.layout.pages.forEach((page, index) => {
        expect(fingerprint(page as LayoutPage)).toBe(fingerprint(final.pages[index] as LayoutPage));
      });
    }
    expect(layoutFingerprint(final as DocumentLayout)).toBe(layoutFingerprint(layout(doc)));
  }, 300_000);

  it('converge beside body page anchors, publishing pages progressively', async () => {
    // A body page anchor on page 5 starts anchor convergence; the page
    // object on page 2 is not a convergence read, so previews still publish.
    const doc = model(200, 2, undefined, 120);
    const source = layoutSourceStore(doc);
    const previews: ProgressiveLayoutPreview[] = [];
    const final = await layoutDocumentProgressively(
      source.bodyLayoutInput,
      createLayoutServices(source, { measureContext: measureContext() }),
      normalizeLayoutOptions(undefined, 0),
      { onPreview: (preview) => { previews.push(preview); } },
    );
    expect(previews.some((preview) => preview.layout.pages.length >= 3)).toBe(true);
    const fingerprint = (page: LayoutPage) => layoutFingerprint({ pages: [page], diagnostics: [] });
    for (const preview of previews) {
      preview.layout.pages.forEach((page, index) => {
        expect(fingerprint(page as LayoutPage)).toBe(fingerprint(final.pages[index] as LayoutPage));
      });
    }
  }, 300_000);

  it('take a page-sized wrapping drawing without exhausting a band', () => {
    const result = layout(model(60, 3, { x: 0, y: 0, w: 612, h: 792, id: 1 }));
    expect(drawingPages(result)).toEqual([3]);
  });

  it('paint on a parity blank page', () => {
    // Page 1 ends a section; the next starts on an odd page, so physical
    // page 2 is a parity blank that holds the page object.
    const result = layout(model(10, 2, undefined, undefined, 4));
    expect(result.pages.length).toBe(3);
    expect(drawingPages(result)).toEqual([2]);
  });

  it('resolve page-relative alignment against the physical page', () => {
    const result = layout(model(10, 2, { x: 72, y: 0, w: 200, h: 120, id: 1, alignBottom: true }));
    const page = result.pages[1]!;
    const drawing = page.layers.header.flatMap((node) => node.kind === 'paragraph' ? node.drawings : [])[0]!;
    expect(Math.round(drawing.flowBounds.yPt)).toBe(792 - 120);
  });

  it('leave a document without page objects as it was', () => {
    // The same document with its anchor in the body: pages 1-3, the drawing
    // on page 1 (DOCX documents never carry page objects).
    const result = layout(model(60));
    expect(result.pages.length).toBe(3);
    expect(drawingPages(result)).toEqual([1]);
  });
});
