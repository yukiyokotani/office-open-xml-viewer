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

function anchorXml(): string {
  // 200 x 120 pt at page (72, 200), square wrap on both sides.
  return '<w:r><w:drawing>'
    + '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
    + 'relativeHeight="251659265" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
    + '<wp:simplePos x="0" y="0"/>'
    + `<wp:positionH relativeFrom="page"><wp:posOffset>${72 * EMU}</wp:posOffset></wp:positionH>`
    + `<wp:positionV relativeFrom="page"><wp:posOffset>${200 * EMU}</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${200 * EMU}" cy="${120 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
    + '<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Anchor"/>'
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="a.png"/><pic:cNvPicPr/></pic:nvPicPr>'
    + '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${200 * EMU}" cy="${120 * EMU}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
    + '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
}

/** Paragraph 0 holds the anchor; the others are one labelled line each. */
function docx(paragraphCount: number): Uint8Array {
  const body = Array.from({ length: paragraphCount }, (_, index) => {
    const label = `P${String(index).padStart(3, '0')} words words words words words words words words`;
    return '<w:p><w:pPr><w:widowControl w:val="0"/>'
      + '<w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/></w:pPr>'
      + (index === 0 ? anchorXml() : `<w:r><w:t>${label}</w:t></w:r>`) + '</w:p>';
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
    + '<w:cols w:space="720"/></w:sectPr></w:body></w:document>';
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
function model(paragraphCount: number, page?: number): DocxDocumentModel {
  const archive = new DocxArchive(docx(paragraphCount));
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
    const fingerprint = (page: LayoutPage) => layoutFingerprint({ pages: [page], diagnostics: [] });
    for (const preview of previews) {
      preview.layout.pages.forEach((page, index) => {
        expect(fingerprint(page as LayoutPage)).toBe(fingerprint(final.pages[index] as LayoutPage));
      });
    }
    expect(layoutFingerprint(final as DocumentLayout)).toBe(layoutFingerprint(layout(doc)));
  }, 300_000);

  it('leave a document without page objects as it was', () => {
    // The same document with its anchor in the body: pages 1-3, the drawing
    // on page 1 (DOCX documents never carry page objects).
    const result = layout(model(60));
    expect(result.pages.length).toBe(3);
    expect(drawingPages(result)).toEqual([1]);
  });
});
