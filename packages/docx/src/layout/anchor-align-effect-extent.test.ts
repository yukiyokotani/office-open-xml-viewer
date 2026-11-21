/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { storeZip } from '../conformance/generate.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import type { DocumentLayout } from './types.js';

// word-align-includes-effect-extent: wp:align positions the extent widened by
// wp:effectExtent (ECMA-376 §20.4.2.6); posOffset positions the bare extent.
// A4 page, margins top 90 / right 36 / bottom 60 / left 50 pt; 200 x 60 pt
// shapes with effectExtent l 1.5 / t 9 / r 9.75 / b 1.5 pt.

const EMU = 12700;
const encoder = new TextEncoder();

type Position =
  | { kind: 'align'; h: string; v: string; relativeFrom?: string }
  | { kind: 'offset' }
  | { kind: 'percent'; percent: number };

function shapeXml(id: number, position: Position): string {
  const choice = (axis: 'H' | 'V') => position.kind === 'align'
    ? `<wp:align>${axis === 'H' ? position.h : position.v}</wp:align>`
    : position.kind === 'percent'
      ? `<wp14:pctPos${axis}Offset>${position.percent * 1000}</wp14:pctPos${axis}Offset>`
      : '<wp:posOffset>0</wp:posOffset>';
  const positionH = choice('H');
  const positionV = choice('V');
  const relativeFrom = position.kind === 'align' ? position.relativeFrom ?? 'margin' : 'margin';
  return '<w:r><w:drawing>'
    + '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
    + `relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">`
    + '<wp:simplePos x="0" y="0"/>'
    + `<wp:positionH relativeFrom="${relativeFrom}">${positionH}</wp:positionH>`
    + `<wp:positionV relativeFrom="${relativeFrom}">${positionV}</wp:positionV>`
    + `<wp:extent cx="${200 * EMU}" cy="${60 * EMU}"/>`
    + `<wp:effectExtent l="${1.5 * EMU}" t="${9 * EMU}" r="${9.75 * EMU}" b="${1.5 * EMU}"/>`
    + `<wp:wrapNone/><wp:docPr id="${id}" name="Shape ${id}"/><wp:cNvGraphicFramePr/>`
    + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
    + '<wps:wsp><wps:cNvSpPr/><wps:spPr>'
    + `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${200 * EMU}" cy="${60 * EMU}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>'
    + '</wps:spPr><wps:bodyPr/></wps:wsp>'
    + '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
}

function docx(
  positions: readonly Position[],
  tbRl = false,
  headerFooterTw: readonly [number, number] = [720, 720],
): Uint8Array {
  const document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    + 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    + 'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"><w:body>'
    + '<w:p><w:r><w:t>Anchors</w:t></w:r>'
    + positions.map((position, index) => shapeXml(index + 1, position)).join('')
    + '</w:p>'
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    + `<w:pgMar w:top="1800" w:right="720" w:bottom="1200" w:left="1000" w:header="${headerFooterTw[0]}" w:footer="${headerFooterTw[1]}" w:gutter="0"/>`
    + (tbRl ? '<w:textDirection w:val="tbRl"/>' : '')
    + '</w:sectPr></w:body></w:document>';
  return storeZip(new Map([
    ['[Content_Types].xml', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '</Types>')],
    ['_rels/.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      + '</Relationships>')],
    ['word/document.xml', encoder.encode(document)],
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

function layout(
  positions: readonly Position[],
  tbRl = false,
  headerFooterTw?: readonly [number, number],
): DocumentLayout {
  const archive = new DocxArchive(docx(positions, tbRl, headerFooterTw));
  let raw: unknown;
  try {
    raw = JSON.parse(new TextDecoder().decode(archive.parse()));
  } finally {
    archive.free();
  }
  const doc = normalizeInternalDocumentModel(raw as never).document;
  return layoutDocument(doc, createLayoutServices(doc, { measureContext: measureContext() }), {
    currentDateMs: 0,
  });
}

function drawingBounds(result: DocumentLayout) {
  return result.pages[0]!.layers.body
    .flatMap((node) => node.kind === 'paragraph' ? node.drawings : [])
    .map((drawing) => drawing.flowBounds);
}

/** A tbRl page is laid out section-logically; its upright physical page is
 * turned a quarter clockwise: physical x = page width − (logical y + h). */
function physicalBounds(result: DocumentLayout) {
  const physicalWidthPt = 595.3;
  return drawingBounds(result).map((bounds) => ({
    xPt: physicalWidthPt - bounds.yPt - bounds.heightPt,
    yPt: bounds.xPt,
    widthPt: bounds.heightPt,
    heightPt: bounds.widthPt,
  }));
}

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('anchored wp:align with wp:effectExtent', () => {
  it('aligns the effect-widened box to the margin edges', () => {
    const [leading, trailing] = drawingBounds(layout([
      { kind: 'align', h: 'left', v: 'top' },
      { kind: 'align', h: 'right', v: 'bottom' },
    ]));
    // Margin box: x 50..559.3, y 90..781.9 (A4 595.3 x 841.9 pt).
    expect(leading!.xPt).toBeCloseTo(50 + 1.5, 6);
    expect(leading!.yPt).toBeCloseTo(90 + 9, 6);
    expect(trailing!.xPt + trailing!.widthPt).toBeCloseTo(595.3 - 36 - 9.75, 6);
    expect(trailing!.yPt + trailing!.heightPt).toBeCloseTo(841.9 - 60 - 1.5, 6);
  });

  it('places a zero posOffset on the bare extent', () => {
    const [shape] = drawingBounds(layout([{ kind: 'offset' }]));
    expect(shape!.xPt).toBeCloseTo(50, 6);
    expect(shape!.yPt).toBeCloseTo(90, 6);
  });

  it('resolves a tbRl section against the physical margins', () => {
    const [leading, trailing, offset] = physicalBounds(layout([
      { kind: 'align', h: 'left', v: 'top' },
      { kind: 'align', h: 'right', v: 'bottom' },
      { kind: 'offset' },
    ], true));
    expect(leading!.xPt).toBeCloseTo(50 + 1.5, 6);
    expect(leading!.yPt).toBeCloseTo(90 + 9, 6);
    expect(trailing!.xPt + trailing!.widthPt).toBeCloseTo(595.3 - 36 - 9.75, 6);
    expect(trailing!.yPt + trailing!.heightPt).toBeCloseTo(841.9 - 60 - 1.5, 6);
    expect(offset!.xPt).toBeCloseTo(50, 6);
    expect(offset!.yPt).toBeCloseTo(90, 6);
  });

  it('offsets a non-zero pctPos by the left/top effect extent', () => {
    const [zero, half] = drawingBounds(layout([
      { kind: 'percent', percent: 0 },
      { kind: 'percent', percent: 50 },
    ]));
    expect(zero!.xPt).toBeCloseTo(50, 6);
    expect(zero!.yPt).toBeCloseTo(90, 6);
    expect(half!.xPt).toBeCloseTo(50 + (595.3 - 86) / 2 + 1.5, 6);
    expect(half!.yPt).toBeCloseTo(90 + (841.9 - 150) / 2 + 9, 6);
  });

  it('insets page-relative vertical inside/outside by half the header and footer distances', () => {
    const [inside, outside] = drawingBounds(layout([
      { kind: 'align', h: 'left', v: 'inside', relativeFrom: 'page' },
      { kind: 'align', h: 'right', v: 'outside', relativeFrom: 'page' },
    ], false, [1000, 500]));
    // header 50 pt, footer 25 pt.
    expect(inside!.yPt).toBeCloseTo(25 + 9, 6);
    expect(outside!.yPt + outside!.heightPt).toBeCloseTo(841.9 - 12.5 - 1.5, 6);
    expect(inside!.xPt).toBeCloseTo(1.5, 6);
    expect(outside!.xPt + outside!.widthPt).toBeCloseTo(595.3 - 9.75, 6);
  });
});
