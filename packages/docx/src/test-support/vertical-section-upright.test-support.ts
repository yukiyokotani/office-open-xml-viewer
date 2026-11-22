/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import type { DocxDocumentModel } from '../types.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import { storeZip } from '../conformance/generate.js';

// Synthetic vertical sections for `word-vertical-section-upright-block-table`,
// in the geometry of its Word controls: A4, margins left 40, right 90, top 60,
// bottom 80 pt, an 18 pt line grid, 10.5 pt text and single-bordered tables
// 72 + 108 pt wide. Table heights come from the test metrics.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const PAGE_W = 11906 / 20;
export const MARGIN = { left: 40, right: 90, top: 60 };
export const PITCH = 18;
/** Physical y of the column end: page height less the 80 pt bottom margin. */
export const COLUMN_END = 16838 / 20 - 80;
export const TABLE_W = 180;
const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);

export const paragraph = (text: string) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
export const P = paragraph('あいうえお');
const cell = (width: number, text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`;
export const table = (rows: number, rowHeightPt?: number) => '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
  + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((edge) => `<w:${edge} w:val="single" w:sz="4"/>`).join('')
  + '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="2160"/></w:tblGrid>'
  + Array.from({ length: rows }, (_, row) =>
    `<w:tr>${rowHeightPt === undefined ? '' : `<w:trPr><w:trHeight w:val="${rowHeightPt * 20}" w:hRule="atLeast"/></w:trPr>`}`
    + `${cell(1440, `A${row + 1}`)}${cell(2160, `B${row + 1}`)}</w:tr>`).join('')
  + '</w:tbl>';
export const T = table(2);

/** A paragraph run anchoring a square-wrapped rectangle at page position
 * (x, y), `width` x `height` pt; `verticalFrom` 'paragraph' makes y relative
 * to the anchoring paragraph instead, and `horizontalFrom` 'character' makes
 * x relative to the anchor's character. */
export const anchoredRectangle = (
  x: number, y: number, width: number, height: number, verticalFrom: 'page' | 'paragraph' = 'page',
  horizontalFrom: 'page' | 'character' = 'page',
) =>
  '<w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
  + ' distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
  + '<wp:simplePos x="0" y="0"/>'
  + `<wp:positionH relativeFrom="${horizontalFrom}"><wp:posOffset>${Math.round(x * 12700)}</wp:posOffset></wp:positionH>`
  + `<wp:positionV relativeFrom="${verticalFrom}"><wp:posOffset>${Math.round(y * 12700)}</wp:posOffset></wp:positionV>`
  + `<wp:extent cx="${width * 12700}" cy="${height * 12700}"/><wp:wrapSquare wrapText="bothSides"/>`
  + '<wp:docPr id="1" name="Rect"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
  + '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
  + '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr/><wps:spPr>'
  + `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${width * 12700}" cy="${height * 12700}"/></a:xfrm>`
  + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp>'
  + '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';

/** A one-row table, one 12 pt wide column, whose cell holds `paragraphs`
 * paragraphs of a PAGE field, in roman numerals unless `format` says
 * otherwise: one character to a line, so the cell's text is as many lines
 * long as the page number has characters. `rowPt` is the row's minimum
 * height. */
export const romanPageTable = (paragraphs: number, rowPt: number, format: 'Roman' | 'Arabic' = 'Roman') => {
  const field = `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>`
    + `<w:fldSimple w:instr=" PAGE \\* ${format} "><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p>`;
  return '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="240"/></w:tblGrid>'
    + `<w:tr><w:trPr><w:trHeight w:val="${rowPt * 20}" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="240" w:type="dxa"/>`
    + '<w:tcMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar></w:tcPr>'
    + `${field.repeat(paragraphs)}</w:tc></w:tr></w:tbl>`;
};

export type Direction = 'tbRl' | 'tbLrV';

function docx(
  direction: Direction,
  body: readonly string[],
  columns = 1,
  linePitch = 360,
  numberingLevel?: string,
  styles?: string,
): Uint8Array {
  const document = `<w:document xmlns:w="${W}"><w:body>${body.join('')}`
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1200" w:right="1800" w:bottom="1600" w:left="800" w:header="720" w:footer="720" w:gutter="0"/>'
    + (columns > 1 ? `<w:cols w:num="${columns}" w:space="425"/>` : '')
    + `<w:textDirection w:val="${direction}"/><w:docGrid w:type="lines" w:linePitch="${linePitch}"/>`
    + '</w:sectPr></w:body></w:document>';
  return storeZip(new Map([
    ['[Content_Types].xml', xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + (numberingLevel === undefined ? ''
        : '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>')
      + (styles === undefined ? ''
        : '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>')
      + '</Types>')],
    ['_rels/.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/document.xml', xml(document)],
    ...(numberingLevel === undefined && styles === undefined ? [] : [
      ['word/_rels/document.xml.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + (numberingLevel === undefined ? '' : `<Relationship Id="rIdNum" Type="${REL}/numbering" Target="numbering.xml"/>`)
        + (styles === undefined ? '' : `<Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/>`)
        + '</Relationships>')],
    ] as Array<[string, Uint8Array]>),
    ...(numberingLevel === undefined ? [] : [
      ['word/numbering.xml', xml(`<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0">${numberingLevel}</w:abstractNum>`
        + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')],
    ] as Array<[string, Uint8Array]>),
    ...(styles === undefined ? [] : [
      ['word/styles.xml', xml(`<w:styles xmlns:w="${W}">${styles}</w:styles>`)],
    ] as Array<[string, Uint8Array]>),
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

export interface Page {
  /** Upright table boxes on the physical page, in source order: the grid's
   * left edge and the physical x extent its vertical rules paint. */
  readonly tables: Array<Readonly<{
    left: number;
    top: number;
    bottom: number;
    paintedLeft: number;
    paintedRight: number;
  }>>;
  /** Line strips per body paragraph, in source order: the strip's distance
   * from the first strip of the page along the block axis, and the physical
   * y at which its inline flow starts. */
  readonly strips: Array<Array<Readonly<{ block: number; inlineStart: number }>>>;
  /** Anchored drawings, in source order: the physical left, top, width and
   * height of their frame. */
  readonly drawings: Array<Readonly<{ left: number; top: number; width: number; height: number }>>;
}

/** `linePitch`: the section's line grid pitch in twentieths of a point.
 * `numberingLevel`: the `w:lvl` of level 0 of the list `w:numId` 1.
 * `styles`: the `w:style` elements of the styles part. */
export function layOutPages(
  direction: Direction,
  body: readonly string[],
  columns = 1,
  linePitch = 360,
  numberingLevel?: string,
  styles?: string,
): Page[] {
  const archive = new DocxArchive(docx(direction, body, columns, linePitch, numberingLevel, styles));
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
  // The section-logical block origin: the right margin in tbRl (lines advance
  // leftward from it), the left margin in tbLrV.
  const blockOrigin = direction === 'tbRl' ? MARGIN.right : MARGIN.left;
  // Each column's strips restart at the block origin.
  return layout.pages.map((laidOut) => {
  const page: Page = { tables: [], strips: [], drawings: [] };
  for (const node of laidOut.layers.body) {
    if (node.kind === 'table') {
      const rules = node.borders.filter((border) => border.from.xPt === border.to.xPt);
      page.tables.push({
        left: node.flowBounds.xPt,
        top: node.flowBounds.yPt,
        bottom: node.flowBounds.yPt + node.advancePt,
        paintedLeft: Math.min(...rules.map((rule) => rule.from.xPt - rule.widthPt / 2)),
        paintedRight: Math.max(...rules.map((rule) => rule.from.xPt + rule.widthPt / 2)),
      });
    } else if (node.kind === 'paragraph') {
      page.strips.push(node.lines.map((line) => ({
        block: line.bounds.yPt - blockOrigin,
        inlineStart: line.bounds.xPt,
      })));
      for (const drawing of node.drawings) {
        // Section-logical frame: x inline (physical y), y block.
        page.drawings.push({
          left: direction === 'tbRl'
            ? PAGE_W - drawing.flowBounds.yPt - drawing.flowBounds.heightPt : drawing.flowBounds.yPt,
          top: drawing.flowBounds.xPt,
          width: drawing.flowBounds.heightPt,
          height: drawing.flowBounds.widthPt,
        });
      }
    }
  }
  return page;
  });
}


/** Loads the parser the fixtures are parsed with. */
export async function loadParser(): Promise<void> {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
}
