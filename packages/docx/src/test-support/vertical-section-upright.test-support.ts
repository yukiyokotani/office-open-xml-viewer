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

export type Direction = 'tbRl' | 'tbLrV';

function docx(direction: Direction, body: readonly string[], columns = 1, linePitch = 360): Uint8Array {
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
}

/** `linePitch`: the section's line grid pitch in twentieths of a point. */
export function layOutPages(
  direction: Direction,
  body: readonly string[],
  columns = 1,
  linePitch = 360,
): Page[] {
  const archive = new DocxArchive(docx(direction, body, columns, linePitch));
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
  const page: Page = { tables: [], strips: [] };
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
    }
  }
  return page;
  });
}


/** Loads the parser the fixtures are parsed with. */
export async function loadParser(): Promise<void> {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
}
