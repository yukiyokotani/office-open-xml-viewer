/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import type { DocxDocumentModel } from './types.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { storeZip } from './conformance/generate.js';

// A vertical merge that continues over rows cut across pages (ECMA-376
// §17.4.84, §17.4.6): the rows table layout is asked for while paginating,
// and what the continuation that opens a page asks of its rows.

const counter = vi.hoisted(() => ({ rows: 0, mergeScans: 0 }));
vi.mock('./layout/table.js', async (original) => {
  const actual = await original<typeof import('./layout/table.js')>();
  return {
    ...actual,
    layoutTable: (...args: Parameters<typeof actual.layoutTable>) => {
      counter.rows += args[0].rows.length;
      return actual.layoutTable(...args);
    },
    widestSideRulesAlongMergePt: (...args: Parameters<typeof actual.widestSideRulesAlongMergePt>) => {
      counter.mergeScans += 1;
      return actual.widestSideRulesAlongMergePt(...args);
    },
  };
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);
const paragraph = (text: string) => `<w:p><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
const BORDERS = '<w:tblBorders>'
  + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((edge) => `<w:${edge} w:val="single" w:sz="4"/>`).join('')
  + '</w:tblBorders>';
/** 60 lines of one word each in a 108 pt cell: taller than a page. */
const TALL = Array.from({ length: 60 }, () => 'wwwwww').join(' ');

function docx(table: string, mode: number | null = null): Uint8Array {
  const document = `<w:document xmlns:w="${W}"><w:body>${table}${paragraph('after')}`
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1200" w:right="1800" w:bottom="1600" w:left="800" w:header="720" w:footer="720" w:gutter="0"/>'
    + '</w:sectPr></w:body></w:document>';
  return storeZip(new Map([
    ['[Content_Types].xml', xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      + '</Types>')],
    ['word/_rels/document.xml.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${REL}/settings" Target="settings.xml"/></Relationships>`)],
    ['word/settings.xml', xml(`<w:settings xmlns:w="${W}"><w:compat>`
      + (mode === null ? '' : `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/>`)
      + '</w:compat></w:settings>')],
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
    measureText(text: string) {
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

/** A two-column table whose first column is one merge over `rowCount` rows
 * and whose second column holds `text` in every row. */
function mergedTable(rowCount: number, text: string, continuationProperties = ''): string {
  const rows = Array.from({ length: rowCount }, (_unused, row) => '<w:tr>'
    + `<w:tc><w:tcPr><w:tcW w:w="1440" w:type="dxa"/>${row === 0 ? '<w:vMerge w:val="restart"/>' : `<w:vMerge/>${continuationProperties}`}</w:tcPr>${paragraph(row === 0 ? 'A' : '')}</w:tc>`
    + `<w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc></w:tr>`).join('');
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${BORDERS}</w:tblPr><w:tblGrid>`
    + `<w:gridCol w:w="1440"/><w:gridCol w:w="2160"/></w:tblGrid>${rows}</w:tbl>`;
}

function layout(table: string, mode: number | null = null) {
  const archive = new DocxArchive(docx(table, mode));
  let model: DocxDocumentModel;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
  return layoutDocument(
    model,
    createLayoutServices(model, { measureContext: measureContext() }),
    { currentDateMs: 0 },
  );
}

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('a vertical merge over rows cut across pages', () => {
  it('lays out a number of rows that grows with the table, not with its square', () => {
    const laidOut = [8, 16, 32].map((rowCount) => {
      counter.rows = 0;
      const pages = layout(mergedTable(rowCount, TALL)).pages.length;
      expect(pages).toBeGreaterThan(rowCount / 2);
      return counter.rows;
    });
    if (process.env.MERGE_WORK_REPORT) throw new Error(`rows laid out: ${laidOut.join(' / ')}`);
    // Every row is cut once. Doubling the rows doubles the work, within
    // the fixed cost of the first and last page.
    expect(laidOut[1]! - laidOut[0]!).toBeLessThanOrEqual(laidOut[0]! * 1.5);
    expect(laidOut[2]! - laidOut[1]!).toBeLessThanOrEqual((laidOut[1]! - laidOut[0]!) * 2.5);
  });

  it('resolves the side rules along a merge once for the cell that holds nested tables', () => {
    // Mode 15: the first column is one merge over N rows and its opening
    // cell holds N small spaced tables.
    const scans = [8, 16, 32].map((count) => {
      const nested = Array.from({ length: count }, () => '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        + `<w:tblCellSpacing w:w="40" w:type="dxa"/>${BORDERS}</w:tblPr><w:tblGrid><w:gridCol w:w="600"/></w:tblGrid>`
        + `<w:tr><w:tc><w:tcPr><w:tcW w:w="600" w:type="dxa"/></w:tcPr>${paragraph('n')}</w:tc></w:tr></w:tbl>${paragraph('')}`).join('');
      const rows = Array.from({ length: count }, (_unused, row) => '<w:tr>'
        + `<w:tc><w:tcPr><w:tcW w:w="1440" w:type="dxa"/>${row === 0 ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>'}</w:tcPr>${row === 0 ? nested : ''}${paragraph('')}</w:tc>`
        + `<w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/></w:tcPr>${paragraph('B')}</w:tc></w:tr>`).join('');
      counter.mergeScans = 0;
      layout(`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${BORDERS}</w:tblPr><w:tblGrid>`
        + `<w:gridCol w:w="1440"/><w:gridCol w:w="2160"/></w:tblGrid>${rows}</w:tbl>`, 15);
      return counter.mergeScans;
    });
    // The merge is scanned at all, and no more often with more nested tables.
    expect(scans[0]).toBeGreaterThan(0);
    expect(scans[1]).toBe(scans[0]);
    expect(scans[2]).toBe(scans[0]);
  });

  it('keeps every row tall enough for its content when the continuations carry large margins', () => {
    // 30 pt top and bottom margins on every continuation of the merge: the
    // cell that opens a page asks for 60 pt of its rows there.
    const margins = '<w:tcMar><w:top w:w="600" w:type="dxa"/><w:bottom w:w="600" w:type="dxa"/></w:tcMar>';
    for (const text of ['B', TALL]) {
      const pages = layout(mergedTable(text === 'B' ? 90 : 6, text, margins)).pages;
      expect(pages.length).toBeGreaterThan(1);
      for (const page of pages) {
        for (const node of page.layers.body) {
          if (node.kind !== 'table') continue;
          let rowEnd = node.flowBounds.yPt;
          for (const row of node.rows) {
            // Rows follow each other without overlap.
            expect(row.flowBounds.yPt).toBeGreaterThanOrEqual(rowEnd - 1e-6);
            rowEnd = row.flowBounds.yPt + row.flowBounds.heightPt;
            const content = row.cells[1]!;
            const blocksPt = content.blocks.reduce((sum, block) => Math.max(sum, block.offsetPt + block.layout.advancePt), 0);
            // The second column's lines end inside their row.
            expect(content.flowBounds.yPt + blocksPt).toBeLessThanOrEqual(rowEnd + 1e-6);
          }
          expect(rowEnd).toBeLessThanOrEqual(node.flowBounds.yPt + node.flowBounds.heightPt + 1e-6);
        }
      }
    }
  });
});
