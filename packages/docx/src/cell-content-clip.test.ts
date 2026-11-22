/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { measureContext } from './test-support/tab-fitting.test-support.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { paintTableLayout } from './paint/canvas-table.js';
import type { LayoutRect, ParagraphLayout, TableLayout } from './layout/types.js';
import type { CanvasPaintContext, PaintCanvas2D } from './paint/types.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const EMU = 12700;
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
  + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
  + 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
  + 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const encode = (text: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);

const drawingml = (positionPt: number) => `<w:r><w:rPr><w:position w:val="${positionPt * 2}"/></w:rPr>`
  + `<w:drawing><wp:inline><wp:extent cx="${40 * EMU}" cy="${30 * EMU}"/><wp:docPr id="1" name="P"/>`
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
  + '<pic:nvPicPr><pic:cNvPr id="1" name="P"/><pic:cNvPicPr/></pic:nvPicPr>'
  + '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
  + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${40 * EMU}" cy="${30 * EMU}"/></a:xfrm>`
  + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
const vml = (positionPt: number) => `<w:r><w:rPr><w:position w:val="${positionPt * 2}"/></w:rPr>`
  + '<w:pict><v:shape id="V" type="#_x0000_t75" style="width:40pt;height:30pt"><v:imagedata r:id="rIdImg" o:title=""/></v:shape></w:pict></w:r>';
const exactLine = (content: string) => '<w:p><w:pPr><w:spacing w:after="0" w:line="480" w:lineRule="exact"/></w:pPr>'
  + `<w:r><w:t xml:space="preserve">Xg </w:t></w:r>${content}</w:p>`;
const rule = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`;

/** One 216pt cell with 0.5pt rules, 5.4pt side margins and 10pt top/bottom
 * margins, holding the given paragraphs. */
function docx(paragraphs: string, body?: string): Uint8Array {
  const table = '<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/>'
    + `<w:tblBorders>${['top', 'left', 'bottom', 'right'].map(rule).join('')}</w:tblBorders><w:tblLayout w:type="fixed"/>`
    + '<w:tblCellMar><w:top w:w="200" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="200" w:type="dxa"/>'
    + '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/></w:tblGrid>'
    + `<w:tr><w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>${paragraphs}</w:tc></w:tr></w:tbl>`;
  const document = `<w:document ${NS}><w:body>${body ?? `<w:p/>${table}<w:p/>`}`
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    + 'w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>';
  return storeZip(new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rIdImg" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/p.png"/></Relationships>')],
    ['word/document.xml', encode(document)],
    ['word/media/p.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47])],
  ]));
}

function layoutPage(bytes: Uint8Array) {
  const archive = new DocxArchive(bytes);
  let parsed;
  try {
    parsed = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse())));
  } finally { archive.free(); }
  const model = parsed.document;
  return layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(5) }), { currentDateMs: 0 }).pages[0]!;
}

function layoutTable(paragraphs: string): TableLayout {
  const page = layoutPage(docx(paragraphs));
  const table = page.layers.body.find((node) => node.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a table');
  return table;
}

type Clip = Readonly<{ rect: LayoutRect; dxPt: number; dyPt: number }>;

/** Paints `table` and returns, for each painted picture, the clips active
 * when it was painted (in page points; the paint applies only translations). */
function paintedPictureClips(table: TableLayout): Array<Readonly<{ bounds: LayoutRect; clips: readonly LayoutRect[] }>> {
  let dx = 0;
  let dy = 0;
  let clips: Clip[] = [];
  let pending: LayoutRect | undefined;
  const stack: Array<Readonly<{ dx: number; dy: number; clips: Clip[] }>> = [];
  const painted: Array<Readonly<{ bounds: LayoutRect; clips: readonly LayoutRect[] }>> = [];
  const ctx = new Proxy({
    save() { stack.push({ dx, dy, clips: [...clips] }); },
    restore() { const top = stack.pop(); if (top) ({ dx, dy, clips } = top); },
    translate(x: number, y: number) { dx += x; dy += y; },
    beginPath() { pending = undefined; },
    rect(xPt: number, yPt: number, widthPt: number, heightPt: number) { pending = { xPt, yPt, widthPt, heightPt }; },
    clip() { if (pending) clips.push({ rect: pending, dxPt: dx, dyPt: dy }); },
    measureText: () => ({ width: 0 }),
  }, {
    get(target, key) {
      return key in target ? target[key as keyof typeof target] : () => undefined;
    },
    set() { return true; },
  }) as unknown as PaintCanvas2D;
  const context: CanvasPaintContext = {
    ctx, scale: 1, dpr: 1,
    resources: {
      paint(_key, kind, bounds) {
        if (kind !== 'image') return;
        painted.push({
          bounds: { ...bounds, xPt: bounds.xPt + dx, yPt: bounds.yPt + dy },
          clips: clips.map((clip) => ({ ...clip.rect, xPt: clip.rect.xPt + clip.dxPt, yPt: clip.rect.yPt + clip.dyPt })),
        });
      },
    },
  };
  paintTableLayout(table, context);
  return painted;
}

describe('word-cell-content-clip', () => {
  it('clips cell paragraphs to the cell inside its top and bottom margins, letting VML pictures escape', () => {
    const table = layoutTable(exactLine(drawingml(5) + vml(-15)));
    const cell = table.rows[0]!.cells[0]!;
    const clip = cell.contentClipBounds!;
    // Horizontally the whole cell; vertically below the 0.5pt top rule and the
    // 10pt top margin, above the 10pt bottom margin and the appended bottom rule.
    expect(clip.xPt).toBeCloseTo(cell.flowBounds.xPt, 6);
    expect(clip.widthPt).toBeCloseTo(cell.flowBounds.widthPt, 6);
    const line = (cell.blocks[0]!.layout as ParagraphLayout).lines[0]!;
    const lineTopPt = cell.flowBounds.yPt + cell.blocks[0]!.offsetPt + (line.bounds.yPt - cell.blocks[0]!.layout.flowBounds.yPt);
    expect(clip.yPt).toBeCloseTo(lineTopPt, 6);
    expect(clip.heightPt).toBeCloseTo(24, 6);

    const [picture, legacy] = paintedPictureClips(table);
    // The DrawingML picture, raised 5pt in the exact 24pt line, crosses the
    // clip top and is painted under the clip; the VML one is not clipped.
    expect(picture!.bounds.yPt).toBeLessThan(clip.yPt);
    expect(picture!.clips).toContainEqual(expect.objectContaining({
      xPt: expect.closeTo(clip.xPt, 6), yPt: expect.closeTo(clip.yPt, 6),
      widthPt: expect.closeTo(clip.widthPt, 6), heightPt: expect.closeTo(clip.heightPt, 6),
    }));
    expect(legacy!.bounds.yPt + legacy!.bounds.heightPt).toBeGreaterThan(clip.yPt + clip.heightPt);
    expect(legacy!.clips).toEqual([]);
  });

  it('keeps a picture that overlaps a neighbouring line inside the cell unclipped by that line', () => {
    const table = layoutTable(exactLine(drawingml(0)) + exactLine(drawingml(0)));
    const [, second] = paintedPictureClips(table);
    const cell = table.rows[0]!.cells[0]!;
    // The second picture rises 10.8pt above its line into the first line; only
    // the cell clip applies, and the picture lies within it.
    expect(second!.clips).toHaveLength(1);
    expect(second!.bounds.yPt).toBeGreaterThan(cell.contentClipBounds!.yPt);
  });
});

describe('word-text-box-content-clip', () => {
  it('clips a fixed horizontal text box to its shape horizontally and inside its insets vertically', () => {
    const box = '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="3810000" cy="1524000"/><wp:docPr id="9" name="T"/>'
      + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
      + '<wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1524000"/></a:xfrm>'
      + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr>'
      + `<wps:txbx><w:txbxContent>${exactLine(drawingml(5))}</w:txbxContent></wps:txbx>`
      + '<wps:bodyPr lIns="36000" tIns="360000" rIns="36000" bIns="36000"><a:noAutofit/></wps:bodyPr>'
      + '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const page = layoutPage(docx('', box));
    const paragraph = page.layers.body.find((node) => node.kind === 'paragraph' && node.textBoxes.length > 0);
    const textBox = paragraph?.kind === 'paragraph' ? paragraph.textBoxes[0] : undefined;
    const clip = textBox?.clipBounds;
    expect(clip).toBeDefined();
    // 300 x 120pt shape; insets 2.83 / 28.35 / 2.83 / 2.83pt.
    expect(clip!.xPt).toBeCloseTo(textBox!.flowBounds.xPt, 6);
    expect(clip!.widthPt).toBeCloseTo(300, 6);
    expect(clip!.yPt).toBeCloseTo(textBox!.flowBounds.yPt + 360000 / 12700, 6);
    expect(clip!.heightPt).toBeCloseTo(120 - (360000 + 36000) / 12700, 6);
  });
});
