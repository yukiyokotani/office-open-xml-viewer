/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { measureContext } from './test-support/tab-fitting.test-support.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { paintTableLayout } from './paint/canvas-table.js';
import { hitTestDocxElementContext } from './element-context.js';
import { createPaintResourceRegistry } from './layout/paint-resources.js';
import { textRunsForPage } from './text-run-projection.js';
import { buildDocxHighlightLayer } from './find-highlight-layer.js';
import { clippedBoxBounds, clippedBoxVisible } from './run-clip.js';
import { clusterBoxes } from './layout/visual-cluster-geometry.js';
import { documentRequiresDomVerticalGlyphLayout } from './vertical-render-capability.js';
import { resolveDocxCommentThreads } from './comments.js';
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
function docx(
  paragraphs: string, body?: string, rowProperties = '', verticalMarginTwips = 200, rules = true, cellProperties = '',
): Uint8Array {
  const table = '<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/>'
    + (rules ? `<w:tblBorders>${['top', 'left', 'bottom', 'right'].map(rule).join('')}</w:tblBorders>` : '')
    + '<w:tblLayout w:type="fixed"/>'
    + `<w:tblCellMar><w:top w:w="${verticalMarginTwips}" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>`
    + `<w:bottom w:w="${verticalMarginTwips}" w:type="dxa"/>`
    + '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/></w:tblGrid>'
    + `<w:tr>${rowProperties ? `<w:trPr>${rowProperties}</w:trPr>` : ''}<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/>${cellProperties}</w:tcPr>${paragraphs}</w:tc></w:tr></w:tbl>`;
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

function layoutDocx(bytes: Uint8Array, context = measureContext(5)) {
  const archive = new DocxArchive(bytes);
  let parsed;
  try {
    parsed = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse())));
  } finally { archive.free(); }
  const model = parsed.document;
  return layoutDocument(model, createLayoutServices(model, { measureContext: context }), { currentDateMs: 0 });
}

function layoutPage(bytes: Uint8Array) {
  return layoutDocx(bytes).pages[0]!;
}

function layoutTable(paragraphs: string): TableLayout {
  const page = layoutPage(docx(paragraphs));
  const table = page.layers.body.find((node) => node.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a table');
  return table;
}

type Clip = Readonly<{ rect: LayoutRect; dxPt: number; dyPt: number }>;
type Painted = Readonly<{ bounds: LayoutRect; clips: readonly LayoutRect[] }>;

/** Paints `table` and returns, in paint order, each painted picture with the
 * clips active when it was painted (in page points; the paint applies only
 * translations), and the number of save and translate calls. */
function paintTable(table: TableLayout): Readonly<{
  painted: Painted[]; saves: number; translates: number;
  inks: Array<Readonly<{ xPt: number; clips: number }>>;
}> {
  let dx = 0;
  let dy = 0;
  let clips: Clip[] = [];
  let pending: LayoutRect | undefined;
  let saves = 0;
  let translates = 0;
  const stack: Array<Readonly<{ dx: number; dy: number; clips: Clip[] }>> = [];
  const painted: Painted[] = [];
  // Rule ink (strokes and filled hairlines) with its page x and active clips.
  const inks: Array<Readonly<{ xPt: number; clips: number }>> = [];
  let pathX = 0;
  const ctx = new Proxy({
    moveTo(x: number) { pathX = x + dx; },
    stroke() { inks.push({ xPt: pathX, clips: clips.length }); },
    fillRect(x: number) { inks.push({ xPt: x + dx, clips: clips.length }); },
    save() { saves++; stack.push({ dx, dy, clips: [...clips] }); },
    restore() { const top = stack.pop(); if (top) ({ dx, dy, clips } = top); },
    translate(x: number, y: number) { translates++; dx += x; dy += y; },
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
  return { painted, saves, translates, inks };
}

/** The painted clip is the content clip widened to whole device pixels. */
function expectPixelWidened(painted: LayoutRect, clip: LayoutRect): void {
  expect(painted.xPt).toBeLessThanOrEqual(clip.xPt);
  expect(painted.xPt).toBeGreaterThan(clip.xPt - 1);
  expect(painted.yPt).toBeLessThanOrEqual(clip.yPt);
  expect(painted.yPt).toBeGreaterThan(clip.yPt - 1);
  expect(painted.xPt + painted.widthPt).toBeGreaterThanOrEqual(clip.xPt + clip.widthPt);
  expect(painted.xPt + painted.widthPt).toBeLessThan(clip.xPt + clip.widthPt + 1);
  expect(painted.yPt + painted.heightPt).toBeGreaterThanOrEqual(clip.yPt + clip.heightPt);
  expect(painted.yPt + painted.heightPt).toBeLessThan(clip.yPt + clip.heightPt + 1);
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

    const [picture, legacy] = paintTable(table).painted;
    // The DrawingML picture, raised 5pt in the exact 24pt line, crosses the
    // clip top and is painted under the clip; the VML one is not clipped.
    expect(picture!.bounds.yPt).toBeLessThan(clip.yPt);
    expect(picture!.clips).toHaveLength(1);
    expectPixelWidened(picture!.clips[0]!, clip);
    expect(legacy!.bounds.yPt + legacy!.bounds.heightPt).toBeGreaterThan(clip.yPt + clip.heightPt);
    expect(legacy!.clips).toEqual([]);
  });

  it('keeps a picture that overlaps a neighbouring line inside the cell unclipped by that line', () => {
    const table = layoutTable(exactLine(drawingml(0)) + exactLine(drawingml(0)));
    const [, second] = paintTable(table).painted;
    const cell = table.rows[0]!.cells[0]!;
    // The second picture rises 10.8pt above its line into the first line; only
    // the cell clip applies, and the picture lies within it.
    expect(second!.clips).toHaveLength(1);
    expect(second!.bounds.yPt).toBeGreaterThan(cell.contentClipBounds!.yPt);
  });

  it('keeps the inline paint order when an earlier VML picture escapes the clip', () => {
    // The lowered VML picture of the first line overlaps the DrawingML picture
    // of the second; the second, later in the cell, stays on top.
    const table = layoutTable(exactLine(vml(-15)) + exactLine(drawingml(0)));
    const [first, second] = paintTable(table).painted;
    expect(first!.clips).toEqual([]);
    expect(second!.clips).toHaveLength(1);
    expect(first!.bounds.yPt + first!.bounds.heightPt).toBeGreaterThan(second!.bounds.yPt);
  });

  it('adds no canvas frames to a text-only cell paragraph', () => {
    // One frame per paragraph, holding both the clip and the translation.
    const { saves, translates } = paintTable(layoutTable(exactLine('')));
    const { saves: tableSaves } = paintTable({ ...layoutTable(exactLine('')), rows: [] });
    expect(saves - tableSaves).toBe(1);
    expect(translates).toBe(1);
  });

  it('keeps an empty content clip empty when widening it to device pixels', () => {
    const table = layoutTable(exactLine(drawingml(5)));
    const cell = table.rows[0]!.cells[0]!;
    const empty = { ...cell.contentClipBounds!, widthPt: 0 };
    const [picture] = paintTable({
      ...table,
      rows: [{ ...table.rows[0]!, cells: [{ ...cell, contentClipBounds: empty }] }],
    }).painted;
    expect(picture!.clips).toEqual([empty]);
  });

  it('opens one more frame for a VML picture than for a DrawingML picture', () => {
    // The text before the VML picture is clipped in one frame; the picture
    // itself, and absent shading, borders and text boxes, open none.
    const drawing = paintTable(layoutTable(exactLine(drawingml(-15)))).saves;
    const legacy = paintTable(layoutTable(exactLine(vml(-15)))).saves;
    expect(legacy - drawing).toBe(1);
  });

  it('keeps a bar tab under the clip when the line holds a VML picture', () => {
    // A bar tab at 225pt in a 216pt cell, beside a VML picture that escapes.
    const table = layoutTable('<w:p><w:pPr><w:tabs><w:tab w:val="bar" w:pos="4500"/></w:tabs>'
      + `<w:spacing w:after="0" w:line="480" w:lineRule="exact"/></w:pPr>${vml(-15)}</w:p>`);
    const cell = table.rows[0]!.cells[0]!;
    const right = cell.flowBounds.xPt + cell.flowBounds.widthPt;
    const { painted, inks } = paintTable(table);
    expect(painted[0]!.clips).toEqual([]);
    const bar = inks.filter((ink) => ink.xPt > right + 1);
    expect(bar.length).toBeGreaterThan(0);
    expect(bar.every((ink) => ink.clips > 0)).toBe(true);
  });

  it('marks VML pictures as escaping only in table cells', () => {
    const page = layoutPage(docx('', `${exactLine(vml(-15))}<w:p/>`));
    const paragraph = page.layers.body[0];
    const placement = paragraph?.kind === 'paragraph'
      ? paragraph.lines[0]!.placements.find((item) => item.kind === 'resource') : undefined;
    expect(placement?.kind).toBe('resource');
    expect(placement).not.toHaveProperty('escapesCellClip');
  });

  it('keeps clipped content out of hit testing and highlights', () => {
    // A picture raised past the cell clip and a run indented past the cell's
    // leading edge.
    const layout = layoutDocx(docx(exactLine(drawingml(5))
      + '<w:p><w:pPr><w:ind w:left="-1200"/></w:pPr><w:r><w:t>Hidden</w:t></w:r></w:p>'));
    const table = layout.pages[0]!.layers.body.find((node) => node.kind === 'table') as TableLayout;
    const clip = table.rows[0]!.cells[0]!.contentClipBounds!;
    const picture = (table.rows[0]!.cells[0]!.blocks[0]!.layout as ParagraphLayout).lines[0]!.placements
      .find((item) => item.kind === 'resource');
    if (picture?.kind !== 'resource') throw new Error('expected the picture');
    const registry = createPaintResourceRegistry([{
      kind: 'image', resourceKey: picture.resourceKey, partPath: 'word/media/p.png',
      mimeType: 'image/png', intrinsicSize: { widthPt: 40, heightPt: 30 },
    }]);
    const pictureXPt = clip.xPt + 40;
    expect(hitTestDocxElementContext(layout, 0, { xPt: pictureXPt, yPt: clip.yPt + 1 }, registry)).not.toBeNull();
    expect(hitTestDocxElementContext(layout, 0, { xPt: pictureXPt, yPt: clip.yPt - 1 }, registry)).toBeNull();

    const runs = textRunsForPage(layout, 0, { scale: 1 });
    const hidden = runs.find((run) => run.text.includes('Hidden'));
    // The run lies wholly before the cell: its clips hide it.
    expect(hidden!.x + hidden!.w).toBeLessThanOrEqual(clip.xPt);
    expect(hidden?.clips?.at(-1)).toEqual({
      x: expect.closeTo(clip.xPt, 9), y: expect.closeTo(clip.yPt, 9),
      width: expect.closeTo(clip.widthPt, 9), height: expect.closeTo(clip.heightPt, 9),
    });
    expect(clippedBoxVisible({ x: hidden!.x, y: hidden!.y, width: hidden!.w, height: hidden!.h }, hidden!.clips)).toBe(false);
    const visible = runs.find((run) => run.text.startsWith('Xg'))!;
    expect(clippedBoxVisible({ x: visible.x, y: visible.y, width: visible.w, height: visible.h }, visible.clips)).toBe(true);
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

const turnedBox = (text: string) => '<w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="635000"/><wp:docPr id="9" name="T"/>'
  + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
  + '<wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm rot="2700000"><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm>'
  + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr><wps:txbx><w:txbxContent>'
  + text + '</w:txbxContent></wps:txbx><wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></wps:bodyPr>'
  + '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

describe('clipped text runs in a turned text box', () => {
  it('carry the cell clip unturned for a turned run', () => {
    const layout = layoutDocx(docx(`<w:p>${turnedBox('<w:p><w:r><w:t>Turned</w:t></w:r></w:p>')}</w:p>`));
    const table = layout.pages[0]!.layers.body.find((node) => node.kind === 'table') as TableLayout;
    const clip = table.rows[0]!.cells[0]!.contentClipBounds!;
    const run = textRunsForPage(layout, 0, { scale: 2 }).find((item) => item.text.includes('Turned'));
    expect(run?.transform).toBeDefined();
    // The cell clip is unturned; the text box's own clip follows it, turned.
    expect(run?.clips?.at(-1)?.transform).toBeDefined();
    expect(run?.clips?.find((item) => !item.transform)).toEqual({
      x: expect.closeTo(2 * clip.xPt, 9), y: expect.closeTo(2 * clip.yPt, 9),
      width: expect.closeTo(2 * clip.widthPt, 9), height: expect.closeTo(2 * clip.heightPt, 9),
    });
  });

  it('report turned text in a zero-height cell as hidden', () => {
    // The single row is exact and zero high.
    const layout = layoutDocx(docx(`<w:p>${turnedBox('<w:p><w:r><w:t>Unseen</w:t></w:r></w:p>')}</w:p>`,
      undefined, '<w:trHeight w:val="0" w:hRule="exact"/>', 0, false));
    const run = textRunsForPage(layout, 0, { scale: 1 }).find((item) => item.text.includes('Unseen'));
    const table = layout.pages[0]!.layers.body.find((node) => node.kind === 'table') as TableLayout;
    expect(table.rows[0]!.cells[0]!.contentClipBounds!.heightPt).toBe(0);
    expect(run?.transform).toBeDefined();
    expect(clippedBoxVisible({ x: run!.x, y: run!.y, width: run!.w, height: run!.h, transform: run!.transform }, run!.clips)).toBe(false);
  });

  it('give no comment highlight for turned text wholly outside its text box', () => {
    // A 45-degree text box, not in a table, and a run indented out of it.
    const box = turnedBox('<w:p><w:pPr><w:ind w:left="-1400"/></w:pPr><w:r><w:t>Turned</w:t></w:r></w:p>'
      + '<w:p><w:r><w:t>Inside</w:t></w:r></w:p>');
    const layout = layoutDocx(docx('', `<w:p>${box}</w:p>`));
    const runs = textRunsForPage(layout, 0, { scale: 1 });
    const turned = runs.find((run) => run.text.includes('Turned'))!;
    const inside = runs.find((run) => run.text.includes('Inside'))!;
    expect(turned.clips?.at(-1)?.transform).toBeDefined();
    const threads = (source: typeof turned.source, runIndex: number) => resolveDocxCommentThreads(
      [{ id: 'c', author: 'A', text: 'x' }] as never,
      [{
        commentId: 'c', source: source!, startRunIndex: runIndex, endRunIndex: runIndex + 1,
        reference: { source: source!, runIndex: runIndex + 1, affinity: 'preceding' },
      }] as never,
      runs,
    );
    expect(JSON.stringify(threads(turned.source, turned.sourceRunIndex!))).not.toContain('"width"');
    const shown = JSON.stringify(threads(inside.source, inside.sourceRunIndex!));
    expect(shown).toContain('"width"');
    expect(shown).toContain('"clips"');
  });
});

describe('overlay clip containers', () => {
  interface FakeElement {
    style: Record<string, string> & { cssText: string };
    children: FakeElement[];
    appendChild(child: FakeElement): void;
  }
  const element = (): FakeElement => {
    const style = { cssText: '' } as Record<string, string> & { cssText: string };
    return {
      style: new Proxy(style, {
        set(target, key: string, value: string) {
          if (key === 'cssText') {
            for (const declaration of value.split(';')) {
              const index = declaration.indexOf(':');
              if (index > 0) target[declaration.slice(0, index).trim()] = declaration.slice(index + 1).trim();
            }
          }
          target[key] = value;
          return true;
        },
      }),
      children: [],
      appendChild(child) { this.children.push(child); },
    };
  };

  it('places a turned clipped run\'s find highlights in one container of its clip', () => {
    vi.stubGlobal('document', { createElement: () => element() });
    try {
      const layer = element();
      const clip = { x: 20, y: 10, width: 40, height: 20 };
      const run = {
        text: 'abc', x: 30, y: 20, w: 30, h: 10, fontSize: 10, font: '10px serif',
        transform: 'rotate(90deg)', clips: [clip],
      };
      buildDocxHighlightLayer(layer as unknown as HTMLDivElement, [run, { ...run }],
        [{ slices: [{ runIndex: 0, start: 0, end: 3 }, { runIndex: 1, start: 1, end: 2 }], active: false }],
        100, 100, () => (text: string) => text.length * 10);
      // One container for the consecutive runs' shared clip.
      expect(layer.children).toHaveLength(1);
      const container = layer.children[0]!;
      expect(container.style.overflow).toBe('clip');
      expect([container.style.left, container.style.top, container.style.width, container.style.height])
        .toEqual(['20%', '10%', '40%', '20%']);
      const [whole, partial] = container.children;
      expect([whole!.style.left, whole!.style.top, whole!.style.width]).toEqual(['25%', '50%', '75%']);
      expect(whole!.style.transform).toBe('matrix(0, 1, -1, 0, 0, 0)');
      // A slice 10px into the run advances down the turned run, not across.
      expect([partial!.style.left, partial!.style.top]).toEqual(['25%', '100%']);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('takes the cluster boxes as final, with no tracking added or taken off', () => {
    vi.stubGlobal('document', { createElement: () => element() });
    try {
      const layer = element();
      // Tracking of 2px: inner boxes already carry it, the final one leaves it out.
      const run = { text: 'ABCD', x: 0, y: 0, w: 26, h: 10, fontSize: 10, font: '10px serif', letterSpacingPx: 2,
        clusterBoxes: [
          { start: 0, end: 1, x: 0, width: 7, rtl: false }, { start: 1, end: 2, x: 7, width: 7, rtl: false },
          { start: 2, end: 3, x: 14, width: 7, rtl: false },
          { start: 3, end: 4, x: 21, width: 5, rtl: false, trackingExcluded: true as const },
        ] };
      buildDocxHighlightLayer(layer as unknown as HTMLDivElement, [run],
        [{ slices: [{ runIndex: 0, start: 3, end: 4 }, { runIndex: 0, start: 1, end: 2 }], active: false }],
        100, 100, () => () => 0);
      const [last, inner] = layer.children;
      expect(Number.parseFloat(last!.style.width!)).toBeCloseTo(5, 9);
      expect(Number.parseFloat(inner!.style.width!)).toBeCloseTo(7, 9);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  /** Natural advances of 16pt glyphs: 'W' and 'M' wide, 'i' narrow. */
  const glyphContext = () => ({
    ...measureContext(),
    measureText: (text: string) => ({
      width: [...text].reduce((sum, character) => sum + ({ W: 15, M: 13, i: 4.5 }[character] ?? 8), 0),
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    }),
  }) as unknown as CanvasRenderingContext2D;

  /** Lays out `paragraphs` in the cell (or `body` as the body), finds `slice`
   * of the run whose text is `text`, and returns that run with the one
   * highlight's box in px at scale 1 (the run's clips left out, so the box is
   * in page units). */
  const findIn = (paragraphs: string, text: string, slice: Readonly<{ start: number; end: number }>, body?: string) => {
    const layout = layoutDocx(docx(paragraphs, body), glyphContext());
    const runs = textRunsForPage(layout, 0, { scale: 1 });
    const runIndex = runs.findIndex((run) => run.text === text);
    const run = { ...runs[runIndex]!, clips: undefined };
    const width = 1000;
    const height = 1000;
    const layer = element();
    buildDocxHighlightLayer(layer as unknown as HTMLDivElement, [run],
      [{ slices: [{ runIndex: 0, ...slice }], active: false }], width, height, () => (value: string) => value.length * 5);
    const [box] = layer.children;
    const percent = (value: string, of: number) => (Number.parseFloat(value) / 100) * of;
    return { run, width: percent(box!.style.width!, width), height: percent(box!.style.height!, height) };
  };

  it('highlights a condensed, floored glyph over the box it is painted in', () => {
    vi.stubGlobal('document', { createElement: () => element() });
    try {
      // 'WiM' at 16pt condensed by 6pt: 'i' floors to no advance and paints at
      // its 4.5pt natural advance.
      const { run, width } = findIn('<w:p><w:r><w:rPr><w:spacing w:val="-120"/><w:sz w:val="32"/></w:rPr>'
        + '<w:t>WiM</w:t></w:r></w:p>', 'WiM', { start: 1, end: 2 });
      expect(run.letterSpacingPx).toBeCloseTo(-6, 9);
      // Its advance box has no width; its glyph box is the drawn glyph.
      expect(run.clusterBoxes![1]!.width).toBeCloseTo(0, 9);
      expect(run.clusterBoxes![1]!.glyphWidth).toBeCloseTo(4.5, 9);
      expect(width).toBeCloseTo(4.5, 6);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('highlights an expanded tate-chu-yoko character within its cell', () => {
    vi.stubGlobal('document', { createElement: () => element() });
    try {
      // '29' at 16pt expanded by 3pt in a tbRl section, set upright as one
      // 16pt cell: each character's box is final, so the highlight keeps to it.
      const body = '<w:p><w:r><w:rPr><w:spacing w:val="60"/><w:sz w:val="32"/>'
        + '<w:eastAsianLayout w:id="1" w:vert="1"/></w:rPr><w:t>29</w:t></w:r></w:p>'
        + '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
        + 'w:left="1440" w:header="0" w:footer="0" w:gutter="0"/><w:textDirection w:val="tbRl"/></w:sectPr></w:pPr></w:p>';
      const { run, width, height } = findIn('', '29', { start: 0, end: 1 }, body);
      // The glyph (8pt across) without the 3pt tracking its advance carries.
      const box = run.clusterBoxes![0]!;
      expect(box.height).toBeCloseTo(11, 6);
      expect(box.glyphHeight).toBeCloseTo(8, 6);
      expect(width).toBeCloseTo(box.width, 6);
      expect(height).toBeCloseTo(8, 6);
      expect(Math.max(width, height)).toBeLessThanOrEqual(16 + 1e-6);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('places tate-chu-yoko find highlights side by side across the turned run', () => {
    vi.stubGlobal('document', { createElement: () => element() });
    try {
      const layer = element();
      // A tbRl run turned 90deg at (100, 50): the cell spans 8px down the
      // column; '２' and '９' stand side by side across it.
      const run = {
        text: '２９', x: 100, y: 50, w: 8, h: 10, fontSize: 8, font: '8px serif', transform: 'rotate(90deg)',
        eastAsianVert: true,
        clusterBoxes: [
          { start: 0, end: 1, x: 0, width: 8, rtl: false, y: 8, height: 6 },
          { start: 1, end: 2, x: 0, width: 8, rtl: false, y: 2, height: 6 },
        ],
      };
      buildDocxHighlightLayer(layer as unknown as HTMLDivElement, [run],
        [{ slices: [{ runIndex: 0, start: 0, end: 1 }, { runIndex: 0, start: 1, end: 2 }], active: false }],
        200, 200, () => () => 0);
      const [two, nine] = layer.children;
      // Along the turned run's cross axis (page x): different lefts, one top.
      expect([two!.style.left, two!.style.top]).toEqual(['46%', '25%']);
      expect([nine!.style.left, nine!.style.top]).toEqual(['49%', '25%']);
      expect([two!.style.width, two!.style.height]).toEqual(['4%', '3%']);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('advances a find slice along an unclipped turned run', () => {
    vi.stubGlobal('document', { createElement: () => element() });
    try {
      const layer = element();
      const run = { text: 'abc', x: 30, y: 40, w: 30, h: 10, fontSize: 10, font: '10px serif', transform: 'rotate(90deg)' };
      buildDocxHighlightLayer(layer as unknown as HTMLDivElement, [run],
        [{ slices: [{ runIndex: 0, start: 1, end: 2 }], active: false }], 100, 100, () => (text: string) => text.length * 10);
      const [box] = layer.children;
      expect([box!.style.left, box!.style.top]).toEqual(['30%', '50%']);
    } finally {
      vi.unstubAllGlobals();
    }
  });
});

describe('clippedBoxVisible', () => {
  const box = { x: 10, y: 10, width: 20, height: 10 };

  it('stays linear for nested unions of overlapping rectangles', () => {
    // 30 clips of 4 rectangles each covering the box: a cross product would
    // hold 4^30 pieces.
    const union = { x: 0, y: 0, width: 100, height: 100, rects: [0, 1, 2, 3].map((index) => (
      { x: index, y: index, width: 90, height: 90 })) };
    const clips = Array.from({ length: 30 }, () => union);
    expect(clippedBoxVisible(box, clips)).toBe(true);
    expect(clippedBoxBounds(box, clips)).toEqual({ x: 10, y: 10, width: 20, height: 10 });
  });

  it('ignores rectangles without area and hides a box under an empty union', () => {
    const union = { x: 0, y: 0, width: 100, height: 100, rects: [
      { x: 0, y: 0, width: 0, height: 0 }, { x: 70, y: 60, width: 30, height: 40 },
    ] };
    expect(clippedBoxVisible(box, [union])).toBe(false);
    expect(clippedBoxVisible(box, [{ ...union, rects: [] }])).toBe(false);
    expect(clippedBoxVisible({ ...box, x: 75, y: 65 }, [union])).toBe(true);
  });
});

describe('a VML text box union with a degenerate rectangle', () => {
  it('gives no comment highlight to text only the degenerate rectangle covers', () => {
    const box = '<w:p><w:r><w:pict><v:group style="width:120pt;height:80pt" coordsize="120,80">'
      + '<v:shapetype id="box" coordsize="1000,1000"/><v:shape type="#box" style="width:100;height:60" filled="f" stroked="f">'
      + '<v:path textboxrect="0,0,0,0;700,600,1000,1000"/><v:textbox inset="0,0,0,0"><w:txbxContent>'
      + '<w:p><w:r><w:t>Hid</w:t></w:r></w:p>'
      + '</w:txbxContent></v:textbox></v:shape></v:group></w:pict></w:r></w:p>';
    const layout = layoutDocx(docx('', box));
    const runs = textRunsForPage(layout, 0, { scale: 1 });
    const hid = runs.find((run) => run.text === 'Hid')!;
    expect(hid.clips?.at(-1)?.rects).toHaveLength(1);
    const threads = resolveDocxCommentThreads([{ id: 'c', author: 'A', text: 'x' }] as never, [{
      commentId: 'c', source: hid.source!, startRunIndex: hid.sourceRunIndex!, endRunIndex: hid.sourceRunIndex! + 1,
      reference: { source: hid.source!, runIndex: hid.sourceRunIndex! + 1, affinity: 'preceding' },
    }] as never, runs);
    expect(JSON.stringify(threads)).not.toContain('"width"');
  });
});

describe('clip unions at scale', () => {
  const box = { x: 40, y: 10, width: 10, height: 10 };

  it('handles a union of 35,000 rectangles without spreading them onto the stack', () => {
    const rects = Array.from({ length: 35_000 }, (_, index) => ({ x: (index % 100) * 2, y: 0, width: 1, height: 100 }));
    expect(() => clippedBoxVisible(box, [{ x: 0, y: 0, width: 200, height: 100, rects }])).not.toThrow();
    expect(clippedBoxVisible(box, [{ x: 0, y: 0, width: 200, height: 100, rects }])).toBe(true);
  });

  it('keeps a box in the gap of a union hidden however often its rectangles repeat', () => {
    // textboxrect="0,0,200,1000;800,0,1000,1000" on a 100pt box: a gap at 20-80.
    const left = { x: 0, y: 0, width: 20, height: 100 };
    const right = { x: 80, y: 0, width: 20, height: 100 };
    const clip = (rects: Array<typeof left>) => [{ x: 0, y: 0, width: 100, height: 100, rects }];
    expect(clippedBoxVisible(box, clip([left, right]))).toBe(false);
    expect(clippedBoxVisible(box, clip([...Array.from({ length: 64 }, () => left), right]))).toBe(false);
    // Rectangles nested in another count as that one.
    expect(clippedBoxVisible(box, clip([left, { x: 2, y: 2, width: 10, height: 10 }, right]))).toBe(false);
  });
});

describe('rotated exact-height cells', () => {
  it('clip to the cell edges, not across the page', () => {
    const lines = Array.from({ length: 14 }, (_, index) => `<w:p><w:r><w:t>L${index}</w:t></w:r></w:p>`).join('');
    const table = '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/><w:tblInd w:w="1000" w:type="dxa"/>'
      + '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'
      + '<w:tr><w:trPr><w:trHeight w:val="600" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>'
      + `<w:textDirection w:val="tbRl"/></w:tcPr>${lines}</w:tc></w:tr></w:tbl><w:p/>`;
    const layout = layoutDocx(docx('', table));
    const laid = layout.pages[0]!.layers.body.find((node) => node.kind === 'table') as TableLayout;
    const cell = laid.rows[0]!.cells[0]!;
    expect(cell.verticalText).toBeDefined();
    expect(cell.clipBounds!.xPt).toBeCloseTo(cell.flowBounds.xPt, 9);
    expect(cell.clipBounds!.widthPt).toBeCloseTo(cell.flowBounds.widthPt, 9);
    // Text runs in it carry the same box as their clip.
    const run = textRunsForPage(layout, 0, { scale: 1 }).find((item) => item.text === 'L0');
    expect(run?.clips?.at(-1)).toMatchObject({
      x: expect.closeTo(cell.flowBounds.xPt, 9), width: expect.closeTo(cell.flowBounds.widthPt, 9),
    });
  });
});

describe('huge VML text rectangle lists', () => {
  it('walk the model for vertical capability without spreading arrays onto the stack', () => {
    const model = { body: [{ type: 'paragraph', runs: [{ type: 'shape',
      vmlTextRects: Array.from({ length: 150_000 }, () => [0, 0, 1, 1]) }] }] };
    expect(() => documentRequiresDomVerticalGlyphLayout(model as never)).not.toThrow();
  });
});

describe('a DrawingML text box without an auto-fit element', () => {
  it('is a fixed box (§21.1.2.1.2 implies noAutofit) and clips its text', () => {
    const box = (autofit: string) => '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="635000"/><wp:docPr id="9" name="T"/>'
      + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
      + '<wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm>'
      + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr><wps:txbx><w:txbxContent>'
      + '<w:p><w:pPr><w:ind w:left="-1400"/></w:pPr><w:r><w:t>Out</w:t></w:r></w:p>'
      + `</w:txbxContent></wps:txbx><wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0">${autofit}</wps:bodyPr>`
      + '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const clipsOf = (autofit: string) => textRunsForPage(layoutDocx(docx('', box(autofit))), 0, { scale: 1 })
      .find((run) => run.text === 'Out')?.clips;
    const explicit = clipsOf('<a:noAutofit/>');
    expect(explicit).toHaveLength(1);
    expect(clipsOf('')).toEqual(explicit);
  });
});

describe('hit testing a zero-width cell', () => {
  it('finds no picture on the edge of a clip without area', () => {
    const table = '<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
      + '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
      + '<w:tblGrid><w:gridCol w:w="0"/><w:gridCol w:w="4320"/></w:tblGrid><w:tr>'
      + `<w:tc><w:tcPr><w:tcW w:w="0" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:ind w:left="-400"/></w:pPr>${drawingml(0)}</w:p></w:tc>`
      + '<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:p/>';
    const layout = layoutDocx(docx('', table));
    const laid = layout.pages[0]!.layers.body.find((node) => node.kind === 'table') as TableLayout;
    const cell = laid.rows[0]!.cells[0]!;
    const clip = cell.contentClipBounds!;
    expect(clip.widthPt).toBe(0);
    const paragraph = cell.blocks[0]!.layout as ParagraphLayout;
    const picture = paragraph.lines[0]!.placements.find((item) => item.kind === 'resource');
    if (picture?.kind !== 'resource') throw new Error('expected the picture');
    const registry = createPaintResourceRegistry([{
      kind: 'image', resourceKey: picture.resourceKey, partPath: 'word/media/p.png',
      mimeType: 'image/png', intrinsicSize: { widthPt: 40, heightPt: 30 },
    }]);
    const middle = clip.yPt + clip.heightPt / 2;
    expect(hitTestDocxElementContext(layout, 0, { xPt: clip.xPt, yPt: middle }, registry)).toBeNull();
  });
});

describe('tate-chu-yoko cluster boxes', () => {
  it('stand the characters upright side by side across the run, each over the cell', () => {
    // '２９' (two 6pt natural advances) in an 8pt cell at x=78, baseline 80:
    // the row is centred on the baseline across the run, first character on
    // the run's far side, as Canvas paints it turned back upright.
    const placement = {
      text: '２９', range: { start: 1, end: 3 }, direction: 'ltr' as const, tateChuYoko: true,
      origin: { xPt: 78, yPt: 80 }, bounds: { xPt: 78, yPt: 72, widthPt: 8, heightPt: 10 },
      clusters: [
        { range: { start: 1, end: 2 }, offset: { xPt: 0, yPt: 0 }, advancePt: 6 },
        { range: { start: 2, end: 3 }, offset: { xPt: 6, yPt: 0 }, advancePt: 6 },
      ],
      paintOps: [{ range: { start: 1, end: 3 }, offset: { xPt: 4, yPt: 0 } }],
    };
    const boxes = clusterBoxes(placement as never, 0, 2, 1, 3, 78);
    expect(boxes.map((box) => [box.xPt, box.widthPt, box.crossPt, box.crossWidthPt])).toEqual([[0, 8, 8, 6], [0, 8, 2, 6]]);
  });

  it('give a floored, condensed glyph the glyph box it is drawn in', () => {
    // 'abc' condensed so that 'b' floors to no advance; its own operation
    // draws it at its 6pt natural advance.
    const placement = {
      text: 'abc', range: { start: 0, end: 3 }, direction: 'ltr' as const,
      origin: { xPt: 0, yPt: 10 }, bounds: { xPt: 0, yPt: 0, widthPt: 8, heightPt: 12 },
      clusters: [
        { range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, advancePt: 4 },
        { range: { start: 1, end: 2 }, offset: { xPt: 4, yPt: 0 }, advancePt: 0 },
        { range: { start: 2, end: 3 }, offset: { xPt: 4, yPt: 0 }, advancePt: 4 },
      ],
      paintOps: [
        { range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 } },
        { range: { start: 1, end: 2 }, offset: { xPt: 4, yPt: 0 }, naturalAdvancePt: 6 },
        { range: { start: 2, end: 3 }, offset: { xPt: 4, yPt: 0 } },
      ],
    };
    const boxes = clusterBoxes(placement as never, 0, 3, 0, 3, 0);
    // Advance boxes from pen to pen; b's glyph box is the drawn glyph.
    expect(boxes.map((box) => [box.xPt, box.widthPt])).toEqual([[0, 4], [4, 0], [4, 4]]);
    expect(boxes.map((box) => [box.glyphXPt, box.glyphWidthPt])).toEqual([[0, 4], [4, 6], [4, 4]]);
  });

  it('end a compressed right-to-left fitted run at its box\'s right edge', () => {
    // אבגד fitted into 12pt with 6pt advances: pitch -4, clusters of 2pt.
    const placement = {
      text: 'אבגד', range: { start: 0, end: 4 }, direction: 'rtl' as const,
      origin: { xPt: 168, yPt: 10 }, bounds: { xPt: 168, yPt: 0, widthPt: 12, heightPt: 12 },
      clusters: [0, 1, 2, 3].map((index) => ({ range: { start: index, end: index + 1 }, offset: { xPt: 2 * index, yPt: 0 }, advancePt: 2 })),
    };
    const boxes = clusterBoxes(placement as never, 0, 4, 0, 4, 168);
    // א (logical first) is painted rightmost, ending at the box edge.
    expect(boxes[0]!.xPt + boxes[0]!.widthPt).toBeCloseTo(12, 9);
    expect(Math.min(...boxes.map((box) => box.xPt))).toBeCloseTo(0, 9);
  });

  it('extend the final cluster of a compressed fitted run to the run\'s end', () => {
    // fitText 12pt over four 6pt characters: pitch -4, so each cluster
    // advances 2pt while the last glyph paints to the box edge.
    const placement = {
      text: 'abcd', range: { start: 0, end: 4 }, direction: 'ltr' as const,
      origin: { xPt: 0, yPt: 10 }, bounds: { xPt: 0, yPt: 0, widthPt: 12, heightPt: 12 },
      clusters: [0, 1, 2, 3].map((index) => ({ range: { start: index, end: index + 1 }, offset: { xPt: 2 * index, yPt: 0 }, advancePt: 2 })),
    };
    const boxes = clusterBoxes(placement as never, 0, 4, 0, 4, 0);
    expect(boxes.at(-1)).toMatchObject({ xPt: 6, widthPt: 6, trackingExcluded: true });
  });
});
