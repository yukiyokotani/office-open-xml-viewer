/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import type { DocxDocumentModel } from './types.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { storeZip } from './conformance/generate.js';
import { paintTextBoxLayout } from './paint/canvas-text.js';
import { elementGeometryForPage } from './layout/text-index.js';

// word-story-anchor-inline: a wp:anchor inside a text box, footnote or
// endnote story lays out like a wp:inline of the same extent, whatever its
// wrap, layer or relative frame (the evidence is in the rule). Synthetic
// reductions of its controls: a DrawingML or VML text box, a footnote or an endnote
// holding M01, a 40x20pt picture anchored 300pt from the column followed by
// Ab, and MEND.

const EMU = 12700;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const V = 'urn:schemas-microsoft-com:vml';
const PNG = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0xf0,
  0x1f, 0x00, 0x05, 0x00, 0x01, 0xff, 0x89, 0x99, 0x3d, 0x1d, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45,
  0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

type Story = 'body' | 'drawingml-textbox' | 'vml-textbox' | 'footnote' | 'endnote';
type Placement = 'inline' | 'wrapNone' | 'wrapSquare' | 'wrapTopAndBottom' | 'behindDoc' | 'page'
  | 'vml-text-square' | 'vml-text-square-mt50' | 'vml-margin' | 'vml-text-none' | 'vml-aligned';

/** Story variants: the content in a table of the story, or a rotated box. */
let variant: Readonly<{
  table?: boolean; rotationDeg?: number;
  /** Size of the VML rect's run (half points). */
  shapeRunHalfPoints?: number;
  /** A default paragraph style other than Normal, at this size (half points). */
  defaultStyleHalfPoints?: number;
  /** Runs replacing the story's object paragraph content. */
  storyRuns?: string;
  /** The whole body, replacing the story-specific one. */
  body?: string;
}> = {};

const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);
const text = (value: string) => `<w:p><w:r><w:t>${value}</w:t></w:r></w:p>`;

const W10 = 'urn:schemas-microsoft-com:office:word';

function vmlRect(placement: Placement): string {
  const frames = placement === 'vml-margin' ? 'margin' : 'text';
  const top = placement === 'vml-text-square-mt50' ? 50 : 0;
  const wrap = placement === 'vml-text-none' ? 'none' : 'square';
  const aligned = placement === 'vml-aligned' ? 'mso-position-horizontal:center;mso-position-vertical:bottom;' : '';
  const runProperties = variant.shapeRunHalfPoints ? `<w:rPr><w:sz w:val="${variant.shapeRunHalfPoints}"/></w:rPr>` : '';
  return `<w:r>${runProperties}<w:pict><v:rect style="position:absolute;margin-left:300pt;margin-top:${top}pt;width:40pt;height:20pt;${aligned}`
    + `mso-position-horizontal-relative:${frames};mso-position-vertical-relative:${frames === 'text' ? 'text' : 'margin'}" `
    + `fillcolor="black" stroked="f"><w10:wrap type="${wrap}"/></v:rect></w:pict></w:r>`;
}

function picture(placement: Placement): string {
  if (placement.startsWith('vml-')) return vmlRect(placement);
  const extent = `<wp:extent cx="${40 * EMU}" cy="${20 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`;
  const graphic = `<a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="p.png"/>`
    + '<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch>'
    + `</pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${40 * EMU}" cy="${20 * EMU}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>';
  if (placement === 'inline') {
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">${extent}`
      + `<wp:docPr id="2" name="d"/>${graphic}</wp:inline></w:drawing></w:r>`;
  }
  const page = placement === 'page';
  const wrap = placement === 'wrapNone' || placement === 'behindDoc' ? '<wp:wrapNone/>'
    : placement === 'wrapTopAndBottom' ? '<wp:wrapTopAndBottom/>' : '<wp:wrapSquare wrapText="bothSides"/>';
  return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" `
    + `behindDoc="${placement === 'behindDoc' ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">`
    + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${page ? 'page' : 'column'}">`
    + `<wp:posOffset>${(page ? 420 : 300) * EMU}</wp:posOffset></wp:positionH>`
    + `<wp:positionV relativeFrom="${page ? 'page' : 'paragraph'}"><wp:posOffset>${(page ? 640 : 0) * EMU}</wp:posOffset></wp:positionV>`
    + `${extent}${wrap}<wp:docPr id="2" name="d"/>${graphic}</wp:anchor></w:drawing></w:r>`;
}

function docx(story: Story, placement: Placement): Uint8Array {
  const paragraphs = text('M01') + `<w:p>${variant.storyRuns ?? picture(placement)}<w:r><w:t>Ab</w:t></w:r></w:p>` + text('MEND');
  const content = variant.table
    ? '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
      + `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${paragraphs}</w:tc></w:tr></w:tbl>${text('END')}`
    : paragraphs;
  let body = text('BODY');
  const notes = story === 'footnote' || story === 'endnote';
  if (notes) {
    body = `<w:p><w:r><w:t>BODY</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${story}Reference w:id="1"/></w:r></w:p>`;
  } else if (story === 'body') {
    body += content;
  } else if (story === 'drawingml-textbox') {
    body += '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
      + 'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
      + '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>'
      + `<wp:positionV relativeFrom="margin"><wp:posOffset>${30 * EMU}</wp:posOffset></wp:positionV>`
      + `<wp:extent cx="${468 * EMU}" cy="${200 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>`
      + `<wp:docPr id="1" name="tb"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/>`
      + `<wps:spPr><a:xfrm${variant.rotationDeg ? ` rot="${variant.rotationDeg * 60000}"` : ''}><a:off x="0" y="0"/><a:ext cx="${468 * EMU}" cy="${200 * EMU}"/></a:xfrm>`
      + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
      + `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>`
      + '<wps:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></wps:bodyPr>'
      + '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>';
  } else {
    body += '<w:p><w:r><w:pict><v:shape id="tb" style="position:absolute;margin-left:0;margin-top:30pt;'
      + `width:468pt;height:200pt;${variant.rotationDeg ? `rotation:${variant.rotationDeg};` : ''}mso-position-horizontal-relative:margin;mso-position-vertical-relative:margin" `
      + `filled="f" stroked="f"><v:textbox inset="0,0,0,0"><w:txbxContent>${content}</w:txbxContent></v:textbox>`
      + '</v:shape></w:pict></w:r></w:p>';
  }
  if (variant.body !== undefined) body = variant.body;
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" `
    + `xmlns:pic="${PIC}" xmlns:wps="${WPS}" xmlns:v="${V}" xmlns:w10="${W10}"><w:body>${body}`
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" '
    + 'w:left="1440" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr></w:body></w:document>';
  const styles = `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>`
    + '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>'
    + (variant.defaultStyleHalfPoints
      ? '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        + '<w:style w:type="paragraph" w:default="1" w:styleId="Big"><w:name w:val="Big"/>'
        + `<w:rPr><w:sz w:val="${variant.defaultStyleHalfPoints}"/></w:rPr></w:style>`
      : '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>')
    + '</w:styles>';
  const settings = `<w:settings xmlns:w="${W}"><w:compat><w:compatSetting w:name="compatibilityMode" `
    + 'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const namespaces = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}" `
    + `xmlns:v="${V}" xmlns:w10="${W10}"`;
  const noteParts: Array<[string, Uint8Array]> = notes ? [
    [`word/${story}s.xml`, xml(`<w:${story}s ${namespaces}>`
      + `<w:${story} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${story}>`
      + `<w:${story} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${story}>`
      + `<w:${story} w:id="1">${content}</w:${story}></w:${story}s>`)],
    [`word/_rels/${story}s.xml.rels`, xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rIdImg" Type="${rel}/image" Target="media/p.png"/></Relationships>`)],
  ] : [];
  return storeZip(new Map([
    ['[Content_Types].xml', xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      + (notes ? `<Override PartName="/word/${story}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${story}s+xml"/>` : '')
      + '</Types>')],
    ['_rels/.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rIdStyles" Type="${rel}/styles" Target="styles.xml"/>`
      + `<Relationship Id="rIdSettings" Type="${rel}/settings" Target="settings.xml"/>`
      + `<Relationship Id="rIdImg" Type="${rel}/image" Target="media/p.png"/>`
      + (notes ? `<Relationship Id="rIdNotes" Type="${rel}/${story}s" Target="${story}s.xml"/>` : '')
      + '</Relationships>')],
    ['word/document.xml', xml(document)],
    ['word/styles.xml', xml(styles)],
    ['word/settings.xml', xml(settings)],
    ['word/media/p.png', PNG],
    ...noteParts,
  ]));
}

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'normal',
    measureText: (value: string) => {
      const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...value].length * px * 0.5,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

function layOut(story: Story, placement: Placement) {
  return layOutDocument(story, placement).pages[0]!;
}

function layOutDocument(story: Story, placement: Placement) {
  const archive = new DocxArchive(docx(story, placement));
  let model: DocxDocumentModel;
  try {
    model = normalizeInternalDocumentModel(
      JSON.parse(new TextDecoder().decode(archive.parse())),
    ).document;
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

describe('word-story-anchor-inline', () => {
  for (const story of ['drawingml-textbox', 'vml-textbox', 'footnote', 'endnote'] as const) {
    for (const placement of ['wrapNone', 'wrapSquare', 'wrapTopAndBottom', 'behindDoc', 'page'] as const) {
      it(`${story}: a ${placement} anchor lays out like the inline picture`, () => {
        expect(JSON.stringify(layOut(story, placement).layers))
          .toBe(JSON.stringify(layOut(story, 'inline').layers));
      });
    }
  }

  it('lays the note content out (the comparison is not vacuous)', () => {
    for (const story of ['footnote', 'endnote'] as const) {
      const layers = JSON.stringify(layOut(story, 'inline').layers);
      expect(layers).toContain('MEND');
      expect(layers).toContain('p.png');
    }
  });

  it('keeps body anchors floating', () => {
    expect(JSON.stringify(layOut('body', 'wrapSquare').layers))
      .not.toBe(JSON.stringify(layOut('body', 'inline').layers));
  });

});

/** Page boxes of the anchored drawings laid out inside the page's stories. */
function storyDrawingBoxes(story: Story, placement: Placement): Array<{ xPt: number; yPt: number }> {
  const boxes: Array<{ xPt: number; yPt: number }> = [];
  const visit = (node: unknown, inStory: boolean): void => {
    if (!node || typeof node !== 'object') return;
    const record = node as Record<string, unknown>;
    if (inStory && record.kind === 'drawing' && record.anchorLayer && record.flowBounds) {
      const bounds = record.flowBounds as { xPt: number; yPt: number; widthPt: number };
      if (bounds.widthPt === 40) boxes.push({ xPt: bounds.xPt, yPt: bounds.yPt });
    }
    for (const [key, value] of Object.entries(record)) {
      if (key === 'source' || key === 'paintOrder') continue;
      visit(value, inStory || key === 'story');
    }
  };
  visit(layOut(story, placement).layers, false);
  // A drawing is retained both on its paragraph and in its line placements.
  return [...new Map(boxes.map((box) => [`${box.xPt},${box.yPt}`, box])).values()];
}

/** Page boxes of the anchored drawings of the body's own paragraphs. */
function bodyHostDrawingBoxes(story: Story, placement: Placement): Array<{ xPt: number; yPt: number }> {
  const boxes = layOut(story, placement).layers.body.flatMap((node) => node.kind === 'paragraph'
    ? node.drawings.filter((drawing) => drawing.anchorLayer && drawing.flowBounds.widthPt === 40)
      .map((drawing) => ({ xPt: drawing.flowBounds.xPt, yPt: drawing.flowBounds.yPt }))
    : []);
  return [...new Map(boxes.map((box) => [`${box.xPt},${box.yPt}`, box])).values()];
}

describe('word-story-vml-page-position', () => {
  it.each(['drawingml-textbox', 'vml-textbox'] as const)('places a %s VML shape from the page origin', (story) => {
    // Text/text square: x = margin-left, y held at the 36pt top margin.
    expect(storyDrawingBoxes(story, 'vml-text-square')).toEqual([{ xPt: 300, yPt: 36 }]);
    expect(storyDrawingBoxes(story, 'vml-text-square-mt50')).toEqual([{ xPt: 300, yPt: 50 }]);
    // Margin frames and wrapNone: the offsets alone, no top-margin floor.
    expect(storyDrawingBoxes(story, 'vml-margin')).toEqual([{ xPt: 300, yPt: 0 }]);
    expect(storyDrawingBoxes(story, 'vml-text-none')).toEqual([{ xPt: 300, yPt: 0 }]);
  });

  it.each(['footnote', 'endnote'] as const)('places a %s VML shape at its page offsets', (story) => {
    // word-note-vml-host-paragraph: the shape leaves its note for an empty
    // Normal paragraph appended to the body, which hosts it at its page offsets.
    expect(storyDrawingBoxes(story, 'vml-text-square')).toEqual([]);
    expect(bodyHostDrawingBoxes(story, 'vml-text-square')).toEqual([{ xPt: 300, yPt: 0 }]);
    expect(bodyHostDrawingBoxes(story, 'vml-text-square-mt50')).toEqual([{ xPt: 300, yPt: 50 }]);
  });

  it.each(['footnote', 'endnote'] as const)('appends one Normal host paragraph for the %s shapes', (story) => {
    const page = layOut(story, 'vml-text-square');
    const body = page.layers.body.filter((node) => node.kind === 'paragraph');
    // BODY, then the appended host paragraph; the note keeps its text.
    expect(body).toHaveLength(2);
    const host = body[1]!;
    expect(host.kind === 'paragraph' && host.drawings.some((drawing) => drawing.anchorLayer)).toBe(true);
    expect(JSON.stringify(page.layers.notes)).toContain('MEND');
    // An inline picture stays in its note: no host paragraph.
    expect(layOut(story, 'inline').layers.body.filter((node) => node.kind === 'paragraph')).toHaveLength(1);
  });
});

/** Device boxes of the filled paths a text box paints (identity base frame). */
function paintedFills(
  story: Story,
  placement: Placement,
  /** Paint the drawings as another producer's page-owned drawings. */
  withoutStoryFlag = false,
): Array<{ xPt: number; yPt: number; clipped: boolean }> {
  type M = { a: number; b: number; c: number; d: number; e: number; f: number };
  let m: M = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  let clipped = false;
  const stack: Array<[M, boolean]> = [];
  let path: Array<[number, number]> = [];
  const fills: Array<{ xPt: number; yPt: number; clipped: boolean }> = [];
  const point = (x: number, y: number): [number, number] => [m.a * x + m.c * y + m.e, m.b * x + m.d * y + m.f];
  const multiply = (n: M) => {
    m = { a: m.a * n.a + m.c * n.b, b: m.b * n.a + m.d * n.b, c: m.a * n.c + m.c * n.d,
      d: m.b * n.c + m.d * n.d, e: m.a * n.e + m.c * n.f + m.e, f: m.b * n.e + m.d * n.f + m.f };
  };
  const methods: Record<string, (...args: number[]) => unknown> = {
    save: () => { stack.push([{ ...m }, clipped]); },
    restore: () => { [m, clipped] = stack.pop() ?? [m, clipped]; },
    clip: () => { clipped = true; },
    translate: (x, y) => multiply({ a: 1, b: 0, c: 0, d: 1, e: x!, f: y! }),
    scale: (x, y) => multiply({ a: x!, b: 0, c: 0, d: y!, e: 0, f: 0 }),
    rotate: (t) => multiply({ a: Math.cos(t!), b: Math.sin(t!), c: -Math.sin(t!), d: Math.cos(t!), e: 0, f: 0 }),
    transform: (a, b, c, d, e, f) => multiply({ a: a!, b: b!, c: c!, d: d!, e: e!, f: f! }),
    setTransform: (a, b, c, d, e, f) => { m = { a: a ?? 1, b: b ?? 0, c: c ?? 0, d: d ?? 1, e: e ?? 0, f: f ?? 0 }; },
    beginPath: () => { path = []; },
    moveTo: (x, y) => { path.push(point(x!, y!)); },
    lineTo: (x, y) => { path.push(point(x!, y!)); },
    rect: (x, y, w, h) => { path.push(point(x!, y!), point(x! + w!, y!), point(x! + w!, y! + h!), point(x!, y! + h!)); },
    fill: () => {
      if (path.length === 0) return;
      const xs = path.map(([x]) => x);
      const ys = path.map(([, y]) => y);
      // The rect's first edge is 40pt long in any rotation.
      const width = path.length >= 2 ? Math.hypot(path[1]![0] - path[0]![0], path[1]![1] - path[0]![1]) : 0;
      if (Math.abs(width - 40) < 0.5) {
        fills.push({ xPt: Math.round(Math.min(...xs) * 100) / 100, yPt: Math.round(Math.min(...ys) * 100) / 100, clipped });
      }
    },
    fillRect: (x, y, w, h) => {
      if (Math.abs(w! - 40) < 0.5) {
        const [px, py] = point(x!, y!);
        fills.push({ xPt: Math.round(px * 100) / 100, yPt: Math.round(py * 100) / 100, clipped });
      }
    },
    measureText: () => ({ width: 0, actualBoundingBoxAscent: 0, actualBoundingBoxDescent: 0,
      fontBoundingBoxAscent: 0, fontBoundingBoxDescent: 0 }),
    getTransform: () => ({ ...m }),
  };
  const ctx = new Proxy({} as Record<string, unknown>, {
    get: (target, key: string) => methods[key] ?? target[key] ?? (() => undefined),
    set: (target, key: string, value) => { target[key] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const page = layOut(story, placement);
  for (const node of page.layers.body) {
    if (node.kind !== 'paragraph') continue;
    for (const textBox of node.textBoxes) {
      const painted = withoutStoryFlag
        ? JSON.parse(JSON.stringify(textBox, (key, value) => key === 'storyPagePosition' ? undefined : value)) as typeof textBox
        : textBox;
      paintTextBoxLayout(painted, { ctx, scale: 1, dpr: 1, resources: { paint: () => {} } });
    }
  }
  return fills;
}

describe('word-story-vml-page-position in tables, rotated boxes and aligned shapes', () => {
  const expected = [{ xPt: 300, yPt: 36, clipped: false }];
  it.each([
    ['a table of the story', { table: true }],
    ['a box rotated by 30 degrees', { rotationDeg: 30 }],
    ['a box rotated by 90 degrees', { rotationDeg: 90 }],
  ] as const)('places and paints the shape at its page offsets in %s', (_name, storyVariant) => {
    variant = storyVariant;
    try {
      for (const story of ['drawingml-textbox', 'vml-textbox'] as const) {
        expect(storyDrawingBoxes(story, 'vml-text-square')).toEqual([{ xPt: 300, yPt: 36 }]);
        expect(paintedFills(story, 'vml-text-square')).toEqual(expected);
      }
    } finally {
      variant = {};
    }
  });

  it.each(['drawingml-textbox', 'vml-textbox'] as const)('aligns a %s shape in its authored frames', (story) => {
    // Centred in the column; a bottom alignment against the text frame takes
    // the page top (WORD_STORY_VML_PAGE_POSITION).
    expect(storyDrawingBoxes(story, 'vml-aligned')).toEqual([{ xPt: 286, yPt: 0 }]);
    expect(paintedFills(story, 'vml-aligned')).toEqual([{ xPt: 286, yPt: 0, clipped: false }]);
  });

  it('paints a direct story shape once, outside the text clip', () => {
    expect(paintedFills('drawingml-textbox', 'vml-text-square')).toEqual(expected);
  });
});

/** A package with an optional footnote and endnote, each holding one VML
 * rect at `margin-left`, and body references to the listed kinds. The two
 * note parts are padded so both rects start at the same part offset. */
function twoNotePackage(notes: Readonly<{ footnote?: number; endnote?: number }>, references: readonly ('footnote' | 'endnote')[]) {
  const rect = (left: number) => `<w:p><w:r><w:pict><v:rect style="position:absolute;margin-left:${left}pt;margin-top:0;`
    + 'width:40pt;height:20pt;mso-position-horizontal-relative:text;mso-position-vertical-relative:text" '
    + `fillcolor="black" stroked="f"><w10:wrap type="none"/></v:rect></w:pict></w:r></w:p>`;
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const namespaces = `xmlns:w="${W}" xmlns:r="${R}" xmlns:v="${V}" xmlns:w10="${W10}"`;
  const part = (kind: 'footnote' | 'endnote', pad: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<w:${kind}s ${namespaces}${pad}><w:${kind} w:id="1">${rect(notes[kind]!)}</w:${kind}></w:${kind}s>`;
  const kinds = (['footnote', 'endnote'] as const).filter((kind) => notes[kind] !== undefined);
  // "footnote" is one character longer than "endnote", twice before the rect.
  const pads = { footnote: '', endnote: '  ' };
  const body = `<w:p><w:r><w:t>BODY</w:t></w:r>${references.map((kind) => `<w:r><w:${kind}Reference w:id="1"/></w:r>`).join('')}</w:p>`;
  const document = `<w:document ${namespaces}><w:body>${body}`
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" '
    + 'w:left="1440" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr></w:body></w:document>';
  const parts = new Map<string, Uint8Array>([
    ['[Content_Types].xml', xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + kinds.map((kind) => `<Override PartName="/word/${kind}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/>`).join('')
      + '</Types>')],
    ['_rels/.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + kinds.map((kind) => `<Relationship Id="rId${kind}" Type="${rel}/${kind}s" Target="${kind}s.xml"/>`).join('')
      + '</Relationships>')],
    ['word/document.xml', xml(document)],
    ...kinds.map((kind) => [`word/${kind}s.xml`, encoder.encode(part(kind, pads[kind]))] as [string, Uint8Array]),
  ]);
  const offsets = kinds.map((kind) => part(kind, pads[kind]).indexOf('<v:rect'));
  const archive = new DocxArchive(storeZip(parts));
  let model: DocxDocumentModel;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), { currentDateMs: 0 });
  const hosted = layout.pages.flatMap((page) => page.layers.body).flatMap((node) => node.kind === 'paragraph'
    ? node.drawings.filter((drawing) => drawing.anchorLayer && drawing.flowBounds.widthPt === 40)
      .map((drawing) => drawing.flowBounds.xPt)
    : []);
  return { offsets, hosted: [...new Set(hosted)].sort((a, b) => a - b), pages: layout.pages.length,
    bodyParagraphs: layout.pages.flatMap((page) => page.layers.body).filter((node) => node.kind === 'paragraph').length };
}

describe('word-note-vml-host-paragraph ownership', () => {
  it('keeps a footnote shape and an endnote shape at the same part offset apart', () => {
    const result = twoNotePackage({ footnote: 300, endnote: 400 }, ['footnote', 'endnote']);
    expect(result.offsets[0]).toBe(result.offsets[1]);
    expect(result.hosted).toEqual([300, 400]);
  });

  it('hosts no shape of an unreferenced note', () => {
    const result = twoNotePackage({ footnote: 300, endnote: 400 }, ['footnote']);
    expect(result.hosted).toEqual([300]);
    const none = twoNotePackage({ endnote: 400 }, []);
    expect(none).toMatchObject({ hosted: [], pages: 1, bodyParagraphs: 1 });
  });
});

describe('word-note-vml-host-paragraph formatting and story shape indexing', () => {
  const hostLine = () => {
    const body = layOut('endnote', 'vml-text-square').layers.body.filter((node) => node.kind === 'paragraph');
    const host = body.at(-1)!;
    return host.kind === 'paragraph' ? host.flowBounds.heightPt : NaN;
  };

  it('formats the host paragraph as an empty Normal paragraph, not with the note run', () => {
    const plain = hostLine();
    try {
      variant = { shapeRunHalfPoints: 160 };
      expect(hostLine()).toBe(plain);
      variant = { defaultStyleHalfPoints: 160 };
      expect(hostLine()).toBe(plain);
    } finally {
      variant = {};
    }
  });

  it.each([0, 30])('indexes a story shape where it paints (text box rotated by %d degrees)', (rotationDeg) => {
    variant = { rotationDeg };
    try {
      for (const story of ['drawingml-textbox', 'vml-textbox'] as const) {
        const entries = elementGeometryForPage(layOutDocument(story, 'vml-text-square'), 0)
          .filter((entry): entry is Extract<typeof entry, { drawing: unknown }> => 'drawing' in entry
            && entry.drawing.flowBounds.widthPt === 40);
        expect(entries).toHaveLength(1);
        const { drawing, pointToPage, clips } = entries[0]!;
        const x = pointToPage.a * drawing.flowBounds.xPt + pointToPage.c * drawing.flowBounds.yPt + pointToPage.e;
        const y = pointToPage.b * drawing.flowBounds.xPt + pointToPage.d * drawing.flowBounds.yPt + pointToPage.f;
        expect([Math.round(x * 100) / 100, Math.round(y * 100) / 100]).toEqual([300, 36]);
        expect(clips).toEqual([]);
      }
    } finally {
      variant = {};
    }
  });

  it('keeps the text box transform and clip for page-owned drawings of other producers', () => {
    variant = { rotationDeg: 30 };
    try {
      const fills = paintedFills('drawingml-textbox', 'vml-text-square', true);
      expect(fills).toHaveLength(1);
      expect(fills[0]).not.toEqual({ xPt: 300, yPt: 36, clipped: false });
      expect(fills[0]!.clipped).toBe(true);
    } finally {
      variant = {};
    }
  });

  it('indexes story shapes in paint order: behind text, story content, then the others', () => {
    const rect = (id: string, z: number) => `<w:r><w:pict><v:rect id="${id}" style="position:absolute;margin-left:72pt;`
      + `margin-top:66pt;width:40pt;height:20pt;z-index:${z};mso-position-horizontal-relative:text;`
      + 'mso-position-vertical-relative:text" fillcolor="black" stroked="f"><w10:wrap type="none"/></v:rect></w:pict></w:r>';
    const order = (runs: string) => {
      variant = { storyRuns: runs };
      try {
        return elementGeometryForPage(layOutDocument('drawingml-textbox', 'inline'), 0).flatMap((entry) => {
          if ('drawing' in entry) {
            return entry.drawing.flowBounds.widthPt === 40 ? [entry.drawing.anchorLayer?.behindDoc ? 'behind' : 'front'] : [];
          }
          return ['inline'];
        });
      } finally {
        variant = {};
      }
    };
    expect(order(picture('inline') + rect('front', 1))).toEqual(['inline', 'front']);
    expect(order(rect('front', 1) + rect('behind', -1))).toEqual(['behind', 'front']);
  });

  it('hits a story shape painted over an image that follows its inline text box', () => {
    const inlineBox = (story: string) => `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">`
      + `<wp:extent cx="${100 * EMU}" cy="${20 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="7" name="tb"/>`
      + `<a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/>`
      + `<a:ext cx="${100 * EMU}" cy="${20 * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>`
      + `<wps:txbx><w:txbxContent><w:p>${story}</w:p></w:txbxContent></wps:txbx>`
      + '<wps:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></wps:bodyPr>'
      + '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const entriesFor = (story: string) => {
      variant = { body: `<w:p>${inlineBox(story)}${picture('inline')}</w:p>` };
      try {
        return elementGeometryForPage(layOutDocument('body', 'inline'), 0);
      } finally {
        variant = {};
      }
    };
    const image = entriesFor('').find((entry): entry is Extract<typeof entry, { placement: unknown }> => 'placement' in entry)!;
    const origin = { x: image.pointToPage.e + image.placement.bounds.xPt, y: image.pointToPage.f + image.placement.bounds.yPt };
    const rect = `<w:r><w:pict><v:rect style="position:absolute;margin-left:${origin.x}pt;margin-top:${origin.y}pt;`
      + 'width:40pt;height:20pt;z-index:1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page" '
      + 'fillcolor="black" stroked="f"><w10:wrap type="none"/></v:rect></w:pict></w:r>';
    const order = entriesFor(rect).flatMap((entry) => 'drawing' in entry
      ? entry.drawing.anchorLayer ? ['rect'] : [] : ['image']);
    // Canvas paints the image, then the box and its story's shape.
    expect(order).toEqual(['image', 'rect']);
  });

  it('aligns story shapes in their authored margin and page frames (excess14 reductions)', () => {
    const rect = (position: string) => '<w:r><w:pict><v:rect style="position:absolute;margin-left:300pt;margin-top:50pt;'
      + `width:40pt;height:20pt;${position}" fillcolor="black" stroked="f"><w10:wrap type="none"/></v:rect></w:pict></w:r>`;
    const place = (story: Story, position: string) => {
      variant = { storyRuns: rect(position) };
      try {
        return story === 'drawingml-textbox' ? storyDrawingBoxes(story, 'inline') : bodyHostDrawingBoxes(story, 'inline');
      } finally {
        variant = {};
      }
    };
    const frames = (h: string, v: string) => `mso-position-horizontal-relative:${h};mso-position-vertical-relative:${v};`;
    // Letter page, 72pt side margins, 36pt top and bottom margins.
    expect(place('drawingml-textbox', `mso-position-horizontal:right;${frames('margin', 'margin')}`)).toEqual([{ xPt: 500, yPt: 50 }]);
    expect(place('drawingml-textbox', `mso-position-horizontal:outside;${frames('page', 'margin')}`)).toEqual([{ xPt: 572, yPt: 50 }]);
    expect(place('drawingml-textbox', `mso-position-vertical:bottom;${frames('margin', 'page')}`)).toEqual([{ xPt: 300, yPt: 772 }]);
    expect(place('drawingml-textbox', `mso-position-vertical:center;${frames('margin', 'margin')}`)).toEqual([{ xPt: 300, yPt: 386 }]);
    expect(place('drawingml-textbox', `mso-top-percent:200;${frames('margin', 'page')}`)).toEqual([{ xPt: 300, yPt: 158.4 }]);
    // A note shape's margin frame ends at the page's note area.
    variant = { storyRuns: rect(`mso-position-vertical:bottom;${frames('margin', 'margin')}`) };
    try {
      const page = layOut('footnote', 'inline');
      const noteTop = Math.min(...page.layers.notes.map((note) => note.flowBounds.yPt));
      expect(bodyHostDrawingBoxes('footnote', 'inline')).toEqual([{ xPt: 300, yPt: noteTop - 20 }]);
    } finally {
      variant = {};
    }
  });

  it('keeps a deleted note shape deleted in the final view', () => {
    const rect = '<w:r><w:pict><v:rect style="position:absolute;margin-left:300pt;margin-top:0;width:40pt;height:20pt;'
      + 'mso-position-horizontal-relative:text;mso-position-vertical-relative:text" fillcolor="black" stroked="f">'
      + '<w10:wrap type="none"/></v:rect></w:pict></w:r>';
    variant = { storyRuns: `<w:del w:id="9" w:author="a" w:date="2026-01-01T00:00:00Z">${rect}</w:del>` };
    try {
      expect(bodyHostDrawingBoxes('footnote', 'inline')).toEqual([]);
      // The note keeps its following text.
      expect(JSON.stringify(layOut('footnote', 'inline').layers.notes)).toContain('Ab');
    } finally {
      variant = {};
    }
  });
});
