/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import type { DocxDocumentModel } from './types.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { storeZip } from './conformance/generate.js';

// word-story-anchor-inline: Windows Word 16.0.20625 lays a wp:anchor inside a
// text box, footnote or endnote story out exactly like a wp:inline of the same
// extent, whatever its wrap, layer or relative frame. Synthetic reductions of
// those controls: a DrawingML or VML text box, a footnote or an endnote
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
  | 'vml-text-square' | 'vml-text-square-mt50' | 'vml-margin' | 'vml-text-none';

const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);
const text = (value: string) => `<w:p><w:r><w:t>${value}</w:t></w:r></w:p>`;

const W10 = 'urn:schemas-microsoft-com:office:word';

function vmlRect(placement: Placement): string {
  const frames = placement === 'vml-margin' ? 'margin' : 'text';
  const top = placement === 'vml-text-square-mt50' ? 50 : 0;
  const wrap = placement === 'vml-text-none' ? 'none' : 'square';
  return `<w:r><w:pict><v:rect style="position:absolute;margin-left:300pt;margin-top:${top}pt;width:40pt;height:20pt;`
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
  const content = text('M01') + `<w:p>${picture(placement)}<w:r><w:t>Ab</w:t></w:r></w:p>` + text('MEND');
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
      + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${468 * EMU}" cy="${200 * EMU}"/></a:xfrm>`
      + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
      + `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>`
      + '<wps:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></wps:bodyPr>'
      + '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>';
  } else {
    body += '<w:p><w:r><w:pict><v:shape id="tb" style="position:absolute;margin-left:0;margin-top:30pt;'
      + 'width:468pt;height:200pt;mso-position-horizontal-relative:margin;mso-position-vertical-relative:margin" '
      + `filled="f" stroked="f"><v:textbox inset="0,0,0,0"><w:txbxContent>${content}</w:txbxContent></v:textbox>`
      + '</v:shape></w:pict></w:r></w:p>';
  }
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" `
    + `xmlns:pic="${PIC}" xmlns:wps="${WPS}" xmlns:v="${V}" xmlns:w10="${W10}"><w:body>${body}`
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" '
    + 'w:left="1440" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr></w:body></w:document>';
  const styles = `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>`
    + '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>';
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
  ).pages[0]!;
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
