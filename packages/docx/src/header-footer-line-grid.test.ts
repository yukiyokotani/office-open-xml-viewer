import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { ParagraphLayout, PaintNode } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// `word-header-footer-off-line-grid`: a §17.6.5 docGrid line pitch applies
// to body lines only; header and footer lines keep their own height.

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => {
      const size = px();
      return {
        width: [...text].length * size,
        fontBoundingBoxAscent: size * 0.8, fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8, actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, letterSpacing: '0px',
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  } as unknown as CanvasRenderingContext2D;
}

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeCtx(); }
};

function para(text: string): BodyElement {
  return {
    type: 'paragraph',
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
      vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph & BodyElement;
}

function lineHeights(gridKind: string, textDirection?: string) {
  const section = {
    pageWidth: 600, pageHeight: 800,
    marginTop: 60, marginRight: 90, marginBottom: 80, marginLeft: 40,
    headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
    docGridType: gridKind, docGridLinePitch: 18,
    ...(textDirection ? { textDirection } : {}),
  } as SectionProps;
  const document = {
    section, body: [para('Body')],
    headers: { default: { body: [para('Head')] }, first: null, even: null },
    footers: { default: { body: [para('Foot')] }, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
  const page = layoutDocument(
    document,
    createLayoutServices(document, { measureContext: makeCtx() }),
    { currentDateMs: 0 },
  ).pages[0]!;
  const firstLine = (nodes: readonly PaintNode[]) =>
    nodes.find((node): node is ParagraphLayout => node.kind === 'paragraph')!.lines[0]!.advancePt;
  return {
    body: firstLine(page.layers.body),
    header: firstLine(page.layers.header),
    footer: firstLine(page.layers.footer),
  };
}

describe('header and footer on a §17.6.5 line grid', () => {
  it.each([undefined, 'tbRl', 'btLr', 'tbLrV'])('keeps %s header and footer lines off a "lines" grid', (direction) => {
    const heights = lineHeights('lines', direction);
    expect(heights.body).toBe(18);
    expect(heights.header).toBeLessThan(18);
    expect(heights.footer).toBe(heights.header);
  });

  it.each(['linesAndChars', 'snapToChars'])('keeps header and footer lines off a %s grid too', (kind) => {
    for (const direction of [undefined, 'tbRl']) {
      const heights = lineHeights(kind, direction);
      expect(heights.body).toBe(18);
      expect(heights.header).toBeLessThan(18);
      expect(heights.footer).toBe(heights.header);
    }
  });
});

// Every paragraph laid out anywhere on the page, by its text, with its first
// line's advance.
function lineAdvancesByText(value: unknown, out = new Map<string, number>(), seen = new Set<object>()) {
  if (value === null || typeof value !== 'object' || seen.has(value)) return out;
  seen.add(value);
  const node = value as Partial<ParagraphLayout>;
  if (node.kind === 'paragraph' && Array.isArray(node.lines) && node.lines.length > 0) {
    const text = node.lines.flatMap((line) => line.placements)
      .flatMap((placement) => placement.kind === 'text' ? [placement.text] : []).join('');
    if (text && !out.has(text)) out.set(text, node.lines[0]!.advancePt);
  }
  for (const child of Object.values(value)) lineAdvancesByText(child, out, seen);
  return out;
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const V = 'urn:schemas-microsoft-com:vml';

function parsedHeaderDocument(gridKind: string, textDirection?: string): DocxDocumentModel {
  const p = (text: string, pPr = '') => `<w:p>${pPr}<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const frame = '<w:pPr><w:framePr w:w="1440" w:hAnchor="text" w:vAnchor="text" w:x="0" w:y="0" w:wrap="around"/></w:pPr>';
  const box = '<w:p><w:r><w:pict><v:shape id="box" type="#_x0000_t202" style="width:72pt;height:36pt">'
    + `<v:textbox inset="0,0,0,0"><w:txbxContent>${p('箱')}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`;
  const cell = '<w:tbl><w:tblPr><w:tblW w:w="2880" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>'
    + `<w:tr><w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>${p('Cell')}</w:tc></w:tr></w:tbl>`;
  const header = p('Plain') + p('Framed', frame) + cell + box;
  const files = new Map<string, string>([
    ['word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${p('Body')}`
      + '<w:sectPr><w:headerReference w:type="default" r:id="hdr"/><w:pgSz w:w="12000" w:h="16000"/>'
      + '<w:pgMar w:top="1200" w:right="1800" w:bottom="1600" w:left="800" w:header="720" w:footer="720"/>'
      + (textDirection ? `<w:textDirection w:val="${textDirection}"/>` : '')
      + `<w:docGrid w:type="${gridKind}" w:linePitch="360"/></w:sectPr></w:body></w:document>`],
    ['word/header1.xml', `<w:hdr xmlns:w="${W}" xmlns:v="${V}">${header}</w:hdr>`],
    ['word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="hdr" Type="${R}/header" Target="header1.xml"/></Relationships>`],
    ['_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
      + '</Types>'],
  ]);
  const archive = new DocxArchive(storeZip(new Map([...files].map(([name, xml]) => [name, new TextEncoder().encode(xml)]))));
  try {
    return normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
}

describe('framed paragraphs, table cells and text boxes in a gridded header', () => {
  beforeAll(async () => {
    await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
  });

  it.each([
    ['lines', undefined], ['linesAndChars', undefined], ['snapToChars', undefined],
    ['lines', 'tbRl'], ['lines', 'btLr'], ['lines', 'tbLrV'],
  ])('keeps framed, cell and text-box text of a %s header (%s) off the grid', (gridKind, direction) => {
    const document = parsedHeaderDocument(gridKind, direction);
    const advances = lineAdvancesByText(layoutDocument(
      document,
      createLayoutServices(document, { measureContext: makeCtx() }),
      { currentDateMs: 0 },
    ).pages[0]!.layers);
    expect(advances.get('Plain')).toBeLessThan(18);
    expect(advances.get('Framed')).toBe(advances.get('Plain'));
    // Table cells and text-box stories inside a header are header content.
    expect(advances.get('Cell')).toBe(advances.get('Plain'));
    expect(advances.get('箱')).toBeLessThan(18);
  });

  it('keeps a compatibility text box hosted by a header paragraph off the grid', () => {
    const section = {
      pageWidth: 600, pageHeight: 800,
      marginTop: 60, marginRight: 90, marginBottom: 80, marginLeft: 40,
      headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
      docGridType: 'lines', docGridLinePitch: 18,
    } as SectionProps;
    const host = para('Host') as DocParagraph & BodyElement;
    host.runs.push({
      type: 'shape', inline: true, widthPt: 72, heightPt: 36,
      anchorXPt: 0, anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: false,
      zOrder: 0, stroke: null, fill: null, subpaths: [],
      textBlocks: [{ text: '箱', fontSizePt: 10, fontFamily: 'NotInMetrics', alignment: 'left' }],
    } as unknown as DocParagraph['runs'][number]);
    const document = {
      section, body: [para('Body')],
      headers: { default: { body: [host] }, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const advances = lineAdvancesByText(layoutDocument(
      document,
      createLayoutServices(document, { measureContext: makeCtx() }),
      { currentDateMs: 0 },
    ).pages[0]!.layers);
    expect(advances.get('箱')).toBeLessThan(18);
    expect(advances.get('Body')).toBe(18);
  });
});
