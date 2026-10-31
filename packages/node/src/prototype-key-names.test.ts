import { describe, expect, it } from 'vitest';
import { storedZip } from '@silurus/ooxml-core/testing';
import { renderSlideNode } from './render';
import { materializeDocxDocument, openDocxDocument } from './docx';
import { materializeXlsxWorkbook } from './xlsx';
import { materializePptxPresentation } from './pptx';
import type { NodeCanvasFactory } from './render';
import type { Presentation } from '@silurus/ooxml-pptx';
import { loadSkiaForTests } from './test-imports';

// Font, preset-geometry, text-warp and shadow-alignment names are document
// strings. Names that are also Object.prototype members ("constructor",
// "__proto__", "toString") used to resolve to inherited properties of the
// lookup tables and made the whole slide render reject. They must behave like
// any other unknown name.
const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;
const factory: NodeCanvasFactory | null = Canvas && loadImage
  ? {
      createCanvas: (w, h) => new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
      loadImage: (buf) => loadImage(buf as Buffer) as unknown as ReturnType<NodeCanvasFactory['loadImage']>,
    }
  : null;

const px = (n: number) => Math.round(n * 9525);

function slideWith(geometry: string, warp: string | null, font: string, shadowAlign: string | null = null): Presentation {
  const shape = {
    type: 'shape', x: px(40), y: px(40), width: px(320), height: px(160), rotation: 0, flipH: false, flipV: false,
    geometry, fill: { type: 'solid', color: '3366CC' }, stroke: null, defaultTextColor: null, custGeom: null,
    adj: null, adj2: null, adj3: null, adj4: null, adj5: null, adj6: null, adj7: null, adj8: null,
    shadow: shadowAlign === null ? null : {
      color: '000000', alpha: 0.5, blur: px(4), dist: px(4), dir: 45, sx: 0.9, sy: 0.9, kx: 10, ky: 0, algn: shadowAlign,
    },
    textBody: {
      verticalAnchor: 'ctr', defaultFontSize: 40, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'none', vert: 'horz', autoFit: 'none',
      ...(warp ? { textWarp: { preset: warp } } : {}),
      paragraphs: [{
        alignment: 'ctr', marL: 0, marR: 0, indent: 0, spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
        bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
        defFontFamily: null, tabStops: [], eaLnBrk: true,
        runs: [{
          type: 'text', text: 'HI', bold: false, italic: null, underline: false, strikethrough: false,
          fontSize: 40, color: '000000', fontFamily: font,
        }],
      }],
    },
  };
  return {
    slideWidth: px(400), slideHeight: px(240), defaultTextColor: null, majorFont: font, minorFont: font,
    slides: [{ index: 0, slideNumber: 1, background: null, elements: [shape] }],
  } as unknown as Presentation;
}

describe.skipIf(!factory)('document names that collide with Object.prototype members', () => {
  for (const name of ['constructor', '__proto__', 'toString']) {
    it(`renders a slide whose font, geometry, warp and shadow alignment are named "${name}"`, async () => {
      const cases = [
        slideWith(name, null, 'Arial'),
        slideWith('rect', name, 'Arial'),
        slideWith('rect', null, name),
        slideWith('rect', null, 'Arial', name),
      ];
      for (const pres of cases) {
        await expect(renderSlideNode(new Canvas(400, 240), pres, 0, { width: 400, dpr: 1, factory: factory! }))
          .resolves.not.toThrow();
      }
    });
  }
});

// Real-WASM packages whose font, style, theme and sheet names are prototype
// member names. Parser maps (fontTable class/pitch/charset) arrive via
// JSON.parse with these names as own keys; every reader must still resolve
// them, and absent names must not resolve to inherited values.
const PROTO_NAMES = ['__proto__', 'constructor', 'toString', 'hasOwnProperty'];
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
const rel = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
const CONTENT_TYPES = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>';

function prototypeNamedDocx(): Uint8Array {
  // Half the names get fontTable entries (own parser-map keys), half do not.
  const fonts = PROTO_NAMES.slice(0, 2).map((name, index) => `<w:font w:name="${name}"><w:charset w:val="${index ? '80' : '86'}"/><w:family w:val="${index ? 'modern' : 'roman'}"/><w:pitch w:val="${index ? 'fixed' : 'variable'}"/></w:font>`).join('');
  const styles = PROTO_NAMES.map((name) => `<w:style w:type="paragraph" w:styleId="${name}"><w:name w:val="${name}"/><w:basedOn w:val="${name === '__proto__' ? 'constructor' : 'Normal'}"/><w:rPr><w:rFonts w:ascii="${name}" w:eastAsia="${name}"/></w:rPr></w:style>`).join('');
  const paragraphs = PROTO_NAMES.map((name) => `<w:p><w:pPr><w:pStyle w:val="${name}"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${name}" w:hAnsi="${name}" w:eastAsia="${name}" w:cs="${name}" w:hint="eastAsia"/></w:rPr><w:t>Text Ā 漢字 ${name}</w:t></w:r></w:p>`).join('');
  return storedZip([
    ['[Content_Types].xml', CONTENT_TYPES],
    ['_rels/.rels', rels(rel('rId1', 'officeDocument', 'word/document.xml'))],
    ['word/_rels/document.xml.rels', rels(rel('rId1', 'styles', 'styles.xml') + rel('rId2', 'fontTable', 'fontTable.xml'))],
    ['word/fontTable.xml', `<w:fonts xmlns:w="${W}">${fonts}</w:fonts>`],
    ['word/styles.xml', `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${styles}</w:styles>`],
    ['word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${paragraphs}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`],
  ]);
}

function prototypeNamedXlsx(): Uint8Array {
  const sheets = PROTO_NAMES.map((name, index) => `<sheet name="${name}" sheetId="${index + 1}" r:id="s${index}"/>`).join('');
  const names = PROTO_NAMES.map((name) => `<definedName name="${name}">'${name}'!$A$1</definedName>`).join('');
  const fonts = PROTO_NAMES.map((name) => `<font><name val="${name}"/><sz val="11"/></font>`).join('');
  return storedZip([
    ['[Content_Types].xml', CONTENT_TYPES],
    ['_rels/.rels', rels(rel('wb', 'officeDocument', 'xl/workbook.xml'))],
    ['xl/workbook.xml', `<workbook xmlns="${S}" xmlns:r="${R}"><sheets>${sheets}</sheets><definedNames>${names}</definedNames></workbook>`],
    ['xl/_rels/workbook.xml.rels', rels(PROTO_NAMES.map((_, index) => rel(`s${index}`, 'worksheet', `worksheets/sheet${index}.xml`)).join('') + rel('st', 'styles', 'styles.xml'))],
    ['xl/styles.xml', `<styleSheet xmlns="${S}"><fonts count="${PROTO_NAMES.length}">${fonts}</fonts><cellXfs count="${PROTO_NAMES.length}">${PROTO_NAMES.map((_, index) => `<xf fontId="${index}"/>`).join('')}</cellXfs><cellStyles count="1"><cellStyle name="constructor" xfId="0"/></cellStyles></styleSheet>`],
    ...PROTO_NAMES.map((name, index) => [`xl/worksheets/sheet${index}.xml`,
      `<worksheet xmlns="${S}"><sheetData><row r="1"><c r="A1" s="${index}" t="inlineStr"><is><t>${name}</t></is></c></row></sheetData></worksheet>`] as const),
  ]);
}

function prototypeNamedPptx(): Uint8Array {
  const runs = PROTO_NAMES.map((name) => `<a:r><a:rPr lang="ja-JP"><a:latin typeface="${name}"/><a:ea typeface="${name}"/><a:cs typeface="${name}"/></a:rPr><a:t>${name} 漢</a:t></a:r>`).join('');
  return storedZip([
    ['[Content_Types].xml', CONTENT_TYPES],
    ['_rels/.rels', rels(rel('p', 'officeDocument', 'ppt/presentation.xml'))],
    ['ppt/presentation.xml', `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst><p:sldSz cx="3657600" cy="2286000"/></p:presentation>`],
    ['ppt/_rels/presentation.xml.rels', rels(rel('slide', 'slide', 'slides/slide1.xml') + rel('theme', 'theme', 'theme/theme1.xml'))],
    ['ppt/theme/theme1.xml', `<a:theme xmlns:a="${A}" name="__proto__"><a:themeElements><a:fontScheme name="constructor"><a:majorFont><a:latin typeface="__proto__"/><a:ea typeface="constructor"/><a:cs typeface="toString"/></a:majorFont><a:minorFont><a:latin typeface="constructor"/><a:ea typeface="__proto__"/><a:cs typeface="hasOwnProperty"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`],
    ['ppt/slides/slide1.xml', `<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld name="toString"><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="__proto__"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3657600" cy="2286000"/></a:xfrm><a:prstGeom prst="constructor"/></p:spPr><p:txBody><a:bodyPr/><a:p>${runs}</a:p><a:p><a:r><a:rPr><a:latin typeface="+mj-lt"/><a:ea typeface="+mn-ea"/></a:rPr><a:t>theme</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`],
  ]);
}

describe.skipIf(!factory)('real WASM packages named with Object.prototype members', () => {
  it('lays out and paints every DOCX page whose fonts and styles use these names', async () => {
    const document = await materializeDocxDocument(prototypeNamedDocx());
    expect(document.fontFamilyClasses).toEqual(JSON.parse('{"__proto__":"roman","constructor":"modern"}'));
    const session = await openDocxDocument(prototypeNamedDocx(), { factory: factory! });
    try {
      expect(session.pageCount).toBeGreaterThan(0);
      for await (const page of session.pages({ width: 400, dpr: 1 })) expect(page.canvas.width).toBe(400);
    } finally {
      await session.close();
    }
  });

  it('parses an XLSX whose sheet, defined, style and font names are these names', async () => {
    const workbook = await materializeXlsxWorkbook(prototypeNamedXlsx());
    expect(workbook.workbookIndex.workbook.sheets.map((sheet) => sheet.name)).toEqual(PROTO_NAMES);
    expect(workbook.workbookIndex.styles.fonts.map((font) => font.name)).toEqual(PROTO_NAMES);
    expect(workbook.worksheets).toHaveLength(PROTO_NAMES.length);
    expect(JSON.stringify(workbook.worksheets.map((sheet) => sheet.rows[0]?.cells[0]?.value)))
      .toBe(JSON.stringify(PROTO_NAMES.map((name) => ({ type: 'text', text: name }))));
  });

  it('renders a PPTX slide whose theme, typefaces and shape names are these names', async () => {
    const presentation = await materializePptxPresentation(prototypeNamedPptx());
    await expect(renderSlideNode(new Canvas(400, 250), presentation, 0, { width: 400, dpr: 1, factory: factory! }))
      .resolves.not.toThrow();
  });
});
