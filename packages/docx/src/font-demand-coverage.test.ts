/// <reference types="node" />
import { readdirSync } from 'node:fs';
import { readFile, readdir } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { CONFORMANCE_CASES } from './conformance/cases.js';
import { generateConformanceDocx, storeZip } from './conformance/generate.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { renderDocumentToCanvas } from './renderer.js';
import { docxFontDemandScalars } from './font-source-load.js';
import { metafileDemandScalars } from '@silurus/ooxml-core/internal/metafile-font-demand';
import { emfTextFixture, metafilePaintedTexts, wmfFixture, wmfRecords } from '@silurus/ooxml-core/testing';
import type { DocxDocumentModel } from './types.js';

// Design §6 invariant: finite DOCX scalar demand is a superset of every
// scalar layout measures or paint draws. A scalar outside the demand belongs
// to a unicode-range subset that never loads, so its glyph would silently
// fall back to a host font. Exercised over the synthetic conformance corpus,
// the public demo samples and targeted synthesized-text cases.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const encoder = new TextEncoder();

const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function docx(body: string, settings = '', media: ReadonlyMap<string, Uint8Array> = new Map()): Uint8Array {
  return storeZip(new Map([
    ['[Content_Types].xml', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Default Extension="wmf" ContentType="image/x-wmf"/><Default Extension="emf" ContentType="image/x-emf"/>'
      + '<Default Extension="png" ContentType="image/png"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '</Types>')],
    ['_rels/.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      + '</Relationships>')],
    ['word/_rels/document.xml.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + [...media.keys()].map((name, index) => `<Relationship Id="rIdM${index}" Type="${R}/image" Target="media/${name}"/>`).join('')
      + '</Relationships>')],
    ...[...media].map(([name, bytes]) => [`word/media/${name}`, bytes] as const),
    ['word/document.xml', encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:document xmlns:w="${W}" xmlns:m="${M}" xmlns:r="${R}"`
      + ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
      + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
      + ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
      + `<w:body>${body}`
      + `<w:sectPr>${settings}<w:pgSz w:w="12240" w:h="15840"/>`
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
      + '</w:sectPr></w:body></w:document>')],
  ]));
}

const field = (instruction: string, cached: string, rPr = '') =>
  `<w:r><w:rPr>${rPr}</w:rPr><w:fldChar w:fldCharType="begin"/></w:r>`
  + `<w:r><w:rPr>${rPr}</w:rPr><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>`
  + `<w:r><w:rPr>${rPr}</w:rPr><w:fldChar w:fldCharType="separate"/></w:r>`
  + `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${cached}</w:t></w:r>`
  + `<w:r><w:rPr>${rPr}</w:rPr><w:fldChar w:fldCharType="end"/></w:r>`;
const p = (content: string, pPr = '') => `<w:p><w:pPr>${pPr}</w:pPr>${content}</w:p>`;
const run = (text: string, rPr = '') => `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** An inline picture of the `index`th media part (§20.4.2.8 inline). */
const picture = (index: number) => '<w:r><w:drawing><wp:inline><wp:extent cx="1905000" cy="952500"/><wp:docPr id="' + (index + 1) + '" name="p"/>'
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
  + `<pic:nvPicPr><pic:cNvPr id="${index + 1}" name="p"/><pic:cNvPicPr/></pic:nvPicPr>`
  + `<pic:blipFill><a:blip r:embed="rIdM${index}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
  + '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
  + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

const TARGETED: Record<string, Uint8Array> = {
  // §22.1 OMML: an unconverted equation is sized by its text fallback.
  'equations': docx(p(run('Area ') + `<m:oMath><m:sSup><m:e><m:r><m:t>πr</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>`)
    + `<m:oMathPara><m:oMath><m:f><m:num><m:r><m:t>∑α</m:t></m:r></m:num><m:den><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>β</m:t></m:r></m:e></m:rad></m:den></m:f></m:oMath></m:oMathPara>`),
  // WMF/EMF text records demand their own scalars when they play (also
  // under a raster part name: the players sniff content).
  'metafile pictures': docx(p(picture(0) + picture(1) + picture(2)), '', new Map([
    ['image1.wmf', wmfFixture([wmfRecords.setWindowExt(1000, 1000), wmfRecords.createFont('Times New Roman'),
      wmfRecords.selectObject(0), wmfRecords.textOut(new Uint8Array([0x80, 0x41, 0xe9]))])],
    ['image2.emf', emfTextFixture('Arial', ['Größe →'])],
    ['image3.png', wmfFixture([wmfRecords.setWindowExt(1000, 1000), wmfRecords.createFont('MS Mincho', { charset: 128 }),
      wmfRecords.selectObject(0), wmfRecords.textOut(new Uint8Array([0x93, 0xfa, 0x96, 0x7b]))])],
  ])),
  // The reviewed defect: a caps DATE field paints its picture literals upper-cased.
  'caps date field': docx(p(field('DATE \\@ "yyyy α"', '2020 old', '<w:caps/>'))),
  'small-caps time field': docx(p(field('TIME \\@ "h:mm \'μ\'"', '9:00', '<w:smallCaps/>'))),
  'caps text': docx(p(run('straße αβγ', '<w:caps/>'))),
  'emphasis marks': docx(p(run('強調', '<w:em w:val="comma"/>') + run('dot', '<w:em w:val="circle"/>'))),
  'page number formats': docx(
    p(field('PAGE \\* roman', 'i') + run(' of ') + field('NUMPAGES \\* ALPHABETIC', 'A')),
    '<w:pgNumType w:fmt="upperLetter" w:start="3"/>'),
  'tab leaders': docx(p(run('Chapter') + '<w:r><w:tab/></w:r>' + run('7'),
    '<w:tabs><w:tab w:val="right" w:leader="middleDot" w:pos="8000"/></w:tabs>')
    + p(run('Next') + '<w:r><w:tab/></w:r>' + run('9'),
      '<w:tabs><w:tab w:val="right" w:leader="underscore" w:pos="8000"/></w:tabs>')),
  // §17.3.3.30 sym and §17.3.2.26 rFonts naming symbol-encoded faces.
  'symbol fonts': docx(p('<w:r><w:sym w:font="Wingdings" w:char="F0E0"/></w:r>'
    + run('lnþ(a', '<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/>')
    + run('\uF0B7\uF06C', '<w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/>')
    + run('ab', '<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/><w:caps/>'))),
  // §17.3.2.30 rtl runs and §17.3.1.6 bidi paragraphs (UAX #9 mirroring).
  'right-to-left text': docx(p(run('مرحبا (عالم) [1] «اقتباس»', '<w:rtl/>') + run(' and (latin) &lt;tail&gt;'), '<w:bidi/>')
    + p(run('שלום {עולם}', '<w:rtl/>'))),
  // §17.18.44 kashida justification inserts U+0640 into justified lines.
  'kashida justification': docx(p(run('بين البيت والمدرسة طريق طويل يمر بين الحقول والبيوت القديمة في القرية الصغيرة '.repeat(4), '<w:rtl/>'),
    '<w:bidi/><w:jc w:val="highKashida"/>')),
  // §17.3.3.25 ruby.
  'ruby': docx(p('<w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="10"/><w:hpsRaise w:val="18"/><w:hpsBaseText w:val="21"/><w:lid w:val="ja-JP"/></w:rubyPr>'
    + '<w:rt><w:r><w:rPr><w:sz w:val="10"/></w:rPr><w:t>かんじ</w:t></w:r></w:rt><w:rubyBase><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>漢字</w:t></w:r></w:rubyBase></w:ruby>'
    + run('の読み'))),
  // §17.18.93 vertical section and cell text directions.
  'vertical section': docx(p(run('縦書き、「括弧」（丸）。ABC 12 ー〜')), '<w:textDirection w:val="tbRl"/>'),
  'vertical table cell': docx('<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
    + '<w:tr><w:trPr><w:trHeight w:val="2000"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:textDirection w:val="tbRlV"/></w:tcPr>' + p(run('縦、「セル」。')) + '</w:tc>'
    + '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:textDirection w:val="btLr"/></w:tcPr>' + p(run('Rotated (cell)')) + '</w:tc></w:tr></w:tbl>' + p('')),
  // A PUA code in a Unicode-encoded face paints as itself.
  'PUA code in a text face': docx(p(run('\uF0B7 item', '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'))),
  'east asian empty marks': docx(p('', '<w:rPr><w:rFonts w:eastAsia="MS Mincho" w:hint="eastAsia"/></w:rPr>')
    + p(run('日本語の文', '<w:rFonts w:eastAsia="MS Mincho"/><w:lang w:eastAsia="ja-JP"/>'))),
};

function parse(bytes: Uint8Array): { model: DocxDocumentModel; metafiles: Uint8Array[] } {
  const archive = new DocxArchive(bytes);
  try {
    const model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
    return { model, metafiles: metafileParts(model, (path) => archive.extract_image(path)) };
  } finally {
    archive.free();
  }
}

/** The bytes of every picture part the model names (the players sniff
 * content, as the image loader does; other parts paint no text). */
function metafileParts(model: unknown, read: (path: string) => Uint8Array): Uint8Array[] {
  const paths = new Set<string>();
  const walk = (value: unknown) => {
    if (value === null || typeof value !== 'object') return;
    for (const [key, child] of Object.entries(value)) {
      if (key === 'imagePath' && typeof child === 'string') paths.add(child);
      else walk(child);
    }
  };
  walk(model);
  const parts: Uint8Array[] = [];
  for (const path of paths) {
    try { parts.push(read(path)); } catch { /* an absent part paints nothing */ }
  }
  return parts;
}

/** Canvas stand-in recording every string measured or painted. */
function recordingCanvas() {
  let font = '10px serif';
  const texts: string[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const noop = () => {};
  const metrics = (text: string) => ({
    width: [...text].length * px() * 0.5,
    fontBoundingBoxAscent: px() * 0.8,
    fontBoundingBoxDescent: px() * 0.2,
    actualBoundingBoxAscent: px() * 0.8,
    actualBoundingBoxDescent: px() * 0.2,
    actualBoundingBoxLeft: 0,
    actualBoundingBoxRight: [...text].length * px() * 0.5,
  }) as TextMetrics;
  const ctx = new Proxy({
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    // Symbol-face probes measure in families no source serves (the route
    // sentinels alone, or a family no host installs): no subset to load.
    measureText: (text: string) => {
      if (!/^[^",]*"__ooxml_(?:route_sentinel_[ab]|symbol_probe_absent)_[0-9a-f]+__"$/u.test(font)) texts.push(text);
      return metrics(text);
    },
    fillText: (text: string) => { texts.push(text); },
    strokeText: (text: string) => { texts.push(text); },
    createLinearGradient: () => ({ addColorStop: noop }),
    createRadialGradient: () => ({ addColorStop: noop }),
    createPattern: () => null,
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
    getImageData: (_x: number, _y: number, w: number, h: number) => ({ data: new Uint8ClampedArray(w * h * 4) }),
  } as Record<string | symbol, unknown>, {
    get(target, key) { return key in target ? target[key] : noop; },
    set(target, key, value) { target[key] = value; return true; },
  });
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, ctx: ctx as unknown as CanvasRenderingContext2D, texts };
}

async function paintedScalars(model: DocxDocumentModel): Promise<Set<number>> {
  const { canvas, ctx, texts } = recordingCanvas();
  const services = createLayoutServices(model, { measureContext: ctx });
  const pages = layoutDocument(model, services).pages.length;
  for (let pageIndex = 0; pageIndex < pages; pageIndex++) {
    await renderDocumentToCanvas(model, canvas, pageIndex, {
      // One layout for every page (a fresh service per page lays the whole
      // document out again for each page).
      dpr: 1, width: 612, currentDate: Date.UTC(2026, 9, 26, 9, 0), layoutServices: services,
    });
  }
  const scalars = new Set<number>();
  for (const text of texts) for (const character of text) scalars.add(character.codePointAt(0)!);
  return scalars;
}

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('DOCX font demand covers every measured or painted scalar', () => {
  const check = async (label: string, bytes: Uint8Array): Promise<'all' | 'finite'> => {
    const { model, metafiles } = parse(bytes);
    const demand = docxFontDemandScalars(model);
    if (demand === 'all') return 'all';
    const demanded = new Set(demand);
    const painted = await paintedScalars(model);
    // Metafile pictures: what the players paint, against the demand their
    // image owner runs before they play.
    for (const metafile of metafiles) {
      const own = metafileDemandScalars(metafile);
      if (own === 'all') return 'all';
      for (const cp of own) demanded.add(cp);
      for (const text of metafilePaintedTexts(metafile)) for (const ch of text) painted.add(ch.codePointAt(0)!);
    }
    const missing = [...painted].filter((cp) => !demanded.has(cp))
      .map((cp) => `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`);
    expect(missing, label).toEqual([]);
    return 'finite';
  };

  it('plays the targeted metafile pictures, including one under a raster part name', () => {
    const painted = parse(TARGETED['metafile pictures']!).metafiles.flatMap(metafilePaintedTexts).join('');
    for (const text of ['日本', 'Größe →', '€Aé']) expect(painted).toContain(text);
  });

  it('targeted synthesized-text cases', async () => {
    for (const [label, bytes] of Object.entries(TARGETED)) {
      expect(await check(label, bytes), label).toBe('finite');
    }
  });

  // Opt-in: OOXML_FONT_DEMAND_CORPUS=<dir> also checks every .docx in it
  // (the private corpus, which is not committed).
  it.skipIf(!process.env.OOXML_FONT_DEMAND_CORPUS)('opt-in corpus', async () => {
    const dir = process.env.OOXML_FONT_DEMAND_CORPUS!;
    const outcomes: string[] = [];
    const [shard, shards] = (process.env.OOXML_FONT_DEMAND_SHARD ?? '0/1').split('/').map(Number) as [number, number];
    const names = readdirSync(dir).filter((file) => file.endsWith('.docx')).sort().filter((_, index) => index % shards === shard);
    for (const name of names) {
      const file = `${dir}/${name}`;
      const started = Date.now();
      process.stdout.write(`corpus file ${name}\n`);
      try {
        outcomes.push(await check(file, new Uint8Array(await readFile(file))));
      } catch (error) {
        if ((error as Error)?.name === 'AssertionError') throw new Error(`${name}: ${(error as Error).message}`);
        outcomes.push('unreadable');
      }
      process.stdout.write(`corpus done ${name} ${Date.now() - started} ms\n`);
    }
    process.stdout.write(`.docx corpus outcomes ${JSON.stringify(outcomes.reduce<Record<string, number>>((counts, outcome) => { counts[outcome] = (counts[outcome] ?? 0) + 1; return counts; }, {}))}\n`);
  }, 21_600_000);

  it('synthetic conformance corpus and public demo samples', async () => {
    const outcomes: string[] = [];
    for (const testCase of CONFORMANCE_CASES) outcomes.push(await check(testCase.id, generateConformanceDocx(testCase)));
    const demo = new URL('../public/demo/', import.meta.url);
    for (const name of (await readdir(demo)).filter((file) => file.endsWith('.docx'))) {
      outcomes.push(await check(name, new Uint8Array(await readFile(new URL(name, demo)))));
    }
    // The invariant is vacuous for 'all' demand; most of the corpus is finite.
    expect(outcomes.filter((outcome) => outcome === 'finite').length).toBeGreaterThan(outcomes.length / 2);
  }, 120_000);
});
