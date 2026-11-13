/// <reference types="node" />
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
import type { DocxDocumentModel } from './types.js';

// Design §6 invariant: finite DOCX scalar demand is a superset of every
// scalar layout measures or paint draws. A scalar outside the demand belongs
// to a unicode-range subset that never loads, so its glyph would silently
// fall back to a host font. Exercised over the synthetic conformance corpus,
// the public demo samples and targeted synthesized-text cases.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const encoder = new TextEncoder();

function docx(body: string, settings = ''): Uint8Array {
  return storeZip(new Map([
    ['[Content_Types].xml', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '</Types>')],
    ['_rels/.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      + '</Relationships>')],
    ['word/document.xml', encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:document xmlns:w="${W}"><w:body>${body}`
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

const TARGETED: Record<string, Uint8Array> = {
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
  'east asian empty marks': docx(p('', '<w:rPr><w:rFonts w:eastAsia="MS Mincho" w:hint="eastAsia"/></w:rPr>')
    + p(run('日本語の文', '<w:rFonts w:eastAsia="MS Mincho"/><w:lang w:eastAsia="ja-JP"/>'))),
};

function parse(bytes: Uint8Array): DocxDocumentModel {
  const archive = new DocxArchive(bytes);
  try {
    return normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
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
      dpr: 1, width: 612, currentDate: Date.UTC(2026, 9, 26, 9, 0), layoutServices: createLayoutServices(model, { measureContext: ctx }),
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
    const model = parse(bytes);
    const demand = docxFontDemandScalars(model);
    if (demand === 'all') return 'all';
    const demanded = new Set(demand);
    const missing = [...await paintedScalars(model)].filter((cp) => !demanded.has(cp))
      .map((cp) => `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`);
    expect(missing, label).toEqual([]);
    return 'finite';
  };

  it('targeted synthesized-text cases', async () => {
    for (const [label, bytes] of Object.entries(TARGETED)) {
      expect(await check(label, bytes), label).toBe('finite');
    }
  });

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
