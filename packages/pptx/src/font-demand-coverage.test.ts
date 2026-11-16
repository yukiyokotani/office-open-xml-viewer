/// <reference types="node" />
import { readFileSync, readdirSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { R, deck, para, picture, run, textBox } from './test-support/synthetic-deck.js';
import { metafileDemandScalars } from '@silurus/ooxml-core/internal/metafile-font-demand';
import { emfTextFixture, metafilePaintedTexts, wmfFixture, wmfRecords } from '@silurus/ooxml-core/testing';
import { initSync, PptxArchive } from './wasm/pptx_parser.js';
import { collectSlideFontDemand } from './font-preload-demand.js';
import { renderSlideWithEmbeddedFonts } from './renderer.js';
import type { Presentation, Slide } from './types';

// Design §6 invariant: finite PPTX scalar demand is a superset of every
// scalar the renderer measures or paints. A scalar outside the demand belongs
// to a unicode-range subset that never loads, so its glyph would silently
// fall back to a host font. Exercised over the public demo deck and targeted
// synthetic slides (charts and raster pictures are exercised by the demo
// deck), equations (with and without a failing math engine) and metafile
// pictures (played on a recording context, against the demand their image
// owner runs before they play); every paint path the collector cannot prove
// declines to 'all' (region maps, unknown elements).

const MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const equation = (text: string) => `<m:oMath xmlns:m="${MATH}"><m:sSup><m:e><m:r><m:t>${text}</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>`;

const TARGETED: Record<string, Uint8Array> = {
  // §22.1 OMML in a14:m: a failed equation paints its MCE preview, its
  // stored fallback text or the unavailable placeholder.
  'equations': deck(textBox(para(run('Area ') + `<mc:AlternateContent xmlns:mc="${MC}" xmlns:a14="${A14}">`
    + `<mc:Choice Requires="a14"><a14:m>${equation('πr')}</a14:m></mc:Choice>`
    + `<mc:Fallback>${run('πr² ≈ ∑')}</mc:Fallback></mc:AlternateContent>`)
    + para(`<a14:m xmlns:a14="${A14}"><m:oMathPara xmlns:m="${MATH}">${equation('αβ')}</m:oMathPara></a14:m>`))),
  // WMF/EMF text records demand their own scalars when they play (also
  // under a raster part name: the players sniff content).
  'metafile pictures': deck(picture(0) + picture(1) + picture(2), new Map([
    ['image1.wmf', wmfFixture([wmfRecords.setWindowExt(1000, 1000), wmfRecords.createFont('Times New Roman'),
      wmfRecords.selectObject(0), wmfRecords.textOut(new Uint8Array([0x80, 0x41, 0xe9]))])],
    ['image2.emf', emfTextFixture('Arial', ['Größe →'])],
    ['image3.png', wmfFixture([wmfRecords.setWindowExt(1000, 1000), wmfRecords.createFont('MS Mincho', { charset: 128 }),
      wmfRecords.selectObject(0), wmfRecords.textOut(new Uint8Array([0x93, 0xfa, 0x96, 0x7b]))])],
  ])),
  'caps and small caps': deck(textBox(para(run('straße αβγ', ' cap="all"') + run('mixed Case', ' cap="small"')))),
  'slide number field': deck(textBox(para('<a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum">'
    + '<a:rPr lang="en-US" sz="2400"/><a:t>‹#›</a:t></a:fld>' + run(' of deck')))),
  'wrapped and autofit text': deck(textBox(para(run('A long paragraph that must wrap across several lines in a narrow box, with punctuation: (a), [b]; "c" — d…')))),
  'east asian line breaking': deck(textBox(para(run('第二段落は日本語の文章、禁則処理と行分割を確認する。')))),
  'line break': deck(textBox(para(run('before') + '<a:br><a:rPr lang="en-US"/></a:br>' + run('after break')))),
  'tabs and hyperlink': deck(textBox(para(run('Name\tValue') + `<a:r><a:rPr lang="en-US" sz="2400"><a:hlinkClick xmlns:r="${R}" r:id=""/></a:rPr><a:t>link</a:t></a:r>`))),
  'table cells': deck('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
    + '<p:xfrm><a:off x="457200" y="3657600"/><a:ext cx="4572000" cy="740664"/></p:xfrm>'
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>'
    + '<a:tblGrid><a:gridCol w="2286000"/><a:gridCol w="2286000"/></a:tblGrid><a:tr h="370332">'
    + `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${para(run('Cell ß'))}</a:txBody><a:tcPr/></a:tc>`
    + `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${para(run('42%'))}</a:txBody><a:tcPr/></a:tc>`
    + '</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>'),
  'auto-number markers': deck(textBox(
    para(run('First'), '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="romanUcPeriod" startAt="3"/></a:pPr>')
    + para(run('Second'), '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="romanUcPeriod" startAt="3"/></a:pPr>')
    + para(run('Third'), '<a:pPr marL="342900" indent="-342900"><a:buAutoNum type="arabicDbPeriod"/></a:pPr>')
    + para(run('Fourth'), '<a:pPr marL="342900" indent="-342900"><a:buAutoNum type="alphaLcParenBoth" startAt="25"/></a:pPr>')
    + para(run('Fifth'), '<a:pPr marL="342900" indent="-342900"><a:buAutoNum type="alphaLcParenBoth" startAt="25"/></a:pPr>')
    + para(run('Sixth'), '<a:pPr marL="342900" indent="-342900"><a:buAutoNum type="alphaLcParenBoth" startAt="25"/></a:pPr>'))),
  'rotated vertical text': deck(textBox(para(run('Rotated text, ninety degrees')), '<a:bodyPr vert="vert270" wrap="square"/>')),
  'east asian vertical text': deck(textBox(para(run('縦書き、「括弧」（丸）。ABC 12', ' altLang="ja-JP"')), '<a:bodyPr vert="eaVert" wrap="square"/>')),
  'stacked vertical text': deck(textBox(para(run('縦書き、「括弧」（丸）。AB 12', ' altLang="ja-JP"')), '<a:bodyPr vert="wordArtVert" wrap="square"/>')),
  'stacked right-to-left columns': deck(textBox(para(run('右から、「左」へ。')) + para(run('二列目')), '<a:bodyPr vert="wordArtVertRtl" wrap="square"/>')),
  'warped text': deck(textBox(para(run('Arched WordArt, ß &amp; 漢字')), '<a:bodyPr wrap="none"><a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp></a:bodyPr>')),
  'right-to-left paragraphs': deck(textBox(para(run('مرحبا (عالم) [1] «اقتباس»'), '<a:pPr rtl="1" algn="r"/>')
    + para(run('(latin) &lt;tail&gt; שלום'), '<a:pPr rtl="1"/>'))),
  'date field': deck(textBox(para('<a:fld id="{00000000-0000-0000-0000-000000000002}" type="datetime1">'
    + '<a:rPr lang="en-US" sz="2400"/><a:t>11/14/2026</a:t></a:fld>' + run(' today')))),
  'east asian and latin mix': deck(textBox(para(run('漢字かなカナ ABC 123', ' altLang="ja-JP"')))),
  // §21.1.2.4.3 buFont / §21.1.2.3.7 latin / §21.1.2.3.10 sym in symbol-encoded faces.
  'symbol bullet and symbol faces': deck(textBox(
    para(run('Item'), '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Symbol"/><a:buChar char="\uF0B7"/></a:pPr>')
    + para(run('Next'), '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Wingdings"/><a:buChar char="§"/></a:pPr>')
    + para('<a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Wingdings"/></a:rPr><a:t>lnþ(\uF0E0</a:t></a:r>'
      + '<a:r><a:rPr lang="en-US" sz="2400"><a:sym typeface="Symbol"/></a:rPr><a:t>a\uF061\uF0AE</a:t></a:r>'))),
};

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
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0, isIdentity: true }),
    getLineDash: () => [],
    getImageData: (_x: number, _y: number, w: number, h: number) => ({ data: new Uint8ClampedArray(Math.max(0, w * h * 4)) }),
  } as Record<string | symbol, unknown>, {
    get(target, key) { return key in target ? target[key] : noop; },
    set(target, key, value) { target[key] = value; return true; },
  });
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, texts };
}

function parse(bytes: Uint8Array): Presentation & { metafiles: Uint8Array[][] } {
  const archive = new PptxArchive(bytes);
  try {
    const pres = JSON.parse(new TextDecoder().decode(archive.parse())) as Presentation;
    return { ...pres, metafiles: pres.slides.map((slide) => pictureParts(slide, (path) => archive.extract_image(path))) };
  } finally {
    archive.free();
  }
}

/** The bytes of every picture part a slide names (pictures, fills, bullets,
 * posters, chart fills): the players sniff content, as the image loader
 * does, and other parts paint no text. */
function pictureParts(slide: Slide, read: (path: string) => Uint8Array): Uint8Array[] {
  const paths = new Set<string>();
  const walk = (value: unknown) => {
    if (value === null || typeof value !== 'object') return;
    for (const [key, child] of Object.entries(value)) {
      if ((key === 'imagePath' || key === 'posterPath') && typeof child === 'string') paths.add(child);
      else walk(child);
    }
  };
  walk(slide);
  const parts: Uint8Array[] = [];
  for (const path of paths) {
    try { parts.push(read(path)); } catch { /* an absent part paints nothing */ }
  }
  return parts;
}

/** A math renderer whose engine fails: equations take their failure paint. */
const failingMath = {
  loadMathJax: () => Promise.reject(new Error('no math engine in this test')),
  mathMLToSvg: () => Promise.reject(new Error('no math engine in this test')),
};

/** The slide without its image paint (pictures, media, picture fills and
 * markers). Node has no image decoder, and an authored blip effect chain
 * fails strictly on the rejected fetch; images paint no served-face text, so
 * the texts painted are the same. */
function withoutImages(slide: Slide): Slide {
  const strip = (value: unknown): unknown => {
    if (Array.isArray(value)) {
      return value.filter((item) => !(item && typeof item === 'object'
        && ((item as { type?: unknown }).type === 'picture' || (item as { type?: unknown }).type === 'media'))).map(strip);
    }
    if (value === null || typeof value !== 'object') return value;
    const record = value as Record<string, unknown>;
    if (record.type === 'blip') return { type: 'none' };
    if (typeof record.imagePath === 'string') return { fillType: 'none' };
    return Object.fromEntries(Object.entries(record).map(([key, child]) => [key, strip(child)]));
  };
  return strip(slide) as Slide;
}

async function paintedScalars(pres: Presentation, slide: Slide): Promise<Set<number>> {
  try {
    return await paintedScalarsOnce(pres, slide);
  } catch (error) {
    if ((error as Error)?.message !== 'Image effect preparation or evaluation failed') throw error;
    return paintedScalarsOnce(pres, withoutImages(slide));
  }
}

async function paintedScalarsOnce(pres: Presentation, slide: Slide): Promise<Set<number>> {
  const { canvas, texts } = recordingCanvas();
  // Without a math engine and with a failing one (MCE previews, fallback
  // text and the unavailable placeholder).
  for (const math of [undefined, failingMath]) {
    await renderSlideWithEmbeddedFonts(canvas, slide, pres.slideWidth, pres.slideHeight, {
      width: 960, dpr: 1,
      defaultTextColor: pres.defaultTextColor,
      majorFont: pres.majorFont, minorFont: pres.minorFont,
      fetchImage: () => Promise.reject(new Error('no images in a finite-demand slide')),
      fetchMedia: () => Promise.reject(new Error('no media in a finite-demand slide')),
      ...(math ? { math } : {}),
    });
  }
  const scalars = new Set<number>();
  for (const text of texts) for (const character of text) scalars.add(character.codePointAt(0)!);
  return scalars;
}

const paintedOf = (texts: readonly string[]) => new Set(texts.flatMap((text) => [...text].map((ch) => ch.codePointAt(0)!)));
const missingScalars = (painted: ReadonlySet<number>, demanded: ReadonlySet<number>) => [...painted].filter((cp) => !demanded.has(cp))
  .map((cp) => `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`);

beforeAll(() => {
  initSync({ module: readFileSync(new URL('./wasm/pptx_parser_bg.wasm', import.meta.url)) });
});

describe('PPTX font demand covers every measured or painted scalar', () => {
  const check = async (label: string, bytes: Uint8Array): Promise<Array<'all' | 'finite'>> => {
    const pres = parse(bytes);
    const outcomes: Array<'all' | 'finite'> = [];
    for (const [index, slide] of pres.slides.entries()) {
      const demand = collectSlideFontDemand(slide, new Set(), Number.MAX_SAFE_INTEGER);
      if (demand === 'all') { outcomes.push('all'); continue; }
      const demanded = new Set(demand);
      expect(missingScalars(await paintedScalars(pres, slide), demanded), `${label} slide ${slide.index}`).toEqual([]);
      // Each metafile picture: what its player paints, against the slide's
      // demand and the demand its image owner runs before it plays.
      let all = false;
      for (const [picture, metafile] of (pres.metafiles[index] ?? []).entries()) {
        const own = metafileDemandScalars(metafile);
        if (own === 'all') { all = true; break; }
        expect(missingScalars(paintedOf(metafilePaintedTexts(metafile)), new Set([...demanded, ...own])),
          `${label} slide ${slide.index} picture ${picture}`).toEqual([]);
      }
      outcomes.push(all ? 'all' : 'finite');
    }
    return outcomes;
  };

  it('plays the targeted metafile pictures, including one under a raster part name', () => {
    const painted = parse(TARGETED['metafile pictures']!).metafiles.flat().flatMap(metafilePaintedTexts).join('');
    for (const text of ['日本', 'Größe →', '€Aé']) expect(painted).toContain(text);
  });

  it('targeted equations take their failure paint', async () => {
    const pres = parse(TARGETED.equations!);
    const painted = String.fromCodePoint(...await paintedScalars(pres, pres.slides[0]!));
    for (const ch of 'πr²≈∑') expect(painted).toContain(ch);
  });

  it('targeted synthetic slides', async () => {
    for (const [label, bytes] of Object.entries(TARGETED)) {
      expect(await check(label, bytes), label).toEqual(['finite']);
    }
  });

  // Opt-in: OOXML_FONT_DEMAND_CORPUS=<dir> also checks every .pptx in it
  // (the private corpus, which is not committed).
  it.skipIf(!process.env.OOXML_FONT_DEMAND_CORPUS)('opt-in corpus', async () => {
    const dir = process.env.OOXML_FONT_DEMAND_CORPUS!;
    const outcomes: string[] = [];
    const [shard, shards] = (process.env.OOXML_FONT_DEMAND_SHARD ?? '0/1').split('/').map(Number) as [number, number];
    const names = readdirSync(dir).filter((file) => file.endsWith('.pptx')).sort().filter((_, index) => index % shards === shard);
    for (const name of names) {
      const file = `${dir}/${name}`;
      const started = Date.now();
      process.stdout.write(`corpus file ${name}\n`);
      try {
        outcomes.push(...await check(file, new Uint8Array(readFileSync(file))));
      } catch (error) {
        if ((error as Error)?.name === 'AssertionError') throw new Error(`${name}: ${(error as Error).message}`);
        outcomes.push('unreadable');
      }
      process.stdout.write(`corpus done ${name} ${Date.now() - started} ms\n`);
    }
    process.stdout.write(`.pptx corpus outcomes ${JSON.stringify(outcomes.reduce<Record<string, number>>((counts, outcome) => { counts[outcome] = (counts[outcome] ?? 0) + 1; return counts; }, {}))}\n`);
  }, 21_600_000);

  it('public demo deck', async () => {
    const demo = new URL('../public/demo/', import.meta.url);
    const outcomes: Array<'all' | 'finite'> = [];
    const finiteKinds = new Set<string>();
    for (const name of readdirSync(demo).filter((file) => file.endsWith('.pptx'))) {
      const bytes = new Uint8Array(readFileSync(new URL(name, demo)));
      const slideOutcomes = await check(name, bytes);
      outcomes.push(...slideOutcomes);
      parse(bytes).slides.forEach((slide, index) => {
        if (slideOutcomes[index] === 'finite') for (const element of slide.elements) finiteKinds.add(element.type);
      });
    }
    // The invariant is vacuous for 'all' demand: the deck's chart and picture
    // slides must be finite, so their painted chart labels are checked.
    expect(outcomes).toContain('finite');
    expect([...finiteKinds]).toEqual(expect.arrayContaining(['chart', 'picture', 'table', 'shape']));
  }, 120_000);
});
