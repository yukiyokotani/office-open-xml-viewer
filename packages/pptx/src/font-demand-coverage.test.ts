/// <reference types="node" />
import { readFileSync, readdirSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { R, deck, para, run, textBox } from './test-support/synthetic-deck.js';
import { initSync, PptxArchive } from './wasm/pptx_parser.js';
import { collectSlideFontDemand } from './font-preload-demand.js';
import { renderSlideWithEmbeddedFonts } from './renderer.js';
import type { Presentation, Slide } from './types';

// Design §6 invariant: finite PPTX scalar demand is a superset of every
// scalar the renderer measures or paints. A scalar outside the demand belongs
// to a unicode-range subset that never loads, so its glyph would silently
// fall back to a host font. Exercised over the public demo deck and targeted
// synthetic slides (charts and raster pictures are exercised by the demo
// deck); every paint path the collector cannot prove declines to 'all'
// (metafiles, region maps, WordArt, stacked vertical text, RTL, equations).

const TARGETED: Record<string, Uint8Array> = {
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

function parse(bytes: Uint8Array): Presentation {
  const archive = new PptxArchive(bytes);
  try {
    return JSON.parse(new TextDecoder().decode(archive.parse())) as Presentation;
  } finally {
    archive.free();
  }
}

async function paintedScalars(pres: Presentation, slide: Slide): Promise<Set<number>> {
  const { canvas, texts } = recordingCanvas();
  await renderSlideWithEmbeddedFonts(canvas, slide, pres.slideWidth, pres.slideHeight, {
    width: 960, dpr: 1,
    defaultTextColor: pres.defaultTextColor,
    majorFont: pres.majorFont, minorFont: pres.minorFont,
    fetchImage: () => Promise.reject(new Error('no images in a finite-demand slide')),
    fetchMedia: () => Promise.reject(new Error('no media in a finite-demand slide')),
  });
  const scalars = new Set<number>();
  for (const text of texts) for (const character of text) scalars.add(character.codePointAt(0)!);
  return scalars;
}

beforeAll(() => {
  initSync({ module: readFileSync(new URL('./wasm/pptx_parser_bg.wasm', import.meta.url)) });
});

describe('PPTX font demand covers every measured or painted scalar', () => {
  const check = async (label: string, bytes: Uint8Array): Promise<Array<'all' | 'finite'>> => {
    const pres = parse(bytes);
    const outcomes: Array<'all' | 'finite'> = [];
    for (const slide of pres.slides) {
      const demand = collectSlideFontDemand(slide, new Set(), Number.MAX_SAFE_INTEGER);
      if (demand === 'all') { outcomes.push('all'); continue; }
      const demanded = new Set(demand);
      const missing = [...await paintedScalars(pres, slide)].filter((cp) => !demanded.has(cp))
        .map((cp) => `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`);
      expect(missing, `${label} slide ${slide.index}`).toEqual([]);
      outcomes.push('finite');
    }
    return outcomes;
  };

  it('targeted synthetic slides', async () => {
    for (const [label, bytes] of Object.entries(TARGETED)) {
      expect(await check(label, bytes), label).toEqual(['finite']);
    }
  });

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
