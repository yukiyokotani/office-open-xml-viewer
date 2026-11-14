import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { collectFontDemandScalars } from '@silurus/ooxml-core/internal/font-scalar-demand';
import { collectSlideFontDemand } from './font-preload-demand.js';
import { PresentationPreflightBuilder, normalizePresentationPreflight } from './presentation-preflight.js';
import type { Slide } from './types';
import type { PresentationBootstrap } from './worker-protocol.js';
import { officeFaceWeight } from './powerpoint-line-metrics.js';

function slide(text: string, run: Partial<TextRunData> = {}): Slide {
  return { index: 0, slideNumber: 27, background: null, elements: [{ type: 'shape', textBody: {
    vert: 'horz', paragraphs: [{ bullet: { type: 'inherit' }, runs: [{ type: 'text', text, ...run }] }],
  } }] } as Slide;
}
function picture(imagePath: string, mimeType: string) {
  return { type: 'picture', imagePath, mimeType, x: 0, y: 0, width: 1, height: 1, rotation: 0, flipH: false, flipV: false, stroke: null };
}
const bootstrap = { slideCount: 2, slideWidth: 100, slideHeight: 100, defaultTextColor: null,
  majorFont: null, minorFont: null, hlinkColor: null, folHlinkColor: null, embeddedFonts: [],
  slides: [{ index: 0 }, { index: 1 }] } as PresentationBootstrap;

describe('complete conservative PPTX paint demand', () => {
  it('applies exactly the shared core scalar rules to its painted texts', () => {
    // One plain run paints its stored text (preparedPowerPointText is the
    // identity here), so the PPTX traversal must agree with the shared
    // collector on that text for every scalar rule and budget edge.
    const committed = new Set([0x41]);
    for (const text of ['AB漢', 'e\u0301', 'א', '\uf0a7', '\u0378', 'ABCDEFG', '']) {
      for (const budget of [1, 2, 12, 10000]) {
        expect(collectSlideFontDemand(slide(text), committed, budget), `${JSON.stringify(text)} @${budget}`)
          .toEqual(collectFontDemandScalars([text, text], { budget, committed }));
      }
    }
  });

  it('includes caps expansion, actual slide number, unchanged language text and measurement sentinels', () => {
    expect(collectSlideFontDemand(slide('§ß\\', { caps: 'all', lang: 'ja-JP' }), new Set(), 10000))
      .toEqual(expect.arrayContaining([0xa7, 0xdf, 0x53, 0x5c, 0x4d, 0x69, 0x20]));
    expect(collectSlideFontDemand(slide('stored', { fieldType: 'slidenum' }), new Set(), 10000))
      .toEqual(expect.arrayContaining([0x32, 0x37]));
  });

  it('declines canonical changes and generated or unmodelled paint paths', () => {
    for (const run of [{ text: 'e\u0301' }, { text: '\ud800' }, { text: '\u0378' }, { text: '\uf0a7' }, { text: '\uf0a7', fontFamilySym: 'Segoe UI Symbol' }]) {
      expect(collectSlideFontDemand(slide(run.text, run), new Set(), 10000)).toBe('all');
    }
    for (const patch of [
      { parseError: 'broken' },
      { elements: [{ type: 'unknown' }] },
      { elements: [{ type: 'chart', chart: { chartType: 'regionMap', series: [] } }] },
      { elements: [picture('ppt/media/image1.emf', 'image/emf')] },
      { elements: [picture('ppt/media/image1.wmf', 'application/octet-stream')] },
      { elements: [picture('ppt/media/image1.pict', 'image/x-pict')] },
      { elements: [picture('ppt/media/image1.png', 'image/png')], background: { fillType: 'image', imagePath: 'ppt/media/bg.emf', mimeType: 'image/emf' } },
      { elements: [{ type: 'media', posterPath: 'ppt/media/poster.wmf', posterMimeType: 'image/wmf', mediaPath: 'ppt/media/m.mp4', mimeType: 'video/mp4' }] },
    ]) {
      expect(collectSlideFontDemand({ ...slide('A'), ...patch } as Slide, new Set(), 10000), JSON.stringify(patch)).toBe('all');
    }
    const body = (value: Slide) => (value.elements[0] as unknown as { textBody: Record<string, unknown> & { paragraphs: Array<Record<string, unknown>> } }).textBody;
    for (const vert of ['wordArtVert', 'wordArtVertRtl']) {
      const stacked = slide('A');
      body(stacked).vert = vert;
      expect(collectSlideFontDemand(stacked, new Set(), 10000), vert).toBe('all');
    }
    const warped = slide('A');
    body(warped).textWarp = { prst: 'textArchUp' };
    expect(collectSlideFontDemand(warped, new Set(), 10000)).toBe('all');
    const equation = slide('A');
    body(equation).paragraphs[0].runs = [{ type: 'math', nodes: [], display: false }];
    expect(collectSlideFontDemand(equation, new Set(), 10000)).toBe('all');
    // A picture fill or picture marker that is a metafile plays its own text records.
    const filled = slide('A');
    (filled.elements[0] as unknown as Record<string, unknown>).fill = { fillType: 'image', imagePath: 'ppt/media/image2.emf', mimeType: 'image/emf' };
    expect(collectSlideFontDemand(filled, new Set(), 10000)).toBe('all');
    const marker = slide('A');
    body(marker).paragraphs[0].bullet = { type: 'blip', imagePath: 'ppt/media/image3.wmf', mimeType: 'image/wmf' };
    expect(collectSlideFontDemand(marker, new Set(), 10000)).toBe('all');
  });

  it('demands nothing for text-free images and media posters', () => {
    const plain = collectSlideFontDemand(slide('A'), new Set(), 10000);
    for (const patch of [
      { elements: [...slide('A').elements, picture('ppt/media/image1.png', 'image/png'), picture('ppt/media/image2.svg', 'image/svg+xml'), picture('ppt/media/image3.JPG', 'image/jpeg')] },
      { elements: [...slide('A').elements, { type: 'media', posterPath: 'ppt/media/poster.png', posterMimeType: 'image/png', mediaPath: 'ppt/media/m.mp4', mimeType: 'video/mp4' }] },
      { elements: [...slide('A').elements, { type: 'media', posterPath: '', posterMimeType: '', mediaPath: 'ppt/media/m.mp3', mimeType: 'audio/mpeg' }] },
      { background: { fillType: 'image', imagePath: 'ppt/media/bg.gif', mimeType: 'image/gif' } },
    ]) {
      expect(collectSlideFontDemand({ ...slide('A'), ...patch } as Slide, new Set(), 10000), JSON.stringify(patch)).toEqual(plain);
    }
    const marker = slide('A');
    (marker.elements[0] as unknown as { textBody: { paragraphs: Array<{ bullet: unknown }> } }).textBody.paragraphs[0].bullet =
      { type: 'blip', imagePath: 'ppt/media/image3.png', mimeType: 'image/png' };
    expect(collectSlideFontDemand(marker, new Set(), 10000)).toEqual(plain);
  });

  it('demands every label an auto-number scheme can paint, in its marker face', () => {
    const numbered = (numType: string, fontFamily: string | null = null) => {
      const value = slide('A');
      (value.elements[0] as unknown as { textBody: { paragraphs: Array<{ bullet: unknown }> } }).textBody.paragraphs[0].bullet =
        { type: 'autoNum', numType, startAt: 3, color: null, fontFamily };
      const demand = collectSlideFontDemand(value, new Set(), 10000);
      expect(demand, numType).not.toBe('all');
      return new Set(demand as number[]);
    };
    const has = (set: Set<number>, text: string) => [...text].every((ch) => set.has(ch.codePointAt(0)!));
    expect(has(numbered('arabicPeriod'), '0123456789.-')).toBe(true);
    expect(has(numbered('arabicParenBoth'), '()0123456789')).toBe(true);
    expect(has(numbered('alphaLcParenR'), 'abcdefghijklmnopqrstuvwxyz)0123456789')).toBe(true);
    expect(has(numbered('alphaUcPeriod'), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ.')).toBe(true);
    expect(has(numbered('romanUcPeriod'), 'MDCLXVI.')).toBe(true);
    expect(has(numbered('romanLcPlain'), 'mdclxvi')).toBe(true);
    expect(has(numbered('arabicDbPeriod'), '０１２３４５６７８９.')).toBe(true);
    // An unknown scheme formats as arabicPeriod.
    expect(has(numbered('thaiAlphaPeriod'), '0123456789.')).toBe(true);
    // A symbol-encoded marker face demands its conversions too.
    expect(has(numbered('arabicPeriod', 'Wingdings'), '\uf030')).toBe(true);
  });

  it('demands rotated and vertical bodies, including eaVert vertical forms', () => {
    const vertical = (vert: string, text: string) => {
      const value = slide(text);
      (value.elements[0] as unknown as { textBody: { vert: string } }).textBody.vert = vert;
      return collectSlideFontDemand(value, new Set(), 10000);
    };
    for (const vert of ['vert', 'vert270', 'mongolianVert']) {
      expect(vertical(vert, 'Ab、「')).toEqual(collectSlideFontDemand(slide('Ab、「'), new Set(), 10000));
    }
    // 、 → U+FE11, 「 → U+FE41, （ → U+FE35 (core vertical-orientation).
    expect(vertical('eaVert', 'あ、「（')).toEqual(expect.arrayContaining([0x3042, 0x3001, 0xfe11, 0x300c, 0xfe41, 0xff08, 0xfe35]));
  });

  it('demands every field\'s painted text', () => {
    expect(collectSlideFontDemand(slide('11/14/2026', { fieldType: 'datetime1' }), new Set(), 10000))
      .toEqual(collectSlideFontDemand(slide('11/14/2026'), new Set(), 10000));
    expect(collectSlideFontDemand(slide('Fußnote', { fieldType: 'unknown', caps: 'all' }), new Set(), 10000))
      .toEqual(expect.arrayContaining([0xdf, 0x53]));
  });

  it('demands what char markers and symbol-encoded faces may paint, not every subset', () => {
    const bulleted = (char: string, fontFamily: string | null) => {
      const value = slide('A');
      (value.elements[0] as { textBody: { paragraphs: Array<{ bullet: unknown }> } }).textBody.paragraphs[0].bullet =
        { type: 'char', char, fontFamily, color: null, sizePct: null };
      return collectSlideFontDemand(value, new Set(), 10000);
    };
    expect(bulleted('•', null)).toEqual(expect.arrayContaining([0x2022, 0x41]));
    // A Symbol bullet: the authored code and its conversion (U+2022) only.
    const symbol = bulleted('\uf0b7', 'Symbol');
    expect(symbol).toEqual(expect.arrayContaining([0xf0b7, 0x2022]));
    expect((symbol as number[]).length).toBeLessThan(8);
    // A symbol-encoded run face and an a:sym face: authored, F0-shifted and converted forms.
    expect(collectSlideFontDemand(slide('l', { fontFamily: 'Wingdings' }), new Set(), 10000))
      .toEqual(expect.arrayContaining([0x6c, 0xf06c, 0x25cf]));
    expect(collectSlideFontDemand(slide('\uf0e0', { fontFamilySym: 'Wingdings' }), new Set(), 10000))
      .toEqual(expect.arrayContaining([0xf0e0, 0x2192]));
    // The family the renderer paints: a ",<generic>" hint dropped, a theme reference resolved.
    expect(bulleted('l', 'Wingdings,Sans-Serif')).toEqual(expect.arrayContaining([0x25cf]));
    const themed = { ...slide('l', { fontFamily: '+mn-lt' }), theme: { minorFont: 'Wingdings' } } as Slide;
    expect(collectSlideFontDemand(themed, new Set(), 10000)).toEqual(expect.arrayContaining([0x25cf]));
    // The byte budget still applies to the symbol scalars.
    const tight = slide('A');
    (tight.elements[0] as { textBody: { paragraphs: Array<{ bullet: unknown }> } }).textBody.paragraphs[0].bullet =
      { type: 'char', char: '\uf0b7', fontFamily: 'Symbol', color: null, sizePct: null };
    expect(collectSlideFontDemand(tight, new Set(), 20)).toBe('all');
  });

  it('keeps bidi mirroring unmodelled and preserves the no-source preflight', () => {
    const rtl = slide('(');
    (rtl.elements[0] as { textBody: { paragraphs: Array<{ rtl?: boolean }> } }).textBody.paragraphs[0].rtl = true;
    expect(collectSlideFontDemand(rtl, new Set(), 10000)).toBe('all');
    expect(collectSlideFontDemand(slide('א('), new Set(), 10000)).toBe('all');
    expect(collectSlideFontDemand(slide('\u202e('), new Set(), 10000)).toBe('all');
    const ordinary = new PresentationPreflightBuilder(bootstrap);
    ordinary.addSlide(slide('A'));
    ordinary.addSlide({ ...slide('B'), index: 1 });
    expect(ordinary.finish()).not.toHaveProperty('fontPreloadDemand');
  });

  it('commits only new scalar deltas and rolls back both scalars and all transitions', () => {
    const builder = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand, officeFaceWeight });
    builder.addSlide(slide('A'));
    expect(builder.currentFontPreloadDemandDelta).toEqual(expect.arrayContaining([65]));
    builder.prepareSlide({ ...slide('B'), index: 1 }).rollback();
    builder.prepareSlide({ ...slide('e\u0301'), index: 1 }).rollback();
    builder.addSlide({ ...slide('C'), index: 1 });
    expect(builder.currentFontPreloadDemandDelta).toEqual([67]);
    const full = builder.finish();
    expect(full.fontPreloadDemand).toEqual(expect.arrayContaining([65, 67]));
    expect(full.fontPreloadDemand).not.toContain(66);
    expect(normalizePresentationPreflight(full).fontPreloadDemand).toEqual(full.fontPreloadDemand);
    expect(() => normalizePresentationPreflight({ ...full, fontPreloadDemand: [0xd800] })).toThrow('font demand');
  });

  it('collapses optional demand growth before exceeding the retained projection allowance', () => {
    expect(collectSlideFontDemand(slide('ABCDEFG'), new Set(), 3)).toBe('all');
    const baseline = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand, officeFaceWeight });
    const first = baseline.prepareSlide(slide('漢'));
    const firstPeak = first.projectedBytes;
    first.commit();
    const second = baseline.prepareSlide({ ...slide('字'), index: 1 });
    const limit = Math.max(firstPeak, second.projectedBytes) + 10;
    second.commit();
    const builder = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand, officeFaceWeight, maxProjectionBytes: limit });
    builder.addSlide(slide(Array.from({ length: 300 }, (_, i) => String.fromCodePoint(0x4e00 + i)).join('')));
    builder.addSlide({ ...slide('字'), index: 1 });
    expect(builder.projectedBytes).toBeLessThanOrEqual(limit);
    expect(builder.finish().fontPreloadDemand).toBeUndefined();
  });
});


it('evicts optional precision under metadata pressure without rejecting or retaining a rollback copy', () => {
  const firstSlide = slide('A');
  const secondSlide = { ...slide('B'), index: 1, notes: 'n'.repeat(200) };
  // Mandatory font-source facts are charged whenever demand is collected; a
  // collector that always declines to 'all' retains no optional precision.
  const ordinary = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: () => 'all', officeFaceWeight });
  ordinary.addSlide(firstSlide);
  const required = ordinary.prepareSlide(secondSlide);
  const limit = required.projectedBytes;
  required.rollback();
  const builder = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand, officeFaceWeight, maxProjectionBytes: limit });
  builder.addSlide(firstSlide);
  expect(builder.currentFontPreloadDemandDelta).toEqual(expect.arrayContaining([65]));
  const prepared = builder.prepareSlide(secondSlide);
  expect(prepared.projectedBytes).toBeLessThanOrEqual(limit);
  prepared.rollback();
  expect(builder.acceptedSlideCount).toBe(1);
  expect(builder.snapshot().slides[0].notes).toBeNull();
  expect(builder.currentFontPreloadDemandDelta).toBe('all');
  builder.addSlide(secondSlide);
  expect(builder.finish().fontPreloadDemand).toBeUndefined();
});


it('does not repair a saturated mandatory projection by subtracting optional cache bytes', () => {
  const ordinary = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: () => 'all', officeFaceWeight });
  ordinary.addSlide(slide('A'));
  const second = { ...slide('B'), index: 1, notes: 'n'.repeat(100) };
  const probe = ordinary.prepareSlide(second);
  const limit = probe.projectedBytes;
  probe.rollback();
  const builder = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand, officeFaceWeight, maxProjectionBytes: limit });
  builder.addSlide(slide('A'));
  const prior = builder.currentFontPreloadDemandDelta;
  expect(() => builder.prepareSlide({ ...second, notes: 'n'.repeat(1000) })).toThrow('hard limit');
  expect(builder.acceptedSlideCount).toBe(1);
  expect(builder.currentFontPreloadDemandDelta).toEqual(prior);
});

it.each([0, 1])('falls back to all when an empty delta array cannot fit its two-byte JSON representation (%i bytes free)', (remaining) => {
  const first = slide('A');
  const second = { ...slide('A'), index: 1, notes: 'n'.repeat(200) };
  const probe = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand, officeFaceWeight });
  probe.addSlide(first);
  const prepared = probe.prepareSlide(second);
  const limit = prepared.projectedBytes - 2 + remaining;
  prepared.rollback();
  const ordinary = new PresentationPreflightBuilder(bootstrap, { maxProjectionBytes: limit });
  ordinary.addSlide(first);
  ordinary.addSlide(second);
  const filtered = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand, officeFaceWeight, maxProjectionBytes: limit });
  filtered.addSlide(first);
  expect(() => filtered.addSlide(second)).not.toThrow();
  expect(filtered.finish().fontPreloadDemand).toBeUndefined();
});
