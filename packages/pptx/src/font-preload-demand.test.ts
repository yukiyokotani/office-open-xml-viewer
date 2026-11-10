import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { collectSlideFontDemand } from './font-preload-demand.js';
import { PresentationPreflightBuilder, normalizePresentationPreflight } from './presentation-preflight.js';
import type { Slide } from './types';
import type { PresentationBootstrap } from './worker-protocol.js';

function slide(text: string, run: Partial<TextRunData> = {}): Slide {
  return { index: 0, slideNumber: 27, background: null, elements: [{ type: 'shape', textBody: {
    vert: 'horz', paragraphs: [{ bullet: { type: 'inherit' }, runs: [{ type: 'text', text, ...run }] }],
  } }] } as Slide;
}
const bootstrap = { slideCount: 2, slideWidth: 100, slideHeight: 100, defaultTextColor: null,
  majorFont: null, minorFont: null, hlinkColor: null, folHlinkColor: null, embeddedFonts: [],
  slides: [{ index: 0 }, { index: 1 }] } as PresentationBootstrap;

describe('complete conservative PPTX paint demand', () => {
  it('includes caps expansion, actual slide number, unchanged language text and measurement sentinels', () => {
    expect(collectSlideFontDemand(slide('§ß\\', { caps: 'all', lang: 'ja-JP' }), new Set(), 10000))
      .toEqual(expect.arrayContaining([0xa7, 0xdf, 0x53, 0x5c, 0x4d, 0x69, 0x20]));
    expect(collectSlideFontDemand(slide('stored', { fieldType: 'slidenum' }), new Set(), 10000))
      .toEqual(expect.arrayContaining([0x32, 0x37]));
  });

  it('declines canonical changes and generated or unmodelled paint paths', () => {
    for (const run of [{ text: 'e\u0301' }, { text: '\ud800' }, { text: '\u0378' }, { text: '\uf0a7' }, { text: 'A', fontFamilySym: 'Symbol' }, { text: 'a', fontFamily: 'Symbol' }, { text: 'A', fieldType: 'unknown' }]) {
      expect(collectSlideFontDemand(slide(run.text, run), new Set(), 10000)).toBe('all');
    }
    for (const patch of [{ parseError: 'broken' }, { elements: [{ type: 'chart' }] }, { elements: [{ type: 'picture' }] }]) {
      expect(collectSlideFontDemand({ ...slide('A'), ...patch } as Slide, new Set(), 10000)).toBe('all');
    }
    const bulleted = slide('A');
    (bulleted.elements[0] as { textBody: { paragraphs: Array<{ bullet: unknown }> } }).textBody.paragraphs[0].bullet = { type: 'char', char: '•' };
    expect(collectSlideFontDemand(bulleted, new Set(), 10000)).toBe('all');
    const vertical = slide('A');
    (vertical.elements[0] as { textBody: { vert: string } }).textBody.vert = 'vert';
    expect(collectSlideFontDemand(vertical, new Set(), 10000)).toBe('all');
  });

  it('keeps bidi mirroring unmodelled and preserves Google-off facts', () => {
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
    const builder = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand });
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
    const baseline = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand });
    const first = baseline.prepareSlide(slide('漢'));
    const firstPeak = first.projectedBytes;
    first.commit();
    const second = baseline.prepareSlide({ ...slide('字'), index: 1 });
    const limit = Math.max(firstPeak, second.projectedBytes) + 10;
    second.commit();
    const builder = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand, maxProjectionBytes: limit });
    builder.addSlide(slide(Array.from({ length: 300 }, (_, i) => String.fromCodePoint(0x4e00 + i)).join('')));
    builder.addSlide({ ...slide('字'), index: 1 });
    expect(builder.projectedBytes).toBeLessThanOrEqual(limit);
    expect(builder.finish().fontPreloadDemand).toBeUndefined();
  });
});


it('evicts optional precision under metadata pressure without rejecting or retaining a rollback copy', () => {
  const firstSlide = slide('A');
  const secondSlide = { ...slide('B'), index: 1, notes: 'n'.repeat(200) };
  const ordinary = new PresentationPreflightBuilder(bootstrap);
  ordinary.addSlide(firstSlide);
  const required = ordinary.prepareSlide(secondSlide);
  const limit = required.projectedBytes;
  required.rollback();
  const builder = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand, maxProjectionBytes: limit });
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
  const ordinary = new PresentationPreflightBuilder(bootstrap);
  ordinary.addSlide(slide('A'));
  const second = { ...slide('B'), index: 1, notes: 'n'.repeat(100) };
  const probe = ordinary.prepareSlide(second);
  const limit = probe.projectedBytes;
  probe.rollback();
  const builder = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand, maxProjectionBytes: limit });
  builder.addSlide(slide('A'));
  const prior = builder.currentFontPreloadDemandDelta;
  expect(() => builder.prepareSlide({ ...second, notes: 'n'.repeat(1000) })).toThrow('hard limit');
  expect(builder.acceptedSlideCount).toBe(1);
  expect(builder.currentFontPreloadDemandDelta).toEqual(prior);
});

it.each([0, 1])('falls back to all when an empty delta array cannot fit its two-byte JSON representation (%i bytes free)', (remaining) => {
  const first = slide('A');
  const second = { ...slide('A'), index: 1, notes: 'n'.repeat(200) };
  const probe = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand });
  probe.addSlide(first);
  const prepared = probe.prepareSlide(second);
  const limit = prepared.projectedBytes - 2 + remaining;
  prepared.rollback();
  const ordinary = new PresentationPreflightBuilder(bootstrap, { maxProjectionBytes: limit });
  ordinary.addSlide(first);
  ordinary.addSlide(second);
  const filtered = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: collectSlideFontDemand, maxProjectionBytes: limit });
  filtered.addSlide(first);
  expect(() => filtered.addSlide(second)).not.toThrow();
  expect(filtered.finish().fontPreloadDemand).toBeUndefined();
});
