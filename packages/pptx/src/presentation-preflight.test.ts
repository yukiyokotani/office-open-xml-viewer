import { describe, expect, it } from 'vitest';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { deserializeWorkerError, serializeWorkerError } from '@silurus/ooxml-core/worker';
import {
  PPTX_MAX_PREFLIGHT_PROJECTION_BYTES,
  PresentationPreflightBuilder,
  assertPresentationPreflightProjectionBytes,
  normalizePresentationBootstrap,
  normalizePresentationPreflight,
} from './presentation-preflight';
import { PptxFontPreloadAccumulator } from './font-requests';
import type { Presentation, Slide } from './types';
import type { PresentationBootstrap } from './worker-protocol';

/** Reference projection: one accumulator fed the whole parsed deck. */
function wholeDeckPreloadNames(pres: Presentation): (string | null | undefined)[] {
  const accumulator = new PptxFontPreloadAccumulator(pres.majorFont, pres.minorFont);
  for (const slide of pres.slides) accumulator.addSlide(slide);
  return accumulator.names();
}

const bootstrap: PresentationBootstrap = {
  slideCount: 2,
  slideWidth: 12_192_000,
  slideHeight: 6_858_000,
  defaultTextColor: '111111',
  majorFont: 'SimSun',
  minorFont: 'Aptos',
  hlinkColor: '0563C1',
  folHlinkColor: null,
  embeddedFonts: [],
  slides: [
    { index: 0, partName: 'ppt/slides/slide1.xml' },
    { index: 1, partName: 'ppt/slides/slide2.xml' },
  ],
};

function slide(index: number, text: string, extra: Partial<Slide> = {}): Slide {
  return {
    index,
    slideNumber: index + 1,
    partName: `ppt/slides/slide${index + 1}.xml`,
    background: null,
    elements: [{
      type: 'shape',
      textBody: {
        paragraphs: [{ runs: [{ type: 'text', text }] }],
      },
    }] as Slide['elements'],
    ...extra,
  };
}

describe('PresentationPreflightBuilder', () => {
  it('retains immutable synchronous facts without retaining a full Slide', () => {
    const first = slide(0, '漢字', {
      notes: 'speaker notes',
      hidden: true,
      elements: [
        ...slide(0, '漢字').elements,
        {
          type: 'media', x: 1, y: 2, width: 3, height: 4,
          rotation: 5, flipH: false, flipV: true, mediaKind: 'video',
          posterPath: 'ppt/media/poster.png', posterMimeType: 'image/png',
          mediaPath: 'ppt/media/movie.mp4', mimeType: 'video/mp4',
        },
      ],
    });
    const second = slide(1, '한국어');
    const builder = new PresentationPreflightBuilder(bootstrap);
    builder.addSlide(first);
    builder.addSlide(second);
    const facts = builder.finish();

    expect(facts.slides[0]).toMatchObject({
      index: 0,
      partName: 'ppt/slides/slide1.xml',
      notes: 'speaker notes',
      hidden: true,
    });
    expect(facts.slides[0].mediaElements).toHaveLength(1);
    expect(facts.slides[0]).not.toHaveProperty('elements');
    expect(Object.isFrozen(facts)).toBe(true);
    expect(Object.isFrozen(facts.slides)).toBe(true);
    expect(Object.isFrozen(facts.slides[0].mediaElements[0])).toBe(true);

    first.notes = 'mutated';
    (first.elements[1] as { mediaPath: string }).mediaPath = 'changed';
    expect(facts.slides[0].notes).toBe('speaker notes');
    expect(facts.slides[0].mediaElements[0].mediaPath).toBe('ppt/media/movie.mp4');
  });

  it('retains detached classic and modern comments only on commented slides', () => {
    const first = slide(0, 'commented', {
      comments: [{
        id: '{ROOT}',
        modernAuthorId: '{ADA}',
        author: 'Ada',
        date: '2026-08-24T12:00:00Z',
        x: 6096000,
        y: 3429000,
        anchors: [{ type: 'drawingElement', elementId: '7', creationId: '{SHAPE}' }],
        status: 'active',
        text: 'Root',
        replies: [{
          id: '{REPLY}', authorId: '{BOB}', author: 'Bob', status: 'active', text: 'Reply',
        }],
      }],
    });
    const builder = new PresentationPreflightBuilder(bootstrap);
    builder.addSlide(first);
    builder.addSlide(slide(1, 'plain'));
    const facts = builder.finish();

    expect(facts.slides[0].comments).toEqual(first.comments);
    expect(facts.slides[1]).not.toHaveProperty('comments');
    expect(Object.isFrozen(facts.slides[0].comments)).toBe(true);
    expect(Object.isFrozen(facts.slides[0].comments?.[0])).toBe(true);
    expect(Object.isFrozen(facts.slides[0].comments?.[0]?.anchors)).toBe(true);
    expect(Object.isFrozen(facts.slides[0].comments?.[0]?.anchors?.[0])).toBe(true);
    expect(Object.isFrozen(facts.slides[0].comments?.[0]?.replies?.[0])).toBe(true);
    first.comments![0]!.text = 'mutated';
    expect(facts.slides[0].comments?.[0]?.text).toBe('Root');
  });

  it('aggregates exactly the same Google-font names as the full Presentation', () => {
    const slides = [slide(0, '漢字 العربية'), slide(1, '한국어 Привет')];
    const builder = new PresentationPreflightBuilder(bootstrap);
    for (const item of slides) builder.addSlide(item);
    const facts = builder.finish();
    const full: Presentation = {
      slideWidth: bootstrap.slideWidth,
      slideHeight: bootstrap.slideHeight,
      slides,
      defaultTextColor: bootstrap.defaultTextColor,
      majorFont: bootstrap.majorFont,
      minorFont: bootstrap.minorFont,
      hlinkColor: bootstrap.hlinkColor ?? undefined,
      folHlinkColor: bootstrap.folHlinkColor ?? undefined,
    };
    expect(facts.fontPreloadNames).toEqual(wholeDeckPreloadNames(full));
  });

  it('accounts for the exact structural JSON projection deterministically', () => {
    const builder = new PresentationPreflightBuilder(bootstrap);
    expect(builder.remainingDescriptorCount).toBe(2);
    builder.addSlide(slide(0, '漢'));
    expect(builder.remainingDescriptorCount).toBe(1);
    const buildingState = {
      slideCount: bootstrap.slideCount,
      slideWidth: bootstrap.slideWidth,
      slideHeight: bootstrap.slideHeight,
      defaultTextColor: bootstrap.defaultTextColor,
      majorFont: bootstrap.majorFont,
      minorFont: bootstrap.minorFont,
      hlinkColor: bootstrap.hlinkColor,
      folHlinkColor: bootstrap.folHlinkColor,
      embeddedFonts: bootstrap.embeddedFonts,
      remainingSlides: [undefined, bootstrap.slides[1]],
      slides: [{
        index: 0,
        partName: 'ppt/slides/slide1.xml',
        notes: null,
        hidden: false,
        mediaElements: [],
      }],
      fontPreloadNames: ['SimSun', 'Aptos', 'Noto Sans SC', 'Noto Serif SC'],
    };
    expect(builder.projectedBytes).toBe(
      new TextEncoder().encode(JSON.stringify(buildingState)).byteLength,
    );
    builder.addSlide(slide(1, 'か'));
    const facts = builder.finish();
    expect(builder.remainingDescriptorCount).toBe(0);
    expect(builder.projectedBytes).toBe(new TextEncoder().encode(JSON.stringify(facts)).byteLength);
  });

  it('prepares transactionally and changes retained state only after Rust ACK commit', () => {
    const builder = new PresentationPreflightBuilder(bootstrap);
    const initialBytes = builder.projectedBytes;
    const first = builder.prepareSlide(slide(0, '漢'));
    expect(first.projectedBytes).toBeGreaterThan(initialBytes);
    expect(builder.acceptedSlideCount).toBe(0);
    expect(builder.remainingDescriptorCount).toBe(2);
    expect(builder.projectedBytes).toBe(initialBytes);
    expect(() => builder.prepareSlide(slide(0, 'duplicate'))).toThrow(/already has a prepared/);

    first.rollback();
    first.rollback();
    expect(() => first.commit()).toThrow(/rolled-back/);
    expect(builder.projectedBytes).toBe(initialBytes);

    const retry = builder.prepareSlide(slide(0, '漢'));
    retry.commit();
    retry.commit();
    expect(() => retry.rollback()).toThrow(/committed/);
    expect(builder.acceptedSlideCount).toBe(1);
    expect(builder.remainingDescriptorCount).toBe(1);
  });

  it('crosses the actual cumulative building-state limit with structured usage', () => {
    const probe = new PresentationPreflightBuilder(bootstrap);
    const candidate = probe.prepareSlide(slide(0, 'candidate notes', { notes: 'x'.repeat(80) }));
    const crossingLimit = candidate.projectedBytes - 1;
    candidate.rollback();
    expect(crossingLimit).toBeGreaterThan(probe.projectedBytes);

    const usage = {
      archiveEntryCount: 12,
      declaredInflatedBytes: 300,
      distinctInflatedBytes: 200,
      operationInflatedBytes: 100,
    };
    const limited = new PresentationPreflightBuilder(bootstrap, {
      maxProjectionBytes: crossingLimit,
    });
    const retainedBeforeFailure = limited.projectedBytes;
    try {
      limited.prepareSlide(slide(0, 'candidate notes', { notes: 'x'.repeat(80) }), usage);
      throw new Error('expected cumulative projection limit');
    } catch (error) {
      expect(error).toBeInstanceOf(OoxmlResourceLimitError);
      const typed = error as OoxmlResourceLimitError;
      expect(typed.details.violation).toMatchObject({
        operation: 'presentation-preflight',
        resource: 'presentation-preflight',
        metric: 'projected-bytes',
        limit: crossingLimit,
        observed: crossingLimit + 1,
        configurable: false,
        usage,
      });
    }
    expect(limited.acceptedSlideCount).toBe(0);
    expect(limited.remainingDescriptorCount).toBe(2);
    expect(limited.projectedBytes).toBe(retainedBeforeFailure);
  });

  it('matches full-presentation fonts when Han, Hangul, and Kana are on separate slides', () => {
    const three: PresentationBootstrap = {
      ...bootstrap,
      slideCount: 3,
      slides: [
        { index: 0, partName: 'ppt/slides/slide1.xml' },
        { index: 1, partName: 'ppt/slides/slide2.xml' },
        { index: 2, partName: 'ppt/slides/slide3.xml' },
      ],
    };
    const slides = [slide(0, '漢'), slide(1, '한'), slide(2, 'か')];
    const builder = new PresentationPreflightBuilder(three);
    for (const item of slides) builder.addSlide(item);
    const facts = builder.finish();
    expect(facts.fontPreloadNames).toEqual(wholeDeckPreloadNames({
      slideWidth: three.slideWidth,
      slideHeight: three.slideHeight,
      slides,
      defaultTextColor: three.defaultTextColor,
      majorFont: three.majorFont,
      minorFont: three.minorFont,
    }));
    expect(facts.fontPreloadNames).toContain('Noto Sans KR');
    expect(facts.fontPreloadNames).toContain('Noto Sans JP');
    // The Han-only slide retains its theme-derived SC route even when later
    // slides introduce stronger Hangul/Kana hints.
    expect(facts.fontPreloadNames).toContain('Noto Sans SC');
  });

  it('rejects identity drift, extra units, and incomplete finalization', () => {
    const builder = new PresentationPreflightBuilder(bootstrap);
    expect(() => builder.addSlide(slide(1, 'wrong order'))).toThrow(/identity/);
    expect(() => builder.finish()).toThrow(/incomplete/);

    const complete = new PresentationPreflightBuilder(bootstrap);
    complete.addSlide(slide(0, 'a'));
    complete.addSlide(slide(1, 'b'));
    expect(() => complete.addSlide(slide(1, 'extra'))).toThrow(/extra slide/);
  });

  it('reports the non-configurable hard ceiling with structured package usage', () => {
    const usage = {
      archiveEntryCount: 7,
      declaredInflatedBytes: 8,
      distinctInflatedBytes: 9,
      operationInflatedBytes: 10,
    };
    expect(() => assertPresentationPreflightProjectionBytes(
      PPTX_MAX_PREFLIGHT_PROJECTION_BYTES + 1,
      usage,
    )).toThrowError(OoxmlResourceLimitError);
    try {
      assertPresentationPreflightProjectionBytes(PPTX_MAX_PREFLIGHT_PROJECTION_BYTES + 1, usage);
    } catch (error) {
      expect((error as OoxmlResourceLimitError).details).toEqual({
        stage: 'parsing',
        violation: {
          format: 'pptx',
          operation: 'presentation-preflight',
          resource: 'presentation-preflight',
          metric: 'projected-bytes',
          limit: PPTX_MAX_PREFLIGHT_PROJECTION_BYTES,
          observed: PPTX_MAX_PREFLIGHT_PROJECTION_BYTES + 1,
          configurable: false,
          usage,
        },
      });
    }
  });

  it('preserves the preflight resource violation across the worker wire', () => {
    let failure: unknown;
    try {
      assertPresentationPreflightProjectionBytes(PPTX_MAX_PREFLIGHT_PROJECTION_BYTES + 1);
    } catch (error) {
      failure = error;
    }
    const restored = deserializeWorkerError(serializeWorkerError(failure));
    expect(restored).toBeInstanceOf(OoxmlResourceLimitError);
    expect((restored as OoxmlResourceLimitError).details.violation).toMatchObject({
      format: 'pptx',
      operation: 'presentation-preflight',
      resource: 'presentation-preflight',
      metric: 'projected-bytes',
      configurable: false,
    });
  });
});

describe('normalizePresentationBootstrap', () => {
  it('copies and freezes the Rust JSON projection', () => {
    const normalized = normalizePresentationBootstrap(bootstrap);
    expect(normalized).toEqual(bootstrap);
    expect(normalized).not.toBe(bootstrap);
    expect(Object.isFrozen(normalized.slides[0])).toBe(true);
  });

  it('copies and freezes validated embedded-font references', () => {
    const embeddedFonts = [{
      fontName: 'Deck Sans',
      style: 'boldItalic' as const,
      partPath: 'ppt/fonts/font1.fntdata',
      contentType: 'application/x-font-ttf' as const,
    }];
    const normalized = normalizePresentationBootstrap({ ...bootstrap, embeddedFonts });
    expect(normalized.embeddedFonts).toEqual(embeddedFonts);
    expect(normalized.embeddedFonts).not.toBe(embeddedFonts);
    expect(Object.isFrozen(normalized.embeddedFonts[0])).toBe(true);
  });

  it('rejects malformed counts and non-canonical ordering', () => {
    expect(() => normalizePresentationBootstrap({ ...bootstrap, slideCount: 1 })).toThrow();
    expect(() => normalizePresentationBootstrap({
      ...bootstrap,
      slides: [{ index: 1 }, { index: 0 }],
    })).toThrow(/slide index/);
  });
});

describe('font-source demand facts (LoadOptions.fonts)', () => {
  const arabicDeck = (index: number, text: string, family: string): Slide => ({
    ...slide(index, text),
    elements: [{ type: 'shape', textBody: { paragraphs: [{ bullet: { type: 'none' },
      runs: [{ type: 'text', text, fontFamily: family, bold: index === 1 }] }] } }] as Slide['elements'],
  });

  it('collects split, tuples and per-family scoped scripts only with a demand collector', () => {
    const ordinary = new PresentationPreflightBuilder(bootstrap);
    ordinary.addSlide(arabicDeck(0, 'العربية', 'Sakkal Majalla'));
    ordinary.addSlide(arabicDeck(1, 'Latin', 'Traditional Arabic'));
    const plain = ordinary.finish();
    expect(plain).not.toHaveProperty('fontPreloadFacts');

    const builder = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: () => 'all' });
    builder.addSlide(arabicDeck(0, 'العربية', 'Sakkal Majalla'));
    // The progressive prefix exposes the committed facts.
    expect(builder.currentFontPreloadFacts?.familyCount).toBe(3);
    builder.addSlide(arabicDeck(1, 'Latin', 'Traditional Arabic'));
    const full = builder.finish();
    // Same names as without sources; the facts only add precision.
    expect(full.fontPreloadNames).toEqual(plain.fontPreloadNames);
    const facts = full.fontPreloadFacts!;
    expect(full.fontPreloadNames.slice(0, facts.familyCount))
      .toEqual(['SimSun', 'Aptos', 'Sakkal Majalla', 'Traditional Arabic']);
    expect(full.fontPreloadNames.slice(facts.familyCount)).toContain('Noto Naskh Arabic');
    // Arabic text proves Sakkal Majalla; Latin-only Traditional Arabic proves
    // nothing; theme fonts govern inherited runs whatever their text.
    expect(facts.substituteScripts).toEqual({ simsun: ['arabic'], aptos: ['arabic'], 'sakkal majalla': ['arabic'] });
    expect(normalizePresentationPreflight(full)).toEqual(full);
  });

  it('rejects malformed facts from a worker', () => {
    const builder = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: () => 'all' });
    builder.addSlide(arabicDeck(0, 'A', 'Arial'));
    builder.addSlide(arabicDeck(1, 'B', 'Arial'));
    const full = builder.finish();
    const facts = full.fontPreloadFacts!;
    for (const bad of [
      { ...facts, familyCount: full.fontPreloadNames.length + 1 },
      { ...facts, substituteScripts: { a: ['latin'] } },
      [],
    ]) {
      expect(() => normalizePresentationPreflight({ ...full, fontPreloadFacts: bad })).toThrow('font demand facts');
    }
  });

  it('keeps facts for families named like Object.prototype members', () => {
    const builder = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: () => 'all' });
    builder.addSlide(arabicDeck(0, 'العربية', '__proto__'));
    builder.addSlide(arabicDeck(1, 'العربية', 'constructor'));
    const full = builder.finish();
    const scripts = full.fontPreloadFacts!.substituteScripts;
    expect(Object.prototype.hasOwnProperty.call(scripts, '__proto__')).toBe(true);
    expect(Object.prototype.hasOwnProperty.call(scripts, 'constructor')).toBe(true);
    // The worker → main round trip keeps them as own entries too.
    const cloned = structuredClone(full);
    const normalized = normalizePresentationPreflight(cloned).fontPreloadFacts!.substituteScripts;
    expect(Object.keys(normalized).sort()).toEqual(Object.keys(scripts).sort());
  });

  it('demands bullet, break and end-of-paragraph faces, not only run families', () => {
    const marked: Slide = {
      ...slide(0, 'Item'),
      elements: [{ type: 'shape', textBody: { paragraphs: [{
        bullet: { type: 'char', char: '•', fontFamily: 'BulletFace' },
        runs: [{ type: 'text', text: 'Item', fontFamily: 'BodyFace' }, { type: 'break', fontFamily: 'BreakFace' }],
        endRunProperties: { fontFamily: 'EndFace' },
      }] } }] as Slide['elements'],
    };
    const builder = new PresentationPreflightBuilder(bootstrap, { collectFontDemand: () => 'all' });
    builder.addSlide(marked);
    builder.addSlide(slide(1, 'B', { elements: [] }));
    const full = builder.finish();
    expect(full.fontPreloadNames.slice(0, full.fontPreloadFacts!.familyCount))
      .toEqual(expect.arrayContaining(['BodyFace', 'BulletFace', 'BreakFace', 'EndFace']));
  });

  it('proves a marker\'s script under the marker\'s own family', () => {
    // Latin body text; an Arabic character bullet in its own face, and an
    // auto-number (ASCII labels) whose face proves nothing.
    const marked: Slide = {
      ...slide(0, 'Latin'),
      elements: [{ type: 'shape', textBody: { paragraphs: [
        { bullet: { type: 'char', char: 'ا', fontFamily: 'Bullet Face' },
          runs: [{ type: 'text', text: 'Latin body', fontFamily: 'Body Face' }] },
        { bullet: { type: 'autoNum', numType: 'arabicPeriod', fontFamily: 'Number Face' },
          runs: [{ type: 'text', text: 'Item', fontFamily: 'Body Face' }] },
      ] } }] as Slide['elements'],
    };
    const builder = new PresentationPreflightBuilder({ ...bootstrap, majorFont: null, minorFont: null },
      { collectFontDemand: () => 'all' });
    builder.addSlide(marked);
    builder.addSlide(slide(1, 'B', { elements: [] }));
    expect(builder.finish().fontPreloadFacts?.substituteScripts).toEqual({ 'bullet face': ['arabic'] });
  });

  it('#1974 keeps slide failures outside the measured projection', () => {
    const healthy = new PresentationPreflightBuilder(bootstrap);
    const broken = new PresentationPreflightBuilder(bootstrap);
    const brokenSlide = slide(1, '', { elements: [], parseError: `ppt/slides/slide2.xml: ${'x'.repeat(5000)}` });
    healthy.addSlide(slide(0, 'a'));
    broken.addSlide(slide(0, 'a'));
    expect(broken.latestPartFailure).toBeUndefined();
    healthy.addSlide(slide(1, '', { elements: [] }));
    broken.addSlide(brokenSlide);
    expect(broken.projectedBytes).toBe(healthy.projectedBytes);
    expect(broken.latestPartFailure).toMatchObject({ kind: 'slide', slideIndex: 1, slideNumber: 2 });
    expect(new TextEncoder().encode(broken.latestPartFailure!.message).length).toBeLessThanOrEqual(1024);
    expect(broken.finish()).toEqual(healthy.finish());
    expect(broken.partFailures).toHaveLength(1);
  });

  it('#1974 a deck exactly at the projection limit still loads with a broken slide', () => {
    const probe = new PresentationPreflightBuilder(bootstrap);
    probe.addSlide(slide(0, 'a'));
    probe.addSlide(slide(1, '', { elements: [] }));
    // The prepare-time peak can exceed the final size; find the smallest limit
    // the healthy deck accepts, then load the broken deck under it.
    let limit = probe.projectedBytes;
    for (;;) {
      try {
        const candidate = new PresentationPreflightBuilder(bootstrap, { maxProjectionBytes: limit });
        candidate.addSlide(slide(0, 'a'));
        candidate.addSlide(slide(1, '', { elements: [] }));
        break;
      } catch {
        limit += 1;
      }
    }
    const broken = new PresentationPreflightBuilder(bootstrap, { maxProjectionBytes: limit });
    broken.addSlide(slide(0, 'a'));
    expect(() => broken.addSlide(slide(1, '', { elements: [], parseError: `ppt/slides/slide2.xml: ${'y'.repeat(4096)}` }))).not.toThrow();
    expect(broken.partFailures.map((f) => f.slideIndex)).toEqual([1]);
  });
});

