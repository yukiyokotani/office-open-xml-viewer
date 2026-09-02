import { describe, expect, it } from 'vitest';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { deserializeWorkerError, serializeWorkerError } from '@silurus/ooxml-core/worker';
import {
  PPTX_MAX_PREFLIGHT_PROJECTION_BYTES,
  PresentationPreflightBuilder,
  assertPresentationPreflightProjectionBytes,
  normalizePresentationBootstrap,
} from './presentation-preflight';
import { pptxFontPreloadNames } from './font-plan';
import type { Presentation, Slide } from './types';
import type { PresentationBootstrap } from './worker-protocol';

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
    expect(facts.fontPreloadNames).toEqual(pptxFontPreloadNames(full));
    expect(facts.fontProviderNames).toEqual(['SimSun', 'Aptos']);
    expect(facts.fontProviderNames).not.toContain('Noto Sans KR');
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
      fontProviderNames: ['SimSun', 'Aptos'],
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
      hardLimitForTesting: crossingLimit,
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
    expect(facts.fontPreloadNames).toEqual(pptxFontPreloadNames({
      slideWidth: three.slideWidth,
      slideHeight: three.slideHeight,
      slides,
      defaultTextColor: three.defaultTextColor,
      majorFont: three.majorFont,
      minorFont: three.minorFont,
    }));
    expect(facts.fontPreloadNames).toContain('Noto Sans KR');
    expect(facts.fontPreloadNames).toContain('Noto Sans JP');
    expect(facts.fontPreloadNames).not.toContain('Noto Sans SC');
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
