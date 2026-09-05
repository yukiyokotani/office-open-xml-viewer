import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { describe, expect, it } from 'vitest';
import type { ParagraphLayoutContext } from '../layout-context.js';
import { createLayoutServices } from '../layout-runtime.js';
import { paragraphAcquisitionInput } from '../parser-model.js';
import type { DocParagraph, DocRun, DocxDocumentModel } from '../types.js';
import type { ParagraphAcquisitionInput } from './text.js';
import type { LayoutServices, SourceRef, WrapExclusion } from './types.js';
import {
  acquireParagraphResult,
  paragraphAcquisitionCacheKey,
  type ParagraphAcquisitionOptions,
} from './paragraph.js';
import {
  accessParagraphWrapRegistry,
} from './paragraph-wrap-registry.js';
import { acquireRegisteredParagraph } from './registered-paragraph-acquisition.js';
import {
  createFieldAcquisitionServicesView,
  createParagraphAcquisitionCacheServicesView,
  paragraphAcquisitionCacheOf,
} from './runtime-state.js';

const source: SourceRef = { story: 'body', storyInstance: 'body', path: [0] };

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '',
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 5,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
    }),
  } as unknown as CanvasRenderingContext2D;
}

function model(): DocxDocumentModel {
  return {
    section: {
      pageWidth: 612, pageHeight: 792,
      marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
    },
    body: [],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as DocxDocumentModel;
}

function textParagraph(value = 'cache me'): DocParagraph {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [{
      type: 'text',
      text: value,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      fontSize: 10,
      color: null,
      fontFamily: 'Test Sans',
      isLink: false,
      background: null,
      vertAlign: null,
      hyperlink: null,
    } as DocRun],
  } as DocParagraph;
}

const context: ParagraphLayoutContext = {
  lineGrid: { active: false, pitchPt: null },
  characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
  rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
  physicalIndentLeftPt: 0,
  physicalIndentRightPt: 0,
  firstIndentPt: 0,
  lineSpacing: null,
  spaceBeforePt: 0,
  spaceAfterPt: 0,
  baseRtl: false,
  isJustified: false,
  stretchLastLine: false,
  tabStops: [],
  hasRuby: false,
  hasEastAsianText: false,
  kinsoku: DEFAULT_KINSOKU_RULES,
  defaultTabPt: 36,
};
const acquisitionMeasureContext = measureContext();
const acquisitionFontFamilyClasses = Object.freeze({});

function scopedServices(): LayoutServices {
  const services = createLayoutServices(model(), { measureContext: measureContext() });
  return createParagraphAcquisitionCacheServicesView(services);
}

function options(
  services: LayoutServices,
  overrides: Partial<ParagraphAcquisitionOptions> = {},
): ParagraphAcquisitionOptions {
  return {
    id: 'paragraph:0',
    source,
    flowDomainId: 'body:page:0:column:0',
    ordinaryFlow: true,
    context: { ...context, tabStops: [...context.tabStops] },
    placement: {
      startYPt: 72,
      paragraphXPt: 72,
      availableWidthPt: 468,
      maximumYPt: 720,
      suppressSpaceBefore: false,
    },
    measurer: {
      context: acquisitionMeasureContext,
      fontFamilyClasses: acquisitionFontFamilyClasses,
    },
    environment: {
      pageIndex: 0,
      totalPages: 1,
      displayPageNumber: 1,
      pageNumberFormat: 'decimal',
      currentDateMs: 100,
      noteNumbers: new Map(),
      pageWritingMode: 'horizontal-tb',
      documentHasEastAsianText: false,
      layoutServices: services,
    },
    exclusions: [],
    paragraphBorderEdges: { top: 'top', bottom: 'bottom' },
    trailingExtentPt: 0,
    ...overrides,
  };
}

describe('paragraph acquisition cache', () => {
  it('reacquires identical geometry after its memo candidate is evicted', () => {
    const services = scopedServices();
    const input = paragraphAcquisitionInput(textParagraph(), source);
    const first = acquireParagraphResult(input, options(services));
    const cache = paragraphAcquisitionCacheOf(services)!;
    for (let index = 0; index < 128; index++) cache.set({}, `evict:${index}`, {});
    const second = acquireParagraphResult(input, options(services));
    expect(second).not.toBe(first);
    expect(second.layout).toEqual(first.layout);
    expect(second.measured).toEqual(first.measured);
    expect(acquireParagraphResult(input, options(services))).toBe(second);
  });

  it('reuses the immutable result across initial and field service views', () => {
    const services = scopedServices();
    const fieldView = createFieldAcquisitionServicesView(services, { totalPages: 1 });
    const input = paragraphAcquisitionInput(textParagraph(), source);
    const first = acquireParagraphResult(input, options(services));
    const firstLayoutLine = first.measured.lines[0]!.layout;
    const firstSegment = firstLayoutLine.segments[0]!;
    const originalMeasuredWidth = firstSegment.measuredWidth;
    let mutationError: unknown;
    try {
      Object.assign(firstSegment, { measuredWidth: originalMeasuredWidth + 100 });
    } catch (error) {
      mutationError = error;
    }
    const second = acquireParagraphResult(input, options(fieldView));

    expect(second).toBe(first);
    expect(mutationError).toBeInstanceOf(TypeError);
    expect(second.measured.lines[0]!.layout.segments[0]!.measuredWidth)
      .toBe(originalMeasuredWidth);
    expect(Object.isFrozen(second)).toBe(true);
    expect(Object.isFrozen(second.measured)).toBe(true);
    expect(Object.isFrozen(second.measured.lines)).toBe(true);
    expect(Object.isFrozen(second.measured.placement)).toBe(true);
    for (const line of second.measured.lines) {
      expect(Object.isFrozen(line)).toBe(true);
      expect(Object.isFrozen(line.layout)).toBe(true);
      expect(Object.isFrozen(line.layout.segments)).toBe(true);
      for (const segment of line.layout.segments) expect(Object.isFrozen(segment)).toBe(true);
    }

    const otherScope = scopedServices();
    expect(acquireParagraphResult(input, options(otherScope))).not.toBe(first);
    expect(acquireParagraphResult(
      paragraphAcquisitionInput(textParagraph(), source),
      options(services),
    )).not.toBe(first);
  });

  it('keeps the wrap-registry transaction outside a cached acquisition', () => {
    const services = scopedServices();
    const input = paragraphAcquisitionInput(textParagraph(), source);
    const owner = {};
    const before = accessParagraphWrapRegistry(owner, 'body:page:0:column:0');
    const first = acquireRegisteredParagraph(owner, input, options(services));
    const afterFirst = accessParagraphWrapRegistry(owner, 'body:page:0:column:0');
    const second = acquireRegisteredParagraph(owner, input, options(services));
    const afterSecond = accessParagraphWrapRegistry(owner, 'body:page:0:column:0');

    expect(second).toBe(first);
    expect(afterFirst).not.toBe(before);
    expect(afterSecond).not.toBe(afterFirst);
  });

  it('does not reuse public anchored drawings across distinct supplied frames', () => {
    const services = scopedServices();
    const input = paragraphAcquisitionInput({
      ...textParagraph(),
      runs: [{
        type: 'image',
        imagePath: 'word/media/image.png',
        mimeType: 'image/png',
        widthPt: 20,
        heightPt: 10,
        anchor: true,
        anchorXPt: 5,
        anchorYPt: 6,
        anchorXRelativeFrom: 'page',
        anchorYRelativeFrom: 'page',
      }],
    } as DocParagraph, source);
    const anchorFrames = {
      page: { xPt: 0, yPt: 0, widthPt: 612, heightPt: 792 },
      margin: { xPt: 72, yPt: 72, widthPt: 468, heightPt: 648 },
      column: { xPt: 72, yPt: 72, widthPt: 468, heightPt: 648 },
      pageParity: 'odd' as const,
    };
    const first = acquireParagraphResult(input, options(services, { anchorFrames }));
    const second = acquireParagraphResult(input, options(services, {
      anchorFrames: {
        ...anchorFrames,
        page: { ...anchorFrames.page, xPt: 100 },
      },
    }));

    expect(first.layout.drawings[0]?.flowBounds.xPt).toBe(5);
    expect(second.layout.drawings[0]?.flowBounds.xPt).toBe(105);
    expect(second).not.toBe(first);
  });

  it('keys every value that can change acquisition output', () => {
    const services = scopedServices();
    const cache = paragraphAcquisitionCacheOf(services);
    expect(cache).toBeDefined();
    const anchoredCompleteTextBox = {
      ...paragraphAcquisitionInput(textParagraph(), source),
      runs: [{
        type: 'shape',
        textBlocks: [],
        anchorAcquisitionInput: { occurrenceId: 'anchor:0' },
        textBoxInput: { kind: 'complete', source, blockCount: 0 },
      }],
    } as unknown as ParagraphAcquisitionInput;
    const base = options(services, {
      anchorCollisions: [],
      anchorCellBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 100 },
      anchorFrames: {
        page: { xPt: 0, yPt: 0, widthPt: 612, heightPt: 792 },
        margin: { xPt: 72, yPt: 72, widthPt: 468, heightPt: 648 },
        column: { xPt: 72, yPt: 72, widthPt: 468, heightPt: 648 },
        pageParity: 'odd',
      },
      acquireCompleteStory: () => ({}) as never,
    });
    const key = (
      optionOverrides: Partial<ParagraphAcquisitionOptions> = {},
      continuation?: Parameters<typeof acquireParagraphResult>[2],
    ) => paragraphAcquisitionCacheKey(
      cache!,
      anchoredCompleteTextBox,
      { ...base, ...optionOverrides },
      continuation,
    );
    const exclusion: WrapExclusion = {
      id: 'float:0',
      wrap: 'square',
      bounds: { xPt: 1, yPt: 2, widthPt: 3, heightPt: 4 },
      polygon: [],
      verticalOwnership: 'host',
    };
    const keys = [
      key(),
      key({ id: 'paragraph:1' }),
      key({ source: { ...source, path: [1] } }),
      key({ flowDomainId: 'other' }),
      key({ ordinaryFlow: false }),
      key({ placement: { ...base.placement, startYPt: 73 } }),
      key({ context: { ...base.context, spaceBeforePt: 1 } }),
      key({ measurer: { ...base.measurer, context: measureContext() } }),
      key({ measurer: { ...base.measurer, fontFamilyClasses: { latin: 'serif' } } }),
      key({ environment: { ...base.environment, pageIndex: 1 } }),
      key({ environment: { ...base.environment, totalPages: 2 } }),
      key({ environment: { ...base.environment, displayPageNumber: 2 } }),
      key({ environment: { ...base.environment, pageNumberFormat: 'upperRoman' } }),
      key({ environment: { ...base.environment, currentDateMs: 101 } }),
      key({ environment: { ...base.environment, noteNumbers: new Map([['footnote:1', 1]]) } }),
      key({ environment: { ...base.environment, noteReferenceNumber: 1 } }),
      key({ environment: { ...base.environment, pageWritingMode: 'vertical-lr' } }),
      key({ environment: { ...base.environment, verticalCJK: true } }),
      key({ environment: { ...base.environment, verticalPageFrame: true } }),
      key({ environment: { ...base.environment, documentHasEastAsianText: true } }),
      key({ environment: { ...base.environment, useFeLayout: true } }),
      key({ environment: {
        ...base.environment,
        balanceSingleByteDoubleByteWidth: true,
      } }),
      key({ environment: {
        ...base.environment,
        resolvedLocalFonts: {
          'test sans': {
            requestedFamily: 'Test Sans',
            ascentPt: 8,
            descentPt: 2,
            lineGapPt: 0,
            unitsPerEm: 10,
          },
        },
      } as never }),
      key({ environment: {
        ...base.environment,
        layoutServices: {
          ...services,
          text: { ...services.text, fingerprint: 'other-text' },
        },
      } }),
      key({ environment: {
        ...base.environment,
        verticalGlyphMeasurement: {
          fingerprint: 'vertical:other',
          measureRunInkExtra: () => 0,
          planRun: () => [],
        },
      } }),
      key({ exclusions: [exclusion] }),
      key({ anchorCollisions: [{
        occurrenceId: 'prior',
        bounds: { xPt: 1, yPt: 2, widthPt: 3, heightPt: 4 },
        horizontalOwnership: 'host',
        verticalOwnership: 'host',
      }] }),
      key({ paragraphBorderEdges: { top: 'none', bottom: 'bottom' } }),
      key({ trailingExtentPt: 1 }),
      key({ containerShading: '#ffffff' }),
      key({ continuesFromPrevious: true }),
      key({ sourceRangeStart: 1 }),
      key({ anchorCellBounds: { xPt: 1, yPt: 0, widthPt: 100, heightPt: 100 } }),
      key({ anchorFrames: {
        ...base.anchorFrames!,
        pageParity: 'even',
      } }),
      key({ acquireCompleteStory: () => ({}) as never }),
      key({}, { boundary: { kind: 'line', index: 1 } } as never),
    ];

    expect(new Set(keys)).toHaveLength(keys.length);
  });
});
