import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import {
  layoutParagraph,
  acquireShapeTextBoxLayout,
  paragraphLayoutFromMeasurement,
  planLine,
  sliceParagraphLayout,
  translateParagraphLayout,
} from './paragraph.js';
import type {
  AcquiredParagraphLayoutInput,
  DrawingPlacement,
  ParagraphLayout,
  ResourcePlacement,
  TabPlacement,
  TextBoxLayout,
  TextPlacement,
} from './types.js';
import { stableFingerprint } from './fingerprint.js';
import type { ParagraphLayoutContext, SectionLayoutContext } from '../layout-context.js';
import type { LayoutImageSeg, LayoutMathSeg, LayoutTabSeg, LayoutTextSeg } from '../line-layout.js';
import type { MeasuredParagraph } from '../paragraph-measure.js';
import { measureParagraph } from '../paragraph-measure.js';
import { createLayoutServices } from '../renderer.js';
import type { DocParagraph } from '../types.js';
import type { AnchorAcquisitionInput } from './anchor-input.js';
import { projectBodyOccurrence } from './occurrence-projection.js';
import { normalizePagePaintNodeAnchors } from './anchor-page-normalization.js';
import { canonicalLogicalToPhysical } from './affine.js';
import { bodyFlowDomainId, createLayoutPage } from './page-factory.js';
import { resolveAnchorFrame } from './anchor-frame.js';

const fontRoute = {
  familyList: '"Test Sans"', scope: 'native', fingerprint: 'test-font-route',
} as const;

const source = { story: 'body', storyInstance: 'body', path: [3] } as const;

const acquisitionContext: ParagraphLayoutContext = {
  lineGrid: { active: false, pitchPt: null },
  characterGrid: { active: false, deltaPt: 0 },
  physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
  lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0,
  baseRtl: false, isJustified: false, stretchLastLine: false,
  tabStops: [], hasRuby: false, hasEastAsianText: false,
  kinsoku: DEFAULT_KINSOKU_RULES, defaultTabPt: 36,
};

function projectMeasuredSegment(
  paragraph: DocParagraph,
  segment: LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg
    | readonly (LayoutTextSeg | LayoutImageSeg | LayoutMathSeg | LayoutTabSeg)[],
  context: ParagraphLayoutContext = acquisitionContext,
  layoutServices?: unknown,
  anchorFrames?: Parameters<typeof paragraphLayoutFromMeasurement>[1]['anchorFrames'],
) {
  const measured = {
    lines: [{
      layout: {
        segments: Array.isArray(segment) ? segment : [segment], height: 10, ascent: 8, descent: 2,
        visibleAscent: 8, visibleDescent: 2, intendedSingle: 10,
        visibleIntendedSingle: 10, xOffset: 0, availWidth: 100,
      },
      topYPt: 10, advancePt: 12,
    }],
    markOnly: false, requestedSpaceBeforePt: 0, requestedSpaceAfterPt: 0,
    uniformRubyAdvancePt: 0, contentStartYPt: 10, contentEndYPt: 22,
    lastLineBelowBaselinePt: 2,
    placement: {
      startYPt: 10, paragraphXPt: 10, availableWidthPt: 100,
      maximumYPt: 500, suppressSpaceBefore: false,
    },
  } as unknown as MeasuredParagraph;
  return paragraphLayoutFromMeasurement(paragraph as never, {
    id: 'measured-segment', source, flowDomainId: 'body', ordinaryFlow: true,
    context,
    placement: measured.placement,
    measurer: {} as never,
    environment: { documentHasEastAsianText: false, layoutServices } as never,
    exclusions: [],
    ...(anchorFrames ? { anchorFrames } : {}),
  }, measured);
}

function retainedAnchor(
  occurrenceId: string,
  overrides: Partial<AnchorAcquisitionInput> = {},
): AnchorAcquisitionInput {
  const missingEdges = {
    topPt: null, topStatus: 'missing', rightPt: null, rightStatus: 'missing',
    bottomPt: null, bottomStatus: 'missing', leftPt: null, leftStatus: 'missing',
  } as const;
  return {
    occurrenceId,
    simplePosition: {
      enabled: false, status: 'valid', xPt: 0, xStatus: 'valid', yPt: 0, yStatus: 'valid',
    },
    horizontal: {
      relativeFrom: 'margin', relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: 2 },
    },
    vertical: {
      relativeFrom: 'paragraph', relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: 3 },
    },
    extent: { widthPt: 20, widthStatus: 'valid', heightPt: 10, heightStatus: 'valid' },
    parentEffectExtent: missingEdges,
    anchorDistances: missingEdges,
    relativeSize: { horizontal: null, vertical: null },
    wrap: {
      kind: 'square', authoredKinds: ['wrapSquare'], side: 'bothSides',
      distances: missingEdges, effectExtent: null, polygon: null,
    },
    behavior: {
      behindDoc: false, behindDocStatus: 'valid',
      relativeHeight: 7, relativeHeightStatus: 'valid',
      locked: false, lockedStatus: 'valid',
      allowOverlap: true, allowOverlapStatus: 'valid',
      layoutInCell: true, layoutInCellStatus: 'valid',
    },
    group: null,
    ...overrides,
  };
}

function text(text: string, start: number, xPt: number): TextPlacement {
  return {
    kind: 'text',
    text,
    range: { start, end: start + text.length },
    origin: { xPt, yPt: 20 },
    bounds: { xPt, yPt: 10, widthPt: text.length * 5, heightPt: 12 },
    advancePt: text.length * 5,
    clusters: [{
      range: { start, end: start + text.length },
      offset: { xPt: 0, yPt: 0 },
      advancePt: text.length * 5,
    }],
    color: { kind: 'explicit', color: '#123456' },
    fontRoute,
    fontSizePt: 10,
    fontWeight: 400,
    fontStyle: 'normal',
    direction: 'ltr',
    paintOps: [{
      text, range: { start, end: start + text.length }, offset: { xPt: 0, yPt: 0 },
      letterSpacingPt: 0, scaleX: 1, direction: 'ltr', kerning: 'auto', writingMode: 'horizontal-tb',
    }],
    decorations: [],
  };
}

function input(overrides: Partial<AcquiredParagraphLayoutInput> = {}): AcquiredParagraphLayoutInput {
  return {
    kind: 'paragraph',
    id: 'paragraph-3',
    source,
    flowDomainId: 'body',
    ordinaryFlow: true,
    flowBounds: { xPt: 10, yPt: 5, widthPt: 100, heightPt: 39 },
    inkBounds: { xPt: 10, yPt: 10, widthPt: 45, heightPt: 24 },
    spacing: { beforePt: 5, afterPt: 6 },
    lines: [
      {
        range: { start: 0, end: 5 },
        bounds: { xPt: 10, yPt: 10, widthPt: 25, heightPt: 12 },
        baselinePt: 20,
        advancePt: 14,
        placements: [text('first', 0, 10)],
      },
      {
        range: { start: 5, end: 9 },
        bounds: { xPt: 10, yPt: 24, widthPt: 20, heightPt: 10 },
        baselinePt: 32,
        advancePt: 14,
        placements: [text('next', 5, 10)],
      },
    ],
    borders: [],
    resources: [],
    drawings: [],
    textBoxes: [],
    events: [],
    exclusions: [],
    ...overrides,
  };
}

describe('layoutParagraph', () => {
  it('retains exact text ranges and distinguishes flow advance from ink height', () => {
    const node = layoutParagraph(input());

    expect(node.lines.map((line) => line.range)).toEqual([
      { start: 0, end: 5 },
      { start: 5, end: 9 },
    ]);
    expect(node.inkBounds.heightPt).toBe(24);
    expect(node.advancePt).toBe(39);
    expect(node.flowBounds.heightPt).toBe(node.advancePt);
    expect(node.lines[1]?.placements[0]).toMatchObject({ text: 'next', origin: { xPt: 10, yPt: 20 } });
  });

  it('charges spacing once and rebases only continuation line geometry', () => {
    const acquired = input();
    const whole = layoutParagraph(acquired);
    const first = sliceParagraphLayout(whole, {
      lineStart: 0, lineEnd: 1, continuesFromPrevious: false, continuesOnNext: true,
    }, 'paragraph-3:0');
    const second = sliceParagraphLayout(whole, {
      lineStart: 1, lineEnd: 2, continuesFromPrevious: true, continuesOnNext: false,
    }, 'paragraph-3:1');

    expect(first.advancePt).toBe(5 + 14);
    expect(second.advancePt).toBe(14 + 6);
    expect(first.advancePt + second.advancePt).toBe(whole.advancePt);
    expect(first.lines[0]).toBe(whole.lines[0]);
    expect(second.lines[0]).not.toBe(whole.lines[1]);
    expect(second.lines[0]?.bounds.yPt).toBe(whole.flowBounds.yPt);
    expect(Object.isFrozen(whole)).toBe(true);
    expect(Object.isFrozen(whole.lines)).toBe(true);
  });

  it('charges the initial jump and only intra-slice inter-line jumps', () => {
    const firstLine = { ...input().lines[0]!, bounds: { ...input().lines[0]!.bounds, yPt: 14 }, advancePt: 10 };
    const secondLine = { ...input().lines[1]!, bounds: { ...input().lines[1]!.bounds, yPt: 30 }, advancePt: 10 };
    const whole = layoutParagraph(input({
      lines: [firstLine, secondLine],
      flowBounds: { xPt: 10, yPt: 5, widthPt: 100, heightPt: 41 },
    }));
    const first = sliceParagraphLayout(whole, {
      lineStart: 0, lineEnd: 1, continuesFromPrevious: false, continuesOnNext: true,
    });
    const second = sliceParagraphLayout(whole, {
      lineStart: 1, lineEnd: 2, continuesFromPrevious: true, continuesOnNext: false,
    });

    expect(first.advancePt).toBe(5 + 4 + 10);
    expect(second.advancePt).toBe(10 + 6);
    expect(first.advancePt + second.advancePt).toBe(whole.advancePt - 6);

    const unsplitLines = sliceParagraphLayout(whole, {
      lineStart: 0, lineEnd: 2, continuesFromPrevious: false, continuesOnNext: false,
    });
    expect(unsplitLines.advancePt).toBe(whole.advancePt);
  });

  it('rebases continuation geometry to the paragraph flow origin without reacquiring it', () => {
    const base = input();
    const placement = {
      ...text('next', 5, 10),
      origin: { xPt: 10, yPt: 32 },
      bounds: { xPt: 10, yPt: 24, widthPt: 20, heightPt: 10 },
      decorations: [{
        kind: 'underline' as const,
        from: { xPt: 10, yPt: 33 }, to: { xPt: 30, yPt: 33 },
        color: '#123456', widthPt: 1, style: 'solid' as const,
      }],
      highlightFragments: [{
        rect: { xPt: 10, yPt: 24, widthPt: 20, heightPt: 10 }, color: '#ffff00',
      }],
      ruby: {
        text: 'ふり', advancePt: 8, authored: { raisePt: 4 },
        paintOps: [{
          text: 'ふり', origin: { xPt: 16, yPt: 20 }, fontRoute,
          fontSizePt: 5, fontWeight: 400, fontStyle: 'normal' as const,
          color: { kind: 'explicit' as const, color: '#123456' },
        }],
      },
      emphasis: {
        authored: 'dot', glyphs: [{
          text: '•', origin: { xPt: 14, yPt: 22 }, fontRoute,
          fontSizePt: 5, fontWeight: 400, fontStyle: 'normal' as const,
          color: { kind: 'explicit' as const, color: '#123456' },
        }], paths: [{
          kind: 'polyline' as const,
          points: [{ xPt: 14, yPt: 22 }, { xPt: 15, yPt: 22 }],
          fill: null, stroke: '#123456', strokeWidthPt: .75,
        }],
      },
      runBorderFragments: [{
        edge: 'top' as const, from: { xPt: 9, yPt: 23 }, to: { xPt: 31, yPt: 23 },
        color: '#123456', widthPt: 1, authoredStyle: 'single', style: 'solid' as const,
      }],
    };
    const drawing = {
      kind: 'drawing', id: 'box:drawing:1', source: { ...source, path: [...source.path, 1] },
      flowDomainId: 'body', ordinaryFlow: false,
      flowBounds: { xPt: 10, yPt: 24, widthPt: 20, heightPt: 10 },
      inkBounds: { xPt: 10, yPt: 24, widthPt: 20, heightPt: 10 },
      advancePt: 0, commands: [],
    } as const;
    const textBox = {
      kind: 'textbox', id: 'box:textbox:1', source: drawing.source,
      flowDomainId: 'textbox', ordinaryFlow: false,
      flowBounds: { xPt: 10, yPt: 24, widthPt: 20, heightPt: 10 },
      inkBounds: { xPt: 10, yPt: 24, widthPt: 20, heightPt: 10 },
      contentBounds: { xPt: 10, yPt: 25, widthPt: 20, heightPt: 8 },
      advancePt: 0, paragraphs: [], writingMode: 'horizontal-tb',
      insets: { topPt: 1, rightPt: 0, bottomPt: 1, leftPt: 0 },
    } as const;
    const secondLine = {
      ...base.lines[1]!,
      placements: [placement, {
        kind: 'drawing' as const, range: { start: 8, end: 9 },
        drawingId: drawing.id, bounds: drawing.inkBounds, advancePt: 0,
      }],
    };
    const whole = layoutParagraph(input({
      lines: [base.lines[0]!, secondLine], drawings: [drawing], textBoxes: [textBox],
      borders: [{
        edge: 'bottom', from: { xPt: 10, yPt: 34 }, to: { xPt: 30, yPt: 34 },
        color: '#000000', widthPt: 1, authoredStyle: 'single', style: 'solid',
      }],
    }));

    const continuation = sliceParagraphLayout(whole, {
      lineStart: 1, lineEnd: 2, continuesFromPrevious: true, continuesOnNext: false,
    });

    expect(continuation.flowBounds.yPt).toBe(5);
    expect(continuation.inkBounds.yPt).toBe(5);
    expect(continuation.lines[0]).toMatchObject({ bounds: { yPt: 5 }, baselinePt: 13 });
    expect(continuation.lines[0]?.placements[0]).toMatchObject({
      origin: { yPt: 13 }, bounds: { yPt: 5 }, decorations: [{ from: { yPt: 14 }, to: { yPt: 14 } }],
      highlightFragments: [{ rect: { yPt: 5 } }],
      ruby: { paintOps: [{ origin: { yPt: 1 } }] },
      emphasis: {
        glyphs: [{ origin: { yPt: 3 } }],
        paths: [{ points: [{ yPt: 3 }, { yPt: 3 }] }],
      },
      runBorderFragments: [{ from: { yPt: 4 }, to: { yPt: 4 } }],
    });
    expect(continuation.drawings[0]).toMatchObject({
      flowBounds: { yPt: 5 }, inkBounds: { yPt: 5 },
    });
    expect(continuation.textBoxes[0]).toMatchObject({
      flowBounds: { yPt: 5 }, inkBounds: { yPt: 5 }, contentBounds: { yPt: 6 },
    });
    expect(continuation.borders[0]).toMatchObject({ from: { yPt: 15 }, to: { yPt: 15 } });
  });

  it('keeps page-owned anchor geometry fixed when its host line is rebased on a continuation', () => {
    const base = input();
    const drawing = {
      kind: 'drawing', id: 'page-anchor', source: { ...source, path: [...source.path, 1] },
      flowDomainId: 'body', ordinaryFlow: false,
      flowBounds: { xPt: 70, yPt: 40, widthPt: 20, heightPt: 10 },
      inkBounds: { xPt: 70, yPt: 40, widthPt: 20, heightPt: 10 },
      advancePt: 0, commands: [{
        kind: 'fill-rect' as const,
        rect: { xPt: 70, yPt: 40, widthPt: 20, heightPt: 10 }, fill: '#000000',
      }],
      anchorLayer: {
        occurrenceId: 'anchor:page', behindDoc: false, relativeHeight: 1,
        coordinateSpace: 'physical-page-points' as const,
        sourceOrder: 1, horizontalOwnership: 'page' as const, verticalOwnership: 'page' as const,
      },
    } as const;
    const secondLine = {
      ...base.lines[1]!,
      placements: [...base.lines[1]!.placements, {
        kind: 'drawing' as const, range: { start: 8, end: 9 },
        drawingId: drawing.id, bounds: drawing.inkBounds, advancePt: 0,
      }],
    };
    const whole = layoutParagraph(input({
      lines: [base.lines[0]!, secondLine], drawings: [drawing],
    }));

    const continuation = sliceParagraphLayout(whole, {
      lineStart: 1, lineEnd: 2, continuesFromPrevious: true, continuesOnNext: false,
    });

    expect(continuation.lines[0]?.bounds.yPt).toBe(whole.flowBounds.yPt);
    expect(continuation.drawings[0]?.flowBounds.yPt).toBe(40);
    expect(continuation.drawings[0]?.commands[0]).toMatchObject({ rect: { yPt: 40 } });
  });

  it('translates complete host-owned frame geometry while preserving nested page-owned anchors', () => {
    const hostDrawing = {
      kind: 'drawing', id: 'host-anchor', source: { ...source, path: [...source.path, 1] },
      flowDomainId: 'body', ordinaryFlow: false,
      flowBounds: { xPt: 10, yPt: 15, widthPt: 20, heightPt: 10 },
      inkBounds: { xPt: 10, yPt: 15, widthPt: 20, heightPt: 10 },
      advancePt: 0, commands: [{
        kind: 'fill-rect' as const,
        rect: { xPt: 10, yPt: 15, widthPt: 20, heightPt: 10 }, fill: '#000000',
      }],
      anchorLayer: {
        occurrenceId: 'anchor:host', behindDoc: false, relativeHeight: 1,
        coordinateSpace: 'physical-page-points' as const,
        sourceOrder: 1, horizontalOwnership: 'host' as const, verticalOwnership: 'host' as const,
      },
    } as const;
    const pageDrawing = {
      ...hostDrawing,
      id: 'page-anchor',
      flowBounds: { xPt: 70, yPt: 40, widthPt: 20, heightPt: 10 },
      inkBounds: { xPt: 70, yPt: 40, widthPt: 20, heightPt: 10 },
      commands: [{
        kind: 'fill-rect' as const,
        rect: { xPt: 70, yPt: 40, widthPt: 20, heightPt: 10 }, fill: '#000000',
      }],
      anchorLayer: {
        ...hostDrawing.anchorLayer,
        occurrenceId: 'anchor:page',
        horizontalOwnership: 'page' as const,
        verticalOwnership: 'page' as const,
      },
    } as const;
    const base = input();
    const whole = layoutParagraph(input({
      lines: [{
        ...base.lines[0]!,
        placements: [
          ...base.lines[0]!.placements,
          {
            kind: 'drawing', range: { start: 4, end: 5 },
            drawingId: hostDrawing.id, bounds: hostDrawing.inkBounds, advancePt: 0,
          },
          {
            kind: 'drawing', range: { start: 5, end: 6 },
            drawingId: pageDrawing.id, bounds: pageDrawing.inkBounds, advancePt: 0,
          },
        ],
      }, base.lines[1]!],
      drawings: [hostDrawing, pageDrawing],
      exclusions: [{
        id: 'page-exclusion', wrap: 'square',
        bounds: pageDrawing.flowBounds,
        polygon: [{ xPt: 70, yPt: 40 }, { xPt: 90, yPt: 50 }],
        anchorOccurrenceId: 'anchor:page', verticalOwnership: 'page',
      }],
    }));

    const translated = translateParagraphLayout(whole, { xPt: 25, yPt: 30 });

    expect(translated.flowBounds).toMatchObject({
      xPt: whole.flowBounds.xPt + 25,
      yPt: whole.flowBounds.yPt + 30,
    });
    expect(translated.lines[0]?.bounds).toMatchObject({
      xPt: whole.lines[0]!.bounds.xPt + 25,
      yPt: whole.lines[0]!.bounds.yPt + 30,
    });
    expect(translated.drawings[0]?.flowBounds).toMatchObject({ xPt: 35, yPt: 45 });
    expect(translated.drawings[0]?.commands[0]).toMatchObject({ rect: { xPt: 35, yPt: 45 } });
    expect(translated.drawings[1]?.flowBounds).toMatchObject({ xPt: 70, yPt: 40 });
    expect(translated.drawings[1]?.commands[0]).toMatchObject({ rect: { xPt: 70, yPt: 40 } });
    expect(translated.lines[0]?.placements.at(-2)).toMatchObject({ bounds: { xPt: 35, yPt: 45 } });
    expect(translated.lines[0]?.placements.at(-1)).toMatchObject({ bounds: { xPt: 70, yPt: 40 } });
    expect(translated.exclusions[0]?.bounds).toMatchObject({ xPt: 70, yPt: 40 });
    expect(translated.exclusions[0]?.polygon[0]).toEqual({ xPt: 70, yPt: 40 });
  });

  it('assigns retained occurrences and the paragraph mark to only their owning slice', () => {
    const firstResource = { kind: 'image', resourceKey: 'image:first', intrinsicSize: { widthPt: 5, heightPt: 5 } } as const;
    const secondResource = { kind: 'chart', resourceKey: 'chart:second', intrinsicSize: { widthPt: 5, heightPt: 5 } } as const;
    const drawing = {
      kind: 'drawing', id: 'drawing:first', source: { ...source, path: [...source.path, 1] },
      flowDomainId: 'body', ordinaryFlow: false,
      flowBounds: { xPt: 0, yPt: 0, widthPt: 5, heightPt: 5 },
      inkBounds: { xPt: 0, yPt: 0, widthPt: 5, heightPt: 5 }, advancePt: 0, commands: [],
    } as const;
    const textBox = {
      kind: 'textbox', id: 'textbox:first', source: drawing.source,
      flowDomainId: 'textbox', ordinaryFlow: false,
      flowBounds: { xPt: 0, yPt: 0, widthPt: 5, heightPt: 5 },
      inkBounds: { xPt: 0, yPt: 0, widthPt: 5, heightPt: 5 }, advancePt: 0,
      paragraphs: [], writingMode: 'horizontal-tb',
      insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
    } as const;
    const base = input();
    const firstLine = {
      ...base.lines[0]!,
      placements: [...base.lines[0]!.placements,
        { kind: 'resource', range: { start: 4, end: 5 }, resourceKey: firstResource.resourceKey, resourceKind: 'image', bounds: base.lines[0]!.bounds, advancePt: 5 } as const,
        { kind: 'drawing', range: { start: 4, end: 5 }, drawingId: drawing.id, bounds: drawing.inkBounds, advancePt: 0 } as const],
    };
    const secondLine = {
      ...base.lines[1]!,
      placements: [...base.lines[1]!.placements,
        { kind: 'resource', range: { start: 8, end: 9 }, resourceKey: secondResource.resourceKey, resourceKind: 'chart', bounds: base.lines[1]!.bounds, advancePt: 5 } as const],
    };
    const whole = layoutParagraph(input({
      lines: [firstLine, secondLine], resources: [firstResource, secondResource],
      drawings: [drawing], textBoxes: [textBox],
      events: [
        { kind: 'break', breakKind: 'page', offset: 2 },
        { kind: 'break', breakKind: 'column', offset: 7 },
      ],
      paragraphMark: { hidden: false, bounds: { xPt: 30, yPt: 24, widthPt: 0, heightPt: 10 } },
    }));
    const first = sliceParagraphLayout(whole, {
      lineStart: 0, lineEnd: 1, continuesFromPrevious: false, continuesOnNext: true,
    });
    const second = sliceParagraphLayout(whole, {
      lineStart: 1, lineEnd: 2, continuesFromPrevious: true, continuesOnNext: false,
    });

    expect(first.resources.map((resource) => resource.resourceKey)).toEqual(['image:first']);
    expect(second.resources.map((resource) => resource.resourceKey)).toEqual(['chart:second']);
    expect(first.drawings.map((item) => item.id)).toEqual(['drawing:first']);
    expect(second.drawings).toEqual([]);
    expect(first.textBoxes.map((item) => item.id)).toEqual(['textbox:first']);
    expect(second.textBoxes).toEqual([]);
    expect(first.events.map((event) => event.offset)).toEqual([2]);
    expect(second.events.map((event) => event.offset)).toEqual([7]);
    expect(first.paragraphMark).toBeUndefined();
    expect(second.paragraphMark).toEqual({
      ...whole.paragraphMark,
      bounds: { ...whole.paragraphMark?.bounds, yPt: whole.flowBounds.yPt },
    });
  });

  it('slices one retained decoration box across paragraph continuations', () => {
    const whole = layoutParagraph(input({
      inkBounds: { xPt: 6, yPt: 8, widthPt: 108, heightPt: 30 },
      shading: { color: '#eeeeee' },
      borders: [
        { edge: 'top', from: { xPt: 6, yPt: 8 }, to: { xPt: 114, yPt: 8 }, color: '#111111', widthPt: 1, authoredStyle: 'single', style: 'solid' },
        { edge: 'right', from: { xPt: 114, yPt: 8 }, to: { xPt: 114, yPt: 38 }, color: '#111111', widthPt: 1, authoredStyle: 'single', style: 'solid' },
        { edge: 'bottom', from: { xPt: 6, yPt: 38 }, to: { xPt: 114, yPt: 38 }, color: '#111111', widthPt: 1, authoredStyle: 'single', style: 'solid' },
        { edge: 'left', from: { xPt: 6, yPt: 8 }, to: { xPt: 6, yPt: 38 }, color: '#111111', widthPt: 1, authoredStyle: 'single', style: 'solid' },
      ],
    }));
    const first = sliceParagraphLayout(whole, {
      lineStart: 0, lineEnd: 1, continuesFromPrevious: false, continuesOnNext: true,
    });
    const second = sliceParagraphLayout(whole, {
      lineStart: 1, lineEnd: 2, continuesFromPrevious: true, continuesOnNext: false,
    });

    expect(first.inkBounds).toEqual({ xPt: 6, yPt: 8, widthPt: 108, heightPt: 16 });
    expect(first.borders).toEqual(expect.arrayContaining([
      expect.objectContaining({ edge: 'top', from: { xPt: 6, yPt: 8 }, to: { xPt: 114, yPt: 8 } }),
      expect.objectContaining({ edge: 'left', from: { xPt: 6, yPt: 8 }, to: { xPt: 6, yPt: 24 } }),
      expect.objectContaining({ edge: 'right', from: { xPt: 114, yPt: 8 }, to: { xPt: 114, yPt: 24 } }),
    ]));
    expect(first.borders.some((border) => border.edge === 'bottom')).toBe(false);
    expect(second.inkBounds).toEqual({ xPt: 6, yPt: 5, widthPt: 108, heightPt: 14 });
    expect(second.borders).toEqual(expect.arrayContaining([
      expect.objectContaining({ edge: 'left', from: { xPt: 6, yPt: 5 }, to: { xPt: 6, yPt: 19 } }),
      expect.objectContaining({ edge: 'right', from: { xPt: 114, yPt: 5 }, to: { xPt: 114, yPt: 19 } }),
      expect.objectContaining({ edge: 'bottom', from: { xPt: 6, yPt: 19 }, to: { xPt: 114, yPt: 19 } }),
    ]));
    expect(second.borders.some((border) => border.edge === 'top')).toBe(false);
  });

  it('retains paint-ready shared dash and double treatments for paragraph borders', () => {
    const segment = {
      text: 'AB', sourceRunIndex: 0, measuredWidth: 10,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
      shapedClusters: [
        { range: { start: 0, end: 1 }, offsetPt: 0, advancePt: 5 },
        { range: { start: 1, end: 2 }, offsetPt: 5, advancePt: 5 },
      ],
    } as unknown as LayoutTextSeg;
    const formatted = {
      type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, tabStops: [],
      runs: [{ type: 'text', text: 'AB', fontSize: 10, fontFamily: 'Test Sans' }],
      borders: {
        top: { style: 'dotDash', width: 2, space: 0, color: '123456' },
        right: null,
        bottom: { style: 'double', width: 3, space: 0, color: '654321' },
        left: null,
        between: null,
      },
    } as unknown as DocParagraph;

    const node = projectMeasuredSegment(formatted, segment);

    expect(node.borders.find((border) => border.edge === 'top')).toMatchObject({
      authoredStyle: 'dotDash', style: 'dashed', dashPatternPt: [2, 4, 6, 4],
    });
    expect(node.borders.find((border) => border.edge === 'bottom')).toMatchObject({
      authoredStyle: 'double', style: 'double', dashPatternPt: [],
    });
  });

  it.each([
    ['numbered', { kind: 'text', role: 'numbering-marker', text: '1.' }],
    ['bidi', { kind: 'text', direction: 'rtl', text: '\u0645\u0631\u062d\u0628\u0627' }],
    ['vertical', { kind: 'text', writingMode: 'vertical-rl', text: '\u7e26' }],
    ['tab leader', { kind: 'tab', leader: 'dot', advancePt: 24 }],
    ['page field', { kind: 'text', role: 'field-result', dependency: 'page', text: '3' }],
  ] as const)('retains %s geometry rather than a source run', (_name, placement) => {
    const line = input().lines[0];
    const node = layoutParagraph(input({
      lines: [{ ...line, placements: [{ ...text('x', 0, 10), ...placement } as TextPlacement] }],
      flowBounds: { xPt: 10, yPt: 5, widthPt: 100, heightPt: 25 },
      inkBounds: { xPt: 10, yPt: 10, widthPt: 25, heightPt: 12 },
      spacing: { beforePt: 5, afterPt: 6 },
    }));

    expect(node.lines[0]?.placements[0]).toMatchObject(placement);
    expect(node.lines[0]?.placements[0]).not.toHaveProperty('run');
    expect(() => stableFingerprint('paragraph', node)).not.toThrow();
  });

  it('retains wrap exclusions, contextual spacing, borders, and a hidden paragraph mark', () => {
    const node = layoutParagraph(input({
      contextualSpacing: true,
      spacing: { beforePt: 0, afterPt: 6 },
      exclusions: [{
        id: 'float-1',
        wrap: 'square',
        bounds: { xPt: 60, yPt: 0, widthPt: 40, heightPt: 30 },
        polygon: [
          { xPt: 60, yPt: 0 }, { xPt: 100, yPt: 0 },
          { xPt: 100, yPt: 30 }, { xPt: 60, yPt: 30 },
        ],
      }],
      borders: [{
        from: { xPt: 10, yPt: 34 },
        to: { xPt: 110, yPt: 34 },
        color: '#000000',
        widthPt: 1,
        authoredStyle: 'single',
        style: 'solid',
      }],
      paragraphMark: { hidden: true, bounds: { xPt: 30, yPt: 24, widthPt: 0, heightPt: 10 } },
    }));

    expect(node.contextualSpacing).toBe(true);
    expect(node.exclusions[0]?.bounds).toEqual({ xPt: 60, yPt: 0, widthPt: 40, heightPt: 30 });
    expect(node.borders).toHaveLength(1);
    expect(node.paragraphMark?.hidden).toBe(true);
  });
});

describe('paragraphLayoutFromMeasurement retained authorities', () => {
  it('retains selected-face anchor-host metrics instead of font-size ratios', () => {
    const paragraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
      tabStops: [], runs: [],
    } as unknown as DocParagraph;
    const segment = {
      text: '', metricOnly: true, measuredWidth: 0, fontSize: 10,
      textShapeRequest: {
        text: '', fontSizePt: 10, fonts: { ascii: 'Test Sans' },
        weight: 400, style: 'normal', measure: true,
      },
      textLayoutService: {
        shape: () => ({ advancePt: 0, ascentPt: 7, descentPt: 3, spans: [], diagnostics: [], graphemeBoundaries: [0] }),
      },
    } as unknown as LayoutTextSeg;
    expect(projectMeasuredSegment(paragraph, segment).lines[0]?.placements[0]).toMatchObject({
      kind: 'anchor-host', sourceMetrics: { ascentPt: 7, descentPt: 3 },
    });
  });

  it('matches one scoped host to one anchored payload and retains one drawing and exclusion', () => {
    const occurrenceId = 'anchor:body:body:3:wp-anchor-1';
    const anchored = retainedAnchor(occurrenceId);
    const anchorParagraph = {
      ...paragraph,
      runs: [
        { type: 'anchorHost', fontSize: 10, anchorOccurrenceId: occurrenceId },
        {
          type: 'image', imagePath: 'word/media/anchor.png', mimeType: 'image/png',
          widthPt: 20, heightPt: 10, anchor: true, anchorAcquisitionInput: anchored,
        },
      ],
    } as unknown as DocParagraph;
    const host = {
      text: '', metricOnly: true, sourceRunIndex: 0, measuredWidth: 0,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
    } as unknown as LayoutTextSeg;

    const node = projectMeasuredSegment(anchorParagraph, host, acquisitionContext, undefined, {
      page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
      margin: { xPt: 10, yPt: 20, widthPt: 180, heightPt: 260 },
      column: { xPt: 10, yPt: 20, widthPt: 90, heightPt: 260 },
      pageParity: 'odd',
    });

    expect(node.drawings).toHaveLength(1);
    expect(node.exclusions).toHaveLength(1);
    expect(node.drawings[0]).toMatchObject({
      flowBounds: { xPt: 12, yPt: 13, widthPt: 20, heightPt: 10 },
      commands: [{ kind: 'resource', resourceKind: 'image' }],
      anchorLayer: {
        behindDoc: false, relativeHeight: 7, sourceOrder: 1,
        coordinateSpace: 'acquired-anchor-points',
        normalization: {
          physicalFrames: {
            page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
            margin: { xPt: 10, yPt: 20, widthPt: 180, heightPt: 260 },
            column: { xPt: 10, yPt: 20, widthPt: 90, heightPt: 260 },
          },
          logicalHostFrames: {
            paragraph: { xPt: 10, yPt: 10, widthPt: 100, heightPt: 12 },
            line: { xPt: 10, yPt: 10, widthPt: 0, heightPt: 12 },
            character: { xPt: 10, yPt: 10, widthPt: 0, heightPt: 12 },
          },
        },
      },
    });
    expect(node.exclusions[0]).toMatchObject({
      wrap: 'square', bounds: { xPt: 12, yPt: 13, widthPt: 20, heightPt: 10 },
    });
  });

  it('keeps public fallback anchors acquired until destination-page normalization', () => {
    const publicParagraph = {
      ...paragraph,
      runs: [{
        type: 'shape', widthPt: 20, heightPt: 10,
        anchorXPt: 4, anchorYPt: 5, anchorXFromMargin: false, anchorYFromPara: true,
        anchorXRelativeFrom: 'page', anchorYRelativeFrom: 'paragraph',
        anchorXAlign: 'center', pctPosV: 0.25,
        zOrder: 9, presetGeometry: 'rect', subpaths: [], fill: null, stroke: null,
      }, {
        type: 'image', imagePath: 'word/media/public.png', mimeType: 'image/png',
        widthPt: 12, heightPt: 8, anchor: true,
        anchorXPt: 3, anchorYPt: 7, anchorXFromMargin: true, anchorYFromPara: false,
        anchorXRelativeFrom: 'margin', anchorYRelativeFrom: 'page',
      }],
    } as unknown as DocParagraph;
    const line = {
      text: '', metricOnly: true, sourceRunIndex: 0, measuredWidth: 0,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
    } as unknown as LayoutTextSeg;
    const node = projectMeasuredSegment(publicParagraph, line, acquisitionContext, undefined, {
      page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
      margin: { xPt: 10, yPt: 20, widthPt: 180, heightPt: 260 },
      column: { xPt: 10, yPt: 20, widthPt: 90, heightPt: 260 },
      pageParity: 'odd',
    });

    expect(node.drawings).toHaveLength(2);
    expect(node.drawings[0]?.anchorLayer).toMatchObject({
      coordinateSpace: 'acquired-anchor-points',
      normalization: {
        acquisition: {
          horizontal: { relativeFrom: 'page', choice: { kind: 'align', value: 'center' } },
          vertical: { relativeFrom: 'paragraph', choice: { kind: 'percent', fraction: 0.25 } },
        },
      },
    });
    expect(node.drawings[1]?.anchorLayer).toMatchObject({
      coordinateSpace: 'acquired-anchor-points',
      normalization: {
        acquisition: {
          horizontal: { relativeFrom: 'margin', choice: { kind: 'offset', valuePt: 3 } },
          vertical: { relativeFrom: 'page', choice: { kind: 'offset', valuePt: 7 } },
        },
      },
    });
  });

  it.each([
    ['page/page', 'page', 'page', 12, 13],
    ['page/host', 'page', 'line', 12, 23],
    ['host/page', 'character', 'page', 160, 13],
    ['host/host', 'character', 'line', 160, 23],
  ] as const)(
    'normalizes projected vertical %s anchor acquisition into physical page points',
    (_name, horizontalFrame, verticalFrame, expectedX, expectedY) => {
      const occurrenceId = `anchor:${horizontalFrame}:${verticalFrame}`;
      const base = retainedAnchor(occurrenceId);
      const anchored = retainedAnchor(occurrenceId, {
        horizontal: horizontalFrame === 'page'
          ? { relativeFrom: 'page', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: 12 } }
          : { relativeFrom: 'character', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: 2 } },
        vertical: verticalFrame === 'page'
          ? { relativeFrom: 'page', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: 13 } }
          : { relativeFrom: 'line', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: 3 } },
        behavior: base.behavior,
      });
      const anchorParagraph = {
        ...paragraph,
        runs: [
          { type: 'anchorHost', fontSize: 10, anchorOccurrenceId: occurrenceId },
          {
            type: 'image', imagePath: 'word/media/anchor.png', mimeType: 'image/png',
            widthPt: 20, heightPt: 10, anchor: true, anchorAcquisitionInput: anchored,
          },
        ],
      } as unknown as DocParagraph;
      const host = {
        text: '', metricOnly: true, sourceRunIndex: 0, measuredWidth: 0,
        fontSize: 10, fontFamily: 'Test Sans', fontRoute,
      } as unknown as LayoutTextSeg;
      const acquired = projectMeasuredSegment(anchorParagraph, host, acquisitionContext, undefined, {
        page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
        margin: { xPt: 10, yPt: 20, widthPt: 180, heightPt: 260 },
        column: { xPt: 10, yPt: 20, widthPt: 90, heightPt: 260 },
        pageParity: 'odd',
      });
      const acquiredDrawing = acquired.drawings[0]!;
      const attachedTextBox: TextBoxLayout = {
        kind: 'textbox', id: `textbox:${occurrenceId}`, source: acquiredDrawing.source,
        flowDomainId: acquired.flowDomainId, ordinaryFlow: false,
        flowBounds: acquiredDrawing.flowBounds, inkBounds: acquiredDrawing.inkBounds,
        advancePt: 0, paragraphs: [], writingMode: 'horizontal-tb',
        insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
      };
      const acquiredWithTextBox: ParagraphLayout = {
        ...acquired,
        drawings: [{ ...acquiredDrawing, textBoxIds: [attachedTextBox.id] }],
        textBoxes: [attachedTextBox],
      };
      const projected = projectBodyOccurrence(acquiredWithTextBox, {
        occurrenceId: `destination:${occurrenceId}`,
        destination: {
          coordinateSpace: 'logical-body-points',
          flowDomainId: bodyFlowDomainId(1, 'vertical', 0),
          translation: { xPt: 10, yPt: 20 },
        },
      });
      const normalized = normalizePagePaintNodeAnchors(projected, {
        currentToPage: canonicalLogicalToPhysical('vertical-rl', 200),
        normalizedFor: {
          physicalPageIndex: 1,
          flowDomainId: bodyFlowDomainId(1, 'vertical', 0),
          regionId: 'vertical',
        },
        destinationFrames: {
          page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
          margin: { xPt: 10, yPt: 20, widthPt: 180, heightPt: 260 },
          column: { xPt: 100, yPt: 0, widthPt: 100, heightPt: 300 },
          pageParity: 'even',
        },
      });
      const drawing = normalized.kind === 'paragraph' ? normalized.drawings[0] : undefined;

      expect(drawing?.anchorLayer?.coordinateSpace).toBe('physical-page-points');
      expect(drawing?.flowBounds).toEqual({
        xPt: expectedX, yPt: expectedY, widthPt: 20, heightPt: 10,
      });
      expect(drawing?.commands[0]).toMatchObject({
        rect: { xPt: expectedX, yPt: expectedY, widthPt: 20, heightPt: 10 },
      });
      if (normalized.kind !== 'paragraph') throw new Error('expected paragraph');
      expect(normalized.textBoxes[0]?.flowBounds).toEqual({
        xPt: expectedX, yPt: expectedY, widthPt: 20, heightPt: 10,
      });
      expect(normalized.exclusions[0]?.bounds).toEqual({
        xPt: expectedX, yPt: expectedY, widthPt: 20, heightPt: 10,
      });
      expect(normalized.exclusions[0]?.polygon).toEqual([
        { xPt: expectedX, yPt: expectedY },
        { xPt: expectedX + 20, yPt: expectedY },
        { xPt: expectedX + 20, yPt: expectedY + 10 },
        { xPt: expectedX, yPt: expectedY + 10 },
      ]);
      expect(normalized.anchorFrames?.[0]).toMatchObject({
        status: 'resolved',
        geometry: { objectFrame: { xPt: expectedX, yPt: expectedY, widthPt: 20, heightPt: 10 } },
      });

      const section = {
        geometry: {
          pageWidth: 200, pageHeight: 300,
          marginTop: 20, marginRight: 10, marginBottom: 20, marginLeft: 10,
          headerDistance: 10, footerDistance: 10,
        },
        columns: [{ xPt: 0, wPt: 300 }],
        grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
        textDirection: 'tbRl', verticalAlignment: 'top',
      } as SectionLayoutContext;
      const page = createLayoutPage({
        pageIndex: 1,
        physicalPage: { widthPt: 200, heightPt: 300, contentTopPt: 0, contentBottomPt: 300 },
        sectionOccurrenceId: 'section:vertical', section,
        sectionRegions: [{
          id: 'vertical', sectionOccurrenceId: 'section:vertical', section,
          writingMode: 'vertical-rl', blockStartPt: 0, blockEndPt: 200,
          columns: [{ inlineStartPt: 0, inlineExtentPt: 300 }],
        }],
        paint: [{
          layer: 'body', node: projected, coordinateSpace: 'logical-body-points',
          logicalBlock: {
            blockStartPt: projected.flowBounds.yPt,
            blockExtentPt: projected.flowBounds.heightPt,
          },
        }],
        readingOrder: [projected],
        pageNumber: { displayNumber: 2, format: 'decimal', sectionOccurrenceId: 'section:vertical' },
      });
      const admitted = page.layers.body[0];
      const admittedDrawing = admitted?.kind === 'paragraph' ? admitted.drawings[0] : undefined;
      expect(admittedDrawing?.anchorLayer?.coordinateSpace).toBe('physical-page-points');
      expect(admittedDrawing?.flowBounds).toEqual(drawing?.flowBounds);
    },
  );

  it.each([
    ['missing behindDoc', { behindDoc: null, behindDocStatus: 'missing' }, 'behavior.behindDoc'],
    ['invalid behindDoc', { behindDoc: null, behindDocStatus: 'invalid' }, 'behavior.behindDoc'],
    ['missing relativeHeight', { relativeHeight: null, relativeHeightStatus: 'missing' }, 'behavior.relativeHeight'],
    ['invalid relativeHeight', { relativeHeight: null, relativeHeightStatus: 'invalid' }, 'behavior.relativeHeight'],
    ['missing locked', { locked: null, lockedStatus: 'missing' }, 'behavior.locked'],
    ['invalid locked', { locked: null, lockedStatus: 'invalid' }, 'behavior.locked'],
    ['missing layoutInCell', { layoutInCell: null, layoutInCellStatus: 'missing' }, 'behavior.layoutInCell'],
    ['invalid layoutInCell', { layoutInCell: null, layoutInCellStatus: 'invalid' }, 'behavior.layoutInCell'],
    ['missing allowOverlap', { allowOverlap: null, allowOverlapStatus: 'missing' }, 'behavior.allowOverlap'],
    ['invalid allowOverlap', { allowOverlap: null, allowOverlapStatus: 'invalid' }, 'behavior.allowOverlap'],
  ] as const)('retains an explicit diagnostic when %s prevents drawing acquisition', (
    _name, behaviorOverride, path,
  ) => {
    const occurrenceId = 'anchor:body:body:3:invalid-required-behavior';
    const base = retainedAnchor(occurrenceId);
    const anchored = retainedAnchor(occurrenceId, {
      behavior: { ...base.behavior, ...behaviorOverride },
    });
    const anchorParagraph = {
      ...paragraph,
      runs: [
        { type: 'anchorHost', fontSize: 10, anchorOccurrenceId: occurrenceId },
        {
          type: 'image', imagePath: 'word/media/anchor.png', mimeType: 'image/png',
          widthPt: 20, heightPt: 10, anchor: true, anchorAcquisitionInput: anchored,
        },
      ],
    } as unknown as DocParagraph;
    const host = {
      text: '', metricOnly: true, sourceRunIndex: 0, measuredWidth: 0,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
    } as unknown as LayoutTextSeg;

    const node = projectMeasuredSegment(anchorParagraph, host, acquisitionContext, undefined, {
      page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
      margin: { xPt: 10, yPt: 20, widthPt: 180, heightPt: 260 },
      column: { xPt: 10, yPt: 20, widthPt: 90, heightPt: 260 },
      pageParity: 'odd',
    });

    expect(node.drawings).toEqual([]);
    expect(node.anchorFrames).toEqual([
      expect.objectContaining({
        status: 'unsupported',
        issues: [expect.objectContaining({ path })],
      }),
    ]);
  });

  it('materializes one outer anchor while retaining each grouped child frame', () => {
    const occurrenceId = 'anchor:body:body:3:wp-anchor-group';
    const group = {
      childSourceId: 'child-0', sourceIndex: 0, sourceCount: 2,
      transformChain: [], childTransform: null,
      resolvedChildFrame: {
        offsetXPt: 2, offsetYPt: 1, widthPt: 5, heightPt: 2,
        rotationDeg: 15, flipH: true, flipV: false,
      },
    } as const;
    const relativeSize = {
      horizontal: {
        relativeFrom: 'page', relativeFromStatus: 'valid',
        fraction: 0.2, fractionStatus: 'valid',
      },
      vertical: {
        relativeFrom: 'page', relativeFromStatus: 'valid',
        fraction: 0.1, fractionStatus: 'valid',
      },
    } as const;
    const first = retainedAnchor(occurrenceId, { group, relativeSize });
    const second = retainedAnchor(occurrenceId, {
      group: {
        ...group,
        childSourceId: 'child-1', sourceIndex: 1,
        resolvedChildFrame: {
          offsetXPt: 10, offsetYPt: 4, widthPt: 4, heightPt: 3,
          rotationDeg: 75, flipH: false, flipV: true,
        },
      },
      relativeSize,
    });
    const groupedParagraph = {
      ...paragraph,
      runs: [
        { type: 'anchorHost', fontSize: 10, anchorOccurrenceId: occurrenceId },
        { type: 'shape', widthPt: 5, heightPt: 2, anchorXPt: 2, anchorYPt: 1, anchorXFromMargin: false, anchorYFromPara: false, anchorAcquisitionInput: first, presetGeometry: 'rect', subpaths: [], fill: null, stroke: null, rotation: 15, flipH: true, flipV: false },
        { type: 'image', imagePath: 'word/media/group.png', mimeType: 'image/png', widthPt: 4, heightPt: 3, anchor: true, anchorAcquisitionInput: second, rotation: 75, flipH: false, flipV: true },
      ],
    } as unknown as DocParagraph;
    const host = {
      text: '', metricOnly: true, sourceRunIndex: 0, measuredWidth: 0,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
    } as unknown as LayoutTextSeg;
    const node = projectMeasuredSegment(groupedParagraph, host, acquisitionContext, {
      text: { shape: () => ({ advancePt: 0, ascentPt: 0, descentPt: 0, spans: [], diagnostics: [], graphemeBoundaries: [0] }) },
    }, {
      page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
      margin: { xPt: 10, yPt: 20, widthPt: 180, heightPt: 260 },
      column: { xPt: 10, yPt: 20, widthPt: 90, heightPt: 260 },
      pageParity: 'odd',
    });

    expect(node.drawings).toHaveLength(1);
    expect(node.exclusions).toHaveLength(1);
    expect(node.drawings[0]).toMatchObject({
      flowBounds: { xPt: 12, yPt: 13, widthPt: 40, heightPt: 30 },
      commands: [
        {
          kind: 'drawingml-shape',
          plan: {
            rect: { x: 16, y: 16, w: 10, h: 6 },
            transform: { rotationDeg: 15, flipH: true, flipV: false },
          },
        },
        {
          kind: 'resource', resourceKind: 'image',
          rect: { xPt: 32, yPt: 25, widthPt: 8, heightPt: 9 },
        },
      ],
    });
    expect(node.exclusions).toHaveLength(1);
    expect(node.exclusions[0]).toMatchObject({
      bounds: { xPt: 12, yPt: 13, widthPt: 40, heightPt: 30 },
    });
  });

  it('rebuilds a complete relative-size anchor transaction for destination extents', () => {
    const occurrenceId = 'anchor:relative-size-transaction';
    const validEdges = {
      topPt: 1, topStatus: 'valid', rightPt: 2, rightStatus: 'valid',
      bottomPt: 3, bottomStatus: 'valid', leftPt: 4, leftStatus: 'valid',
    } as const;
    const relativeSize = {
      horizontal: {
        relativeFrom: 'page', relativeFromStatus: 'valid',
        fraction: 0.2, fractionStatus: 'valid',
      },
      vertical: {
        relativeFrom: 'page', relativeFromStatus: 'valid',
        fraction: 0.1, fractionStatus: 'valid',
      },
    } as const;
    const group = {
      childSourceId: 'child-0', sourceIndex: 0, sourceCount: 2,
      transformChain: [], childTransform: null,
      resolvedChildFrame: {
        offsetXPt: 2, offsetYPt: 1, widthPt: 5, heightPt: 2,
        rotationDeg: 15, flipH: true, flipV: false,
      },
    } as const;
    const wrap = {
      kind: 'tight', authoredKinds: ['wrapTight'], side: 'bothSides',
      distances: validEdges, effectExtent: validEdges,
      polygon: {
        edited: true,
        coordinateSpace: { width: 21600, height: 21600 } as const,
        points: [
          { x: 0, y: 0, rawX: '0', rawY: '0' },
          { x: 21600, y: 0, rawX: '21600', rawY: '0' },
          { x: 10800, y: 21600, rawX: '10800', rawY: '21600' },
        ],
        invalidPointCount: 0,
      },
    } as const;
    const first = retainedAnchor(occurrenceId, {
      group, relativeSize, parentEffectExtent: validEdges, wrap,
    });
    const second = retainedAnchor(occurrenceId, {
      group: {
        ...group,
        childSourceId: 'child-1', sourceIndex: 1,
        resolvedChildFrame: {
          offsetXPt: 10, offsetYPt: 4, widthPt: 4, heightPt: 3,
          rotationDeg: 75, flipH: false, flipV: true,
        },
      },
      relativeSize,
    });
    const groupedParagraph = {
      ...paragraph,
      runs: [
        { type: 'anchorHost', fontSize: 10, anchorOccurrenceId: occurrenceId },
        { type: 'shape', widthPt: 5, heightPt: 2, anchorXPt: 2, anchorYPt: 1, anchorXFromMargin: false, anchorYFromPara: false, anchorAcquisitionInput: first, presetGeometry: 'rect', subpaths: [], fill: null, stroke: null, rotation: 15, flipH: true, flipV: false },
        { type: 'image', imagePath: 'word/media/group.png', mimeType: 'image/png', widthPt: 4, heightPt: 3, anchor: true, anchorAcquisitionInput: second, rotation: 75, flipH: false, flipV: true },
      ],
    } as unknown as DocParagraph;
    const host = {
      text: '', metricOnly: true, sourceRunIndex: 0, measuredWidth: 0,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
    } as unknown as LayoutTextSeg;
    const acquired = projectMeasuredSegment(groupedParagraph, host, acquisitionContext, {
      text: { shape: () => ({ advancePt: 0, ascentPt: 0, descentPt: 0, spans: [], diagnostics: [], graphemeBoundaries: [0] }) },
    }, {
      page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
      margin: { xPt: 10, yPt: 20, widthPt: 180, heightPt: 260 },
      column: { xPt: 10, yPt: 20, widthPt: 90, heightPt: 260 },
      pageParity: 'odd',
    });
    const sourceDrawing = acquired.drawings[0]!;
    let measuredFont = '10px sans-serif';
    const measureContext = {
      get font() { return measuredFont; },
      set font(value: string) { measuredFont = value; },
      measureText(value: string) {
        const size = Number.parseFloat(/([\d.]+)px/u.exec(measuredFont)?.[1] ?? '10');
        return {
          width: [...value].length * size / 2,
          fontBoundingBoxAscent: size * .8, fontBoundingBoxDescent: size * .2,
          actualBoundingBoxAscent: size * .8, actualBoundingBoxDescent: size * .2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const textServices = createLayoutServices({
      section: {
        pageWidth: 200, pageHeight: 300,
        marginTop: 20, marginRight: 10, marginBottom: 20, marginLeft: 10,
        headerDistance: 10, footerDistance: 10,
        titlePage: false, evenAndOddHeaders: false,
      },
      body: [], headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as never, { measureContext });
    const acquiredTextBox = acquireShapeTextBoxLayout({
      textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
      textAnchor: 't', textBlocks: [{
        text: 'AB', fontSizePt: 10, color: '112233', alignment: 'left',
        runs: [{ text: 'AB', fontSizePt: 10, color: '112233' }],
      }],
    } as never, sourceDrawing.flowBounds, {
      id: 'relative-size-textbox', source: sourceDrawing.source,
      flowDomainId: acquired.flowDomainId, context: acquisitionContext,
      measurer: { context: measureContext, fontFamilyClasses: {} },
      environment: {
        pageIndex: 0, totalPages: 1, documentHasEastAsianText: false,
        layoutServices: textServices,
      },
    });
    if (!acquiredTextBox) throw new Error('expected acquired non-empty textbox');
    const acquiredTextParagraph = acquiredTextBox.paragraphs[0]!;
    const acquiredTextLine = acquiredTextParagraph.lines[0]!;
    const acquiredText = acquiredTextLine.placements.find((placement) => placement.kind === 'text');
    if (!acquiredText || acquiredText.kind !== 'text') throw new Error('expected retained text');
    const retainedText: TextPlacement = {
      ...acquiredText,
      advancePt: 7,
      clusters: acquiredText.clusters.map((cluster, index) => ({
        ...cluster,
        offset: { xPt: index + 2, yPt: index + 1 },
        advancePt: index + 6,
      })),
      paintOps: acquiredText.paintOps.map((operation, index) => ({
        ...operation,
        offset: { xPt: index + 3, yPt: index + 2 },
        letterSpacingPt: index + 1,
        scaleX: 0.8,
      })),
      characterSpacingPt: 1,
      fitText: { regionIndex: 0, perGapPt: 2, trailingPadPt: 1 },
      positionPt: 2,
      ownedTrailingSlackPt: 1,
      ruby: {
        text: 'ル', advancePt: 4, authored: { raisePt: 2 },
        paintOps: [{
          text: 'ル', origin: {
            xPt: acquiredText.origin.xPt + 1,
            yPt: acquiredText.origin.yPt - 2,
          },
          fontRoute: acquiredText.fontRoute, fontSizePt: 6,
          fontWeight: 400, fontStyle: 'normal', color: acquiredText.color,
          inkBounds: { xMinPt: 0, xMaxPt: 4, ascentPt: 5, descentPt: 1 },
        }],
      },
      emphasis: {
        authored: 'dot',
        glyphs: [{
          text: '•', origin: {
            xPt: acquiredText.origin.xPt + 2,
            yPt: acquiredText.origin.yPt - 3,
          },
          fontRoute: acquiredText.fontRoute, fontSizePt: 5,
          fontWeight: 400, fontStyle: 'normal', color: acquiredText.color,
        }],
        paths: [{
          kind: 'polyline' as const,
          points: [
            { xPt: acquiredText.origin.xPt + 1, yPt: acquiredText.origin.yPt - 1 },
            { xPt: acquiredText.origin.xPt + 3, yPt: acquiredText.origin.yPt - 1 },
          ],
          fill: '#000000', stroke: null, strokeWidthPt: 0.5,
        }],
      },
    };
    const tabPlacement: TabPlacement = {
      kind: 'tab' as const, range: { start: 2, end: 3 },
      bounds: { xPt: 24, yPt: 16, widthPt: 4, heightPt: 8 },
      advancePt: 4, leader: 'dot' as const,
      leaderGlyphs: [{
        text: '.', origin: { xPt: 25, yPt: 22 },
        fontRoute: acquiredText.fontRoute, fontSizePt: 10,
        fontWeight: 400, fontStyle: 'normal' as const, color: acquiredText.color,
      }],
    };
    const resourcePlacement: ResourcePlacement = {
      kind: 'resource' as const, range: { start: 3, end: 4 },
      resourceKey: 'image:inline', resourceKind: 'image' as const,
      bounds: { xPt: 28, yPt: 16, widthPt: 5, heightPt: 8 }, advancePt: 5,
    };
    const inlineDrawingPlacement: DrawingPlacement = {
      kind: 'drawing' as const, range: { start: 4, end: 5 },
      drawingId: 'inline:drawing',
      bounds: { xPt: 33, yPt: 16, widthPt: 6, heightPt: 8 }, advancePt: 6,
    };
    const textBoxParagraph: ParagraphLayout = {
      ...acquiredTextParagraph,
      lines: [{
        ...acquiredTextLine,
        placements: [retainedText, tabPlacement, resourcePlacement, inlineDrawingPlacement],
      }],
    };
    const textBox: TextBoxLayout = {
      ...acquiredTextBox,
      clipBounds: { xPt: 14, yPt: 15, widthPt: 20, heightPt: 12 },
      paragraphs: [textBoxParagraph],
    };
    const acquiredTransaction: ParagraphLayout = {
      ...acquired,
      drawings: [{
        ...sourceDrawing,
        clipBounds: { xPt: 13, yPt: 14, widthPt: 30, heightPt: 20 },
        clip: {
          kind: 'polygon',
          points: [{ xPt: 12, yPt: 13 }, { xPt: 52, yPt: 43 }],
        },
        textBoxIds: [textBox.id],
      }],
      textBoxes: [textBox],
    };
    const destinationFrames = {
      page: { xPt: 0, yPt: 0, widthPt: 300, heightPt: 400 },
      margin: { xPt: 10, yPt: 20, widthPt: 280, heightPt: 360 },
      column: { xPt: 10, yPt: 20, widthPt: 140, heightPt: 360 },
      paragraph: { xPt: 10, yPt: 10, widthPt: 100, heightPt: 12 },
      line: { xPt: 10, yPt: 10, widthPt: 0, heightPt: 12 },
      character: { xPt: 10, yPt: 10, widthPt: 0, heightPt: 12 },
      pageParity: 'even' as const,
    };
    const expected = resolveAnchorFrame({ acquisition: first, frames: destinationFrames });
    if (expected.status !== 'resolved') throw new Error('expected resolved destination anchor');

    const normalized = normalizePagePaintNodeAnchors(acquiredTransaction, {
      currentToPage: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      normalizedFor: { physicalPageIndex: 1, flowDomainId: 'body', regionId: 'body' },
      destinationFrames,
    });
    if (normalized.kind !== 'paragraph') throw new Error('expected paragraph');
    const drawing = normalized.drawings[0]!;

    expect(drawing.flowBounds).toEqual(expected.geometry.objectFrame);
    expect(drawing.inkBounds).toEqual(expected.geometry.inkBounds);
    expect(normalized.lines.flatMap((line) => line.placements).find((placement) => (
      placement.kind === 'drawing' && placement.drawingId === drawing.id
    ))).toMatchObject({ bounds: expected.geometry.inkBounds });
    expect(drawing.commands).toMatchObject([
      { kind: 'drawingml-shape', plan: { rect: { x: 18, y: 17, w: 15, h: 8 } } },
      { kind: 'resource', rect: { xPt: 42, yPt: 29, widthPt: 12, heightPt: 12 } },
    ]);
    expect(drawing.clipBounds).toEqual({ xPt: 13.5, yPt: 14.333333333333334, widthPt: 45, heightPt: 26.666666666666664 });
    expect(drawing.clip).toEqual({
      kind: 'polygon', points: [{ xPt: 12, yPt: 13 }, { xPt: 72, yPt: 53 }],
    });
    const scaleX = expected.geometry.objectFrame.widthPt / sourceDrawing.flowBounds.widthPt;
    const scaleY = expected.geometry.objectFrame.heightPt / sourceDrawing.flowBounds.heightPt;
    const expectedPoint = (point: { xPt: number; yPt: number }) => ({
      xPt: expected.geometry.objectFrame.xPt
        + (point.xPt - sourceDrawing.flowBounds.xPt) * scaleX,
      yPt: expected.geometry.objectFrame.yPt
        + (point.yPt - sourceDrawing.flowBounds.yPt) * scaleY,
    });
    const expectedRect = (rect: { xPt: number; yPt: number; widthPt: number; heightPt: number }) => ({
      ...expectedPoint(rect), widthPt: rect.widthPt * scaleX, heightPt: rect.heightPt * scaleY,
    });
    expect(normalized.textBoxes[0]).toMatchObject({
      flowBounds: expected.geometry.objectFrame,
      inkBounds: expected.geometry.objectFrame,
      clipBounds: { xPt: 15, yPt: 15.666666666666666, widthPt: 30, heightPt: 16 },
      contentBounds: expectedRect(acquiredTextBox.contentBounds!),
      paragraphs: [{
        flowBounds: expectedRect(acquiredTextParagraph.flowBounds),
        inkBounds: expectedRect(acquiredTextParagraph.inkBounds),
      }],
    });
    const normalizedLine = normalized.textBoxes[0]!.paragraphs[0]!.lines[0]!;
    const normalizedText = normalizedLine.placements[0];
    if (normalizedText?.kind !== 'text') throw new Error('expected normalized text');
    expect(normalizedText).toMatchObject({
      range: retainedText.range,
      origin: expectedPoint(retainedText.origin),
      bounds: expectedRect(retainedText.bounds),
      advancePt: retainedText.advancePt * scaleX,
      clusters: retainedText.clusters.map((cluster) => ({
        range: cluster.range,
        offset: { xPt: cluster.offset.xPt * scaleX, yPt: cluster.offset.yPt * scaleY },
        advancePt: cluster.advancePt * scaleX,
      })),
      paintOps: retainedText.paintOps.map((operation) => ({
        range: operation.range,
        offset: { xPt: operation.offset.xPt * scaleX, yPt: operation.offset.yPt * scaleY },
        letterSpacingPt: operation.letterSpacingPt * scaleX,
        scaleX: operation.scaleX,
      })),
      characterSpacingPt: retainedText.characterSpacingPt! * scaleX,
      fitText: {
        ...retainedText.fitText,
        perGapPt: retainedText.fitText!.perGapPt * scaleX,
        trailingPadPt: retainedText.fitText!.trailingPadPt * scaleX,
      },
      positionPt: retainedText.positionPt! * scaleY,
      ownedTrailingSlackPt: retainedText.ownedTrailingSlackPt! * scaleX,
      fontSizePt: retainedText.fontSizePt,
      ruby: {
        advancePt: retainedText.ruby!.advancePt * scaleX,
        paintOps: [{
          origin: expectedPoint(retainedText.ruby!.paintOps[0]!.origin),
          fontSizePt: retainedText.ruby!.paintOps[0]!.fontSizePt,
          inkBounds: retainedText.ruby!.paintOps[0]!.inkBounds,
        }],
      },
      emphasis: {
        glyphs: [{
          origin: expectedPoint(retainedText.emphasis!.glyphs![0]!.origin),
          fontSizePt: retainedText.emphasis!.glyphs![0]!.fontSizePt,
        }],
        paths: [{
          points: retainedText.emphasis!.paths![0]!.points.map(expectedPoint),
          strokeWidthPt: retainedText.emphasis!.paths![0]!.strokeWidthPt,
        }],
      },
    });
    const [normalizedTab, normalizedResource, normalizedInlineDrawing] = normalizedLine.placements.slice(1);
    expect(normalizedTab).toMatchObject({
      kind: 'tab', bounds: expectedRect(tabPlacement.bounds!),
      advancePt: tabPlacement.advancePt * scaleX,
      leaderGlyphs: [{
        origin: expectedPoint(tabPlacement.leaderGlyphs![0]!.origin),
        fontSizePt: tabPlacement.leaderGlyphs![0]!.fontSizePt,
      }],
    });
    expect(normalizedResource).toMatchObject({
      kind: 'resource', bounds: expectedRect(resourcePlacement.bounds),
      advancePt: resourcePlacement.advancePt * scaleX,
    });
    expect(normalizedInlineDrawing).toMatchObject({
      kind: 'drawing', bounds: expectedRect(inlineDrawingPlacement.bounds),
      advancePt: inlineDrawingPlacement.advancePt * scaleX,
    });
    expect(normalized.exclusions[0]).toEqual(expect.objectContaining({
      bounds: expected.geometry.wrapBounds,
      polygon: expected.geometry.wrap.polygon?.points,
    }));
    expect(normalized.anchorFrames?.[0]).toEqual(expected);
  });

  const paragraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
    tabStops: [],
    runs: [{
      type: 'text', text: 'AB', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Test Sans',
      background: null, vertAlign: null,
    }],
  } as unknown as DocParagraph;

  const measureContext = {
    font: '', letterSpacing: '0px', fontKerning: 'auto',
    measureText(text: string) {
      return {
        width: [...text].length * 5,
        actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
        fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;

  it('acquires ordinary CJK as complete service-shaped grapheme clusters', () => {
    const cjkParagraph = {
      ...paragraph,
      runs: [{ ...paragraph.runs[0], text: '国語' }],
    } as DocParagraph;
    const services = createLayoutServices({
      section: {
        pageWidth: 612, pageHeight: 792,
        marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
        headerDistance: 36, footerDistance: 36,
        titlePage: false, evenAndOddHeaders: false,
      },
      body: [],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    }, { measureContext });
    const measured = measureParagraph(
      cjkParagraph,
      acquisitionContext,
      {
        startYPt: 0, paragraphXPt: 0, availableWidthPt: 100,
        maximumYPt: 500, suppressSpaceBefore: false,
      },
      { context: measureContext, fontFamilyClasses: {} },
      {
        pageIndex: 0, totalPages: 1, documentHasEastAsianText: true,
        layoutServices: services,
      },
    );
    const node = paragraphLayoutFromMeasurement(cjkParagraph as never, {
      id: 'ordinary-cjk', source, flowDomainId: 'body', ordinaryFlow: true,
      context: acquisitionContext, placement: measured.placement,
      measurer: { context: measureContext, fontFamilyClasses: {} },
      environment: {
        pageIndex: 0, totalPages: 1, documentHasEastAsianText: true,
        layoutServices: services,
      },
      exclusions: [],
    }, measured);
    const placement = node.lines[0]?.placements.find((candidate) => candidate.kind === 'text');

    expect(placement?.kind === 'text' ? placement.clusters : []).toEqual([
      expect.objectContaining({ range: { start: 0, end: 1 } }),
      expect.objectContaining({ range: { start: 1, end: 2 } }),
    ]);
  });

  it('rejects a visible segment that bypassed authoritative cluster shaping', () => {
    const measured = measureParagraph(
      paragraph,
      acquisitionContext,
      {
        startYPt: 0, paragraphXPt: 0, availableWidthPt: 100,
        maximumYPt: 500, suppressSpaceBefore: false,
      },
      { context: measureContext, fontFamilyClasses: {} },
      { pageIndex: 0, totalPages: 1, documentHasEastAsianText: false },
    );

    expect(() => paragraphLayoutFromMeasurement(paragraph as never, {
      id: 'missing-clusters', source, flowDomainId: 'body', ordinaryFlow: true,
      context: acquisitionContext, placement: measured.placement,
      measurer: { context: measureContext, fontFamilyClasses: {} },
      environment: { pageIndex: 0, totalPages: 1, documentHasEastAsianText: false },
      exclusions: [],
    }, measured)).toThrow(/authoritative grapheme clusters/i);
  });

  it.each([
    ['docGrid pitch', { snapToCharacterGrid: true }, 2],
    ['fitText pitch', {
      snapToCharacterGrid: true, fitTextRegionIndex: 0,
      fitTextPerGapPx: 3, fitTextTrailingPadPx: 1,
    }, 3],
  ] as const)('retains the measured %s in both clusters and paint operations', (_name, overrides, pitchPt) => {
    const segment = {
      text: 'AB', sourceRunIndex: 0, measuredWidth: pitchPt === 2 ? 14 : 17,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
      shapedClusters: [
        { range: { start: 0, end: 1 }, offsetPt: 0, advancePt: 5 },
        { range: { start: 1, end: 2 }, offsetPt: 5, advancePt: 5 },
      ],
      ...overrides,
    } as unknown as LayoutTextSeg;
    const node = projectMeasuredSegment(paragraph, segment, {
      ...acquisitionContext,
      characterGrid: { active: true, deltaPt: 2 },
    });
    const placement = node.lines[0]?.placements[0];

    expect(placement).toMatchObject({
      kind: 'text',
      paintOps: [expect.objectContaining({ text: 'AB', letterSpacingPt: pitchPt })],
    });
    expect(placement?.kind === 'text' ? placement.clusters : []).toEqual([
      expect.objectContaining({ offset: { xPt: 0, yPt: 0 }, advancePt: 5 + pitchPt }),
      expect.objectContaining({
        offset: { xPt: 5 + pitchPt, yPt: 0 },
        advancePt: 5 + pitchPt + (pitchPt === 3 ? 1 : 0),
      }),
    ]);
  });

  it('resolves display math as m:jc, then m:defJc, then centerGroup', () => {
    const mathParagraph = {
      ...paragraph,
      runs: [{ type: 'math', nodes: [], display: true, fontSize: 10, resourceKey: 'math:display' }],
    } as unknown as DocParagraph;
    const math = {
      mathNodes: [], mathResourceKey: 'math:display', display: true,
      fallbackText: 'x', measuredWidth: 10, mathAscent: 8, mathDescent: 2,
      fontSize: 10,
    } as unknown as LayoutMathSeg;

    const documentDefault = projectMeasuredSegment(mathParagraph, math, {
      ...acquisitionContext, mathDefJc: 'right',
    });
    const perRun = projectMeasuredSegment(mathParagraph, { ...math, jc: 'left' }, {
      ...acquisitionContext, mathDefJc: 'right',
    });
    const specificationDefault = projectMeasuredSegment(mathParagraph, math);

    expect(documentDefault.lines[0]?.placements[0]).toMatchObject({ bounds: { xPt: 100 } });
    expect(perRun.lines[0]?.placements[0]).toMatchObject({ bounds: { xPt: 10 } });
    expect(specificationDefault.lines[0]?.placements[0]).toMatchObject({ bounds: { xPt: 55 } });
  });

  it('retains a shaped numbering marker and uses its suffix-resolved body offset', () => {
    const numbered = {
      ...paragraph,
      indentLeft: 12, indentFirst: -12,
      numbering: {
        numId: 1, level: 0, format: 'decimal', text: '1.',
        indentLeft: 12, tab: 12, suff: 'tab', jc: 'left',
      },
      numberingMarkerShapeInput: {
        fontSizePt: 10, fonts: { ascii: 'Test Sans' }, weight: 400,
        style: 'normal', complexScript: false,
      },
    } as unknown as DocParagraph & { numberingMarkerShapeInput: object };
    const markerServices = {
      text: {
        shape(request: { text: string }) {
          const advancePt = request.text.length * 5;
          return {
            text: request.text,
            spans: [{
              text: request.text, start: 0, end: request.text.length,
              script: 'ascii', breakBefore: true,
              font: {
                requestedFamily: 'Test Sans', resolvedFamily: 'Test Sans', route: fontRoute,
                source: 'native', weight: 400, style: 'normal', diagnostics: [], genericFamily: 'sans-serif',
              },
              fontRoute, advancePt, ascentPt: 8, descentPt: 2,
            }],
            advancePt, ascentPt: 8, descentPt: 2,
            graphemeBoundaries: [0, request.text.length],
            clusters: [...request.text].map((_character, index) => ({
              range: { start: index, end: index + 1 }, offsetPt: index * 5, advancePt: 5,
            })),
            diagnostics: [],
          };
        },
      },
    };
    const body = {
      text: 'AB', sourceRunIndex: 0, measuredWidth: 10,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
      shapedClusters: [
        { range: { start: 0, end: 1 }, offsetPt: 0, advancePt: 5 },
        { range: { start: 1, end: 2 }, offsetPt: 5, advancePt: 5 },
      ],
    } as unknown as LayoutTextSeg;
    const node = projectMeasuredSegment(numbered as never, body, {
      ...acquisitionContext,
      physicalIndentLeftPt: 12, firstIndentPt: -12,
    }, markerServices);

    expect(node.lines[0]?.placements).toEqual([
      expect.objectContaining({
        kind: 'text', role: 'numbering-marker', text: '1.',
        range: { start: -2, end: 0 }, origin: { xPt: 10, yPt: expect.any(Number) },
      }),
      expect.objectContaining({
        kind: 'text', text: 'AB', range: { start: 0, end: 2 },
        origin: { xPt: 22, yPt: expect.any(Number) },
      }),
    ]);
  });

  it.each([
    ['ruby', { hasRuby: true }],
    ['docGrid', { lineGrid: { active: true, pitchPt: 12 } }],
  ] as const)('centers visible ink in the retained %s line advance', (_name, override) => {
    const segment = {
      text: 'AB', sourceRunIndex: 0, measuredWidth: 10,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
      shapedClusters: [
        { range: { start: 0, end: 1 }, offsetPt: 0, advancePt: 5 },
        { range: { start: 1, end: 2 }, offsetPt: 5, advancePt: 5 },
      ],
    } as unknown as LayoutTextSeg;
    const node = projectMeasuredSegment(paragraph, segment, {
      ...acquisitionContext,
      lineSpacing: { rule: 'auto', value: 1 },
      ...override,
    });

    expect(node.lines[0]?.baselinePt).toBe(19);
  });

  it('applies the decimal auto-tab only to numeric content without an explicit tab', () => {
    const decimalContext = {
      ...acquisitionContext,
      tabStops: [{ pos: 50, alignment: 'decimal' as const, leader: 'none' as const }],
    };
    const measuredText = (value: string, sourceRunIndex = 0) => ({
      text: value, sourceRunIndex, measuredWidth: value.length * 5,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
      shapedClusters: [...value].map((_character, index) => ({
        range: { start: index, end: index + 1 }, offsetPt: index * 5, advancePt: 5,
      })),
    }) as unknown as LayoutTextSeg;
    const decimalParagraph = (value: string) => ({
      ...paragraph,
      tabStops: decimalContext.tabStops,
      runs: [{ ...(paragraph.runs[0] as object), text: value }],
    }) as unknown as DocParagraph;

    const numeric = projectMeasuredSegment(decimalParagraph('123'), measuredText('123'), decimalContext);
    const nonnumeric = projectMeasuredSegment(decimalParagraph('abc'), measuredText('abc'), decimalContext);
    const numericField = projectMeasuredSegment({
      ...decimalParagraph(''),
      runs: [{
        type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '123',
        bold: false, italic: false, underline: false, strikethrough: false,
        fontSize: 10, color: null, fontFamily: 'Test Sans', background: null, vertAlign: null,
      }],
    } as unknown as DocParagraph, measuredText('123'), decimalContext);
    const explicitParagraph = {
      ...decimalParagraph('123'),
      runs: [
        { type: 'ptab', alignment: 'left', relativeTo: 'margin', leader: 'none', fontSize: 10 },
        { ...(paragraph.runs[0] as object), text: '123' },
      ],
    } as unknown as DocParagraph;
    const explicit = projectMeasuredSegment(explicitParagraph, [{
      isTab: true, sourceRunIndex: 0, measuredWidth: 20,
      fontSize: 10, leader: 'none', bold: false, italic: false,
    } as unknown as LayoutTabSeg, measuredText('123', 1)], decimalContext);

    expect(numeric.lines[0]?.placements[0]).toMatchObject({ origin: { xPt: 45 } });
    expect(numericField.lines[0]?.placements[0]).toMatchObject({
      role: 'field-result', origin: { xPt: 45 },
    });
    expect(nonnumeric.lines[0]?.placements[0]).toMatchObject({ origin: { xPt: 10 } });
    expect(explicit.lines[0]?.placements).toEqual([
      expect.objectContaining({ kind: 'tab', bounds: expect.objectContaining({ xPt: 10 }) }),
      expect.objectContaining({ kind: 'text', origin: { xPt: 30, yPt: expect.any(Number) } }),
    ]);
  });

  it('uses one logical occurrence domain for text, breaks, resources, math, and tabs', () => {
    const textSegment = (value: string, sourceRunIndex: number) => ({
      text: value, sourceRunIndex, measuredWidth: value.length * 5,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
      shapedClusters: [...value].map((_character, index) => ({
        range: { start: index, end: index + 1 }, offsetPt: index * 5, advancePt: 5,
      })),
    }) as unknown as LayoutTextSeg;
    const occurrenceParagraph = {
      ...paragraph,
      runs: [
        { ...(paragraph.runs[0] as object), text: 'A' },
        { type: 'break', breakType: 'page' },
        { type: 'image', imagePath: 'word/media/a.png', mimeType: 'image/png', widthPt: 6, heightPt: 6 },
        { type: 'math', nodes: [], display: false, fontSize: 10, resourceKey: 'math:occurrence' },
        { type: 'ptab', alignment: 'left', relativeTo: 'margin', leader: 'none', fontSize: 10 },
        { ...(paragraph.runs[0] as object), text: 'B' },
      ],
    } as unknown as DocParagraph;
    const node = projectMeasuredSegment(occurrenceParagraph, [
      textSegment('A', 0),
      {
        imagePath: 'word/media/a.png', sourceRunIndex: 2, anchor: false,
        measuredWidth: 6, widthPt: 6, heightPt: 6,
      } as unknown as LayoutImageSeg,
      {
        mathNodes: [], mathResourceKey: 'math:occurrence', sourceRunIndex: 3,
        display: false, fallbackText: 'xy', measuredWidth: 10,
        mathAscent: 8, mathDescent: 2, fontSize: 10,
      } as unknown as LayoutMathSeg,
      {
        isTab: true, sourceRunIndex: 4, measuredWidth: 5,
        fontSize: 10, leader: 'none', bold: false, italic: false,
      } as unknown as LayoutTabSeg,
      textSegment('B', 5),
    ]);

    expect(node.lines[0]?.placements.map((placement) => placement.range)).toEqual([
      { start: 0, end: 1 },
      { start: 2, end: 3 },
      { start: 3, end: 5 },
      { start: 5, end: 6 },
      { start: 6, end: 7 },
    ]);
    expect(node.events).toEqual([{ kind: 'break', breakKind: 'page', offset: 1 }]);
    expect(node.lines[0]?.range).toEqual({ start: 0, end: 7 });
  });
});

describe('planLine visual geometry', () => {
  const measuredText = (textValue: string, start: number, widthPt: number, rtl = false) => ({
    kind: 'text' as const,
    text: textValue,
    range: { start, end: start + textValue.length },
    measuredWidthPt: widthPt,
    fontRoute,
    fontSizePt: 10,
    fontWeight: 400,
    fontStyle: 'normal' as const,
    color: { kind: 'explicit' as const, color: '#000000' },
    direction: rtl ? 'rtl' as const : 'ltr' as const,
    basePaintOps: [{
      text: textValue, range: { start, end: start + textValue.length }, offset: { xPt: 0, yPt: 0 },
      letterSpacingPt: 0, scaleX: 1, direction: rtl ? 'rtl' as const : 'ltr' as const,
      kerning: 'auto' as const, writingMode: 'horizontal-tb' as const,
    }],
    decorations: [],
    clusters: [{
      range: { start, end: start + textValue.length },
      offset: { xPt: 0, yPt: 0 }, advancePt: widthPt,
    }],
  });

  it.each([
    ['left', false, [10, 30]],
    ['center', false, [40, 60]],
    ['right', false, [70, 90]],
    ['left', true, [70, 90]],
  ] as const)('resolves %s alignment with baseRtl=%s into final origins', (alignment, baseRtl, xs) => {
    const line = planLine({
      paragraphXPt: 10,
      availableWidthPt: 100,
      alignment,
      baseRtl,
      isFirstLine: true,
      isLastLine: true,
      stretchLastLine: false,
      line: {
        range: { start: 0, end: 4 },
        topPt: 5,
        baselinePt: 15,
        advancePt: 14,
        xOffsetPt: 0,
        availableWidthPt: 100,
        endsWithBreak: false,
        segments: [measuredText('aa', 0, 20), measuredText('bb', 2, 20)],
      },
    });

    expect(line.placements.map((placement) => placement.kind === 'text' ? placement.origin.xPt : -1)).toEqual(xs);
    expect(line.placements.map((placement) => placement.kind === 'text' ? placement.range : null)).toEqual([
      { start: 0, end: 2 }, { start: 2, end: 4 },
    ]);
  });

  it('materializes justified trailing slack and tab geometry without paint decisions', () => {
    const line = planLine({
      paragraphXPt: 0,
      availableWidthPt: 60,
      alignment: 'both',
      baseRtl: false,
      isFirstLine: false,
      isLastLine: false,
      stretchLastLine: false,
      line: {
        range: { start: 0, end: 3 }, topPt: 0, baselinePt: 10, advancePt: 12,
        xOffsetPt: 0, availableWidthPt: 60, endsWithBreak: false,
        segments: [
          measuredText('a ', 0, 10),
          { kind: 'tab', range: { start: 2, end: 3 }, measuredWidthPt: 10, leader: 'dot', fontSizePt: 10 },
          measuredText('b', 2, 10),
        ],
      },
    });

    expect(line.placements).toEqual([
      expect.objectContaining({ kind: 'text', text: 'a ', origin: { xPt: 0, yPt: 10 } }),
      expect.objectContaining({ kind: 'tab', advancePt: 10, bounds: expect.objectContaining({ xPt: 40 }) }),
      expect.objectContaining({ kind: 'text', text: 'b', origin: { xPt: 50, yPt: 10 } }),
    ]);
    expect(() => stableFingerprint('planned-line', line)).not.toThrow();
  });

  it('retains contextual text plus uniform spacing for internal CJK justification gaps', () => {
    const line = planLine({
      paragraphXPt: 5,
      availableWidthPt: 40,
      alignment: 'both',
      baseRtl: false,
      isFirstLine: true,
      isLastLine: false,
      stretchLastLine: false,
      line: {
        range: { start: 0, end: 2 }, topPt: 2, baselinePt: 12, advancePt: 14,
        xOffsetPt: 0, availableWidthPt: 40, endsWithBreak: false,
        segments: [{
          ...measuredText('観察', 0, 20),
          clusters: [
            { range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, advancePt: 10 },
            { range: { start: 1, end: 2 }, offset: { xPt: 10, yPt: 0 }, advancePt: 10 },
          ],
        }],
      },
    });

    expect(line.placements[0]).toMatchObject({
      kind: 'text',
      origin: { xPt: 5, yPt: 12 },
      bounds: { xPt: 5, yPt: 2, widthPt: 40, heightPt: 14 },
      clusters: [
        { range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, advancePt: 10 },
        { range: { start: 1, end: 2 }, offset: { xPt: 30, yPt: 0 }, advancePt: 10 },
      ],
      paintOps: [
        expect.objectContaining({
          text: '観察', range: { start: 0, end: 2 },
          offset: { xPt: 0, yPt: 0 }, letterSpacingPt: 20,
        }),
      ],
    });
  });

  it('maps code-point justification cuts to UTF-16 shaped slices without splitting combining text', () => {
    const value = '𠮟か\u3099観';
    const line = planLine({
      paragraphXPt: 0, availableWidthPt: 64, alignment: 'both', baseRtl: false,
      isFirstLine: true, isLastLine: false, stretchLastLine: false,
      line: {
        range: { start: 10, end: 15 }, topPt: 0, baselinePt: 10, advancePt: 12,
        xOffsetPt: 0, availableWidthPt: 64, endsWithBreak: false,
        segments: [{
          ...measuredText(value, 10, 34),
          clusters: [
            { range: { start: 10, end: 12 }, offset: { xPt: 0, yPt: 0 }, advancePt: 12 },
            { range: { start: 12, end: 14 }, offset: { xPt: 12, yPt: 0 }, advancePt: 10 },
            { range: { start: 14, end: 15 }, offset: { xPt: 22, yPt: 0 }, advancePt: 12 },
          ],
        }],
      },
    });

    expect(line.placements[0]).toMatchObject({
      kind: 'text',
      paintOps: [
        expect.objectContaining({ text: '𠮟', range: { start: 10, end: 12 }, offset: { xPt: 0, yPt: 0 } }),
        expect.objectContaining({ text: 'か\u3099', range: { start: 12, end: 14 }, offset: { xPt: 27, yPt: 0 } }),
        expect.objectContaining({ text: '観', range: { start: 14, end: 15 }, offset: { xPt: 52, yPt: 0 } }),
      ],
    });
  });

  it.each([
    ['positive LTR', false, 15, undefined, 122],
    ['positive RTL', true, 15, undefined, 172],
    ['hanging LTR', false, -10, undefined, 97],
    ['hanging RTL', true, -10, undefined, 197],
    ['number tab LTR', false, -10, { bodyOffsetPt: 12 }, 119],
    ['number space LTR', false, -10, { bodyOffsetPt: 12 }, 119],
    ['number nothing LTR', false, -10, { bodyOffsetPt: 12 }, 119],
    ['number tab RTL', true, -10, { bodyOffsetPt: 12 }, 175],
    ['number space RTL', true, -10, { bodyOffsetPt: 12 }, 175],
    ['number nothing RTL', true, -10, { bodyOffsetPt: 12 }, 175],
  ] as const)('owns first-line region geometry for %s with float x-offset', (
    _name, baseRtl, firstLineIndentPt, numbering, expectedX,
  ) => {
    const line = planLine({
      paragraphXPt: 100, availableWidthPt: 100, alignment: 'left', baseRtl,
      isFirstLine: true, isLastLine: true, stretchLastLine: false,
      firstLineIndentPt,
      ...(numbering ? { numbering } : {}),
      line: {
        range: { start: 0, end: 2 }, topPt: 0, baselinePt: 10, advancePt: 12,
        xOffsetPt: 7, availableWidthPt: 100, endsWithBreak: false,
        segments: [measuredText('aa', 0, 20)],
      },
    });

    expect(line.placements[0]).toMatchObject({ origin: { xPt: expectedX, yPt: 10 } });
  });

  it.each([
    ['decimal auto-tab', { decimalAutoTabPt: 50 }, 130],
    ['display math left', { displayMathJustification: 'left' }, 107],
    ['display math centerGroup', { displayMathJustification: 'centerGroup' }, 147],
    ['display math right', { displayMathJustification: 'right' }, 187],
  ] as const)('owns the %s line-origin override', (_name, override, expectedX) => {
    const line = planLine({
      paragraphXPt: 100, availableWidthPt: 100, alignment: 'right', baseRtl: false,
      isFirstLine: false, isLastLine: true, stretchLastLine: false,
      ...override,
      line: {
        range: { start: 0, end: 2 }, topPt: 0, baselinePt: 10, advancePt: 12,
        xOffsetPt: 7, availableWidthPt: 100, endsWithBreak: false,
        segments: [measuredText('aa', 0, 20)],
      },
    });
    expect(line.placements[0]).toMatchObject({ origin: { xPt: expectedX, yPt: 10 } });
  });

  it('does not insert slack across a style boundary that continues a shaped grapheme', () => {
    const first = measuredText('か', 0, 10);
    const combining = {
      ...measuredText('\u3099', 1, 0),
      breakBefore: false,
    };
    const tail = {
      ...measuredText('観察', 2, 20),
      clusters: [
        { range: { start: 2, end: 3 }, offset: { xPt: 0, yPt: 0 }, advancePt: 10 },
        { range: { start: 3, end: 4 }, offset: { xPt: 10, yPt: 0 }, advancePt: 10 },
      ],
    };
    const line = planLine({
      paragraphXPt: 0, availableWidthPt: 40, alignment: 'both', baseRtl: false,
      isFirstLine: true, isLastLine: false, stretchLastLine: false,
      line: {
        range: { start: 0, end: 4 }, topPt: 0, baselinePt: 10, advancePt: 12,
        xOffsetPt: 0, availableWidthPt: 40, endsWithBreak: false,
        segments: [first, combining, tail],
      },
    });

    expect(line.placements.map((placement) => placement.kind === 'text' ? placement.origin.xPt : -1))
      .toEqual([0, 10, 15]);
    expect(line.placements[2]).toMatchObject({
      kind: 'text',
      paintOps: [
        expect.objectContaining({
          text: '観察', range: { start: 2, end: 4 },
          offset: { xPt: 0, yPt: 0 }, letterSpacingPt: 5,
        }),
      ],
    });
  });

  it('retains the acquisition-shaped kashida string as a contextual paint operation', () => {
    const arabic = {
      ...measuredText('سلام', 0, 20, true),
      textLayoutService: {
        fingerprint: 'kashida-service', localMetrics: {},
        resolve(): never { throw new Error('not used'); },
        shape(request: { text: string }) {
        const advancePt = [...request.text].length * 5;
        return {
          advancePt, ascentPt: 8, descentPt: 2, diagnostics: [], spans: [],
          graphemeBoundaries: [0, request.text.length],
        };
        },
      },
      textShapeRequest: {
        text: 'سلام', fontSizePt: 10, fonts: { complexScript: 'Test Sans' },
        weight: 400, style: 'normal' as const, complexScript: true,
      },
    };
    const line = planLine({
      paragraphXPt: 0, availableWidthPt: 35, alignment: 'highKashida', baseRtl: true,
      isFirstLine: true, isLastLine: false, stretchLastLine: false,
      line: {
        range: { start: 0, end: 4 }, topPt: 0, baselinePt: 10, advancePt: 12,
        xOffsetPt: 0, availableWidthPt: 35, endsWithBreak: false,
        segments: [arabic],
      },
    });
    const placement = line.placements[0];

    expect(placement).toMatchObject({
      kind: 'text', text: 'سلام',
      paintOps: [expect.objectContaining({ text: expect.stringContaining('ـ') })],
    });
  });

  it('retains complete authored typography facts independently from effective geometry', () => {
    const typographyInput = {
      sourceText: 'AB',
      underline: {
        val: { status: 'valid', raw: 'double', value: 'double' },
        color: { status: 'valid', raw: 'auto', value: 'auto' },
        themeColor: { status: 'valid', raw: 'accent2', value: 'accent2' },
        themeTint: { status: 'valid', raw: '66', value: '66' },
        themeShade: { status: 'missing', raw: null, value: null },
      },
      strike: true, doubleStrike: false, caps: false, smallCaps: true, colorAuto: false,
      verticalAlign: { status: 'valid', raw: 'superscript', value: 'super' },
      positionPt: { status: 'valid', raw: '4', value: 2 },
      snapToGrid: null, characterSpacingPt: null, characterScale: null,
      kerningThresholdPt: null,
      emphasis: { status: 'valid', raw: 'dot', value: 'dot' },
      languages: { eastAsia: null, bidi: null },
      eastAsianLayout: {
        vert: null, vertCompress: null, combine: null,
        combineBrackets: { status: 'missing', raw: null, value: null },
      },
    } as const;
    const baseRun = {
      type: 'text', text: 'AB', bold: false, italic: false, underline: true,
      strikethrough: true, fontSize: 8, color: null, fontFamily: 'Test Sans',
      background: null, vertAlign: 'super', typographyInput,
    } as const;
    const formattedParagraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
      tabStops: [], runs: [baseRun],
    } as unknown as DocParagraph;
    const textLayoutService = {
      fingerprint: 'authoritative-typography-test', localMetrics: {},
      resolve(): never { throw new Error('not used'); },
      shape(request: { text: string; fontSizePt: number }) {
        const advancePt = [...request.text].length * 5;
        const lowLine = request.text === '_';
        const inkBounds = {
          xMinPt: 0, xMaxPt: advancePt,
          ascentPt: lowLine ? 0 : request.fontSizePt * 0.7,
          descentPt: lowLine ? 1 : request.fontSizePt * 0.2,
        };
        return {
          advancePt, ascentPt: request.fontSizePt * 0.8,
          descentPt: request.fontSizePt * 0.2, inkBounds,
          diagnostics: [],
          spans: [{
            text: request.text, start: 0, end: request.text.length,
            script: 'ascii', breakBefore: true,
            font: {
              requestedFamily: 'Test Sans', resolvedFamily: 'Test Sans', route: fontRoute,
              source: 'native', weight: 400, style: 'normal', diagnostics: [], genericFamily: 'sans-serif',
            },
            fontRoute, advancePt, ascentPt: request.fontSizePt * 0.8,
            descentPt: request.fontSizePt * 0.2, inkBounds,
          }],
          graphemeBoundaries: [0, request.text.length],
          clusters: [...request.text].map((_character, index) => ({
            range: { start: index, end: index + 1 }, offsetPt: index * 5, advancePt: 5,
          })),
        };
      },
    };
    const segment = {
      text: 'AB', sourceRunIndex: 0, measuredWidth: 10,
      fontFamily: 'Test Sans', fontRoute,
      shapedClusters: [
        { range: { start: 0, end: 1 }, offsetPt: 0, advancePt: 5 },
        { range: { start: 1, end: 2 }, offsetPt: 5, advancePt: 5 },
      ],
      smallCaps: true, underline: true, underlineStyle: 'double', underlineColor: 'auto',
      strikethrough: true, doubleStrikethrough: false,
      emphasisMark: 'dot', position: 2, vertAlign: null, fontSize: 10,
      textLayoutService,
      textShapeRequest: {
        text: 'AB', fontSizePt: 10, fonts: { ascii: 'Test Sans' },
        weight: 400, style: 'normal', measure: true,
      },
    } as unknown as LayoutTextSeg;
    const node = projectMeasuredSegment(formattedParagraph, segment);
    const placement = node.lines[0]?.placements[0];

    expect(placement).toMatchObject({
      kind: 'text', fontSizePt: 8, positionPt: 2,
      typography: {
        smallCaps: true, strike: true, doubleStrike: false,
        underline: {
          val: { status: 'valid', raw: 'double', value: 'double' },
          themeColor: { status: 'valid', raw: 'accent2', value: 'accent2' },
        },
      },
    });
    expect(placement).not.toHaveProperty('unsupportedGeometry');
    expect(placement).toMatchObject({
      decorations: [
        { kind: 'underline' }, { kind: 'underline' }, { kind: 'strikethrough' },
      ],
      emphasis: {
        authored: 'dot',
        glyphs: [
          { text: '•', fontRoute, inkBounds: {} },
          { text: '•', fontRoute, inkBounds: {} },
        ],
      },
    });
  });
});
