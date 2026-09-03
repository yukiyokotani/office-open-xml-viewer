import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import {
  layoutParagraph,
  paragraphLayoutFromMeasurement,
  planLine,
  sliceParagraphLayout,
  translateParagraphLayout,
} from './paragraph.js';
import type { AcquiredParagraphLayoutInput, TextPlacement } from './types.js';
import { stableFingerprint } from './fingerprint.js';
import type { ParagraphLayoutContext } from '../layout-context.js';
import type { LayoutImageSeg, LayoutMathSeg, LayoutTabSeg, LayoutTextSeg } from '../line-layout.js';
import type { MeasuredParagraph } from '../paragraph-measure.js';
import { measureParagraph } from '../paragraph-measure.js';
import { createLayoutServices } from '../layout-runtime.js';
import type { DocParagraph } from '../types.js';
import type { AnchorAcquisitionInput } from './anchor-input.js';
import type { VerticalGlyphMeasurementService } from './measurement-capabilities.js';

const fontRoute = {
  familyList: '"Test Sans"', scope: 'native', fingerprint: 'test-font-route',
} as const;

const source = { story: 'body', storyInstance: 'body', path: [3] } as const;

const acquisitionContext: ParagraphLayoutContext = {
  lineGrid: { active: false, pitchPt: null },
  characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
  rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
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
  paragraphBorderEdges?: Parameters<typeof paragraphLayoutFromMeasurement>[1]['paragraphBorderEdges'],
  verticalGlyphMeasurement?: VerticalGlyphMeasurementService,
  verticalPageFrame = false,
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
    environment: {
      documentHasEastAsianText: false, layoutServices, verticalGlyphMeasurement,
      pageWritingMode: verticalPageFrame ? 'vertical-rl' : 'horizontal-tb',
      verticalPageFrame,
    } as never,
    exclusions: [],
    ...(anchorFrames ? { anchorFrames } : {}),
    ...(paragraphBorderEdges ? { paragraphBorderEdges } : {}),
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
    const acquired = input({ bookmarkStarts: ['destination'] });
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
    expect(first.bookmarkStarts).toEqual(['destination']);
    expect(second.bookmarkStarts).toBeUndefined();
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
      advancePt: 0,
      story: {
        story: 'textbox' as const,
        flowBounds: { xPt: 10, yPt: 25, widthPt: 0, heightPt: 0 },
        inkBounds: { xPt: 10, yPt: 25, widthPt: 0, heightPt: 0 },
        blocks: [],
        advancePt: 0,
        diagnostics: [],
      },
      transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      writingMode: 'horizontal-tb',
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

  it('characterizes continuation Y translation as the general translation with zero X delta', () => {
    const whole = layoutParagraph(input());
    const deltaYPt = whole.flowBounds.yPt - whole.lines[1]!.bounds.yPt;

    const continuation = sliceParagraphLayout(whole, {
      lineStart: 1, lineEnd: 2, continuesFromPrevious: true, continuesOnNext: false,
    });
    const generallyTranslated = translateParagraphLayout(whole, { xPt: 0, yPt: deltaYPt });

    expect(continuation.lines[0]).toEqual(generallyTranslated.lines[1]);
    expect(continuation.lines[0]!.bounds.xPt).toBe(whole.lines[1]!.bounds.xPt);
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
        sourceOrder: 1, horizontalOwnership: 'host' as const, verticalOwnership: 'host' as const,
        cellContainment: true as const,
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
        occurrenceId: 'anchor:page',
        behindDoc: hostDrawing.anchorLayer.behindDoc,
        relativeHeight: hostDrawing.anchorLayer.relativeHeight,
        sourceOrder: hostDrawing.anchorLayer.sourceOrder,
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
      cellContainmentBounds: hostDrawing.flowBounds,
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
    expect(translated.cellContainmentBounds).toEqual(translated.drawings[0]!.flowBounds);
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
      anchorLayer: {
        occurrenceId: 'anchor:first', behindDoc: false, relativeHeight: 1,
        sourceOrder: 1, horizontalOwnership: 'host' as const, verticalOwnership: 'host' as const,
        cellContainment: true as const,
      },
    } as const;
    const textBox = {
      kind: 'textbox', id: 'textbox:first', source: drawing.source,
      flowDomainId: 'textbox', ordinaryFlow: false,
      flowBounds: { xPt: 0, yPt: 0, widthPt: 5, heightPt: 5 },
      inkBounds: { xPt: 0, yPt: 0, widthPt: 5, heightPt: 5 }, advancePt: 0,
      story: {
        story: 'textbox' as const,
        flowBounds: { xPt: 0, yPt: 0, widthPt: 0, heightPt: 0 },
        inkBounds: { xPt: 0, yPt: 0, widthPt: 0, heightPt: 0 },
        blocks: [],
        advancePt: 0,
        diagnostics: [],
      },
      transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      writingMode: 'horizontal-tb',
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
      cellContainmentBounds: drawing.flowBounds,
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
    expect(first.cellContainmentBounds).toEqual(first.drawings[0]!.flowBounds);
    expect(second.cellContainmentBounds).toBeUndefined();
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

  it('extends paragraph shading through visible border spacing', () => {
    const segment = {
      text: 'AB', sourceRunIndex: 0, measuredWidth: 10,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
      shapedClusters: [
        { range: { start: 0, end: 1 }, offsetPt: 0, advancePt: 5 },
        { range: { start: 1, end: 2 }, offsetPt: 5, advancePt: 5 },
      ],
    } as unknown as LayoutTextSeg;
    const edge = (space: number, style = 'single') => ({
      style, width: 2, space, color: '123456',
    });
    const formatted = {
      type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, tabStops: [],
      runs: [{ type: 'text', text: 'AB', fontSize: 10, fontFamily: 'Test Sans' }],
      shading: 'E0E0E0',
      borders: {
        top: edge(1),
        right: edge(3),
        bottom: edge(5, 'none'),
        left: edge(2),
        between: null,
      },
    } as unknown as DocParagraph;

    const node = projectMeasuredSegment(formatted, segment);

    // ECMA-376 §17.3.1.31: shading owns the retained border box, including
    // each visible §17.3.1.7 spacing interval. A `none` edge contributes none.
    expect(node.shading).toEqual({ color: '#E0E0E0' });
    expect(node.inkBounds).toEqual({
      xPt: 8, yPt: 9, widthPt: 105, heightPt: 13,
    });
  });

  it('uses a retained between edge and suppresses an unowned bottom edge', () => {
    const segment = {
      text: 'AB', sourceRunIndex: 0, measuredWidth: 10,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
      shapedClusters: [
        { range: { start: 0, end: 1 }, offsetPt: 0, advancePt: 5 },
        { range: { start: 1, end: 2 }, offsetPt: 5, advancePt: 5 },
      ],
    } as unknown as LayoutTextSeg;
    const edge = (space: number) => ({
      style: 'single', width: 2, space, color: '123456',
    });
    const formatted = {
      type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, tabStops: [],
      runs: [{ type: 'text', text: 'AB', fontSize: 10, fontFamily: 'Test Sans' }],
      shading: 'E0E0E0',
      borders: {
        top: edge(9),
        right: null,
        bottom: edge(7),
        left: null,
        between: edge(2),
      },
    } as unknown as DocParagraph;

    const node = projectMeasuredSegment(
      formatted,
      segment,
      acquisitionContext,
      undefined,
      undefined,
      { top: 'between', bottom: 'none' },
    );

    expect(node.inkBounds).toEqual({
      xPt: 10, yPt: 8, widthPt: 100, heightPt: 14,
    });
    expect(node.borders.map((border) => border.edge)).toEqual(['between']);
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
  it('flows consecutive wp:inline WPS shapes at their paragraph pen positions', () => {
    const inlineShape = () => ({
      type: 'shape' as const,
      inline: true,
      widthPt: 20,
      heightPt: 10,
      anchorXPt: 0,
      anchorYPt: 0,
      anchorXFromMargin: false,
      anchorYFromPara: false,
      zOrder: 0,
      subpaths: [],
      presetGeometry: 'rect',
      fill: { fillType: 'solid' as const, color: '009989' },
      stroke: null,
      textBlocks: [],
    });
    const tabRun = {
      ...(paragraph.runs[0] as Extract<DocParagraph['runs'][number], { type: 'text' }>),
      text: '\t',
    };
    const inlineParagraph = {
      ...paragraph,
      runs: [
        inlineShape(), tabRun,
        inlineShape(), tabRun,
        inlineShape(),
      ],
    } as unknown as DocParagraph;
    const context: ParagraphLayoutContext = {
      ...acquisitionContext,
      tabStops: [
        { pos: 30, alignment: 'left', leader: 'none' },
        { pos: 60, alignment: 'left', leader: 'none' },
      ],
    };
    const placement = {
      startYPt: 10,
      paragraphXPt: 10,
      availableWidthPt: 100,
      maximumYPt: 500,
      suppressSpaceBefore: false,
    };
    const measurer = { context: measureContext, fontFamilyClasses: {} };
    const environment = {
      documentHasEastAsianText: false,
      pageWritingMode: 'horizontal-tb' as const,
      pageIndex: 0,
      totalPages: 1,
    };
    const measured = measureParagraph(
      inlineParagraph as never,
      context,
      placement,
      measurer,
      environment,
    );
    const node = paragraphLayoutFromMeasurement(inlineParagraph as never, {
      id: 'inline-wps', source, flowDomainId: 'body', ordinaryFlow: true,
      context, placement, measurer, environment,
      exclusions: [],
    } as never, measured);

    const placements = node.lines.flatMap((line) => line.placements)
      .filter((item) => item.kind === 'drawing');
    expect(placements.map((item) => item.bounds.xPt)).toEqual([10, 40, 70]);
    expect(placements.map((item) => item.advancePt)).toEqual([20, 20, 20]);
    expect(node.drawings.map((drawing) => drawing.flowBounds.xPt)).toEqual([10, 40, 70]);
    expect(node.drawings.every((drawing) => drawing.anchorLayer === undefined)).toBe(true);
  });

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

  it('retains selected-face ink on text paint operations for page admission', () => {
    const textParagraph = {
      ...paragraph,
      runs: [{ ...(paragraph.runs[0] as object), text: 'A' }],
    } as unknown as DocParagraph;
    const segment = {
      text: 'A', sourceRunIndex: 0, measuredWidth: 7,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
      shapedClusters: [{
        range: { start: 0, end: 1 }, offsetPt: 0, advancePt: 7,
      }],
      selectedFaceInkBounds: {
        xMinPt: 0, xMaxPt: 7, ascentPt: 7, descentPt: 3,
      },
    } as unknown as LayoutTextSeg;

    expect(projectMeasuredSegment(textParagraph, segment)
      .lines[0]?.placements[0]).toMatchObject({
      kind: 'text',
      paintOps: [{
        inkBounds: { xMinPt: 0, xMaxPt: 7, ascentPt: 7, descentPt: 3 },
      }],
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
      anchorLayer: { behindDoc: false, relativeHeight: 7, sourceOrder: 1 },
    });
    expect(node.exclusions[0]).toMatchObject({
      wrap: 'square', bounds: { xPt: 12, yPt: 13, widthPt: 20, heightPt: 10 },
    });
  });

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

  it('retains balanced authored-space width in LTR cluster geometry', () => {
    const segment = {
      text: 'AB  ', sourceRunIndex: 0, measuredWidth: 18,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
      shapedClusters: [
        { range: { start: 0, end: 1 }, offsetPt: 0, advancePt: 5 },
        { range: { start: 1, end: 2 }, offsetPt: 5, advancePt: 5 },
        { range: { start: 2, end: 3 }, offsetPt: 10, advancePt: 3 },
        { range: { start: 3, end: 4 }, offsetPt: 13, advancePt: 3 },
      ],
      snapToCharacterGrid: true,
      widthBalanceGridDeltaFactor: 0.5,
      widthBalanceSpaceSequence: true,
      widthBalanceSpaceAdjustmentPt: 2,
    } as unknown as LayoutTextSeg;
    const node = projectMeasuredSegment(paragraph, segment, {
      ...acquisitionContext,
      characterGrid: { active: true, kind: 'linesAndChars', pitchPt: 12, deltaPt: -1 },
    });
    const placement = node.lines[0]?.placements[0];
    if (placement?.kind !== 'text') throw new Error('expected retained text placement');

    expect(placement.clusters).toEqual([
      expect.objectContaining({ offset: { xPt: 0, yPt: 0 }, advancePt: 4.5 }),
      expect.objectContaining({ offset: { xPt: 4.5, yPt: 0 }, advancePt: 4.5 }),
      expect.objectContaining({ offset: { xPt: 9, yPt: 0 }, advancePt: 4.5 }),
      expect.objectContaining({ offset: { xPt: 13.5, yPt: 0 }, advancePt: 4.5 }),
    ]);
    const retainedEnd = placement.clusters.at(-1)!;
    expect(retainedEnd.offset.xPt + retainedEnd.advancePt).toBe(placement.advancePt);
  });

  it('uses balanced trailing-space clusters for the RTL physical-leading offset', () => {
    const segment = {
      text: 'نص  ', sourceRunIndex: 0, measuredWidth: 20,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
      shapedClusters: [
        { range: { start: 0, end: 1 }, offsetPt: 0, advancePt: 5 },
        { range: { start: 1, end: 2 }, offsetPt: 5, advancePt: 5 },
        { range: { start: 2, end: 3 }, offsetPt: 10, advancePt: 3 },
        { range: { start: 3, end: 4 }, offsetPt: 13, advancePt: 3 },
      ],
      rtl: true,
      direction: 'rtl',
      widthBalanceSpaceSequence: true,
      widthBalanceSpaceAdjustmentPt: 2,
    } as unknown as LayoutTextSeg;
    const node = projectMeasuredSegment(paragraph, segment, {
      ...acquisitionContext,
      baseRtl: true,
    });
    const placement = node.lines[0]?.placements[0];
    if (placement?.kind !== 'text') throw new Error('expected retained text placement');

    expect(placement.clusters.slice(-2)).toEqual([
      expect.objectContaining({ offset: { xPt: 10, yPt: 0 }, advancePt: 5 }),
      expect.objectContaining({ offset: { xPt: 15, yPt: 0 }, advancePt: 5 }),
    ]);
    expect(placement.origin.xPt - placement.bounds.xPt).toBe(10);
    expect(placement.paintOps.map((operation) => operation.range)).toEqual([
      { start: 0, end: 2 },
      { start: 2, end: 4 },
    ]);
  });

  it.each([
    ['fitText', { fitTextPerGapPx: 1 }, 13],
    ['tate-chu-yoko', { tateChuYoko: true }, 10],
  ] as const)('keeps %s override geometry outside balanced-space adjustment', (
    _name,
    override,
    measuredWidth,
  ) => {
    const segment = {
      text: 'A  ', sourceRunIndex: 0, measuredWidth,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
      shapedClusters: [
        { range: { start: 0, end: 1 }, offsetPt: 0, advancePt: 4 },
        { range: { start: 1, end: 2 }, offsetPt: 4, advancePt: 3 },
        { range: { start: 2, end: 3 }, offsetPt: 7, advancePt: 3 },
      ],
      widthBalanceSpaceSequence: true,
      widthBalanceSpaceAdjustmentPt: 2,
      ...override,
    } as unknown as LayoutTextSeg;
    const node = projectMeasuredSegment(paragraph, segment);
    const placement = node.lines[0]?.placements[0];
    if (placement?.kind !== 'text') throw new Error('expected retained text placement');

    const retainedEnd = placement.clusters.at(-1)!;
    expect(retainedEnd.offset.xPt + retainedEnd.advancePt).toBe(placement.advancePt);
    expect(placement.advancePt).toBe(measuredWidth);
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

  it('plans a directional vertical-page anchor in an upright physical paint frame', () => {
    const occurrenceId = 'anchor:body:body:3:vertical-right-arrow';
    const anchored = retainedAnchor(occurrenceId, {
      horizontal: {
        relativeFrom: 'page', relativeFromStatus: 'valid',
        choice: { kind: 'offset', valuePt: 40 },
      },
      vertical: {
        relativeFrom: 'page', relativeFromStatus: 'valid',
        choice: { kind: 'offset', valuePt: 50 },
      },
      extent: {
        widthPt: 80, widthStatus: 'valid', heightPt: 30, heightStatus: 'valid',
      },
      wrap: {
        kind: 'none', authoredKinds: [], side: null,
        distances: retainedAnchor(occurrenceId).wrap.distances,
        effectExtent: null, polygon: null,
      },
    });
    const anchorParagraph = {
      ...paragraph,
      runs: [
        { type: 'anchorHost', fontSize: 10, anchorOccurrenceId: occurrenceId },
        {
          type: 'shape', widthPt: 80, heightPt: 30,
          anchorXPt: 40, anchorYPt: 50,
          anchorXFromMargin: false, anchorYFromPara: false,
          anchorAcquisitionInput: anchored,
          presetGeometry: 'rightArrow', subpaths: [], fill: null, stroke: null,
        },
      ],
    } as unknown as DocParagraph;
    const host = {
      text: '', metricOnly: true, sourceRunIndex: 0, measuredWidth: 0,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute,
    } as unknown as LayoutTextSeg;

    const node = projectMeasuredSegment(
      anchorParagraph,
      host,
      acquisitionContext,
      undefined,
      {
        // Vertical layout retains the physical 300 x 200 page as a 300 x 200
        // anchor-reference frame; page.height is the physical paint width.
        page: { xPt: 0, yPt: 0, widthPt: 300, heightPt: 200 },
        margin: { xPt: 10, yPt: 20, widthPt: 280, heightPt: 160 },
        column: { xPt: 10, yPt: 20, widthPt: 140, heightPt: 160 },
        pageParity: 'odd',
      },
      undefined,
      undefined,
      true,
    );

    expect(node.drawings[0]).toMatchObject({
      orientation: 'upright-physical',
      flowBounds: { xPt: 50, yPt: 80, widthPt: 30, heightPt: 80 },
      transform: { a: 0, b: -1, c: 1, d: 0, e: 65, f: 120 },
      commands: [{
        kind: 'drawingml-shape',
        plan: {
          geometry: { kind: 'preset', name: 'rightArrow' },
          rect: { x: -40, y: -15, w: 80, h: 30 },
        },
      }],
    });
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

  it('resolves a public VML text shape against its authored margin alignment', () => {
    const watermark = {
      ...paragraph,
      runs: [{
        type: 'shape',
        widthPt: 40,
        heightPt: 20,
        anchorXPt: 0,
        anchorYPt: 0,
        anchorXFromMargin: true,
        anchorYFromPara: false,
        anchorXAlign: 'center',
        anchorYAlign: 'center',
        anchorXRelativeFrom: 'margin',
        anchorYRelativeFrom: 'margin',
        behindDoc: true,
        zOrder: 0,
        subpaths: [],
        presetGeometry: 'rect',
        fill: { fillType: 'solid', color: 'c0c0c0' },
        stroke: null,
        rotation: 315,
        fillOpacity: 0.5,
        vmlTextPathInput: {
          string: 'DRAFT',
          textPathOk: true,
          on: true,
          fitShape: true,
          fitPath: false,
          trim: false,
          xScale: false,
          fontSizePt: 1,
        },
      }],
    } as unknown as DocParagraph;
    const host = {
      text: '',
      metricOnly: true,
      sourceRunIndex: 0,
      measuredWidth: 0,
      fontSize: 10,
      fontFamily: 'Test Sans',
      fontRoute,
    } as unknown as LayoutTextSeg;
    const textService = {
      shape: () => ({
        advancePt: 5,
        ascentPt: 0.8,
        descentPt: 0.2,
        spans: [{
          text: 'DRAFT',
          start: 0,
          end: 5,
          script: 'ascii',
          breakBefore: true,
          advancePt: 5,
          ascentPt: 0.8,
          descentPt: 0.2,
          font: {
            family: 'Test Sans',
            route: fontRoute,
            weight: 400,
            style: 'normal',
            diagnostics: [],
          },
          fontRoute,
        }],
        diagnostics: [],
        graphemeBoundaries: [0, 1, 2, 3, 4, 5],
      }),
    };

    const node = projectMeasuredSegment(
      watermark,
      host,
      acquisitionContext,
      { text: textService },
      {
        page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
        margin: { xPt: 20, yPt: 30, widthPt: 160, heightPt: 240 },
        column: { xPt: 20, yPt: 30, widthPt: 80, heightPt: 240 },
        pageParity: 'odd',
      },
    );

    expect(node.drawings[0]).toMatchObject({
      flowBounds: { xPt: 80, yPt: 140, widthPt: 40, heightPt: 20 },
      inkBounds: { xPt: 80, yPt: 140, widthPt: 40, heightPt: 20 },
      anchorLayer: {
        behindDoc: true,
        horizontalOwnership: 'page',
        verticalOwnership: 'page',
      },
      commands: [{
        kind: 'watermark-text',
        rect: { xPt: 80, yPt: 140, widthPt: 40, heightPt: 20 },
        text: 'DRAFT',
        fitShape: true,
      }],
    });
  });

  it('resolves public anchor margin strips and mirrors inside alignment by page parity', () => {
    const host = {
      text: '',
      metricOnly: true,
      sourceRunIndex: 0,
      measuredWidth: 0,
      fontSize: 10,
      fontFamily: 'Test Sans',
      fontRoute,
    } as unknown as LayoutTextSeg;
    const acquire = (
      pageParity: 'odd' | 'even',
      overrides: Record<string, unknown>,
    ) => {
      const anchored = {
        ...paragraph,
        runs: [{
          type: 'shape',
          widthPt: 10,
          heightPt: 10,
          anchorXPt: 0,
          anchorYPt: 0,
          anchorXFromMargin: false,
          anchorYFromPara: false,
          zOrder: 0,
          subpaths: [],
          presetGeometry: 'rect',
          fill: null,
          stroke: null,
          ...overrides,
        }],
      } as unknown as DocParagraph;
      return projectMeasuredSegment(
        anchored,
        host,
        acquisitionContext,
        undefined,
        {
          page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
          margin: { xPt: 20, yPt: 30, widthPt: 160, heightPt: 240 },
          column: { xPt: 20, yPt: 30, widthPt: 80, heightPt: 240 },
          pageParity,
        },
      ).drawings[0]!.flowBounds;
    };

    expect(acquire('odd', {
      anchorXRelativeFrom: 'margin',
      anchorXAlign: 'inside',
      anchorYRelativeFrom: 'page',
    }).xPt).toBe(20);
    expect(acquire('even', {
      anchorXRelativeFrom: 'margin',
      anchorXAlign: 'inside',
      anchorYRelativeFrom: 'page',
    }).xPt).toBe(170);
    expect(acquire('odd', {
      anchorXRelativeFrom: 'insideMargin',
      anchorXAlign: 'center',
      anchorYRelativeFrom: 'page',
    }).xPt).toBe(5);
    expect(acquire('even', {
      anchorXRelativeFrom: 'insideMargin',
      anchorXAlign: 'center',
      anchorYRelativeFrom: 'page',
    }).xPt).toBe(185);
    expect(acquire('odd', {
      anchorXRelativeFrom: 'page',
      anchorYRelativeFrom: 'topMargin',
      anchorYAlign: 'bottom',
    }).yPt).toBe(20);
    expect(acquire('odd', {
      anchorXRelativeFrom: 'page',
      anchorYRelativeFrom: 'bottomMargin',
      anchorYAlign: 'top',
    }).yPt).toBe(270);
    expect(acquire('odd', {
      anchorXRelativeFrom: 'page',
      anchorYRelativeFrom: 'line',
      anchorYPt: 3,
    }).yPt).toBe(13);
  });

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
        pageWritingMode: 'horizontal-tb',
        layoutServices: services,
      },
    );
    const node = paragraphLayoutFromMeasurement(cjkParagraph as never, {
      id: 'ordinary-cjk', source, flowDomainId: 'body', ordinaryFlow: true,
      context: acquisitionContext, placement: measured.placement,
      measurer: { context: measureContext, fontFamilyClasses: {} },
      environment: {
        pageIndex: 0, totalPages: 1, documentHasEastAsianText: true,
        pageWritingMode: 'horizontal-tb',
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
      {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: false,
      },
    );

    expect(() => paragraphLayoutFromMeasurement(paragraph as never, {
      id: 'missing-clusters', source, flowDomainId: 'body', ordinaryFlow: true,
      context: acquisitionContext, placement: measured.placement,
      measurer: { context: measureContext, fontFamilyClasses: {} },
      environment: {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: false,
      },
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
      characterGrid: { active: true, kind: 'linesAndChars', pitchPt: 12, deltaPt: 2 },
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
      math: true, mathResourceKey: 'math:display', display: true,
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

  it('retains the vertical glyph plan instead of reclassifying text at paint time', () => {
    const value = 'A（ー）B';
    const verticalParagraph = {
      ...paragraph,
      runs: [{ ...(paragraph.runs[0] as object), text: value }],
    } as unknown as DocParagraph;
    const segment = {
      text: value, sourceRunIndex: 0, measuredWidth: 50,
      fontSize: 10, fontFamily: 'Test Sans', fontRoute, verticalRun: true,
      // A negative East Asian pitch narrows each retained cell. Upright/rotate
      // glyphs are already positioned by those cell origins; only the
      // contextual sideways pieces still need Canvas letter spacing.
      fitTextPerGapPx: -6,
      shapedClusters: [...value].map((_character, index) => ({
        range: { start: index, end: index + 1 }, offsetPt: index * 10, advancePt: 10,
      })),
    } as unknown as LayoutTextSeg;
    const verticalGlyphMeasurement: VerticalGlyphMeasurementService = {
      fingerprint: 'vertical:retained-plan-test',
      measureRunInkExtra: () => 0,
      planRun: () => [
        {
          text: 'A', range: { start: 0, end: 1 }, orientation: 'sideways',
          originPt: 0, advancePt: 10, drawOffsetPt: { xPt: 0, yPt: 4 }, verticalFeature: false,
          blockAxisInkBounds: { startPt: -4, endPt: 6 },
        },
        {
          text: '︵', range: { start: 1, end: 2 }, orientation: 'upright',
          originPt: 15, advancePt: 10, drawOffsetPt: { xPt: 0, yPt: -2 }, verticalFeature: false,
          blockAxisInkBounds: { startPt: -5, endPt: 5 },
        },
        {
          text: 'ー', range: { start: 2, end: 3 }, orientation: 'rotate',
          originPt: 25, advancePt: 10, drawOffsetPt: { xPt: 0, yPt: 0 }, verticalFeature: false,
          blockAxisInkBounds: { startPt: -5, endPt: 5 },
        },
        {
          text: '）', range: { start: 3, end: 4 }, orientation: 'upright',
          originPt: 35, advancePt: 10, drawOffsetPt: { xPt: 0, yPt: 0 }, verticalFeature: true,
          blockAxisInkBounds: { startPt: -5, endPt: 5 },
        },
        {
          text: 'B', range: { start: 4, end: 5 }, orientation: 'sideways',
          originPt: 40, advancePt: 10, drawOffsetPt: { xPt: 0, yPt: 4 }, verticalFeature: false,
          blockAxisInkBounds: { startPt: -4, endPt: 6 },
        },
      ],
    };

    const node = projectMeasuredSegment(
      verticalParagraph, segment, acquisitionContext,
      undefined, undefined, undefined, verticalGlyphMeasurement,
    );
    const placement = node.lines[0]?.placements[0];

    expect(placement).toMatchObject({
      kind: 'text',
      text: value,
      paintOps: [
        expect.objectContaining({
          text: 'A', range: { start: 0, end: 1 }, glyphOrientation: 'sideways',
          offset: { xPt: 0, yPt: 0 }, glyphOffsetPt: { xPt: 0, yPt: 4 },
          letterSpacingPt: -6,
        }),
        expect.objectContaining({
          text: '︵', range: { start: 1, end: 2 }, glyphOrientation: 'upright',
          offset: { xPt: 15, yPt: 0 }, glyphOffsetPt: { xPt: 0, yPt: -2 },
          letterSpacingPt: 0,
        }),
        expect.objectContaining({
          text: 'ー', range: { start: 2, end: 3 }, glyphOrientation: 'rotate',
          offset: { xPt: 25, yPt: 0 },
          letterSpacingPt: 0,
        }),
        expect.objectContaining({
          text: '）', range: { start: 3, end: 4 }, glyphOrientation: 'upright',
          offset: { xPt: 35, yPt: 0 }, verticalFeature: true, letterSpacingPt: 0,
        }),
        expect.objectContaining({
          text: 'B', range: { start: 4, end: 5 }, glyphOrientation: 'sideways',
          offset: { xPt: 40, yPt: 0 }, glyphOffsetPt: { xPt: 0, yPt: 4 },
          letterSpacingPt: -6,
        }),
      ],
    });
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
        math: true, mathResourceKey: 'math:occurrence', sourceRunIndex: 3,
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

  it('retains one continuous dotted underline cadence across adjacent source runs', () => {
    const dotted = (fromXPt: number, toXPt: number) => ({
      kind: 'underline' as const,
      authoredStyle: 'dotted',
      from: { xPt: fromXPt, yPt: 16 },
      to: { xPt: toXPt, yPt: 16 },
      color: '#000000',
      widthPt: 1,
      style: 'dotted' as const,
      dashPatternPt: [1, 2] as const,
    });
    const first = { ...measuredText('aa', 0, 20), decorations: [dotted(0, 20)] };
    const second = { ...measuredText('bb', 2, 20), decorations: [dotted(20, 40)] };

    const line = planLine({
      paragraphXPt: 0,
      availableWidthPt: 100,
      alignment: 'left',
      baseRtl: false,
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
        segments: [first, second],
      },
    });

    expect(line.placements).toMatchObject([
      { kind: 'text', decorations: [{ from: { xPt: 0 }, to: { xPt: 40 } }] },
      { kind: 'text', decorations: [] },
    ]);
  });

  it('uses one safe baseline for a solid underline spanning adjacent source runs', () => {
    const underline = (fromXPt: number, toXPt: number, yPt: number) => ({
      kind: 'underline' as const,
      from: { xPt: fromXPt, yPt }, to: { xPt: toXPt, yPt },
      color: '#000000', widthPt: 1, style: 'solid' as const,
    });
    const first = { ...measuredText('aa', 0, 20), decorations: [underline(0, 20, 16)] };
    const second = { ...measuredText('bb', 2, 20), decorations: [underline(20, 40, 17)] };

    const line = planLine({
      paragraphXPt: 0, availableWidthPt: 100, alignment: 'left', baseRtl: false,
      isFirstLine: true, isLastLine: true, stretchLastLine: false,
      line: {
        range: { start: 0, end: 4 }, topPt: 5, baselinePt: 15, advancePt: 14,
        xOffsetPt: 0, availableWidthPt: 100, endsWithBreak: false,
        segments: [first, second],
      },
    });

    expect(line.placements).toMatchObject([
      { kind: 'text', decorations: [{ from: { xPt: 0, yPt: 17 }, to: { xPt: 40, yPt: 17 } }] },
      { kind: 'text', decorations: [] },
    ]);
  });

  it('keeps a solid underline continuous across floating-precision retained run seams', () => {
    const seamPt = 300;
    const driftedSeamPt = seamPt + Number.EPSILON * seamPt * 2;
    const underline = (fromXPt: number, toXPt: number, yPt: number) => ({
      kind: 'underline' as const,
      from: { xPt: fromXPt, yPt }, to: { xPt: toXPt, yPt },
      color: '#ff3333', widthPt: 1, style: 'solid' as const,
    });
    const first = {
      ...measuredText('left  ', 0, seamPt),
      decorations: [underline(0, seamPt, 16)],
    };
    const second = {
      ...measuredText('right', 6, 20),
      decorations: [underline(driftedSeamPt, 320, 17)],
    };

    const line = planLine({
      paragraphXPt: 0, availableWidthPt: 400, alignment: 'left', baseRtl: false,
      isFirstLine: true, isLastLine: true, stretchLastLine: false,
      line: {
        range: { start: 0, end: 11 }, topPt: 5, baselinePt: 15, advancePt: 14,
        xOffsetPt: 0, availableWidthPt: 400, endsWithBreak: false,
        segments: [first, second],
      },
    });

    expect(line.placements).toMatchObject([
      { kind: 'text', decorations: [{ from: { xPt: 0, yPt: 17 }, to: { xPt: 320, yPt: 17 } }] },
      { kind: 'text', decorations: [] },
    ]);
  });

  it('keeps a one-point underline gap across adjacent runs', () => {
    const underline = (fromXPt: number, toXPt: number) => ({
      kind: 'underline' as const,
      from: { xPt: fromXPt, yPt: 16 }, to: { xPt: toXPt, yPt: 16 },
      color: '#000000', widthPt: 1, style: 'solid' as const,
    });
    const first = { ...measuredText('aa', 0, 20), decorations: [underline(0, 20)] };
    const second = { ...measuredText('bb', 2, 20), decorations: [underline(21, 40)] };

    const line = planLine({
      paragraphXPt: 0, availableWidthPt: 100, alignment: 'left', baseRtl: false,
      isFirstLine: true, isLastLine: true, stretchLastLine: false,
      line: {
        range: { start: 0, end: 4 }, topPt: 5, baselinePt: 15, advancePt: 14,
        xOffsetPt: 0, availableWidthPt: 100, endsWithBreak: false,
        segments: [first, second],
      },
    });

    expect(line.placements.map((placement) =>
      placement.kind === 'text' ? placement.decorations?.length : 0)).toEqual([1, 1]);
  });

  it('retains one wave cadence across adjacent underlined source runs', () => {
    const wave = (fromXPt: number, toXPt: number) => ({
      kind: 'underline' as const,
      authoredStyle: 'wave',
      from: { xPt: fromXPt, yPt: 16 }, to: { xPt: toXPt, yPt: 16 },
      color: '#000000', widthPt: 1, style: 'wavy' as const,
      path: [
        { xPt: fromXPt, yPt: 15.5 },
        { xPt: toXPt, yPt: 16.5 },
      ],
    });
    const first = { ...measuredText('aa', 0, 20), decorations: [wave(0, 20)] };
    const second = { ...measuredText('bb', 2, 20), decorations: [wave(20, 40)] };

    const line = planLine({
      paragraphXPt: 0, availableWidthPt: 100, alignment: 'left', baseRtl: false,
      isFirstLine: true, isLastLine: true, stretchLastLine: false,
      line: {
        range: { start: 0, end: 4 }, topPt: 5, baselinePt: 15, advancePt: 14,
        xOffsetPt: 0, availableWidthPt: 100, endsWithBreak: false,
        segments: [first, second],
      },
    });

    const decorations = line.placements.flatMap((placement) =>
      placement.kind === 'text' ? placement.decorations : []);
    expect(decorations).toHaveLength(1);
    expect(decorations[0]).toMatchObject({
      style: 'wavy', from: { xPt: 0, yPt: 16 }, to: { xPt: 40, yPt: 16 },
    });
    expect(decorations[0]?.path?.filter((point) => point.xPt === 20)).toHaveLength(1);
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

  it('retains an underline across the full advance of an underlined tab run', () => {
    const line = planLine({
      paragraphXPt: 0, availableWidthPt: 100, alignment: 'left', baseRtl: false,
      isFirstLine: true, isLastLine: true, stretchLastLine: false,
      line: {
        range: { start: 0, end: 1 }, topPt: 0, baselinePt: 10, advancePt: 12,
        xOffsetPt: 0, availableWidthPt: 100, endsWithBreak: false,
        segments: [{
          kind: 'tab', range: { start: 0, end: 1 }, measuredWidthPt: 72,
          leader: 'none', fontSizePt: 10,
          underline: {
            base: { ascentPt: 8, descentPt: 2, inkBounds: { xMinPt: 0, xMaxPt: 5, ascentPt: 7, descentPt: 1 } },
            probe: { ascentPt: 8, descentPt: 2, inkBounds: { xMinPt: 0, xMaxPt: 5, ascentPt: -2, descentPt: 3 } },
            color: '#123456', authoredStyle: 'single',
          },
        }],
      },
    });

    expect(line.placements[0]).toMatchObject({
      kind: 'tab', advancePt: 72,
      decorations: [{
        kind: 'underline', color: '#123456',
        from: { xPt: 0 }, to: { xPt: 72 },
      }],
    });
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

  it('retains every RTL trailing-space source slice after internal justification', () => {
    const value = 'نص  ';
    const start = 246;
    const base = measuredText(value, start, 30, true);
    const line = planLine({
      paragraphXPt: 0,
      availableWidthPt: 80,
      alignment: 'both',
      baseRtl: true,
      isFirstLine: true,
      isLastLine: false,
      stretchLastLine: false,
      line: {
        range: { start: start - 4, end: start + value.length + 2 },
        topPt: 0,
        baselinePt: 10,
        advancePt: 12,
        xOffsetPt: 0,
        availableWidthPt: 100,
        endsWithBreak: false,
        segments: [
          {
            ...measuredText('قبل ', start - 4, 20, true),
            clusters: [...'قبل '].map((_character, index) => ({
              range: { start: start - 4 + index, end: start - 3 + index },
              offset: { xPt: index * 5, yPt: 0 },
              advancePt: 5,
            })),
          },
          {
            ...base,
            clusters: [...value].map((_character, index) => ({
              range: { start: start + index, end: start + index + 1 },
              offset: { xPt: index * 5, yPt: 0 },
              advancePt: 5,
            })),
          },
          {
            ...measuredText('24', start + value.length, 10, false),
            clusters: [...'24'].map((_character, index) => ({
              range: {
                start: start + value.length + index,
                end: start + value.length + index + 1,
              },
              offset: { xPt: index * 5, yPt: 0 },
              advancePt: 5,
            })),
          },
        ],
      },
    });
    const placement = line.placements.find((candidate) => candidate.range.start === start);
    expect(placement?.kind).toBe('text');
    if (placement?.kind !== 'text') throw new Error('expected retained text placement');

    expect(placement.range).toEqual({ start: 246, end: 250 });
    expect(placement.paintOps.map((operation) => operation.range)).toEqual([
      { start: 246, end: 248 },
      { start: 248, end: 249 },
      { start: 249, end: 250 },
    ]);
    expect(placement.paintOps.map((operation) => operation.text)).toEqual([
      'نص',
      ' ',
      ' ',
    ]);
  });

  it('shifts cluster-aligned vertical paint operations during internal justification', () => {
    const base = measuredText('観察', 0, 20);
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
          ...base,
          clusters: [
            { range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, advancePt: 10 },
            { range: { start: 1, end: 2 }, offset: { xPt: 10, yPt: 0 }, advancePt: 10 },
          ],
          basePaintOps: [
            {
              ...base.basePaintOps[0]!, text: '観', range: { start: 0, end: 1 },
              offset: { xPt: 5, yPt: 0 }, writingMode: 'vertical-rl', glyphOrientation: 'upright',
            },
            {
              ...base.basePaintOps[0]!, text: '察', range: { start: 1, end: 2 },
              offset: { xPt: 15, yPt: 0 }, writingMode: 'vertical-rl', glyphOrientation: 'upright',
            },
          ],
        }],
      },
    });

    expect(line.placements[0]).toMatchObject({
      kind: 'text',
      clusters: [
        { range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 } },
        { range: { start: 1, end: 2 }, offset: { xPt: 30, yPt: 0 } },
      ],
      paintOps: [
        expect.objectContaining({
          text: '観', range: { start: 0, end: 1 }, offset: { xPt: 5, yPt: 0 },
          writingMode: 'vertical-rl', glyphOrientation: 'upright',
        }),
        expect.objectContaining({
          text: '察', range: { start: 1, end: 2 }, offset: { xPt: 35, yPt: 0 },
          writingMode: 'vertical-rl', glyphOrientation: 'upright',
        }),
      ],
    });
  });

  it('intersects internal justification cuts with several retained paint operations', () => {
    const base = measuredText('観察林', 0, 30);
    const line = planLine({
      paragraphXPt: 0, availableWidthPt: 60, alignment: 'both', baseRtl: false,
      isFirstLine: true, isLastLine: false, stretchLastLine: false,
      line: {
        range: { start: 0, end: 3 }, topPt: 0, baselinePt: 10, advancePt: 12,
        xOffsetPt: 0, availableWidthPt: 60, endsWithBreak: false,
        segments: [{
          ...base,
          clusters: [
            { range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, advancePt: 10 },
            { range: { start: 1, end: 2 }, offset: { xPt: 10, yPt: 0 }, advancePt: 10 },
            { range: { start: 2, end: 3 }, offset: { xPt: 20, yPt: 0 }, advancePt: 10 },
          ],
          basePaintOps: [
            {
              ...base.basePaintOps[0]!, text: '観察', range: { start: 0, end: 2 },
              offset: { xPt: 0, yPt: 0 },
            },
            {
              ...base.basePaintOps[0]!, text: '林', range: { start: 2, end: 3 },
              offset: { xPt: 20, yPt: 0 },
            },
          ],
        }],
      },
    });

    expect(line.placements[0]).toMatchObject({
      kind: 'text',
      clusters: [
        { range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 } },
        { range: { start: 1, end: 2 }, offset: { xPt: 25, yPt: 0 } },
        { range: { start: 2, end: 3 }, offset: { xPt: 50, yPt: 0 } },
      ],
      paintOps: [
        expect.objectContaining({ text: '観', range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 } }),
        expect.objectContaining({ text: '察', range: { start: 1, end: 2 }, offset: { xPt: 25, yPt: 0 } }),
        expect.objectContaining({ text: '林', range: { start: 2, end: 3 }, offset: { xPt: 50, yPt: 0 } }),
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

  it('retains superscript baseline geometry across the main/worker clone boundary', () => {
    const superscript = {
      ...measuredText('S', 0, 5),
      verticalAlign: 'super' as const,
      basePaintOps: [{
        ...measuredText('S', 0, 5).basePaintOps[0],
        offset: { xPt: 0, yPt: -3.5 },
      }],
    };
    const main = planLine({
      paragraphXPt: 0, availableWidthPt: 100, alignment: 'left', baseRtl: false,
      isFirstLine: true, isLastLine: true, stretchLastLine: false,
      line: {
        range: { start: 0, end: 1 }, topPt: 0, baselinePt: 10, advancePt: 12,
        xOffsetPt: 0, availableWidthPt: 100, endsWithBreak: false,
        segments: [superscript],
      },
    });
    const workerClone = structuredClone(main);

    expect(workerClone).toEqual(main);
    expect(workerClone.placements[0]).toMatchObject({
      kind: 'text',
      verticalAlign: 'super',
      paintOps: [{ offset: { xPt: 0, yPt: -3.5 } }],
    });
    expect(stableFingerprint('superscript-line', workerClone))
      .toBe(stableFingerprint('superscript-line', main));
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
        // Canvas reports a low-line wholly below the alphabetic baseline with
        // a signed ascent. The aggregate deliberately mirrors the old lossy
        // shape union so this test requires the selected probe span itself.
        const spanInkBounds = {
          xMinPt: 0, xMaxPt: advancePt,
          ascentPt: lowLine ? -2.5 : request.fontSizePt * 0.7,
          descentPt: lowLine ? 3 : request.fontSizePt * 0.2,
        };
        const inkBounds = lowLine ? { ...spanInkBounds, ascentPt: 0 } : spanInkBounds;
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
            descentPt: request.fontSizePt * 0.2, inkBounds: spanInkBounds,
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
      selectedFaceFontBox: { ascentPt: 8, descentPt: 2 },
      selectedFaceInkBounds: {
        xMinPt: 0, xMaxPt: 10, ascentPt: 7, descentPt: 2,
      },
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
        { kind: 'underline', widthPt: .5 },
        { kind: 'underline', widthPt: .5 },
        { kind: 'strikethrough' },
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
