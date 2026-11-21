import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { acquireParagraphLayout, layoutParagraph, sliceParagraphLayout } from './paragraph.js';
import type { AcquiredParagraphLayoutInput, InlineResourceLayout, LayoutServices, ParagraphPlacement, WrapExclusion } from './types.js';
import type { DocParagraph, DocRun, DocxTextRun } from '../types.js';
import type { ParagraphLayoutContext } from '../layout-context.js';
import { paragraphAcquisitionInput } from '../parser-model.js';
import { paintParagraphLayout } from '../paint/canvas-text.js';
import type { CanvasPaintResourcePainter } from '../paint/types.js';
import type { AnchorAcquisitionInput } from './anchor-input.js';
import type { WrapOracle } from '../paragraph-measure.js';

const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
const bounds = { xPt: 10, yPt: 20, widthPt: 12, heightPt: 10 } as const;
const fontRoute = { familyList: 'Test Sans', scope: 'native', fingerprint: 'test-font-route' } as const;
const noPaintResources: CanvasPaintResourcePainter = {
  paint(resourceKey, kind): never {
    throw new Error(`Unexpected ${kind} paint resource: ${resourceKey}`);
  },
};

const validEdges = (topPt: number, rightPt: number, bottomPt: number, leftPt: number) => ({
  topPt, topStatus: 'valid', rightPt, rightStatus: 'valid',
  bottomPt, bottomStatus: 'valid', leftPt, leftStatus: 'valid',
} as const);

function sameParagraphAnchor(
  occurrenceId: string,
  wrap: AnchorAcquisitionInput['wrap'],
): AnchorAcquisitionInput {
  return {
    occurrenceId,
    simplePosition: {
      enabled: false, status: 'valid', xPt: 0, xStatus: 'valid', yPt: 0, yStatus: 'valid',
    },
    horizontal: {
      relativeFrom: 'column', relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: 82 },
    },
    vertical: {
      relativeFrom: 'paragraph', relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: 0 },
    },
    extent: { widthPt: 30, widthStatus: 'valid', heightPt: 20, heightStatus: 'valid' },
    parentEffectExtent: validEdges(2, 3, 4, 1),
    anchorDistances: validEdges(5, 6, 7, 8),
    relativeSize: { horizontal: null, vertical: null },
    wrap,
    behavior: {
      behindDoc: false, behindDocStatus: 'valid',
      relativeHeight: 1, relativeHeightStatus: 'valid',
      locked: false, lockedStatus: 'valid',
      allowOverlap: true, allowOverlapStatus: 'valid',
      layoutInCell: true, layoutInCellStatus: 'valid',
    },
    group: null,
  };
}

const sameParagraphContext: ParagraphLayoutContext = {
  lineGrid: { active: false, pitchPt: null }, characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
  rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
  physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
  lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0, baseRtl: false,
  isJustified: false, stretchLastLine: false, tabStops: [], hasRuby: false,
  hasEastAsianText: false, kinsoku: DEFAULT_KINSOKU_RULES, defaultTabPt: 36,
};

let sameParagraphFont = '10px Test Sans';
const sameParagraphMeasureContext = {
  get font() { return sameParagraphFont; }, set font(value: string) { sameParagraphFont = value; },
  letterSpacing: '0px', fontKerning: 'auto',
  measureText(value: string) {
    return { width: [...value].length * 5, actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2, fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2 } as TextMetrics;
  },
} as unknown as CanvasRenderingContext2D;

const sameParagraphTextRun = (value: string): DocRun => ({
  type: 'text', text: value, bold: false, italic: false, underline: false,
  strikethrough: false, fontSize: 10, color: null, fontFamily: 'Test Sans',
  isLink: false, background: null, vertAlign: null, hyperlink: null,
} as DocxTextRun & { type: 'text' });

const sameParagraphTextService: LayoutServices['text'] = {
  fingerprint: 'same-paragraph-fixed-width', localMetrics: {},
  resolve(request) {
    const requestedFamily = request.fonts.ascii ?? 'Test Sans';
    return {
      requestedFamily, resolvedFamily: requestedFamily, route: fontRoute,
      source: 'native', weight: request.weight ?? 400, style: request.style ?? 'normal',
      diagnostics: [], genericFamily: 'sans-serif',
    };
  },
  shape(request) {
    let offset = 0;
    const clusters = [...request.text].map((character) => {
      const start = offset;
      offset += character.length;
      return { range: { start, end: offset }, offsetPt: start * 5, advancePt: 5 };
    });
    return {
      text: request.text, advancePt: clusters.length * 5, ascentPt: 8, descentPt: 2,
      spans: [{
        text: request.text, start: 0, end: request.text.length, script: 'ascii',
        breakBefore: true,
        font: {
          requestedFamily: 'Test Sans', resolvedFamily: 'Test Sans', route: fontRoute,
          source: 'native', weight: request.weight ?? 400, style: request.style ?? 'normal',
          diagnostics: [], genericFamily: 'sans-serif',
        },
        fontRoute, advancePt: clusters.length * 5, ascentPt: 8, descentPt: 2,
      }],
      graphemeBoundaries: [0, ...clusters.map((cluster) => cluster.range.end)],
      clusters, diagnostics: [],
    };
  },
};

const sameParagraphOccurrenceId = 'anchor:body:body:0:same-paragraph';

function acquireRuns(
  runs: readonly unknown[],
  pathIndex: number,
  id: string,
) {
  const paragraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs,
  } as unknown as DocParagraph;
  const runSource = { ...source, path: [pathIndex] } as const;
  return acquireParagraphLayout(paragraphAcquisitionInput(paragraph, runSource), {
    id,
    source: runSource,
    flowDomainId: 'body',
    ordinaryFlow: true,
    context: sameParagraphContext,
    placement: {
      startYPt: 10,
      paragraphXPt: 10,
      availableWidthPt: 100,
      maximumYPt: 300,
      suppressSpaceBefore: false,
    },
    measurer: { context: sameParagraphMeasureContext, fontFamilyClasses: {} },
    environment: {
      pageIndex: 0,
      totalPages: 1,
      pageWritingMode: 'horizontal-tb',
      documentHasEastAsianText: false,
      layoutServices: {
        text: sameParagraphTextService,
        images: {
          fingerprint: 'images',
          resolve: () => ({ widthPt: 30, heightPt: 20, mimeType: 'image/png' }),
        },
        math: {
          fingerprint: 'math',
          resolve: (resourceKey) => ({
            resourceKey, widthEm: 1, ascentEm: .8, descentEm: .2, diagnostics: [],
          }),
        },
      },
    },
    exclusions: [],
    anchorFrames: {
      page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
      margin: { xPt: 10, yPt: 10, widthPt: 180, heightPt: 280 },
      column: { xPt: 10, yPt: 10, widthPt: 100, heightPt: 280 },
      pageParity: 'odd',
    },
  });
}

function acquireSameParagraph(
  anchorAcquisitionInput: AnchorAcquisitionInput,
  overrides: Readonly<{
    wrap?: WrapOracle;
    exclusions?: Parameters<typeof acquireParagraphLayout>[1]['exclusions'];
    verticalPageFrame?: boolean;
    pageWritingMode?: 'vertical-rl' | 'vertical-lr';
  }> | undefined = undefined,
) {
  const paragraph: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [
      { type: 'anchorHost', fontSize: 10, fontFamily: 'Test Sans', anchorOccurrenceId: sameParagraphOccurrenceId },
      {
        type: 'image', imagePath: 'word/media/owned.png', mimeType: 'image/png',
        widthPt: 30, heightPt: 20, anchor: true, anchorAcquisitionInput,
      },
      sameParagraphTextRun('abcdefghijklmnopqrstuvwxyzabcdefghi'),
    ] as DocRun[],
  };
  return acquireParagraphLayout(paragraphAcquisitionInput(paragraph, source), {
    id: `same-paragraph:${anchorAcquisitionInput.wrap.kind}`,
    source, flowDomainId: 'body', ordinaryFlow: true,
    context: sameParagraphContext,
    placement: {
      startYPt: 10, paragraphXPt: 10, availableWidthPt: 100,
      maximumYPt: 300, suppressSpaceBefore: false,
      ...(overrides?.wrap ? { wrap: overrides.wrap } : {}),
    },
    measurer: { context: sameParagraphMeasureContext, fontFamilyClasses: {} },
    environment: {
      pageIndex: 0, totalPages: 1, documentHasEastAsianText: false,
      pageWritingMode: overrides?.pageWritingMode ?? (
        overrides?.verticalPageFrame ? 'vertical-rl' : 'horizontal-tb'
      ),
      ...(overrides?.verticalPageFrame ? { verticalPageFrame: true } : {}),
      layoutServices: {
        text: sameParagraphTextService,
        images: {
          fingerprint: 'images',
          resolve: () => ({ widthPt: 30, heightPt: 20, mimeType: 'image/png' }),
        },
        math: {
          fingerprint: 'math',
          resolve: (resourceKey) => ({
            resourceKey, widthEm: 1, ascentEm: .8, descentEm: .2, diagnostics: [],
          }),
        },
      },
    },
    exclusions: overrides?.exclusions ?? [],
    anchorFrames: {
      page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
      margin: { xPt: 10, yPt: 10, widthPt: 180, heightPt: 280 },
      column: { xPt: 10, yPt: 10, widthPt: 100, heightPt: 280 },
      pageParity: 'odd',
    },
  });
}

function externalExclusion(
  id: string,
  bounds: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>,
  anchorOccurrenceId?: string,
): WrapExclusion {
  return {
    id,
    wrap: 'topAndBottom',
    bounds,
    polygon: [
      { xPt: bounds.xPt, yPt: bounds.yPt },
      { xPt: bounds.xPt + bounds.widthPt, yPt: bounds.yPt },
      { xPt: bounds.xPt + bounds.widthPt, yPt: bounds.yPt + bounds.heightPt },
      { xPt: bounds.xPt, yPt: bounds.yPt + bounds.heightPt },
    ],
    ...(anchorOccurrenceId ? { anchorOccurrenceId, verticalOwnership: 'page' as const } : {}),
  };
}

const cases: ReadonlyArray<readonly [string, ParagraphPlacement | undefined, InlineResourceLayout | undefined, string | undefined]> = [
  ['text', { kind: 'text', text: 'a', range: { start: 0, end: 1 }, origin: { xPt: 10, yPt: 30 }, bounds, advancePt: 6, clusters: [{ range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, advancePt: 6 }], paintOps: [{ text: 'a', range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, letterSpacingPt: 0, scaleX: 1, direction: 'ltr', kerning: 'auto', writingMode: 'horizontal-tb' }], color: { kind: 'explicit', color: '#000' }, fontRoute, fontSizePt: 10, fontWeight: 400, fontStyle: 'normal', direction: 'ltr', decorations: [] }, undefined, undefined],
  ['anchorHost', { kind: 'anchor-host', range: { start: 0, end: 0 }, bounds, baselinePt: 30, sourceMetrics: { ascentPt: 8, descentPt: 2 } }, undefined, undefined],
  ['image', { kind: 'resource', range: { start: 0, end: 1 }, resourceKey: 'image:body:0', resourceKind: 'image', bounds, advancePt: 12 }, { kind: 'image', resourceKey: 'image:body:0', intrinsicSize: { widthPt: 12, heightPt: 10 } }, undefined],
  ['chart', { kind: 'resource', range: { start: 0, end: 1 }, resourceKey: 'chart:body:0', resourceKind: 'chart', bounds, advancePt: 12 }, { kind: 'chart', resourceKey: 'chart:body:0', intrinsicSize: { widthPt: 12, heightPt: 10 } }, undefined],
  ['break', undefined, undefined, 'column'],
  ['field', { kind: 'text', role: 'field-result', dependency: 'total-pages', text: '7', range: { start: 0, end: 1 }, origin: { xPt: 10, yPt: 30 }, bounds, advancePt: 6, clusters: [{ range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, advancePt: 6 }], paintOps: [{ text: '7', range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, letterSpacingPt: 0, scaleX: 1, direction: 'ltr', kerning: 'auto', writingMode: 'horizontal-tb' }], color: { kind: 'explicit', color: '#000' }, fontRoute, fontSizePt: 10, fontWeight: 400, fontStyle: 'normal', direction: 'ltr', decorations: [] }, undefined, undefined],
  ['shape/textbox', { kind: 'drawing', range: { start: 0, end: 1 }, drawingId: 'shape-0', bounds, advancePt: 12 }, undefined, undefined],
  ['math', { kind: 'resource', range: { start: 0, end: 1 }, resourceKey: 'math:body:0', resourceKind: 'math', bounds, advancePt: 12 }, { kind: 'math', resourceKey: 'math:body:0', intrinsicSize: { widthPt: 12, heightPt: 10 } }, undefined],
  ['ptab', { kind: 'tab', range: { start: 0, end: 1 }, leader: 'none', advancePt: 12, bounds }, undefined, undefined],
  ['picture bullet', { kind: 'resource', range: { start: -1, end: 0 }, resourceKey: 'image:bullet:0', resourceKind: 'picture-bullet', bounds, advancePt: 12 }, { kind: 'picture-bullet', resourceKey: 'image:bullet:0', intrinsicSize: { widthPt: 12, heightPt: 10 } }, undefined],
];

describe('paragraph run resource projection', () => {
  it('retains inline geometry and surrounding text without registering an unavailable resource', () => {
    const node = acquireRuns([
      sameParagraphTextRun('A'),
      {
        type: 'unavailableDrawing',
        resourceKind: 'image',
        widthPt: 24,
        heightPt: 12,
      },
      sameParagraphTextRun('B'),
    ], 4, 'unavailable-inline');

    const placements = node.lines.flatMap((line) => line.placements);
    const drawingPlacement = placements.find((placement) => placement.kind === 'drawing');
    const textPlacements = placements.filter((placement) => placement.kind === 'text');

    expect(drawingPlacement).toMatchObject({
      kind: 'drawing',
      bounds: { widthPt: 24, heightPt: 12 },
      advancePt: 24,
    });
    expect(textPlacements.map((placement) => placement.kind === 'text' ? placement.text : ''))
      .toEqual(['A', 'B']);
    expect(textPlacements[1]?.origin.xPt).toBe(39);
    expect(node.resources).toEqual([]);
    expect(node.drawings).toHaveLength(1);
    expect(node.drawings[0]).toMatchObject({
      commands: [{ kind: 'noop' }],
      diagnostics: [{
        code: 'MISSING_RESOURCE',
        severity: 'warning',
        source: { story: 'body', storyInstance: 'body', path: [4, 1] },
      }],
    });
    expect(Object.isFrozen(node.drawings[0]?.diagnostics)).toBe(true);
    expect(JSON.stringify(node)).not.toContain('imagePath');
  });

  it('retains anchored placement and wrap geometry when the drawing resource is unavailable', () => {
    const occurrenceId = 'unavailable-anchor';
    const anchor = sameParagraphAnchor(occurrenceId, {
      kind: 'square',
      authoredKinds: ['wrapSquare'],
      side: 'bothSides',
      distances: validEdges(5, 6, 7, 8),
      effectExtent: validEdges(2, 3, 4, 1),
      polygon: null,
    });
    const node = acquireRuns([
      {
        type: 'anchorHost',
        fontSize: 10,
        fontFamily: 'Test Sans',
        __anchorOccurrenceId: occurrenceId,
      },
      {
        type: 'unavailableDrawing',
        resourceKind: 'image',
        widthPt: 30,
        heightPt: 20,
        __anchorAcquisition: anchor,
      },
      sameParagraphTextRun('abcdefghijklmnopqrstuvwxyzabcdefghi'),
    ], 5, 'unavailable-anchor');

    expect(node.resources).toEqual([]);
    expect(node.drawings).toHaveLength(1);
    expect(node.drawings[0]).toMatchObject({
      flowBounds: { xPt: 92, yPt: 10, widthPt: 30, heightPt: 20 },
      commands: [{ kind: 'noop' }],
      diagnostics: [{
        code: 'MISSING_RESOURCE',
        severity: 'warning',
        source: { story: 'body', storyInstance: 'body', path: [5, 1] },
      }],
      anchorLayer: {
        occurrenceId: 'anchor:body:body:5:unavailable-anchor',
      },
    });
    expect(node.exclusions[0]).toMatchObject({
      anchorOccurrenceId: 'anchor:body:body:5:unavailable-anchor',
      bounds: { xPt: 83, yPt: 3, widthPt: 48, heightPt: 38 },
    });
    expect(node.lines.length).toBeGreaterThan(1);
  });

  it('retains upright physical orientation for an anchored image in a vertical page frame', () => {
    const anchor = sameParagraphAnchor(sameParagraphOccurrenceId, {
      kind: 'none', authoredKinds: ['wrapNone'], side: 'bothSides',
      distances: validEdges(0, 0, 0, 0), effectExtent: null, polygon: null,
    });

    const vertical = acquireSameParagraph(anchor, { verticalPageFrame: true });
    const horizontal = acquireSameParagraph(anchor);

    // `word-vertical-section-flow-anchor-frames`: the column positionH
    // starts at the physical image of the paragraph's block-start edge
    // (x = 300 - 10), so the 82 pt offset lands at physical x = 372, i.e.
    // logical y = 300 - 372 - 30.
    expect(vertical.drawings[0]).toMatchObject({
      orientation: 'upright-physical',
      flowBounds: { xPt: 10, yPt: -102, widthPt: 20, heightPt: 30 },
      transform: { a: 0, b: -1, c: 1, d: 0, e: 20, f: -87 },
    });
    expect(vertical.drawings[0]?.commands[0]).toMatchObject({
      kind: 'resource', resourceKind: 'image',
      rect: { xPt: -15, yPt: -10, widthPt: 30, heightPt: 20 },
    });
    expect(vertical.drawings[0]?.commands[0]).not.toHaveProperty('orientation');
    // Both physical references follow the anchor paragraph in a vertical
    // frame: paragraph positionV on the logical inline axis and the column
    // positionH, which starts at the paragraph, on the logical block axis.
    expect(vertical.drawings[0]?.anchorLayer).toMatchObject({
      horizontalOwnership: 'host',
      verticalOwnership: 'host',
    });
    expect(horizontal.drawings[0]?.anchorLayer).toMatchObject({
      horizontalOwnership: 'page',
      verticalOwnership: 'host',
    });
    expect(horizontal.drawings[0]?.commands[0]).not.toHaveProperty('orientation');
  });

  it('reflows a host paragraph around its parser-owned square anchor to a fixed point', () => {
    const missingEdges = validEdges(0, 0, 0, 0);
    const square = sameParagraphAnchor(sameParagraphOccurrenceId, {
      kind: 'square', authoredKinds: ['wrapSquare'], side: 'bothSides',
      distances: validEdges(5, 6, 7, 8),
      effectExtent: validEdges(2, 3, 4, 1), polygon: null,
    });
    const none = sameParagraphAnchor(sameParagraphOccurrenceId, {
      kind: 'none', authoredKinds: ['wrapNone'], side: 'bothSides',
      distances: missingEdges, effectExtent: null, polygon: null,
    });
    const wrapped = acquireSameParagraph(square);
    const unwrapped = acquireSameParagraph(none);

    expect(wrapped.lines[0]?.bounds.yPt).toBeGreaterThan(unwrapped.lines[0]!.bounds.yPt);
    expect(wrapped.advancePt).toBeGreaterThan(unwrapped.advancePt);
    expect(wrapped.drawings).toHaveLength(1);
    expect(wrapped.exclusions).toHaveLength(1);
    expect(wrapped.drawings[0]?.anchorLayer?.occurrenceId).toBe(sameParagraphOccurrenceId);
    expect(wrapped.exclusions[0]).toMatchObject({
      anchorOccurrenceId: sameParagraphOccurrenceId,
      bounds: { xPt: 83, yPt: 3, widthPt: 48, heightPt: 38 },
    });
    expect(new Set(wrapped.drawings.map((drawing) =>
      drawing.anchorLayer?.occurrenceId))).toEqual(new Set([sameParagraphOccurrenceId]));
    expect(new Set(wrapped.exclusions.map((exclusion) => exclusion.anchorOccurrenceId)))
      .toEqual(new Set([sameParagraphOccurrenceId]));
  });

  it('remeasures square wrapping with the authored left, right, and largest side', () => {
    const acquireSide = (side: 'left' | 'right' | 'largest') => acquireSameParagraph(
      sameParagraphAnchor(sameParagraphOccurrenceId, {
        kind: 'square', authoredKinds: ['wrapSquare'], side,
        distances: validEdges(5, 6, 7, 8),
        effectExtent: validEdges(2, 3, 4, 1), polygon: null,
      }),
    );

    const left = acquireSide('left');
    const right = acquireSide('right');
    const largest = acquireSide('largest');

    expect(left.exclusions[0]).toMatchObject({ wrapSide: 'left' });
    expect(right.exclusions[0]).toMatchObject({ wrapSide: 'right' });
    expect(largest.exclusions[0]).toMatchObject({ wrapSide: 'largest' });
    expect(left.lines[0]?.bounds.yPt).toBe(41);
    expect(largest.lines.map((line) => line.bounds)).toEqual(left.lines.map((line) => line.bounds));
    expect(right.lines[0]?.bounds.yPt).toBeGreaterThanOrEqual(41);
  });

  it('rejects conflicting placement and retained exclusion wrap authorities', () => {
    const square = sameParagraphAnchor(sameParagraphOccurrenceId, {
      kind: 'square', authoredKinds: ['wrapSquare'], side: 'bothSides',
      distances: validEdges(5, 6, 7, 8),
      effectExtent: validEdges(2, 3, 4, 1), polygon: null,
    });
    const passThrough: WrapOracle = {
      lineWindow: (input) => ({
        topYPt: input.topYPt,
        xOffsetPt: 0,
        maximumWidthPt: input.maximumWidthPt,
      }),
      skipTopAndBottomBands: ({ yPt }) => yPt,
    };

    expect(() => acquireSameParagraph(square, { wrap: passThrough }))
      .toThrow(/conflicting paragraph wrap authorities/i);
  });

  it('measures and retains externally authoritative geometry for a matching occurrence ID', () => {
    const square = sameParagraphAnchor(sameParagraphOccurrenceId, {
      kind: 'square', authoredKinds: ['wrapSquare'], side: 'bothSides',
      distances: validEdges(0, 0, 0, 0), effectExtent: null, polygon: null,
    });
    const external = externalExclusion(
      'external-matching-occurrence',
      { xPt: 10, yPt: 10, widthPt: 100, heightPt: 50 },
      sameParagraphOccurrenceId,
    );

    const acquired = acquireSameParagraph(square, { exclusions: [external] });

    expect(acquired.lines[0]?.bounds.yPt).toBeGreaterThanOrEqual(60);
    expect(acquired.exclusions).toEqual([external]);
  });

  it('preserves externally authoritative page geometry in continuation slices', () => {
    const square = sameParagraphAnchor(sameParagraphOccurrenceId, {
      kind: 'square', authoredKinds: ['wrapSquare'], side: 'bothSides',
      distances: validEdges(0, 0, 0, 0), effectExtent: null, polygon: null,
    });
    const external = externalExclusion(
      'external-continuation-occurrence',
      { xPt: 10, yPt: 10, widthPt: 100, heightPt: 50 },
      sameParagraphOccurrenceId,
    );
    const acquired = acquireSameParagraph(square, { exclusions: [external] });

    const continuation = sliceParagraphLayout(acquired, {
      lineStart: 1,
      lineEnd: acquired.lines.length,
      continuesFromPrevious: true,
      continuesOnNext: false,
    });

    expect(continuation.drawings).toEqual([]);
    expect(continuation.exclusions).toEqual([external]);
  });

  it('rejects duplicate externally authoritative occurrence IDs', () => {
    const square = sameParagraphAnchor(sameParagraphOccurrenceId, {
      kind: 'square', authoredKinds: ['wrapSquare'], side: 'bothSides',
      distances: validEdges(0, 0, 0, 0), effectExtent: null, polygon: null,
    });
    const first = externalExclusion(
      'external-first', { xPt: 10, yPt: 200, widthPt: 10, heightPt: 10 },
      sameParagraphOccurrenceId,
    );
    const duplicate = externalExclusion(
      'external-duplicate', { xPt: 30, yPt: 200, widthPt: 10, heightPt: 10 },
      sameParagraphOccurrenceId,
    );

    expect(() => acquireSameParagraph(square, { exclusions: [first, duplicate] }))
      .toThrow(/duplicate external paragraph exclusion occurrence.*same-paragraph/i);
  });

  it('preserves mixed keyed and unkeyed external exclusion order', () => {
    const square = sameParagraphAnchor(sameParagraphOccurrenceId, {
      kind: 'square', authoredKinds: ['wrapSquare'], side: 'bothSides',
      distances: validEdges(0, 0, 0, 0), effectExtent: null, polygon: null,
    });
    const exclusions = [
      externalExclusion('unkeyed-before', { xPt: 0, yPt: 200, widthPt: 5, heightPt: 5 }),
      externalExclusion('keyed-matching', { xPt: 10, yPt: 200, widthPt: 5, heightPt: 5 }, sameParagraphOccurrenceId),
      externalExclusion('unkeyed-after', { xPt: 20, yPt: 200, widthPt: 5, heightPt: 5 }),
      externalExclusion('keyed-other', { xPt: 30, yPt: 200, widthPt: 5, heightPt: 5 }, 'external:other'),
    ];

    const acquired = acquireSameParagraph(square, { exclusions });

    expect(acquired.exclusions).toEqual(exclusions);
  });

  it.each(cases)('retains the %s arm without a DocRun reference', (_name, placement, resource, event) => {
    const input: AcquiredParagraphLayoutInput = {
      kind: 'paragraph', id: `p-${_name}`, source, flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 20 },
      inkBounds: { xPt: 10, yPt: 20, widthPt: 12, heightPt: 10 },
      spacing: { beforePt: 0, afterPt: 0 },
      lines: placement ? [{ range: { start: 0, end: 1 }, bounds, baselinePt: 30, advancePt: 20, placements: [placement] }] : [],
      borders: [], resources: resource ? [resource] : [], drawings: [], textBoxes: [], exclusions: [],
      events: event ? [{ kind: 'break', breakKind: event as 'line' | 'page' | 'column', offset: 0 }] : [],
    };
    const node = layoutParagraph(input);

    expect(node.lines[0]?.placements[0]).toEqual(placement);
    expect(node.resources[0]).toEqual(resource);
    expect(node.events[0]).toMatchObject(event ? { kind: 'break', breakKind: event } : {});
    expect(node).not.toHaveProperty('paragraph');
    expect(node).not.toHaveProperty('runs');
  });

  it('acquires every actual DocRun arm through the shared line engine before finalization', () => {
    const text = (value: string): DocRun => ({
      type: 'text', text: value, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Test Sans',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocxTextRun & { type: 'text' });
    const runs: DocRun[] = [
      text('A'),
      { type: 'anchorHost', fontSize: 10, fontFamily: 'Test Sans' },
      { type: 'image', imagePath: 'word/media/a.png', mimeType: 'image/png', widthPt: 12, heightPt: 10 },
      { type: 'chart', chart: { title: 'chart' } as never, widthPt: 20, heightPt: 12, anchor: false },
      { type: 'break', breakType: 'line' },
      { type: 'break', breakType: 'page' },
      { type: 'break', breakType: 'column' },
      { type: 'field', fieldType: 'numPages', instruction: 'NUMPAGES', fallbackText: '9', bold: false,
        italic: false, underline: false, strikethrough: false, fontSize: 10, color: null,
        fontFamily: 'Test Sans', background: null, vertAlign: null },
      { type: 'shape', widthPt: 30, heightPt: 20, anchorXPt: 4, anchorYPt: 5,
        anchorXFromMargin: false, anchorYFromPara: true, zOrder: 1, subpaths: [], fill: null,
        stroke: null, textBlocks: [{ text: 'box', fontSizePt: 10, alignment: 'left' }] },
      { type: 'math', nodes: [], display: false, fontSize: 10, resourceKey: 'math:body:0:9' } as unknown as DocRun,
      { type: 'ptab', alignment: 'right', relativeTo: 'margin', leader: 'dot', fontSize: 10 },
    ];
    const paragraph: DocParagraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 3, spaceAfter: 4, lineSpacing: null, tabStops: [], runs,
      numbering: {
        numId: 1, level: 0, format: 'bullet', text: '', indentLeft: 0, tab: 18, suff: 'tab',
        picBulletImagePath: 'word/media/bullet.png', picBulletMimeType: 'image/png',
        picBulletWidthPt: 6, picBulletHeightPt: 6,
      },
      shading: 'ffeeaa',
      borders: {
        top: { style: 'single', color: '111111', width: 1, space: 0 },
        right: null,
        bottom: { style: 'double', color: '222222', width: 2, space: 0 },
        left: null,
        between: null,
      },
    };
    const context: ParagraphLayoutContext = {
      lineGrid: { active: false, pitchPt: null }, characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
      rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
      physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
      lineSpacing: null, spaceBeforePt: 3, spaceAfterPt: 4, baseRtl: false,
      isJustified: false, stretchLastLine: false, tabStops: [], hasRuby: false,
      hasEastAsianText: false, kinsoku: DEFAULT_KINSOKU_RULES, defaultTabPt: 36,
    };
    let font = '10px Test Sans';
    const measureContext = {
      get font() { return font; }, set font(value: string) { font = value; },
      letterSpacing: '0px', fontKerning: 'auto',
      measureText(value: string) {
        return { width: [...value].length * 5, actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const services: LayoutServices = {
      text: {
        fingerprint: 'text',
        localMetrics: {},
        resolve(request) {
          const requestedFamily = request.fonts.ascii ?? request.fonts.highAnsi ?? 'Test Sans';
          return {
            requestedFamily, resolvedFamily: requestedFamily, route: fontRoute,
            source: 'native', weight: request.weight ?? 400, style: request.style ?? 'normal',
            diagnostics: [], genericFamily: request.genericFamily ?? 'sans-serif',
          };
        },
        shape(request) {
          let clusterOffset = 0;
          const clusters = [...request.text].map((character) => {
            const start = clusterOffset;
            clusterOffset += character.length;
            return { range: { start, end: clusterOffset }, offsetPt: start * 5, advancePt: character.length * 5 };
          });
          return { text: request.text, spans: [{
            text: request.text, start: 0, end: request.text.length, script: 'ascii', breakBefore: true,
            font: { requestedFamily: 'Test Sans', resolvedFamily: 'Test Sans', route: fontRoute,
              source: 'native', weight: request.weight ?? 400, style: request.style ?? 'normal',
              diagnostics: [], genericFamily: 'sans-serif' },
            fontRoute, advancePt: request.text.length * 5,
            ascentPt: 8, descentPt: 2,
          }], advancePt: request.text.length * 5, ascentPt: 8, descentPt: 2,
            graphemeBoundaries: [0, request.text.length], clusters, diagnostics: [] };
        },
      },
      images: { fingerprint: 'images', resolve: () => ({ widthPt: 12, heightPt: 10, mimeType: 'image/png' }) },
      math: { fingerprint: 'math', resolve: (resourceKey) => ({ resourceKey, widthEm: 1, ascentEm: .8, descentEm: .2, diagnostics: [] }) },
    };

    const snapshot = paragraphAcquisitionInput(paragraph, source);
    paragraph.shading = '000000';
    (paragraph.runs[0] as Extract<DocRun, { type: 'text' }>).text = 'mutated';
    const node = acquireParagraphLayout(snapshot, {
      id: 'all-runs', source, flowDomainId: 'body', ordinaryFlow: true,
      context,
      placement: { startYPt: 10, paragraphXPt: 10, availableWidthPt: 300, maximumYPt: 700, suppressSpaceBefore: false },
      measurer: { context: measureContext, fontFamilyClasses: {} },
      environment: {
        pageIndex: 2, totalPages: 9, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: false, layoutServices: services,
      },
      exclusions: [],
    });

    const placements = node.lines.flatMap((line) => line.placements);
    expect(placements.map((placement) => placement.kind)).toEqual(expect.arrayContaining([
      'text', 'anchor-host', 'resource', 'tab', 'drawing',
    ]));
    expect(node.resources.map((resource) => resource.kind)).toEqual(expect.arrayContaining([
      'image', 'chart', 'math', 'picture-bullet',
    ]));
    expect(placements).toEqual(expect.arrayContaining([
      expect.objectContaining({ resourceKind: 'image', sourceRunIndex: 2 }),
      expect.objectContaining({ resourceKind: 'chart', sourceRunIndex: 3 }),
      expect.objectContaining({ resourceKind: 'math', sourceRunIndex: 9 }),
    ]));
    expect(node.events).toEqual(expect.arrayContaining([
      expect.objectContaining({ breakKind: 'line' }),
      expect.objectContaining({ breakKind: 'page' }),
      expect.objectContaining({ breakKind: 'column' }),
    ]));
    expect(placements).toEqual(expect.arrayContaining([
      expect.objectContaining({ role: 'field-result', dependency: 'total-pages', text: '9' }),
    ]));
    const textBoxParagraph = node.textBoxes[0]?.story.blocks[0];
    expect(textBoxParagraph?.kind).toBe('paragraph');
    expect(textBoxParagraph?.kind === 'paragraph'
      ? textBoxParagraph.lines[0]?.placements[0] : undefined).toMatchObject({ text: 'box' });
    expect(node.textBoxes[0]?.insets).toEqual({
      topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0,
    });
    expect(node.textBoxes[0]?.flowBounds).toEqual(node.drawings[0]?.flowBounds);
    expect(textBoxParagraph?.flowBounds).toMatchObject({
      xPt: 4, yPt: 15, widthPt: 30,
    });
    expect(JSON.stringify(node)).not.toContain('imagePath');
    expect(JSON.stringify(node)).not.toContain('fallbackText');
    expect(node.resources.every((resource) => !('data' in resource))).toBe(true);
    expect(node.lines.flatMap((line) => line.placements)).toEqual(expect.arrayContaining([
      expect.objectContaining({ resourceKind: 'picture-bullet', resourceKey: expect.stringContaining('bullet.png') }),
    ]));
    expect(node.shading).toEqual({ color: '#ffeeaa' });
    expect(node.borders.map((border) => border.edge)).toEqual(['top', 'bottom']);
    expect(node.lines[0]?.placements).toEqual(expect.arrayContaining([
      expect.objectContaining({ text: 'A' }),
    ]));
    expect(snapshot).not.toBe(paragraph);
    expect(snapshot.numberingMarkerShapeInput).toMatchObject({ fontSizePt: expect.any(Number) });
    expect(snapshot.runs.every((run, index) => run !== paragraph.runs[index])).toBe(true);
    expect(Object.isFrozen(snapshot)).toBe(true);

    const markParagraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [], runs: [],
      defaultFontSize: 14, markVanish: true,
      paragraphMarkFontFacts: { bold: true },
    } as unknown as DocParagraph;
    const markSnapshot = paragraphAcquisitionInput(markParagraph, { ...source, path: [1] });
    const markNode = acquireParagraphLayout(markSnapshot, {
      id: 'hidden-mark', source: { ...source, path: [1] }, flowDomainId: 'body', ordinaryFlow: true,
      context: { ...context, spaceBeforePt: 0, spaceAfterPt: 0 },
      placement: { startYPt: 10, paragraphXPt: 10, availableWidthPt: 100, maximumYPt: 100, suppressSpaceBefore: false },
      measurer: { context: measureContext, fontFamilyClasses: {} },
      environment: {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: false, layoutServices: services,
      },
      exclusions: [],
    });
    expect(markSnapshot.paragraphMarkShapeInput?.fontSizePt).toBe(14);
    expect(markNode.paragraphMark?.hidden).toBe(true);
    expect(markNode.advancePt).toBeGreaterThan(0);

    const wrapParagraph: DocParagraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [], runs: [text('wrap')],
    };
    const wrapNode = acquireParagraphLayout(paragraphAcquisitionInput(wrapParagraph, { ...source, path: [2] }), {
      id: 'wrapped', source: { ...source, path: [2] }, flowDomainId: 'body', ordinaryFlow: true,
      context: { ...context, spaceBeforePt: 0, spaceAfterPt: 0 },
      placement: { startYPt: 10, paragraphXPt: 10, availableWidthPt: 100, maximumYPt: 100, suppressSpaceBefore: false },
      measurer: { context: measureContext, fontFamilyClasses: {} },
      environment: {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: false, layoutServices: services,
      },
      exclusions: [{
        id: 'float', wrap: 'square', bounds: { xPt: 10, yPt: 10, widthPt: 85, heightPt: 25 },
        polygon: [{ xPt: 10, yPt: 10 }, { xPt: 95, yPt: 10 }, { xPt: 95, yPt: 35 }, { xPt: 10, yPt: 35 }],
      }],
    });
    expect(wrapNode.lines[0]?.bounds.yPt).toBeGreaterThanOrEqual(35);
    const wrapLastLine = wrapNode.lines.at(-1)!;
    expect(wrapNode.advancePt).toBe(
      wrapLastLine.bounds.yPt + wrapLastLine.advancePt - wrapNode.flowBounds.yPt,
    );

    const acquireVariant = (
      variant: DocParagraph,
      variantContext: ParagraphLayoutContext,
      id: string,
      widthPt = 100,
      exclusions: Parameters<typeof acquireParagraphLayout>[1]['exclusions'] = [],
    ) => acquireParagraphLayout(paragraphAcquisitionInput(variant, { ...source, path: [3] }), {
      id, source: { ...source, path: [3] }, flowDomainId: 'body', ordinaryFlow: true,
      context: variantContext,
      placement: { startYPt: 10, paragraphXPt: 10, availableWidthPt: widthPt, maximumYPt: 300, suppressSpaceBefore: false },
      measurer: { context: measureContext, fontFamilyClasses: {} },
      environment: {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: true, layoutServices: services,
      },
      exclusions,
    });

    const bidiRun = text('مرحبا\t12');
    Object.assign(bidiRun, { rtl: true, langBidi: 'ar-SA' });
    const mixedBidi = acquireVariant({
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
      runs: [bidiRun, text(' LTR')],
    }, context, 'mixed-bidi');
    const mixedPlacements = mixedBidi.lines[0]?.placements ?? [];
    expect(mixedPlacements.some((placement) => placement.kind === 'tab')).toBe(true);
    expect(mixedPlacements.filter((placement) => placement.kind === 'text').map((placement) => placement.range))
      .toEqual(expect.arrayContaining([
        { start: 0, end: 5 }, { start: 6, end: 8 }, { start: 8, end: 9 }, { start: 9, end: 12 },
      ]));

    const distributed = acquireVariant({
      alignment: 'distribute', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
      runs: [text('観察：［')],
    }, { ...context, isJustified: true, stretchLastLine: true, hasEastAsianText: true }, 'cjk', 80);
    const cjk = distributed.lines[0]?.placements.find((placement) => placement.kind === 'text');
    expect(cjk).toMatchObject({
      clusters: expect.arrayContaining([expect.objectContaining({ range: { start: 0, end: 1 } })]),
      paintOps: [expect.objectContaining({ text: '観察：［' })],
    });
    const retainedPaintCalls: string[] = [];
    const retainedPaintContext = {
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: 'left',
      textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
      save() {}, restore() {}, translate() {}, scale() {}, setLineDash() {}, fillRect() {}, strokeRect() {},
      beginPath() {}, moveTo() {}, lineTo() {}, stroke() {},
      fillText(value: string) { retainedPaintCalls.push(value); },
      measureText() { throw new Error('retained paint must not measure'); },
    } as unknown as CanvasRenderingContext2D;
    expect(() => paintParagraphLayout(distributed, {
      ctx: retainedPaintContext, scale: 2, dpr: 2, resources: noPaintResources,
    })).not.toThrow();
    expect(retainedPaintCalls).toContain('観察：［');

    const decimal = acquireVariant({
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
      tabStops: [{ pos: 50, alignment: 'decimal', leader: 'none' }], runs: [text('123')],
    }, { ...context, tabStops: [{ pos: 50, alignment: 'decimal', leader: 'none' }] }, 'decimal');
    expect(decimal.lines[0]?.placements[0]).toMatchObject({ origin: { xPt: 45, yPt: expect.any(Number) } });

    const numberedText = acquireVariant({
      alignment: 'left', indentLeft: 12, indentRight: 0, indentFirst: -12,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, tabStops: [], runs: [text('body')],
      numbering: {
        numId: 1, level: 0, format: 'decimal', text: '1.', indentLeft: 12,
        tab: 12, suff: 'tab', jc: 'left', fontFamily: 'Test Sans',
      },
    }, { ...context, physicalIndentLeftPt: 12, firstIndentPt: -12 }, 'numbered-text');
    expect(numberedText.lines[0]?.placements).toEqual([
      expect.objectContaining({
        kind: 'text', role: 'numbering-marker', text: '1.', range: { start: -2, end: 0 },
        origin: { xPt: 10, yPt: expect.any(Number) },
      }),
      expect.objectContaining({
        kind: 'text', text: 'body', range: { start: 0, end: 4 },
        origin: { xPt: 22, yPt: expect.any(Number) },
      }),
    ]);

    const math = acquireVariant({
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
      runs: [{ type: 'math', nodes: [], display: true, fontSize: 10, jc: 'right', resourceKey: 'math:right' } as unknown as DocRun],
    }, context, 'math-right');
    expect(math.lines[0]?.placements[0]).toMatchObject({
      kind: 'resource', resourceKey: 'math:right', bounds: { xPt: 100 },
    });
  });
});
