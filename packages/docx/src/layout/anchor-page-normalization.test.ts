import { describe, expect, it } from 'vitest';
import { resolveAnchorFrame } from './anchor-frame.js';
import type { AnchorAcquisitionInput } from './anchor-input.js';
import {
  canonicalLogicalToPhysical,
  composeAffine,
  mapAffineRect,
  quarterTurnAffine,
  translationAffine,
} from './affine.js';
import {
  normalizeAnchorReferenceFrames,
  normalizePagePaintNodeAnchors,
} from './anchor-page-normalization.js';
import type { ParagraphLayout, TableLayout, TextBoxLayout } from './types.js';

const missingEdges = {
  topPt: null, topStatus: 'missing', rightPt: null, rightStatus: 'missing',
  bottomPt: null, bottomStatus: 'missing', leftPt: null, leftStatus: 'missing',
} as const;

function acquisition(
  choice: 'offset' | 'align' | 'percent' | 'relative-size',
): AnchorAcquisitionInput {
  const horizontal = choice === 'align'
    ? { relativeFrom: 'character', relativeFromStatus: 'valid' as const, choice: { kind: 'align' as const, value: 'right' } }
    : choice === 'percent'
      ? { relativeFrom: 'character', relativeFromStatus: 'valid' as const, choice: { kind: 'percent' as const, fraction: 0.25 } }
      : { relativeFrom: 'character', relativeFromStatus: 'valid' as const, choice: { kind: 'offset' as const, valuePt: 2 } };
  const vertical = choice === 'align'
    ? { relativeFrom: 'line', relativeFromStatus: 'valid' as const, choice: { kind: 'align' as const, value: 'bottom' } }
    : choice === 'percent'
      ? { relativeFrom: 'line', relativeFromStatus: 'valid' as const, choice: { kind: 'percent' as const, fraction: 0.4 } }
      : { relativeFrom: 'line', relativeFromStatus: 'valid' as const, choice: { kind: 'offset' as const, valuePt: 3 } };
  return {
    occurrenceId: `anchor:${choice}`,
    simplePosition: { enabled: false, status: 'valid', xPt: 0, xStatus: 'valid', yPt: 0, yStatus: 'valid' },
    horizontal, vertical,
    extent: { widthPt: 20, widthStatus: 'valid', heightPt: 10, heightStatus: 'valid' },
    parentEffectExtent: missingEdges, anchorDistances: missingEdges,
    relativeSize: choice === 'relative-size' ? {
      horizontal: { relativeFrom: 'character', relativeFromStatus: 'valid', fraction: 0.5, fractionStatus: 'valid' },
      vertical: { relativeFrom: 'line', relativeFromStatus: 'valid', fraction: 0.5, fractionStatus: 'valid' },
    } : { horizontal: null, vertical: null },
    wrap: { kind: 'none', authoredKinds: [], side: null, distances: missingEdges, effectExtent: null, polygon: null },
    behavior: {
      behindDoc: false, behindDocStatus: 'valid', relativeHeight: 1, relativeHeightStatus: 'valid',
      locked: false, lockedStatus: 'valid', allowOverlap: true, allowOverlapStatus: 'valid',
      layoutInCell: true, layoutInCellStatus: 'valid',
    },
    group: null,
  };
}

describe('page anchor reference-frame normalization', () => {
  it.each(['vertical-rl', 'vertical-lr'] as const)(
    'replays offset, align, percent, and relative size against physical %s host extents',
    (writingMode) => {
      const matrix = canonicalLogicalToPhysical(writingMode, 200);
      const logicalHostFrames = {
        paragraph: { xPt: 11, yPt: 23, widthPt: 70, heightPt: 31 },
        line: { xPt: 17, yPt: 29, widthPt: 53, heightPt: 13 },
        character: { xPt: 19, yPt: 31, widthPt: 7, heightPt: 11 },
      };
      const destinationFrames = {
        page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
        margin: { xPt: 10, yPt: 20, widthPt: 180, heightPt: 260 },
        column: { xPt: 100, yPt: 0, widthPt: 100, heightPt: 300 },
        pageParity: 'even' as const,
      };
      for (const choice of ['offset', 'align', 'percent', 'relative-size'] as const) {
        const authored = acquisition(choice);
        const normalized = normalizeAnchorReferenceFrames({
          acquisition: authored,
          pageParity: 'odd',
          physicalFrames: destinationFrames,
          logicalHostFrames,
        }, matrix, destinationFrames);
        const expectedFrames = {
          ...destinationFrames,
          paragraph: mapAffineRect(matrix, logicalHostFrames.paragraph),
          line: mapAffineRect(matrix, logicalHostFrames.line),
          character: mapAffineRect(matrix, logicalHostFrames.character),
        };

        expect(resolveAnchorFrame({ acquisition: authored, frames: normalized }))
          .toEqual(resolveAnchorFrame({ acquisition: authored, frames: expectedFrames }));
      }
    },
  );

  it('normalizes acquired anchors below table placement and a turned text-box frame', () => {
    const authored = acquisition('offset');
    const logicalHostFrames = {
      paragraph: { xPt: 0, yPt: 0, widthPt: 50, heightPt: 20 },
      line: { xPt: 1, yPt: 3, widthPt: 30, heightPt: 9 },
      character: { xPt: 2, yPt: 4, widthPt: 5, heightPt: 7 },
    };
    const physicalFrames = {
      page: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
      margin: { xPt: 10, yPt: 20, widthPt: 180, heightPt: 260 },
      column: { xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 },
    };
    const drawingBounds = { xPt: 4, yPt: 6, widthPt: 20, heightPt: 10 };
    const paragraph: ParagraphLayout = {
      kind: 'paragraph', id: 'nested-paragraph',
      source: { story: 'body', storyInstance: 'body', path: [0] },
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: { xPt: 0, yPt: 0, widthPt: 50, heightPt: 20 },
      inkBounds: { xPt: 0, yPt: 0, widthPt: 50, heightPt: 20 },
      advancePt: 20, spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
      lines: [], borders: [], resources: [], textBoxes: [], events: [], exclusions: [],
      drawings: [{
        kind: 'drawing', id: 'nested-anchor', source: { story: 'body', storyInstance: 'body', path: [0, 1] },
        flowDomainId: 'body', ordinaryFlow: false, flowBounds: drawingBounds, inkBounds: drawingBounds,
        advancePt: 0, commands: [{ kind: 'fill-rect', rect: drawingBounds, fill: '#000000' }],
        anchorLayer: {
          occurrenceId: 'anchor:offset', behindDoc: false, relativeHeight: 1, sourceOrder: 0,
          horizontalOwnership: 'host', verticalOwnership: 'host',
          coordinateSpace: 'acquired-anchor-points',
          normalization: {
            acquisition: authored, pageParity: 'odd', physicalFrames, logicalHostFrames,
          },
        },
      }],
    };
    const tableBounds = { xPt: 40, yPt: 50, widthPt: 100, heightPt: 40 };
    const table: TableLayout = {
      kind: 'table', id: 'nested-table', source: paragraph.source,
      flowDomainId: 'body', ordinaryFlow: true, flowBounds: tableBounds, inkBounds: tableBounds,
      advancePt: 40, columnWidthsPt: [100], borders: [], rows: [{
        kind: 'table-row', id: 'nested-row', source: paragraph.source,
        flowDomainId: 'body', ordinaryFlow: true, flowBounds: tableBounds, inkBounds: tableBounds,
        advancePt: 40, heightPt: 40, contentHeightPt: 20, cells: [{
          kind: 'table-cell', id: 'nested-cell', source: paragraph.source,
          flowDomainId: 'body', ordinaryFlow: true, flowBounds: tableBounds, inkBounds: tableBounds,
          contentBounds: { xPt: 42, yPt: 52, widthPt: 96, heightPt: 36 },
          advancePt: 40, verticalMerge: 'none', vAlign: 'top',
          blocks: [{ layout: paragraph, offsetPt: 3, advancePt: 20 }],
        }],
      }],
    };
    const textBox: TextBoxLayout = {
      kind: 'textbox', id: 'turned-box', source: paragraph.source,
      flowDomainId: 'body', ordinaryFlow: false,
      flowBounds: { xPt: 70, yPt: 90, widthPt: 40, heightPt: 20 },
      inkBounds: { xPt: 70, yPt: 90, widthPt: 40, heightPt: 20 },
      advancePt: 0, paragraphs: [paragraph], writingMode: 'vertical-rl', verticalMode: 'vert',
      insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
    };
    const region = canonicalLogicalToPhysical('vertical-rl', 200);
    const context = {
      currentToPage: region,
      normalizedFor: { physicalPageIndex: 1, flowDomainId: 'body', regionId: 'vertical' },
      destinationFrames: { ...physicalFrames, pageParity: 'even' as const },
    };
    const normalizedTable = normalizePagePaintNodeAnchors(table, context);
    const tableDrawing = normalizedTable.rows[0]!.cells[0]!.blocks[0]!.layout;
    if (tableDrawing.kind !== 'paragraph') throw new Error('expected paragraph');
    const tableCurrent = composeAffine(region, translationAffine(42, 53));
    const tableExpected = resolveAnchorFrame({
      acquisition: authored,
      frames: normalizeAnchorReferenceFrames(
        { acquisition: authored, pageParity: 'odd', physicalFrames, logicalHostFrames },
        tableCurrent,
        context.destinationFrames,
      ),
    });
    if (tableExpected.status !== 'resolved') throw new Error('expected resolved table anchor');
    expect(tableDrawing.drawings[0]?.flowBounds).toEqual(tableExpected.geometry.objectFrame);

    const normalizedTextBox = normalizePagePaintNodeAnchors(textBox, context);
    const textBoxCurrent = composeAffine(
      region,
      composeAffine(translationAffine(90, 100), quarterTurnAffine(1)),
    );
    const textBoxExpected = resolveAnchorFrame({
      acquisition: authored,
      frames: normalizeAnchorReferenceFrames(
        { acquisition: authored, pageParity: 'odd', physicalFrames, logicalHostFrames },
        textBoxCurrent,
        context.destinationFrames,
      ),
    });
    if (textBoxExpected.status !== 'resolved') throw new Error('expected resolved textbox anchor');
    expect(normalizedTextBox.paragraphs[0]?.drawings[0]?.flowBounds)
      .toEqual(textBoxExpected.geometry.objectFrame);
  });
});
