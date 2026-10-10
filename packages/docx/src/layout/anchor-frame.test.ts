import { describe, expect, it } from 'vitest';
import type {
  AnchorAcquisitionInput,
  AnchorEdgesInput,
} from './anchor-input.js';
import {
  resolveAnchorFrame,
  type AnchorFrameInput,
} from './anchor-frame.js';
import { createFloatWrapOracle } from '../paragraph-measure.js';

const missingEdges = (): AnchorEdgesInput => ({
  topPt: null,
  topStatus: 'missing',
  rightPt: null,
  rightStatus: 'missing',
  bottomPt: null,
  bottomStatus: 'missing',
  leftPt: null,
  leftStatus: 'missing',
});

const validEdges = (
  topPt: number,
  rightPt: number,
  bottomPt: number,
  leftPt: number,
): AnchorEdgesInput => ({
  topPt,
  topStatus: 'valid',
  rightPt,
  rightStatus: 'valid',
  bottomPt,
  bottomStatus: 'valid',
  leftPt,
  leftStatus: 'valid',
});

function anchor(
  overrides: Partial<AnchorAcquisitionInput> = {},
): AnchorAcquisitionInput {
  return {
    occurrenceId: 'anchor-1',
    simplePosition: {
      enabled: false,
      status: 'valid',
      xPt: 0,
      xStatus: 'valid',
      yPt: 0,
      yStatus: 'valid',
    },
    horizontal: {
      relativeFrom: 'page',
      relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: 0 },
    },
    vertical: {
      relativeFrom: 'page',
      relativeFromStatus: 'valid',
      choice: { kind: 'offset', valuePt: 0 },
    },
    extent: {
      widthPt: 20,
      heightPt: 10,
      widthStatus: 'valid',
      heightStatus: 'valid',
    },
    parentEffectExtent: missingEdges(),
    anchorDistances: missingEdges(),
    relativeSize: { horizontal: null, vertical: null },
    wrap: {
      kind: 'square',
      authoredKinds: ['wrapSquare'],
      side: 'bothSides',
      distances: missingEdges(),
      effectExtent: null,
      polygon: null,
    },
    behavior: {
      behindDoc: false,
      behindDocStatus: 'valid',
      relativeHeight: 1,
      relativeHeightStatus: 'valid',
      locked: false,
      lockedStatus: 'valid',
      allowOverlap: true,
      allowOverlapStatus: 'valid',
      layoutInCell: true,
      layoutInCellStatus: 'valid',
    },
    group: null,
    ...overrides,
  };
}

function input(
  acquisition: AnchorAcquisitionInput = anchor(),
): AnchorFrameInput {
  return {
    acquisition,
    frames: {
      page: { xPt: 10, yPt: 20, widthPt: 600, heightPt: 800 },
      margin: { xPt: 60, yPt: 80, widthPt: 500, heightPt: 680 },
      column: { xPt: 70, yPt: 80, widthPt: 240, heightPt: 680 },
      paragraph: { xPt: 70, yPt: 120, widthPt: 240, heightPt: 60 },
      line: { xPt: 75, yPt: 135, widthPt: 220, heightPt: 14 },
      character: { xPt: 123, yPt: 135, widthPt: 7, heightPt: 14 },
      pageParity: 'odd',
    },
  };
}

function resolved(result: ReturnType<typeof resolveAnchorFrame>) {
  expect(result.status).toBe('resolved');
  if (result.status !== 'resolved') throw new Error('expected resolved anchor frame');
  return result;
}

describe('retained anchor frame geometry', () => {
  it.each([
    ['missing behindDoc', { behindDoc: null, behindDocStatus: 'missing' }, 'missing-required-behavior', 'behavior.behindDoc'],
    ['invalid behindDoc', { behindDoc: null, behindDocStatus: 'invalid' }, 'invalid-required-behavior', 'behavior.behindDoc'],
    ['missing relativeHeight', { relativeHeight: null, relativeHeightStatus: 'missing' }, 'missing-required-behavior', 'behavior.relativeHeight'],
    ['invalid relativeHeight', { relativeHeight: null, relativeHeightStatus: 'invalid' }, 'invalid-required-behavior', 'behavior.relativeHeight'],
    ['missing locked', { locked: null, lockedStatus: 'missing' }, 'missing-required-behavior', 'behavior.locked'],
    ['invalid locked', { locked: null, lockedStatus: 'invalid' }, 'invalid-required-behavior', 'behavior.locked'],
    ['missing layoutInCell', { layoutInCell: null, layoutInCellStatus: 'missing' }, 'missing-required-behavior', 'behavior.layoutInCell'],
    ['invalid layoutInCell', { layoutInCell: null, layoutInCellStatus: 'invalid' }, 'invalid-required-behavior', 'behavior.layoutInCell'],
    ['missing allowOverlap', { allowOverlap: null, allowOverlapStatus: 'missing' }, 'missing-required-behavior', 'behavior.allowOverlap'],
    ['invalid allowOverlap', { allowOverlap: null, allowOverlapStatus: 'invalid' }, 'invalid-required-behavior', 'behavior.allowOverlap'],
  ] as const)('rejects %s before producing drawable anchor geometry', (
    _name, behaviorOverride, code, path,
  ) => {
    const base = anchor();
    const result = resolveAnchorFrame(input(anchor({
      behavior: { ...base.behavior, ...behaviorOverride },
    })));

    expect(result.status).toBe('unsupported');
    expect(result.issues).toEqual([expect.objectContaining({ code, path })]);
  });

  it('uses positionH and positionV when the optional simplePos attribute is absent', () => {
    const base = anchor();
    const result = resolved(resolveAnchorFrame(input(anchor({
      simplePosition: {
        ...base.simplePosition,
        enabled: null,
        status: 'missing',
      },
    }))));

    expect(result.geometry.objectFrame).toMatchObject({ xPt: 10, yPt: 20 });
    expect(result.axes.horizontal).toMatchObject({ status: 'resolved', choiceKind: 'offset' });
    expect(result.axes.vertical).toMatchObject({ status: 'resolved', choiceKind: 'offset' });
  });

  it('resolves character and line references from their explicit frames', () => {
    const result = resolved(resolveAnchorFrame(input(anchor({
      horizontal: {
        relativeFrom: 'character',
        relativeFromStatus: 'valid',
        choice: { kind: 'offset', valuePt: 3 },
      },
      vertical: {
        relativeFrom: 'line',
        relativeFromStatus: 'valid',
        choice: { kind: 'offset', valuePt: 4 },
      },
    }))));

    expect(result.geometry.objectFrame).toEqual({
      xPt: 126,
      yPt: 139,
      widthPt: 20,
      heightPt: 10,
    });
    expect(result.axes.horizontal).toMatchObject({
      status: 'resolved',
      relativeFrom: 'character',
      referenceFrame: 'character',
      choiceKind: 'offset',
    });
    expect(result.axes.vertical).toMatchObject({
      status: 'resolved',
      relativeFrom: 'line',
      referenceFrame: 'line',
      choiceKind: 'offset',
    });
  });

  it('does not substitute column for a missing character frame', () => {
    const value = input(anchor({
      horizontal: {
        relativeFrom: 'character',
        relativeFromStatus: 'valid',
        choice: { kind: 'offset', valuePt: 2 },
      },
    }));
    value.frames.character = null;

    const result = resolveAnchorFrame(value);

    expect(result.status).toBe('unsupported');
    expect(result.axes.horizontal).toMatchObject({
      status: 'unsupported',
      relativeFrom: 'character',
      issueCode: 'missing-reference-frame',
    });
  });

  it('does not substitute paragraph for a missing line frame', () => {
    const value = input(anchor({
      vertical: {
        relativeFrom: 'line',
        relativeFromStatus: 'valid',
        choice: { kind: 'offset', valuePt: 2 },
      },
    }));
    value.frames.line = null;

    const result = resolveAnchorFrame(value);

    expect(result.status).toBe('unsupported');
    expect(result.axes.vertical).toMatchObject({
      status: 'unsupported',
      relativeFrom: 'line',
      issueCode: 'missing-reference-frame',
    });
  });

  it('preserves exact zero percent positioning', () => {
    const result = resolved(resolveAnchorFrame(input(anchor({
      horizontal: {
        relativeFrom: 'margin',
        relativeFromStatus: 'valid',
        choice: { kind: 'percent', fraction: 0 },
      },
      vertical: {
        relativeFrom: 'page',
        relativeFromStatus: 'valid',
        choice: { kind: 'percent', fraction: 0 },
      },
    }))));

    expect(result.geometry.objectFrame).toEqual({
      xPt: 60,
      yPt: 20,
      widthPt: 20,
      heightPt: 10,
    });
    expect(result.axes.horizontal).toMatchObject({
      choiceKind: 'percent',
      choiceValue: 0,
    });
  });

  it('uses wp:extent when Word does not support an exact-zero relative size', () => {
    const result = resolved(resolveAnchorFrame(input(anchor({
      relativeSize: {
        horizontal: {
          relativeFrom: 'page',
          relativeFromStatus: 'valid',
          fraction: 0,
          fractionStatus: 'valid',
        },
        vertical: null,
      },
    }))));

    expect(result.geometry.objectFrame.widthPt).toBe(20);
    expect(result.geometry.size.horizontal).toEqual({
      source: 'extent',
      valuePt: 20,
      relativeFrom: 'page',
      referenceFrame: null,
      fraction: 0,
      compatibilityFallback: 'word-zero-relative-size',
    });
  });

  it('rejects a negative relative size instead of treating it as an extent fallback', () => {
    const result = resolveAnchorFrame(input(anchor({
      relativeSize: {
        horizontal: {
          relativeFrom: 'page',
          relativeFromStatus: 'valid',
          fraction: -0.1,
          fractionStatus: 'valid',
        },
        vertical: null,
      },
    })));

    expect(result.status).toBe('unsupported');
    expect(result.issues).toContainEqual(expect.objectContaining({
      code: 'invalid-relative-size-fraction',
      path: 'relativeSize.horizontal.fraction',
    }));
  });

  it('uses simple positioning only when explicitly enabled', () => {
    const result = resolved(resolveAnchorFrame(input(anchor({
      simplePosition: {
        enabled: true,
        status: 'valid',
        xPt: -5,
        xStatus: 'valid',
        yPt: 7,
        yStatus: 'valid',
      },
      horizontal: {
        relativeFrom: null,
        relativeFromStatus: 'missing',
        choice: { kind: 'missing' },
      },
      vertical: {
        relativeFrom: null,
        relativeFromStatus: 'missing',
        choice: { kind: 'missing' },
      },
    }))));

    expect(result.geometry.objectFrame).toMatchObject({ xPt: 5, yPt: 27 });
    expect(result.axes.horizontal).toMatchObject({
      status: 'resolved',
      referenceFrame: 'page',
      choiceKind: 'simple-position',
    });
  });

  it('reports a valid zero outer extent as unsupported on either anchor axis', () => {
    for (const axis of ['horizontal', 'vertical'] as const) {
      const acquisition = anchor({
        extent: {
          ...anchor().extent,
          [axis === 'horizontal' ? 'widthPt' : 'heightPt']: 0,
        },
      });
      const result = resolveAnchorFrame(input(acquisition));
      expect(result.status).toBe('unsupported');
      expect(result.axes[axis]).toMatchObject({
        status: 'unsupported', issueCode: 'invalid-size',
      });
    }
  });

  it('requires authored axis facts instead of falling back to page or zero', () => {
    const missingBase = resolveAnchorFrame(input(anchor({
      horizontal: {
        relativeFrom: null,
        relativeFromStatus: 'missing',
        choice: { kind: 'offset', valuePt: 2 },
      },
    })));
    const missingChoice = resolveAnchorFrame(input(anchor({
      horizontal: {
        relativeFrom: 'page',
        relativeFromStatus: 'valid',
        choice: { kind: 'missing' },
      },
    })));

    expect(missingBase.status).toBe('unsupported');
    expect(missingBase.axes.horizontal).toMatchObject({
      status: 'unsupported',
      issueCode: 'missing-relative-from',
    });
    expect(missingChoice.status).toBe('unsupported');
    expect(missingChoice.axes.horizontal).toMatchObject({
      status: 'unsupported',
      issueCode: 'missing-axis-choice',
    });
  });

  it('uses explicit page parity for inside and outside placement', () => {
    const value = input(anchor({
      horizontal: {
        relativeFrom: 'margin',
        relativeFromStatus: 'valid',
        choice: { kind: 'align', value: 'inside' },
      },
    }));
    value.frames.pageParity = 'even';

    const result = resolved(resolveAnchorFrame(value));

    expect(result.geometry.objectFrame.xPt).toBe(540);
    expect(result.axes.horizontal).toMatchObject({
      choiceKind: 'align',
      choiceValue: 'inside',
      pageParity: 'even',
    });
  });

  it('maps tight and through polygons exactly in the fixed Office coordinate space', () => {
    const result = resolved(resolveAnchorFrame(input(anchor({
      extent: {
        widthPt: 216,
        heightPt: 108,
        widthStatus: 'valid',
        heightStatus: 'valid',
      },
      parentEffectExtent: validEdges(0.2, 0.3, 0.4, 0.1),
      anchorDistances: validEdges(1, 2, 3, 4),
      wrap: {
        kind: 'through',
        authoredKinds: ['wrapThrough'],
        side: 'largest',
        distances: {
          ...missingEdges(),
          leftPt: 0.5,
          leftStatus: 'valid',
          rightPt: 0.6,
          rightStatus: 'valid',
        },
        effectExtent: null,
        polygon: {
          edited: true,
          coordinateSpace: { width: 21600, height: 21600 },
          points: [
            { x: 0, y: 0, rawX: '0', rawY: '0' },
            { x: 21600, y: 10800, rawX: '21600', rawY: '10800' },
            { x: 24000, y: 21600, rawX: '24000', rawY: '21600' },
          ],
          invalidPointCount: 0,
        },
      },
    }))));

    expect(result.geometry.inkBounds).toEqual({
      xPt: 9.9,
      yPt: 19.8,
      widthPt: 216.4,
      heightPt: 108.60000000000001,
    });
    expect(result.geometry.wrap).toMatchObject({
      kind: 'through',
      side: 'largest',
      coordinateSpace: { width: 21600, height: 21600 },
      distances: { topPt: 1, rightPt: 0.6, bottomPt: 3, leftPt: 0.5 },
      polygon: {
        points: [
          { xPt: 10, yPt: 20 },
          { xPt: 226, yPt: 74 },
          { xPt: 250, yPt: 128 },
        ],
      },
    });
    expect(result.geometry.wrapBounds).toEqual({
      xPt: 9.5,
      yPt: 19,
      widthPt: 241.1,
      heightPt: 112,
    });
  });

  it('hands a zero-signed-area bow-tie to the compiled line oracle without late rejection', () => {
    const result = resolved(resolveAnchorFrame(input(anchor({
      extent: {
        widthPt: 100,
        heightPt: 100,
        widthStatus: 'valid',
        heightStatus: 'valid',
      },
      wrap: {
        kind: 'through',
        authoredKinds: ['wrapThrough'],
        side: 'bothSides',
        distances: missingEdges(),
        effectExtent: null,
        polygon: {
          edited: false,
          coordinateSpace: { width: 21600, height: 21600 },
          points: [
            { x: 0, y: 0, rawX: '0', rawY: '0' },
            { x: 21600, y: 21600, rawX: '21600', rawY: '21600' },
            { x: 0, y: 21600, rawX: '0', rawY: '21600' },
            { x: 21600, y: 0, rawX: '21600', rawY: '0' },
          ],
          invalidPointCount: 0,
        },
      },
    }))));
    const bounds = result.geometry.wrapBounds!;
    const polygon = result.geometry.wrap.polygon!.points;
    const oracle = createFloatWrapOracle([{
      kind: 'shape', mode: 'square', authoredWrap: 'through', wrapPolygon: polygon,
      imageKey: 'bow-tie', imageX: 10, imageY: 20, imageW: 100, imageH: 100,
      xLeft: bounds.xPt, xRight: bounds.xPt + bounds.widthPt,
      yTop: bounds.yPt, yBottom: bounds.yPt + bounds.heightPt,
      side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
      paraId: 0,
    }]);

    expect(oracle.lineWindow({
      topYPt: 40,
      minimumStartWidthPt: 1,
      probeHeightPt: 10,
      paragraphXPt: 0,
      maximumWidthPt: 120,
      columnXPt: 0,
      columnWidthPt: 120,
    })).toMatchObject({ topYPt: 40 });
  });

  it('does not degrade an invalid tight polygon to square wrapping', () => {
    const result = resolveAnchorFrame(input(anchor({
      wrap: {
        kind: 'tight',
        authoredKinds: ['wrapTight'],
        side: 'bothSides',
        distances: missingEdges(),
        effectExtent: null,
        polygon: {
          edited: false,
          coordinateSpace: { width: 21600, height: 21600 },
          points: [
            { x: 0, y: 0, rawX: '0', rawY: '0' },
            { x: null, y: 21600, rawX: 'bad', rawY: '21600' },
            { x: 21600, y: 21600, rawX: '21600', rawY: '21600' },
          ],
          invalidPointCount: 1,
        },
      },
    })));

    expect(result.status).toBe('unsupported');
    expect(result.issues).toContainEqual(expect.objectContaining({
      code: 'invalid-wrap-polygon',
    }));
  });

  it('reports an invalid wrapping-child distance instead of using the anchor value', () => {
    const result = resolveAnchorFrame(input(anchor({
      anchorDistances: validEdges(1, 2, 3, 4),
      wrap: {
        kind: 'square',
        authoredKinds: ['wrapSquare'],
        side: 'bothSides',
        distances: {
          ...missingEdges(),
          leftPt: null,
          leftStatus: 'invalid',
        },
        effectExtent: null,
        polygon: null,
      },
    })));

    expect(result.status).toBe('unsupported');
    expect(result.issues).toContainEqual(expect.objectContaining({
      code: 'invalid-distance',
      path: 'wrap.distances.left',
    }));
  });

  it('retains parent and wrapping-child effect extents with unambiguous provenance', () => {
    const result = resolved(resolveAnchorFrame(input(anchor({
      parentEffectExtent: validEdges(1, 2, 3, 4),
      wrap: {
        kind: 'square',
        authoredKinds: ['wrapSquare'],
        side: 'bothSides',
        distances: missingEdges(),
        effectExtent: validEdges(0.1, 0.2, 0.3, 0.4),
        polygon: null,
      },
    }))));

    expect(result.geometry.parentEffectExtent).toEqual({
      topPt: 1,
      rightPt: 2,
      bottomPt: 3,
      leftPt: 4,
    });
    expect(result.geometry.inkBounds).toEqual({
      xPt: 6,
      yPt: 19,
      widthPt: 26,
      heightPt: 14,
    });
    expect(result.geometry.wrap).toMatchObject({
      effectExtentSource: 'wrap-child',
      effectExtent: {
        topPt: 0.1,
        rightPt: 0.2,
        bottomPt: 0.3,
        leftPt: 0.4,
      },
    });
    expect(result.geometry.wrapBounds).toEqual({
      xPt: 9.6,
      yPt: 19.9,
      widthPt: 20.599999999999998,
      heightPt: 10.4,
    });
  });

  it.each([
    ['missing x', { xPt: null, xStatus: 'missing', yPt: 0, yStatus: 'valid' }, 'missing-simple-coordinate', 'simplePosition.x'],
    ['invalid x', { xPt: null, xStatus: 'invalid', yPt: 0, yStatus: 'valid' }, 'invalid-simple-position', 'simplePosition.x'],
    ['missing y', { xPt: 0, xStatus: 'valid', yPt: null, yStatus: 'missing' }, 'missing-simple-coordinate', 'simplePosition.y'],
    ['invalid y', { xPt: 0, xStatus: 'valid', yPt: null, yStatus: 'invalid' }, 'invalid-simple-position', 'simplePosition.y'],
  ] as const)('diagnoses %s independently when simple positioning is enabled', (
    _name, coordinates, code, path,
  ) => {
    const result = resolveAnchorFrame(input(anchor({
      simplePosition: {
        enabled: true,
        status: 'valid',
        ...coordinates,
      },
    })));

    expect(result.status).toBe('unsupported');
    expect(result.issues).toContainEqual(expect.objectContaining({
      code,
      path,
    }));
  });

  it('keeps group transforms as cloned metadata without baking them into the frame', () => {
    const transform = {
      offsetXEmu: 914400,
      offsetYEmu: 1828800,
      extentWidthEmu: 2743200,
      extentHeightEmu: 3657600,
      childOffsetXEmu: 100,
      childOffsetYEmu: 200,
      childExtentWidthEmu: 300,
      childExtentHeightEmu: 400,
      rotationUnits: 60000,
      flipH: true,
      flipV: false,
    };
    const acquisition = anchor({
      group: {
        childSourceId: 'child-1',
        sourceIndex: 0,
        sourceCount: 1,
        transformChain: [transform],
        childTransform: transform,
        resolvedChildFrame: {
          offsetXPt: 1, offsetYPt: 2, widthPt: 3, heightPt: 4,
          rotationDeg: 1, flipH: true, flipV: false,
        },
      },
    });

    const result = resolved(resolveAnchorFrame(input(acquisition)));
    transform.offsetXEmu = 0;

    expect(result.geometry.objectFrame).toEqual({
      xPt: 10,
      yPt: 20,
      widthPt: 20,
      heightPt: 10,
    });
    expect(result.geometry.transform).toMatchObject({
      groupApplication: 'parser-resolved-child-frame',
      group: {
        childSourceId: 'child-1',
        transformChain: [{ offsetXEmu: 914400 }],
      },
    });
  });

  it('returns deeply frozen structured-clone-safe plain data detached from inputs', () => {
    const value = input();
    const result = resolveAnchorFrame(value);

    (value.frames.page as { xPt: number }).xPt = 999;

    expect(structuredClone(result)).toEqual(result);
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result.axes)).toBe(true);
    if (result.status === 'resolved') {
      expect(result.geometry.objectFrame.xPt).toBe(10);
      expect(Object.isFrozen(result.geometry)).toBe(true);
      expect(Object.isFrozen(result.geometry.wrap)).toBe(true);
    }
  });
});
