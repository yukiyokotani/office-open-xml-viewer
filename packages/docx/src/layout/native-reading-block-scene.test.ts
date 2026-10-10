import { describe, it } from 'vitest';
import assert from 'node:assert/strict';
import type { ParagraphAcquisitionInput } from './text.js';
import { snapshotPlainData } from './plain-data.js';
import { acquireNativeReadingBlockScene, placeNativeReadingBlockScene, NativeReadingSceneError } from './native-reading-block-scene.js';

const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
const occurrence = 'invented-group';
function raw() {
  const acquisition = (index: number) => ({
    occurrenceId: occurrence,
    extent: { widthStatus: 'valid', heightStatus: 'valid', widthPt: 80, heightPt: 40 },
    relativeSize: { horizontal: null, vertical: null },
    wrap: { kind: 'through', polygon: null },
    group: {
      sourceIndex: index, sourceCount: 2, childSourceId: `invented-child-${index}`,
      transformChain: [], childTransform: null,
      resolvedChildFrame: { offsetXPt: index === 0 ? -10 : 30, offsetYPt: 5,
        widthPt: 20, heightPt: 10, rotationDeg: index === 0 ? 90 : 0,
        flipH: index === 0, flipV: false },
    },
  });
  return { runs: [
    { type: 'text', text: 'retained paragraph text', anchorAcquisitionInput: undefined },
    { type: 'anchorHost', anchorOccurrenceId: occurrence, anchorAcquisitionInput: undefined },
    ...[0, 1].map(index => ({
      type: 'shape', presetGeometry: 'rect', subpaths: [], fill: { fillType: 'solid', color: index === 0 ? 'FF0000' : '0000FF' },
      stroke: '000000', strokeWidth: 2,
      widthPt: 20, heightPt: 10, anchorXPt: 0, anchorYPt: 0,
      anchorXFromMargin: false, anchorYFromPara: false, zOrder: index,
      anchorAcquisitionInput: acquisition(index),
    })),
  ] };
}
function scene(input = raw(), budget = 2000) {
  return acquireNativeReadingBlockScene(snapshotPlainData(input, 'invented paragraph') as unknown as ParagraphAcquisitionInput,
    occurrence, source, 'invented-scene', 'invented-flow', budget);
}
function blocked(fn: () => unknown, expected: string) {
  assert.throws(fn, error => error instanceof NativeReadingSceneError && error.blocker === expected);
}

describe('complete native reading block scene', () => {
  it('retains all members in order, child transforms and paragraph text without a frame guess', () => {
    const input = raw(), before = JSON.stringify(input), acquired = scene(input);
    assert.deepEqual(acquired.memberRunIndices, [2, 3]);
    assert.deepEqual(acquired.childSourceIds, ['invented-child-0', 'invented-child-1']);
    assert.equal(acquired.authoredWrap, 'through');
    assert.equal(acquired.drawing.commands.length, 2);
    const first = acquired.drawing.commands[0];
    assert.equal(first.kind, 'drawingml-shape');
    if (first.kind !== 'drawingml-shape') throw new Error('missing member');
    assert.deepEqual(first.plan.transform, { rotationDeg: 90, flipH: true, flipV: false });
    assert.equal(first.plan.rect.x, -10);
    // Actual rotated miter paint extends left of the positive authored group.
    assert.equal(acquired.drawing.inkBounds.xPt, -6);
    assert.equal(acquired.drawing.flowBounds.xPt, -6);
    assert.equal(acquired.drawing.flowBounds.widthPt, 86);
    assert.equal(acquired.drawing.flowBounds.heightPt, 41);
    assert.equal(JSON.stringify(input), before);
  });

  it('rejects missing, duplicate and reordered member ownership atomically', () => {
    const missing = raw(); missing.runs.pop(); blocked(() => scene(missing), 'member-coverage');
    const duplicate = raw(); duplicate.runs[3] = duplicate.runs[2]; blocked(() => scene(duplicate), 'member-coverage');
    const reversed = raw(); [reversed.runs[2], reversed.runs[3]] = [reversed.runs[3], reversed.runs[2]];
    blocked(() => scene(reversed), 'member-coverage');
  });

  it('rejects unsupported later resource/text members rather than returning a valid first member', () => {
    const resource = raw(); Object.assign(resource.runs[3], { type: 'unavailableDrawing', resourceKind: 'image' });
    blocked(() => scene(resource), 'resource');
    const text = raw(); Object.assign(text.runs[3], { textBoxInput: { kind: 'complete', source, blockCount: 1 } });
    blocked(() => scene(text), 'owned-text');
    const imageFill = raw(); Object.assign(imageFill.runs[3], { fill: { fillType: 'image', imagePath: 'invented.png' } });
    blocked(() => scene(imageFill), 'resource');
  });

  it('rejects unsealed input, absent host, inconsistent extent and stored contour', () => {
    blocked(() => acquireNativeReadingBlockScene(raw() as unknown as ParagraphAcquisitionInput, occurrence, source, 'scene', 'flow', 2000), 'source-ownership');
    const host = raw(); host.runs.splice(1, 1); blocked(() => scene(host), 'source-ownership');
    const extent = raw(); Object.assign(extent.runs[3], { anchorAcquisitionInput: { ...extent.runs[2].anchorAcquisitionInput,
      group: extent.runs[3].anchorAcquisitionInput?.group, extent: { widthStatus: 'valid', heightStatus: 'valid', widthPt: 90, heightPt: 40 } } });
    blocked(() => scene(extent), 'source-ownership');
    const contour = raw(); Object.assign(contour.runs[2].anchorAcquisitionInput?.wrap ?? {}, { polygon: { points: [] } });
    blocked(() => scene(contour), 'source-ownership');
  });

  it('rejects oversized input before planning and unsupported paint without source edits', () => {
    const input = raw(), before = JSON.stringify(input); blocked(() => scene(input, 3), 'budget');
    assert.equal(JSON.stringify(input), before);
    blocked(() => scene(raw(), 25), 'budget');
    const pattern = raw(); Object.assign(pattern.runs[3], { fill: { fillType: 'pattern', preset: 'pct25', fg: '000000', bg: 'FFFFFF' } });
    blocked(() => scene(pattern), 'paint');
  });

  it('relocates the whole paint scene into one slot and rejects rather than clips overflow', () => {
    const acquired = scene(), before = JSON.stringify(acquired);
    const placed = placeNativeReadingBlockScene(acquired, { xPt: 100, yPt: 200, widthPt: 86, heightPt: 41 });
    assert.deepEqual(placed.flowBounds, { xPt: 100, yPt: 200, widthPt: 86, heightPt: 41 });
    const first = placed.commands[0], second = placed.commands[1];
    if (first.kind !== 'drawingml-shape' || second.kind !== 'drawingml-shape') throw new Error('lost member');
    assert.equal(first.plan.rect.x, 96); assert.equal(second.plan.rect.x - first.plan.rect.x, 40);
    assert.deepEqual(first.plan.transform, acquired.drawing.commands[0].kind === 'drawingml-shape' ? acquired.drawing.commands[0].plan.transform : null);
    blocked(() => placeNativeReadingBlockScene(acquired, { xPt: 100, yPt: 200, widthPt: 85, heightPt: 41 }), 'placement');
    blocked(() => placeNativeReadingBlockScene(acquired, { xPt: 100, yPt: 200, widthPt: 86, heightPt: 40 }), 'placement');
    assert.equal(JSON.stringify(acquired), before);
  });
});
