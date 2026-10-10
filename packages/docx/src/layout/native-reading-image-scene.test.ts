import type { DocxDocumentModel } from '../types.js';
import { projectDocumentSnapshotResources } from './production-paint-resources.js';
import { paintDrawingLayout } from '../paint/canvas-drawing.js';
import { createCanvasPaintResourcePainter } from '../paint/canvas-page.js';
import { canonicalCanvasPaintResourceHandlers } from '../paint/canonical-resource-handlers.js';
import { createPaintResourceSession } from '../paint/resource-session.js';
import type { CanvasPaintContext } from '../paint/types.js';
import { describe, it } from 'vitest';
import assert from 'node:assert/strict';
import type { ParagraphAcquisitionInput } from './text.js';
import type { ImagePaintResourceDescriptor } from './types.js';
import { sealPlainData, snapshotPlainData } from './plain-data.js';
import { createPaintResourceRegistry, indexSealedPaintResourceDescriptors } from './paint-resources.js';
import { imageResourceKey } from './source-key.js';
import { acquireNativeReadingImageBlockScene } from './native-reading-image-scene.js';
import { NativeReadingSceneError, placeNativeReadingBlockScene } from './native-reading-block-scene.js';
const source = { story: 'body', storyInstance: 'body', path: [1] } as const;
const key = imageResourceKey({ ...source, path: [1, 2] }, 'invented.jpeg');
function raw() { return { runs: [
  { type: 'text', text: 'retained full source paragraph' },
  { type: 'anchorHost', anchorOccurrenceId: 'invented-image-anchor' },
  { type: 'image', imagePath: 'invented.jpeg', mimeType: 'image/jpeg', widthPt: 40, heightPt: 20,
    rotation: 90, flipH: true, alpha: 0.75, srcRect: { l: -0.5, t: 0, r: 0, b: 0 },
    anchorAcquisitionInput: { occurrenceId: 'invented-image-anchor', group: null,
      extent: { widthPt: 40, heightPt: 20, widthStatus: 'valid', heightStatus: 'valid' },
      relativeSize: { horizontal: null, vertical: null }, wrap: { kind: 'tight', polygon: null } },
  },
] }; }
function descriptor(): ImagePaintResourceDescriptor { return {
  kind: 'image', resourceKey: key, partPath: 'invented.jpeg', mimeType: 'image/jpeg', intrinsicSize: { widthPt: 40, heightPt: 20 },
  rotation: 90, flipH: true, alpha: 0.75, srcRect: { l: -0.5, t: 0, r: 0, b: 0 },
}; }
function acquire(input = raw(), resource = descriptor()) {
  return acquireNativeReadingImageBlockScene(snapshotPlainData(input, 'invented image paragraph') as unknown as ParagraphAcquisitionInput,
    'invented-image-anchor', source, 'invented-image-scene', 'invented-flow', createPaintResourceRegistry([resource]), 100);
}

describe('complete standalone image reading scene', () => {
  it('native_image_null_effect acquires the exact plain raster scene from the real serialized-null descriptor owner', () => {
    // Rust ImageRun.color_replace_from has no skip_serializing_if: None is
    // serialized null, while Some remains an authored effect requiring proof.
    const omitted = raw(), serialized = raw();
    Object.assign(serialized.runs[2], { colorReplaceFrom: null });
    const before = JSON.stringify(serialized);
    const registryFor = (input: ReturnType<typeof raw>) => projectDocumentSnapshotResources({
      section: { pageWidth: 612, pageHeight: 792, marginTop: 72, marginBottom: 72, marginLeft: 72, marginRight: 72 },
      headers: {}, footers: {}, body: [{ type: 'paragraph', runs: [] }, { type: 'paragraph', ...input }],
    } as unknown as DocxDocumentModel, undefined, []).paintResources;
    const sceneFor = (input: ReturnType<typeof raw>, registry: ReturnType<typeof registryFor>) =>
      acquireNativeReadingImageBlockScene(snapshotPlainData(input, 'invented serialized image paragraph') as unknown as ParagraphAcquisitionInput,
        'invented-image-anchor', source, 'invented-image-scene', 'invented-flow', registry, 100);
    const registry = registryFor(serialized), control = registryFor(omitted);
    const scene = sceneFor(serialized, registry);
    assert.deepEqual(scene, sceneFor(omitted, control));
    assert.deepEqual(registry.descriptors, control.descriptors);
    assert.equal(Object.hasOwn(registry.resolve(key, 'image'), 'colorReplaceFrom'), false);
    assert.equal(JSON.stringify(serialized), before);
    assert.equal((serialized.runs[2] as unknown as { colorReplaceFrom: null }).colorReplaceFrom, null);
    for (const colorReplaceFrom of ['', 'FFFFFF']) {
      const effected = raw(); Object.assign(effected.runs[2], { colorReplaceFrom });
      const effectRegistry = registryFor(effected);
      assert.equal((effectRegistry.resolve(key, 'image') as ImagePaintResourceDescriptor).colorReplaceFrom, colorReplaceFrom);
      assert.throws(() => sceneFor(effected, effectRegistry), error => error instanceof NativeReadingSceneError
        && error.blocker === 'resource' && /color effects/.test(error.message));
    }
  });

  it('retains the exact image resource/crop/rotation binding and whole rotated allocation', () => {
    const input = raw(), before = JSON.stringify(input), scene = acquire(input);
    assert.equal(scene.authoredWrap, 'tight'); assert.deepEqual(scene.memberRunIndices, [2]);
    assert.deepEqual(scene.drawing.source.path, [1, 2]);
    assert.equal(scene.drawing.commands.length, 1);
    const command = scene.drawing.commands[0];
    if (command.kind !== 'resource') throw new Error('image replaced');
    assert.equal(command.resourceKey, key); assert.deepEqual(command.rect, { xPt: 0, yPt: 0, widthPt: 40, heightPt: 20 });
    assert.ok(Math.abs(scene.drawing.flowBounds.widthPt - 20) < 1e-12);
    assert.equal(scene.drawing.flowBounds.heightPt, 40);
    const placed = placeNativeReadingBlockScene(scene, { xPt: 80, yPt: 90, widthPt: 21, heightPt: 40 });
    assert.equal(placed.flowBounds.xPt, 80); assert.equal(placed.flowBounds.yPt, 90);
    const placedCommand = placed.commands[0];
    if (placedCommand.kind !== 'resource') throw new Error('image replaced during translation');
    assert.equal(placedCommand.nativeImagePlan, command.nativeImagePlan);
    // A consuming realm must reseal cloned source descriptors with the existing
    // authority; no bitmap or WeakSet identity is transported.
    const clonedDescriptors = structuredClone(createPaintResourceRegistry([descriptor()]).descriptors);
    sealPlainData(clonedDescriptors, 'invented worker descriptor input');
    const clonedRegistry = indexSealedPaintResourceDescriptors(clonedDescriptors);
    const clonedScene = acquireNativeReadingImageBlockScene(snapshotPlainData(structuredClone(input), 'invented worker paragraph') as unknown as ParagraphAcquisitionInput,
      'invented-image-anchor', source, 'invented-image-scene', 'invented-flow', clonedRegistry, 100);
    assert.deepEqual(clonedScene.drawing.commands, scene.drawing.commands);
    const image = { width: 100, height: 80 } as CanvasImageSource;
    const resources = createCanvasPaintResourcePainter(createPaintResourceSession(clonedRegistry, [{ kind: 'image', resourceKey: key, handle: image }]), canonicalCanvasPaintResourceHandlers);
    const calls: { name: string; args: unknown[]; alpha?: number }[] = [];
    let alpha = 1; const stack: number[] = [];
    const ctx = { get globalAlpha() { return alpha; }, set globalAlpha(n: number) { alpha = n; },
      save() { stack.push(alpha); }, restore() { alpha = stack.pop()!; },
      translate(...args: number[]) { calls.push({ name: 'translate', args }); },
      rotate(...args: number[]) { calls.push({ name: 'rotate', args }); },
      scale(...args: number[]) { calls.push({ name: 'scale', args }); },
      drawImage(...args: unknown[]) { calls.push({ name: 'drawImage', args, alpha }); },
      measureText() { throw new Error('paint must not measure'); },
    };
    const context = { ctx, resources, scale: 1, dpr: 1 } as unknown as CanvasPaintContext;
    paintDrawingLayout(placed, context);
    const second = placeNativeReadingBlockScene(scene, { xPt: 180, yPt: 290, widthPt: 21, heightPt: 40 });
    paintDrawingLayout(second, context);
    assert.deepEqual(calls.filter(c => c.name === 'translate').map(c => c.args), [[90, 110], [190, 310]]);
    const draws = calls.filter(c => c.name === 'drawImage');
    assert.equal(draws.length, 2); assert.deepEqual(draws[0], draws[1]);
    assert.equal(draws[0].args[0], image);
    const expected = [0, 0, 100, 80, -20 + 40 / 3, -10, 80 / 3, 20];
    for (let i = 0; i < expected.length; i++) assert.ok(Math.abs(Number(draws[0].args[i + 1]) - expected[i]) < 1e-12);
    assert.equal(draws[0].alpha, 0.75);
    assert.equal(JSON.stringify(input), before);
    // This is allocation/ownership only; no decoded-image or pixel proof.
    assert.equal(scene.observedPaintOperations, 0);
  });
  it('rejects descriptor drift and alternative color/SVG ownership rather than using the wrong handle', () => {
    assert.throws(() => acquireNativeReadingImageBlockScene(snapshotPlainData(raw(), 'invented missing resource') as unknown as ParagraphAcquisitionInput, 'invented-image-anchor', source, 'scene', 'flow', createPaintResourceRegistry([]), 100), error => error instanceof NativeReadingSceneError && error.blocker === 'resource');
    assert.throws(() => acquire(raw(), { ...descriptor(), rotation: 0 }), /descriptor mismatch/);
    assert.throws(() => acquire(raw(), { ...descriptor(), srcRect: undefined }), /crop presence/);
    assert.throws(() => acquire(raw(), { ...descriptor(), srcRect: { l: 0, t: 0, r: 0, b: 0 } }), /crop descriptor/);
    assert.throws(() => acquire(raw(), { ...descriptor(), svgImagePath: 'invented.svg' }), /substitution/);
    assert.throws(() => acquire(raw(), { ...descriptor(), colorReplaceFrom: 'FFFFFF' }), /color effects/);
    assert.throws(() => acquire(raw(), { ...descriptor(), partPath: 'other.jpeg' }), /ownership mismatch/);
    const overflow = raw(); Object.assign(overflow.runs[2], { srcRect: { l: -1e308, t: 0, r: -1e308, b: 0 } });
    assert.throws(() => acquire(overflow, { ...descriptor(), srcRect: { l: -1e308, t: 0, r: -1e308, b: 0 } }), error => error instanceof NativeReadingSceneError && error.blocker === 'geometry');
  });
  it('requires exactly one payload and host and preserves strict stored/relative geometry boundaries', () => {
    const duplicate = raw(); duplicate.runs.push(duplicate.runs[2]); assert.throws(() => acquire(duplicate), /exactly one/);
    const missingHost = raw(); missingHost.runs.splice(1, 1); assert.throws(() => acquire(missingHost), /host and payload/);
    const relative = raw(); Object.assign(relative.runs[2].anchorAcquisitionInput?.relativeSize ?? {}, { horizontal: { fraction: 0.5 } });
    assert.throws(() => acquire(relative), /page-dependent/);
    const contour = raw(); Object.assign(contour.runs[2].anchorAcquisitionInput?.wrap ?? {}, { polygon: { points: [] } });
    assert.throws(() => acquire(contour), /automatic authored contour/);
  });
});
