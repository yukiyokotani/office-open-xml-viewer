import type { ParagraphAcquisitionInput } from './text.js';
import type { ImagePaintResourceDescriptor, PaintResourceRegistry, SourceRef } from './types.js';
import type { NativeReadingBlockScene } from './native-reading-block-scene.js';
import { NativeReadingSceneError } from './native-reading-block-scene.js';
import { isDeepFrozenPlainDataRoot, snapshotPlainData } from './plain-data.js';
import { imageResourceKey } from './source-key.js';
import { acquireNativeReadingImagePlan, nativeReadingImageFrame } from './native-reading-image-frame.js';

/** Complete standalone native image scene. The registry owns the exact source
 * part and transform/crop/effect descriptor; a resource command is retained,
 * never replaced by a silhouette or placeholder. Actual decoded availability
 * must be rechecked by the page-paint owner before exposing approximate pages.
 * This narrow class excludes grouped images, SVG substitution and color
 * effects; those cannot silently use an unrelated decoded resource. */
export function acquireNativeReadingImageBlockScene(
  paragraph: ParagraphAcquisitionInput, occurrenceId: string, source: SourceRef,
  id: string, flowDomainId: string, registry: PaintResourceRegistry, maximumRuns: number,
): NativeReadingBlockScene {
  const fail = (blocker: NativeReadingSceneError['blocker'], message: string): never => { throw new NativeReadingSceneError(blocker, message); };
  if (!Number.isSafeInteger(maximumRuns) || maximumRuns < 1 || paragraph.runs.length > maximumRuns) return fail('budget', 'image run allowance exceeded');
  if (!isDeepFrozenPlainDataRoot(paragraph)) return fail('source-ownership', 'image paragraph must be internally sealed');
  let hosts = 0, memberIndex = -1;
  for (let index = 0; index < paragraph.runs.length; index++) {
    const run = paragraph.runs[index];
    if (run.type === 'anchorHost' && run.anchorOccurrenceId === occurrenceId) hosts++;
    if ((run.type === 'image' || run.type === 'shape' || run.type === 'chart' || run.type === 'unavailableDrawing')
      && run.anchorAcquisitionInput?.occurrenceId === occurrenceId) {
      if (memberIndex !== -1) return fail('member-coverage', 'standalone image must own exactly one payload');
      memberIndex = index;
    }
  }
  const run = paragraph.runs[memberIndex];
  if (hosts !== 1 || !run || run.type !== 'image' || !run.anchorAcquisitionInput) return fail('source-ownership', 'one image host and payload are required');
  const input = run.anchorAcquisitionInput;
  if (input.group || input.wrap.polygon !== null || (input.wrap.kind !== 'tight' && input.wrap.kind !== 'through')) return fail('source-ownership', 'requires standalone automatic authored contour');
  if (input.relativeSize.horizontal !== null || input.relativeSize.vertical !== null) return fail('geometry', 'page-dependent image extent');
  const width = input.extent.widthPt, height = input.extent.heightPt;
  if (input.extent.widthStatus !== 'valid' || input.extent.heightStatus !== 'valid' || width === null || height === null
    || width <= 0 || height <= 0 || !Number.isFinite(width) || !Number.isFinite(height)) return fail('geometry', 'invalid image allocation');
  const imageSource = { ...source, path: [...source.path, memberIndex] };
  const key = imageResourceKey(imageSource, run.imagePath);
  let descriptor: ImagePaintResourceDescriptor;
  try { descriptor = registry.resolve(key, 'image') as ImagePaintResourceDescriptor; }
  catch { return fail('resource', 'image descriptor is unavailable for its source key'); }
  if (!isDeepFrozenPlainDataRoot(descriptor) || descriptor.kind !== 'image' || descriptor.resourceKey !== key
    || descriptor.partPath !== run.imagePath || descriptor.mimeType !== run.mimeType) return fail('resource', 'image descriptor ownership mismatch');
  if (run.svgImagePath != null || descriptor.svgImagePath !== undefined || run.duotone != null || descriptor.duotone !== undefined || run.colorReplaceFrom != null || descriptor.colorReplaceFrom !== undefined) return fail('resource', 'image substitution/color effects require their decoded-owner proof');
  for (const field of ['rotation', 'flipH', 'flipV', 'alpha'] as const) if (descriptor[field] !== run[field]) return fail('resource', `image ${field} descriptor mismatch`);
  if ((descriptor.srcRect === undefined) !== (run.srcRect == null)) return fail('resource', 'image crop presence mismatch');
  if (descriptor.srcRect && run.srcRect) for (const edge of ['l', 't', 'r', 'b'] as const) if (descriptor.srcRect[edge] !== run.srcRect[edge]) return fail('resource', 'image crop descriptor mismatch');
  const rect = { xPt: 0, yPt: 0, widthPt: width, heightPt: height };
  let nativeImagePlan, bounds;
  try {
    nativeImagePlan = acquireNativeReadingImagePlan(descriptor, width, height);
    bounds = nativeReadingImageFrame(nativeImagePlan, rect);
  } catch (error) { return fail('geometry', error instanceof Error ? error.message : 'image projection failed'); }
  return snapshotPlainData({ occurrenceId, authoredWrap: input.wrap.kind, memberRunIndices: [memberIndex], childSourceIds: [],
    drawing: { kind: 'drawing' as const, id, source: imageSource, flowDomainId,
      flowBounds: bounds, inkBounds: bounds, advancePt: bounds.heightPt, ordinaryFlow: true,
      commands: [{ kind: 'resource' as const, resourceKind: 'image' as const, resourceKey: key, rect, nativeImagePlan }],
    }, observedPaintOperations: 0,
  }, 'complete native image reading scene');
}
