import type { ParagraphAcquisitionInput } from './text.js';
import type { AcquiredParagraphLayoutInput, PaintResourceRegistry, LayoutRect, SourceRef } from './types.js';
import { acquireNativeReadingBlockScene, placeNativeReadingBlockScene, NativeReadingSceneError, type NativeReadingBlockScene } from './native-reading-block-scene.js';
import { acquireNativeReadingImageBlockScene } from './native-reading-image-scene.js';

/** Source-owned requests only. Authored wrap/geometry remain in the acquisition
 * tree; an ordinary missing contour never becomes an implicit reading request. */
export function nativeReadingOccurrenceIds(paragraph: ParagraphAcquisitionInput): readonly string[] {
  const occurrences = new Map<string, boolean>();
  for (const run of paragraph.runs) {
    if (run.type !== 'image' && run.type !== 'shape' && run.type !== 'chart' && run.type !== 'unavailableDrawing') continue;
    const acquisition = run.anchorAcquisitionInput;
    if (!acquisition) continue;
    const requested = acquisition.nativeReadingRelocation === 'completeScene';
    const previous = occurrences.get(acquisition.occurrenceId);
    if (previous !== undefined && previous !== requested) throw new NativeReadingSceneError('source-ownership', 'mixed placement requests within one occurrence');
    occurrences.set(acquisition.occurrenceId, requested);
  }
  return Object.freeze([...occurrences].filter(([, requested]) => requested).map(([id]) => id));
}

/** Complete paragraph plus scenes is one indivisible reading block. This is an
 * explicit changed-placement policy, not Office's automatic exclusion contour.
 * Keeping text and all scenes together avoids inventing source continuation
 * boundaries. Oversize blocks fail instead of truncating, scaling or clipping.
 * The body paginator may move this whole block to a fresh column. */
export function appendNativeReadingScenes(
  paragraph: ParagraphAcquisitionInput, input: AcquiredParagraphLayoutInput,
  frame: LayoutRect, registry: PaintResourceRegistry | undefined, maximumWork = 16384,
): AcquiredParagraphLayoutInput {
  const occurrences = nativeReadingOccurrenceIds(paragraph);
  if (occurrences.length === 0) return input;
  if (input.source.story !== 'body' || !input.ordinaryFlow || input.continuation || input.clipBounds)
    throw new NativeReadingSceneError('placement', 'reading blocks require complete unclipped body paragraphs');
  if (paragraph.runs.length > maximumWork || occurrences.length > maximumWork)
    throw new NativeReadingSceneError('budget', 'reading paragraph source allowance exceeded');
  const exclusionWork = input.exclusions.length * (occurrences.length + 1);
  if (!Number.isSafeInteger(exclusionWork) || exclusionWork > maximumWork)
    throw new NativeReadingSceneError('budget', 'reading exclusion scan allowance exceeded');
  // The disclosed reading policy places complete scenes in ordinary flow; it
  // does not reconstruct Word's unavailable contour. Retain prior exclusions
  // and move below each intersecting allocation, charging every gap. Sorting
  // by bottom makes one monotonic pass sufficient, independent of input order.
  const exclusions = input.exclusions.map(exclusion => {
    const b = exclusion.bounds;
    if (![b.xPt,b.yPt,b.widthPt,b.heightPt,b.xPt+b.widthPt,b.yPt+b.heightPt].every(Number.isFinite)
      || b.widthPt < 0 || b.heightPt < 0)
      throw new NativeReadingSceneError('placement', 'invalid reading exclusion bounds');
    return b;
  }).filter(b=>b.widthPt > 0 && b.heightPt > 0).sort((a,b)=>(a.yPt+a.heightPt)-(b.yPt+b.heightPt));
  const scenes = acquireRequestedNativeReadingScenes(paragraph, input.source, input.id, input.flowDomainId, registry, maximumWork-exclusionWork);
  const height = scenes.reduce((total, scene) => total + scene.drawing.flowBounds.heightPt, 0);
  let fullHeight = input.flowBounds.heightPt + height;
  if (![frame.xPt, frame.yPt, frame.widthPt, frame.heightPt, height, fullHeight].every(Number.isFinite)
    || frame.widthPt <= 0 || frame.heightPt <= 0 || fullHeight > frame.heightPt)
    throw new NativeReadingSceneError('placement', 'complete reading paragraph exceeds a full column');
  let top = input.flowBounds.yPt + input.flowBounds.heightPt - input.spacing.afterPt;
  const drawings = [...input.drawings], lines = [...input.lines];
  const end = lines.at(-1)?.range.end ?? 0;
  let left = input.inkBounds.xPt, right = left + input.inkBounds.widthPt;
  let inkTop = input.inkBounds.yPt, bottom = inkTop + input.inkBounds.heightPt;
  for (const scene of scenes) {
    const beforeExclusions = top, width = scene.drawing.flowBounds.widthPt;
    for (const e of exclusions) {
      if (input.flowBounds.xPt < e.xPt+e.widthPt && input.flowBounds.xPt+width > e.xPt
        && top < e.yPt+e.heightPt && top+scene.drawing.flowBounds.heightPt > e.yPt)
        top = e.yPt+e.heightPt;
    }
    fullHeight += top-beforeExclusions;
    // The existing indivisible selector can move a gap-charged candidate to a
    // fresh column, then reacquire it against that column's exclusions. A gap
    // larger than a full column remains unsupported; never bypass admission.
    if (!Number.isFinite(fullHeight) || fullHeight > frame.heightPt)
      throw new NativeReadingSceneError('placement', 'complete reading paragraph exceeds a full column');
    const drawing = placeNativeReadingBlockScene(scene, { xPt: input.flowBounds.xPt, yPt: top,
      widthPt: Math.min(input.flowBounds.widthPt, frame.widthPt), heightPt: scene.drawing.flowBounds.heightPt });
    const b = drawing.flowBounds;
    drawings.push(drawing);
    lines.push({ range: { start: end, end }, bounds: b, baselinePt: top + b.heightPt, advancePt: b.heightPt,
      placements: [{ kind: 'drawing', drawingId: drawing.id, range: { start: end, end }, bounds: drawing.inkBounds, advancePt: b.heightPt }] });
    left = Math.min(left, drawing.inkBounds.xPt); right = Math.max(right, drawing.inkBounds.xPt + drawing.inkBounds.widthPt);
    inkTop = Math.min(inkTop, drawing.inkBounds.yPt); bottom = Math.max(bottom, drawing.inkBounds.yPt + drawing.inkBounds.heightPt);
    top += b.heightPt;
  }
  return { ...input, flowBounds: { ...input.flowBounds, heightPt: fullHeight },
    inkBounds: { xPt: left, yPt: inkTop, widthPt: right - left, heightPt: bottom - inkTop }, drawings, lines,
    nativeReadingRelocations: Object.freeze(scenes.map(scene => scene.drawing.id)),
    ...(paragraph.runs.some(run => {
      if (run.type !== 'image') return false;
      const metadata = run.anchorAcquisitionInput?.nativePictureMetadata;
      return [metadata?.inactiveFillCarrier, metadata?.inactiveLineCarrier].some(property =>
        property != null && property.retention !== 'ignoredZeroIndex');
    }) ? { nativeReadingInactivePictureData: true } : {}),
  };
}

/** Used by page-anchor prescan to defer only a fully acquired scene. */
export function acquireRequestedNativeReadingScenes(paragraph: ParagraphAcquisitionInput, source: SourceRef, id: string, flowDomainId: string, registry: PaintResourceRegistry | undefined, maximumWork = 16384): readonly NativeReadingBlockScene[] {
  const occurrences = nativeReadingOccurrenceIds(paragraph);
  if (paragraph.runs.length * occurrences.length > maximumWork) throw new NativeReadingSceneError('budget', 'reading occurrence scan allowance exceeded');
  const scenes: NativeReadingBlockScene[] = [];
  let remaining = maximumWork;
  // All scenes acquire before any changed candidate is returned. One unsupported
  // later resource/member therefore cannot expose a valid prefix of the block.
  for (const occurrence of occurrences) {
    const first = paragraph.runs.find(run => (run.type === 'shape' || run.type === 'image' || run.type === 'chart' || run.type === 'unavailableDrawing')
      && run.anchorAcquisitionInput?.occurrenceId === occurrence);
    const drawingId = `${id}:reading:${occurrence}`;
    if (first?.type === 'image') {
      if (!registry) throw new NativeReadingSceneError('resource', 'reading image registry is unavailable');
      scenes.push(acquireNativeReadingImageBlockScene(paragraph, occurrence, source, drawingId, flowDomainId, registry, remaining));
    } else scenes.push(acquireNativeReadingBlockScene(paragraph, occurrence, source, drawingId, flowDomainId, remaining));
    remaining -= paragraph.runs.length + scenes.at(-1)!.observedPaintOperations;
    if (remaining < 0) throw new NativeReadingSceneError('budget', 'reading scene work allowance exceeded');
  }
  return Object.freeze(scenes);
}
