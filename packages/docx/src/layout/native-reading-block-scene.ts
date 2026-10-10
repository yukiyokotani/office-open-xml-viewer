import { GeometryWorkBudgetError } from '@silurus/ooxml-core';
import type { AnchorAcquisitionInput } from './anchor-input.js';
import type { ParagraphAcquisitionInput, ParagraphShapeRun, TextLayoutService } from './text.js';
import type { DrawingLayout, DrawingPaintCommand, LayoutRect, SourceRef } from './types.js';
import { isDeepFrozenPlainDataRoot, snapshotPlainData } from './plain-data.js';
import { planShapeDrawing, type ShapeDrawingPlanResult } from './shape-drawing-plan.js';
import { retainedAnchorChildFrame } from './retained-anchor-child-frame.js';
import { NativeWrapPaintBudgetError, deriveNativeWrapPaintExtent } from './native-wrap-paint-bounds.js';
import { translateDrawing } from './retained-geometry-translation.js';

export type NativeReadingSceneBlocker =
  | 'source-ownership' | 'member-coverage' | 'geometry' | 'resource'
  | 'owned-text' | 'clip' | 'paint' | 'budget' | 'placement';

export class NativeReadingSceneError extends Error {
  readonly blocker: NativeReadingSceneBlocker;
  constructor(blocker: NativeReadingSceneBlocker, message: string) {
    super(`Native reading scene ${blocker}: ${message}`);
    this.name = 'NativeReadingSceneError';
    this.blocker = blocker;
  }
}
function fail(blocker: NativeReadingSceneBlocker, message: string): never {
  throw new NativeReadingSceneError(blocker, message);
}

/** Same planner as ordinary anchored groups. Resolved child rotation/reflection
 * belongs to the child, not the outer group. Do not apply the raw transform
 * chain again: native acquisition has already resolved it into this frame. */
export function planRetainedAnchorShape(
  run: ParagraphShapeRun,
  acquisition: AnchorAcquisitionInput,
  rect: LayoutRect,
  text?: TextLayoutService,
  imageFillResourceKey?: string,
  maximumGeometryWork?: number,
): ShapeDrawingPlanResult {
  const child = acquisition.group?.resolvedChildFrame;
  return planShapeDrawing(child ? {
    ...run, rotation: child.rotationDeg, flipH: child.flipH, flipV: child.flipV,
  } : run, rect, text, run.vmlTextPathInput, imageFillResourceKey, maximumGeometryWork);
}

export interface NativeReadingBlockScene {
  readonly occurrenceId: string;
  readonly authoredWrap: 'tight' | 'through';
  readonly memberRunIndices: readonly number[];
  readonly childSourceIds: readonly string[];
  readonly drawing: DrawingLayout;
  readonly observedPaintOperations: number;
}

/** Library reading policy: relocate the complete group into body flow.
 * Authored tight/through contours remain source facts and are not evaluated.
 * Acquire one complete vector group from the retained paragraph. Every declared
 * source member must occur exactly once, in source order; unsupported content
 * aborts the whole scene. The original paragraph and authored wrap facts stay
 * intact. No resource, textbox, clip, unknown command or unsupported shape is
 * converted to a placeholder. This boundary does not authorize parsing an
 * unsupported native record or publishing pages/notices; those owners must opt
 * in and complete their independent atomic publication gates.
 *
 * Bounds come from the retained shared geometry, not an anchor/frame silhouette.
 * The authored extent is additionally reserved to retain empty group space.
 * Relative sizes require page-dependent resolution and are outside this local
 * scene class; the caller cannot silently discard them. */
export function acquireNativeReadingBlockScene(
  paragraph: ParagraphAcquisitionInput,
  occurrenceId: string,
  source: SourceRef,
  id: string,
  flowDomainId: string,
  maximumWork: number,
): NativeReadingBlockScene {
  if (!Number.isSafeInteger(maximumWork) || maximumWork < 1) fail('budget', 'invalid work allowance');
  if (!isDeepFrozenPlainDataRoot(paragraph)) fail('source-ownership', 'paragraph must be internally sealed');
  if (!occurrenceId || !id || !flowDomainId) fail('source-ownership', 'missing retained identity');
  if (paragraph.runs.length > maximumWork) fail('budget', 'paragraph run allowance exceeded');
  let hosts = 0;
  const members: { run: Extract<ParagraphAcquisitionInput['runs'][number], { type: 'shape' | 'image' | 'chart' | 'unavailableDrawing' }>; runIndex: number; acquisition: AnchorAcquisitionInput }[] = [];
  for (let runIndex = 0; runIndex < paragraph.runs.length; runIndex++) {
    const run = paragraph.runs[runIndex];
    if (run.type === 'anchorHost' && run.anchorOccurrenceId === occurrenceId) hosts++;
    if (run.type !== 'shape' && run.type !== 'image' && run.type !== 'chart' && run.type !== 'unavailableDrawing') continue;
    if (run.anchorAcquisitionInput?.occurrenceId === occurrenceId) {
      members.push({ run, runIndex, acquisition: run.anchorAcquisitionInput });
    }
  }
  if (hosts !== 1 || members.length === 0) fail('source-ownership', 'one host and a nonempty group are required');
  const first = members[0].acquisition;
  const count = first.group?.sourceCount;
  if (!Number.isSafeInteger(count) || count !== members.length) fail('member-coverage', 'declared group coverage is incomplete');
  const width = first.extent.widthPt, height = first.extent.heightPt;
  if (first.extent.widthStatus !== 'valid' || first.extent.heightStatus !== 'valid'
    || width === null || height === null || width <= 0 || height <= 0
    || !Number.isFinite(width) || !Number.isFinite(height)) fail('geometry', 'invalid authored group extent');
  if (first.wrap.kind !== 'tight' && first.wrap.kind !== 'through') fail('source-ownership', 'requires retained authored tight/through');
  if (first.wrap.polygon !== null) fail('source-ownership', 'stored contour requires the ordinary contour route');
  const outer: LayoutRect = { xPt: 0, yPt: 0, widthPt: width, heightPt: height };
  const commands: DrawingPaintCommand[] = [], childSourceIds: string[] = [];
  const seenChildren = new Set<string>();
  let work = paragraph.runs.length, resolvedWork = 0;
  for (let index = 0; index < members.length; index++) {
    const { run, acquisition } = members[index];
    const group = acquisition.group;
    if (!group || group.sourceIndex !== index || group.sourceCount !== count
      || !group.childSourceId || seenChildren.has(group.childSourceId)) fail('member-coverage', 'members must be unique and ordered');
    seenChildren.add(group.childSourceId); childSourceIds.push(group.childSourceId);
    if (acquisition.extent.widthStatus !== 'valid' || acquisition.extent.heightStatus !== 'valid'
      || acquisition.extent.widthPt !== width || acquisition.extent.heightPt !== height
      || acquisition.wrap.kind !== first.wrap.kind || acquisition.wrap.polygon !== null) fail('source-ownership', 'members disagree on their outer owner');
    if (acquisition.relativeSize.horizontal !== null || acquisition.relativeSize.vertical !== null) fail('geometry', 'page-dependent relative extent');
    const child = group.resolvedChildFrame;
    if (![child.offsetXPt, child.offsetYPt, child.widthPt, child.heightPt, child.rotationDeg].every(Number.isFinite)
      || child.widthPt < 0 || child.heightPt < 0 || typeof child.flipH !== 'boolean' || typeof child.flipV !== 'boolean') fail('geometry', 'invalid resolved child frame');
    if (run.type !== 'shape') fail('resource', `member ${index} requires ${run.type} acquisition and decoded-resource bounds`);
    if (run.textBoxInput || run.textBlocks?.length || run.textPath || run.vmlTextPathInput) fail('owned-text', `member ${index} requires complete story/glyph ink acquisition`);
    if (run.fill?.fillType === 'image') fail('resource', `member ${index} has an image fill`);
    // Charge the input before the shared planner clones it, not after painting.
    work += 1 + (run.adjValues?.length ?? 0) + (run.strokeCustomDash?.length ?? 0);
    for (const path of run.subpaths) { work += path.length + 1; if (work + resolvedWork + commands.length >= maximumWork) fail('budget', 'shape source allowance exceeded'); }
    if (work + resolvedWork + commands.length >= maximumWork) fail('budget', 'shape source allowance exceeded');
    if (run.subpathPaint && run.subpathPaint.length !== run.subpaths.length) fail('paint', 'per-path paint ownership mismatch');
    let planned: ShapeDrawingPlanResult;
    try { planned = planRetainedAnchorShape(run, acquisition, retainedAnchorChildFrame(acquisition, outer), undefined, undefined, maximumWork - work - resolvedWork - commands.length - 1); }
    catch (error) { fail(error instanceof GeometryWorkBudgetError ? 'budget' : 'paint', error instanceof Error ? error.message : 'geometry resolution failed'); }
    if (planned.status !== 'planned' || planned.command.kind !== 'drawingml-shape') fail('paint', `member ${index} cannot retain a vector command`);
    resolvedWork += planned.command.plan.resolvedGeometry.workUnits;
    if (work + resolvedWork + commands.length + 1 > maximumWork) fail('budget', 'resolved geometry allowance exceeded');
    commands.push(planned.command);
  }
  const drawing = snapshotPlainData({
    kind: 'drawing' as const, id, source, flowDomainId,
    flowBounds: outer, inkBounds: outer, advancePt: 0, ordinaryFlow: true, commands,
  }, 'native reading scene');
  if (maximumWork - work < 1) fail('budget', 'no work remains for shared geometry enclosure');
  let painted;
  try { painted = deriveNativeWrapPaintExtent(drawing, maximumWork - work - resolvedWork - commands.length - 1); }
  catch (error) { fail(error instanceof NativeWrapPaintBudgetError ? 'budget' : 'paint', error instanceof Error ? error.message : 'shared painter failed'); }
  const paint = painted.bounds;
  const left = Math.min(0, paint?.xPt ?? 0), top = Math.min(0, paint?.yPt ?? 0);
  const right = Math.max(width, paint ? paint.xPt + paint.widthPt : width);
  const bottom = Math.max(height, paint ? paint.yPt + paint.heightPt : height);
  const reserved = { xPt: left, yPt: top, widthPt: right - left, heightPt: bottom - top };
  if (!Object.values(reserved).every(Number.isFinite)) fail('geometry', 'non-finite reserved scene');
  return snapshotPlainData({
    occurrenceId, authoredWrap: first.wrap.kind, memberRunIndices: members.map(member => member.runIndex), childSourceIds,
    drawing: { ...drawing, flowBounds: reserved, inkBounds: paint ?? outer, advancePt: reserved.heightPt },
    observedPaintOperations: painted.observedOperations,
  }, 'complete native reading scene');
}

/** Place the complete scene into one resolved horizontal flow slot. The exact
 * available slot is supplied by the paginator for this page/variant. Reject
 * overflow; never shrink, clip or split a member to force admission. Commands
 * preserve their relative geometry/order, translated together into the slot. */
export function placeNativeReadingBlockScene(scene: NativeReadingBlockScene, slot: LayoutRect): DrawingLayout {
  if (!isDeepFrozenPlainDataRoot(scene)) fail('source-ownership', 'scene must be internally sealed');
  if (!Object.values(slot).every(Number.isFinite) || slot.widthPt < 0 || slot.heightPt < 0) fail('placement', 'invalid resolved flow slot');
  const bounds = scene.drawing.flowBounds;
  if (bounds.widthPt > slot.widthPt || bounds.heightPt > slot.heightPt) fail('placement', 'complete scene exceeds the resolved flow slot');
  const placed = translateDrawing(scene.drawing, { xPt: slot.xPt - bounds.xPt, yPt: slot.yPt - bounds.yPt });
  if (!Object.values(placed.flowBounds).every(Number.isFinite) || !Object.values(placed.inkBounds).every(Number.isFinite)) fail('placement', 'translated scene overflow');
  return snapshotPlainData(placed, 'placed native reading scene', scene);
}
