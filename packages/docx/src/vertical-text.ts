import { graphemeClusterOffsets } from '@silurus/ooxml-core';
import {
  bidiLevelChanges,
  CONDENSED_CONTEXT_GRAPHEMES,
  hasCursiveJoining,
  textSpacingAdvance,
} from './character-spacing.js';
// ECMA-376 §17.6.20 vertical writing (`<w:textDirection w:val="tbRl">`) — the
// glyph-level primitives for rendering a page that has been laid out in the
// SWAPPED logical coordinate space and rotated +90° into physical space by the
// renderer's page transform (see `renderDocumentToCanvas`).
//
// After the page transform, a normal `ctx.fillText(text, x, baseline)` paints
// the run flowing DOWNWARD in physical space (logical +x → physical +y), which
// is exactly the character progression a vertical line wants — but every glyph
// is lying on its right side (rotated +90° CW with the page). The per-glyph
// orientation is decided by the Unicode UAX#50 Vertical_Orientation (vo)
// property (core `verticalOrientation`), NOT by an ad-hoc CJK-vs-Latin guess:
//   • vo=U  (upright): CJK ideographs, kana, Hangul, fullwidth forms. Drawn with
//     a local −90° counter-rotation about the glyph's own centre, cancelling the
//     page rotation so it stands UPRIGHT while still advancing down the line.
//   • vo=Tu (transform, fallback upright): 、。，．！？ and small kana. When
//     this DOM canvas/font proves that `vert` changes the code point, the original
//     character is drawn with its featured A/D placement. Otherwise the existing
//     FE10–FE12 / upright base-glyph machinery remains the fallback.
//   • vo=Tr (transform, fallback rotate): the fullwidth brackets （「」〈〉【】… and
//     the white lenticular brackets 〖〗 have a U+FE1x/FE3x vertical presentation
//     form (core `verticalBracketFormSubstitute`) present in the substitute fonts;
//     UAX#50 §5 makes Tr "substitute a vertical glyph, ROTATE only as fallback", so
//     we SUBSTITUTE and draw them upright (Word/PowerPoint-verified, #969). A
//     reachable Tr code point instead uses its original `vert` glyph and featured
//     cell. Unreachable Tr points keep the substitution/geometric fallbacks:
//       – ROTATE (plain): the quotes “” and the fullwidth colon ：— drawn CENTRED on
//         the column via a plain `fillText` in the +90° page frame. The rotation IS
//         the font's designed vertical form for these (font-verified: the quotes'
//         comma-hooks match, and the colon's FE13 side-by-side dots fall out of the
//         base rotation since its FE13 form is absent from most render fonts).
//       – ROTATE (plain): unreachable long-stroke marks ー 〜 ～ use the UAX #50 Tr
//         fallback with no mirror/shear. Main-thread and worker/skia output may differ
//         because only the DOM path can verify the font's real `vert` design.
//       – UPRIGHT: the fullwidth semicolon ；, whose FE14 form is an upright dot-over-
//         comma, not a rotation (issue #969 follow-up; core `verticalTrUprightFallback`).
//   • vo=R  (rotated): Latin letters, Western digits, Latin punctuation. Stay
//     SIDEWAYS (rotated with the page) — the conventional "縦中横 not applied"
//     appearance — drawn as an ordinary contextual `fillText` at the alphabetic
//     baseline, preserving the browser's shaping/advance for the run.
//
// This module owns ONLY the pure geometry + classification; the renderer wires
// it into the whole-run glyph draw sites, the anchor/inline/float image draws,
// and the text-selection overlay behind the `verticalCJK` flag, so the
// horizontal path stays byte-identical.
//
// SCOPE (issue #771). Implemented: +90° page rotation; vo-driven upright(U) /
// substituted-upright(Tu) / rotated(Tr) / sideways(R) glyph draw; anchor images
// resolved against the physical page then projected into the logical flow
// (PDF-verified centroid); inline/anchored/float image uprighting; and the
// vertical text-layer transform. Still approximated / deferred (flagged inline):
// an unreachable U+FF0E vertical alternate retains the upright base glyph;
// paragraph-relative vertical anchors are a follow-up. `btLr` shares the +90°
// page FRAME but bypasses this module's
// upright/substitute glyph handling entirely (issue #988 re-adjudication: every
// glyph rides the page rotation — see BodyAcquisitionState.verticalAllRotated).

import { wordPreservesVerticalTuCorner } from './layout/script-compatibility.js';
import type { WritingMode } from './layout/types.js';
import {
  verticalOrientation,
  verticalFormSubstitute,
  verticalBracketFormSubstitute,
  verticalTrUprightFallback,
  measureVerticalVertGlyph,
  verticalVertGlyphReachable,
  withVertFeature,
  emBoxCenterAboveBaseline,
  inkCenterAboveMiddlePx,
} from '@silurus/ooxml-core';

// Browser acceptance tests load this source module through Vite's `/src` bridge;
// re-export the exact production feature operations so they exercise the same
// capability gate and restoration contract instead of test-local approximations.
export { verticalVertGlyphReachable, withVertFeature } from '@silurus/ooxml-core';

/** How a code point is painted inside the +90°-rotated vertical page:
 *   - `upright`  — counter-rotated −90° to stand up (vo=U, and vo=Tu).
 *   - `rotate`   — left rotated with the page but CENTRED on the column axis,
 *                  the UAX#50 Tr fallback for full-width brackets / 長音符.
 *   - `sideways` — left rotated with the page at the alphabetic baseline (vo=R,
 *                  Latin/digits). */
export type VerticalDrawMode = 'upright' | 'rotate' | 'sideways';

/**
 * The draw mode for a code point in vertical text, from its UAX#50
 * Vertical_Orientation (vo). Single source of truth: core `verticalOrientation`.
 *
 *   U  → upright   (stand the glyph up)
 *   Tu → upright   (draw upright; the caller substitutes a vertical form glyph
 *                   via {@link verticalFormSubstitute} when one exists so the
 *                   comma/full stop land in the upper-right of the cell)
 *   Tr → rotate    (rotate 90° CW — the fallback when the element/CSS route or
 *                   that glyph's vertical coverage is unavailable — centred)
 *   R  → sideways  (leave rotated with the page: Latin/digits)
 *
 * @param cp A Unicode scalar value (e.g. from `String.prototype.codePointAt`).
 */
export function verticalDrawMode(cp: number): VerticalDrawMode {
  const vo = verticalOrientation(cp);
  if (vo === 'U' || vo === 'Tu') return 'upright';
  if (vo === 'Tr') return 'rotate';
  return 'sideways'; // vo === 'R'
}

/**
 * True when `cp` stands UPRIGHT in vertical text (UAX#50 vo ∈ {U, Tu}). Kept for
 * callers that only need the upright/not-upright split; new code should prefer
 * {@link verticalDrawMode} which also distinguishes the Tr (rotate) case.
 *
 * @param cp A Unicode scalar value (e.g. from `String.prototype.codePointAt`).
 */
export function isUprightVerticalGlyph(cp: number): boolean {
  return verticalDrawMode(cp) === 'upright';
}

/**
 * Split a run's text into maximal runs of same-draw-mode code points (UAX#50 vo,
 * via {@link verticalDrawMode}), so the vertical draw path can counter-rotate
 * the UPRIGHT segments per glyph, rotate the Tr segments, and draw the SIDEWAYS
 * (Latin/digit) segments as a single contextual `fillText`. Preserves surrogate
 * pairs (iterates by code point) and returns the pieces in logical order.
 *
 * @param text The run's text.
 * @returns Ordered pieces, each `{ text, mode }`.
 */
export function splitVerticalOrientationRuns(
  text: string,
): Array<{ text: string; mode: VerticalDrawMode }> {
  const pieces: Array<{ text: string; mode: VerticalDrawMode }> = [];
  let cur = '';
  let curMode: VerticalDrawMode | null = null;
  for (const ch of text) {
    const cp = ch.codePointAt(0) ?? 0;
    const mode = verticalDrawMode(cp);
    if (curMode === null) {
      curMode = mode;
      cur = ch;
    } else if (mode === curMode) {
      cur += ch;
    } else {
      pieces.push({ text: cur, mode: curMode });
      cur = ch;
      curMode = mode;
    }
  }
  if (cur !== '' && curMode !== null) {
    pieces.push({ text: cur, mode: curMode });
  }
  return pieces;
}

type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
type VertCapability = (cp: number) => boolean;
const NO_VERT_CAPABILITY: VertCapability = () => false;

/**
 * True for a vo=Tr code point that takes the GEOMETRIC ROTATE fallback in
 * {@link drawVerticalRun} — i.e. `mode==='rotate'` with NO substituted vertical
 * bracket form and NOT the upright-fallback semicolon. These are the marks drawn
 * by a plain `fillText` in the +90° page frame: ー 〜 ～,
 * the quotes “”, and the colon ：. This is the SINGLE predicate shared by the
 * paint path and the {@link verticalRunInkExtraPx} measure path (issue #1014), so
 * the two agree on which glyphs get the ink-sized cell.
 *
 * @param cp A Unicode scalar value.
 */
function isVerticalRotateFallback(cp: number): boolean {
  return (
    verticalDrawMode(cp) === 'rotate' &&
    verticalBracketFormSubstitute(cp) === null &&
    !verticalTrUprightFallback(cp)
  );
}

/** Only UAX #50 transform classes may replace the manual fallback with a
 * feature-selected glyph. U and R remain byte-identical regardless of what a
 * capability callback reports. */
function isVerticalVertCandidate(cp: number): boolean {
  const vo = verticalOrientation(cp);
  return vo === 'Tu' || vo === 'Tr';
}

/**
 * Along-column ink geometry of a vo=Tr rotate-fallback glyph (issue #1014). The
 * glyph is painted by a plain `fillText` in the +90°-rotated page frame, so its
 * HORIZONTAL ink extent maps onto the ALONG-COLUMN axis (the advance axis). Read
 * the tight horizontal ink box with a `center`/`middle` alignment:
 *   - `extentPx` = actualBoundingBoxLeft + actualBoundingBoxRight — the ink width
 *     along the column (used to size the cell so the ink cannot spill past it).
 *   - `shiftPx`  = (actualBoundingBoxLeft − actualBoundingBoxRight)/2 — the local
 *     along-column shift that re-centres the ink on the (grown) cell, since a
 *     `center` draw centres the glyph's ADVANCE and an under-reported advance is
 *     off-centre from the ink.
 * Returns `null` when the Canvas does not report `actualBoundingBox*` (older
 * engines / node mocks) so callers degrade to the advance-sized, advance-centred
 * draw exactly as before this metric existed.
 */
function verticalRotateInkGeometry(
  ctx: Ctx2D,
  ch: string,
): { extentPx: number; shiftPx: number } | null {
  const prevAlign = ctx.textAlign;
  const prevBaseline = ctx.textBaseline;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  const m = ctx.measureText(ch);
  ctx.textAlign = prevAlign;
  ctx.textBaseline = prevBaseline;
  const l = m.actualBoundingBoxLeft;
  const r = m.actualBoundingBoxRight;
  if (
    typeof l !== 'number' ||
    typeof r !== 'number' ||
    !Number.isFinite(l) ||
    !Number.isFinite(r)
  ) {
    return null;
  }
  return { extentPx: l + r, shiftPx: (l - r) / 2 };
}

interface RoutedVerticalGlyphCell {
  naturalPx: number;
  vert: ReturnType<typeof measureVerticalVertGlyph> | null;
  rotateInkShiftPx: number;
}

export interface PlannedVerticalGlyphCell {
  readonly range: Readonly<{ start: number; end: number }>;
  readonly text: string;
  readonly orientation: 'upright' | 'rotate' | 'sideways';
  readonly originPt: number;
  readonly advancePt: number;
  readonly drawOffsetPt: Readonly<{ xPt: number; yPt: number }>;
  readonly verticalFeature: boolean;
  readonly blockAxisInkBounds?: Readonly<{ startPt: number; endPt: number }>;
  /** WORD_CONDENSED_SPACING_ADVANCE_FLOOR took this cell's advance (or, for
   * an unsplit cursive piece, one of its graphemes' advances) to zero. */
  readonly floored?: true;
  /** Drawn advance of a floored sideways cluster painted on its own (its
   * isolated shaped width), for direction-aware placement of its ink. */
  readonly naturalAdvancePt?: number;
}

function plannedVerticalBlockAxisInkBounds(
  ctx: Ctx2D,
  text: string,
  orientation: PlannedVerticalGlyphCell['orientation'],
  drawOffsetPt: Readonly<{ xPt: number; yPt: number }>,
  charScale: number,
  writingMode: WritingMode,
  verticalFeature: boolean,
): PlannedVerticalGlyphCell['blockAxisInkBounds'] {
  const previousAlign = ctx.textAlign;
  const previousBaseline = ctx.textBaseline;
  const measure = (): TextMetrics => {
    ctx.textAlign = orientation === 'sideways' ? 'left' : 'center';
    ctx.textBaseline = orientation === 'sideways' ? 'alphabetic' : 'middle';
    return ctx.measureText(text);
  };
  let metrics: TextMetrics;
  try {
    metrics = verticalFeature
      ? withVertFeature(ctx as CanvasRenderingContext2D, measure)
      : measure();
  } finally {
    ctx.textAlign = previousAlign;
    ctx.textBaseline = previousBaseline;
  }
  if (orientation === 'upright') {
    if (!Number.isFinite(metrics.actualBoundingBoxLeft)
      || !Number.isFinite(metrics.actualBoundingBoxRight)) return undefined;
    const leftPt = metrics.actualBoundingBoxLeft;
    const rightPt = metrics.actualBoundingBoxRight;
    // Paint applies rotate(-90deg) after the page transform. In vertical-rl,
    // w:w scales glyph-local y (the cross axis), not local x (the logical block
    // axis). Other frames scale local x, exactly as canvas-text does.
    const blockScale = writingMode === 'vertical-rl' ? 1 : charScale;
    const firstPt = -(drawOffsetPt.xPt - leftPt) * blockScale;
    const secondPt = -(drawOffsetPt.xPt + rightPt) * blockScale;
    return Object.freeze({
      startPt: Math.min(firstPt, secondPt),
      endPt: Math.max(firstPt, secondPt),
    });
  }
  if (!Number.isFinite(metrics.actualBoundingBoxAscent)
    || !Number.isFinite(metrics.actualBoundingBoxDescent)) return undefined;
  const ascentPt = metrics.actualBoundingBoxAscent;
  const descentPt = metrics.actualBoundingBoxDescent;
  const firstPt = drawOffsetPt.yPt - ascentPt;
  const secondPt = drawOffsetPt.yPt + descentPt;
  return Object.freeze({
    startPt: Math.min(firstPt, secondPt),
    endPt: Math.max(firstPt, secondPt),
  });
}

// UAX #50 §3.1 / Vertical_Orientation=Tu permits an upright base glyph when the
// vertical alternate is unavailable. OpenType `vert` selects a distinct glyph;
// `vrt2` additionally supplies rotation and must not be stacked with `vert`.
// ECMA-376 §§17.6.20, 17.18.93 specify flow, not a punctuation translation.
//
// Windows Word 16.0 build 16.0.20430: seven reference-catalog CJK families at
// 8/12/20/32pt, section tbRl and ideographic VML textboxes, with horizontal
// counterexamples, select different glyph IDs for ．、。，. The U+FF0E alternate
// can change outline dimensions as well as position. Neither horizontal Canvas
// ink bounds nor the catalog's font-wide metrics identify that alternate or its
// vertical origin. No Unicode vertical presentation character exists for FF0E.
// Therefore the reachable-feature path below keeps the font's design, while
// the unreachable path uses ordinary upright placement, including its measured
// along-column ink centering. It deliberately does not synthesize a corner
// glyph by translating the horizontal one. This fallback has residuals in all
// seven tested families; worker/offscreen hosts can differ from the DOM path.
// The FE10–FE12 fallbacks for 、。， retain their existing designed placement.

/** The single cell router shared by layout measurement and paint. */
function routedVerticalGlyphCell(
  ctx: Ctx2D,
  ch: string,
  cp: number,
  vertCapability: VertCapability,
  growRotateInk: boolean,
): RoutedVerticalGlyphCell {
  const plainAdvance = ctx.measureText(ch).width;
  if (isVerticalVertCandidate(cp) && vertCapability(cp)) {
    const vert = measureVerticalVertGlyph(ctx, ch);
    return { naturalPx: vert.cellAdvancePx, vert, rotateInkShiftPx: 0 };
  }
  if (growRotateInk && isVerticalRotateFallback(cp)) {
    const geom = verticalRotateInkGeometry(ctx, ch);
    if (geom !== null && geom.extentPx > plainAdvance) {
      return {
        naturalPx: geom.extentPx,
        vert: null,
        rotateInkShiftPx: geom.shiftPx,
      };
    }
  }
  return { naturalPx: plainAdvance, vert: null, rotateInkShiftPx: 0 };
}

/** Resolve the legacy UAX #50/`vert` routing into immutable glyph cells for the
 * retained layout pipeline. The returned plan contains every paint decision and
 * metric; retained paint never probes fonts or remeasures text. */
export function planVerticalRunWithCapability(
  ctx: Ctx2D,
  text: string,
  fontPt: number,
  letterSpacingPt: number,
  charScale = 1,
  growTrRotateInk = false,
  vertCapability: VertCapability = NO_VERT_CAPABILITY,
  writingMode: WritingMode = 'vertical-rl',
  characterGridDeltaPt = 0,
  characterAdjustmentPt?: (start: number, end: number) => number,
): readonly PlannedVerticalGlyphCell[] {
  const cells: PlannedVerticalGlyphCell[] = [];
  const emBoxCenterPt = emBoxCenterAboveBaseline(ctx, text, fontPt);
  let ax = 0;
  let sourceOffset = 0;
  const charSpacingPt = letterSpacingPt - characterGridDeltaPt;
  // WORD_CONDENSED_SPACING_ADVANCE_FLOOR, only when acquisition opts in by
  // supplying `characterAdjustmentPt` (punctuation compression and width
  // balance of a source range, floored with the cell): no character advance
  // below zero, and the reported advances include the adjustment. Without it
  // the plan is the ordinary unfloored one. Unfloored cells keep their centring.
  const mayFloor = characterAdjustmentPt !== undefined
    && (letterSpacingPt < 0 || characterGridDeltaPt < 0);
  const unitAdvance = (unit: string, start: number, scaledNaturalPt: number): Readonly<{
    advancePt: number;
    cellPt: number;
    floored: boolean;
  }> => {
    const spacedPt = scaledNaturalPt
      + textSpacingAdvance(unit, characterGridDeltaPt, charSpacingPt, letterSpacingPt);
    if (!mayFloor) return { advancePt: spacedPt, cellPt: spacedPt, floored: false };
    const rawPt = spacedPt + characterAdjustmentPt!(start, start + unit.length);
    if (rawPt < 0) return { advancePt: 0, cellPt: 0, floored: true };
    return { advancePt: rawPt, cellPt: spacedPt, floored: false };
  };
  for (const piece of splitVerticalOrientationRuns(text)) {
    if (piece.mode === 'sideways') {
      const drawOffsetPt = { xPt: 0, yPt: emBoxCenterPt };
      const offsets = [...new Set([0, ...graphemeClusterOffsets(piece.text), piece.text.length])]
        .sort((a, b) => a - b);
      // Contextual grapheme advances (bounded preceding context, as line
      // fitting measures them): an over-condensed cluster stops the pen at its
      // floor, which a shared Canvas letterSpacing cannot express, so the
      // piece then paints each cluster from its own origin. Cursive text stays
      // one contextual cell (Canvas does not track inside cursive scripts)
      // advancing by its graphemes' floored sum.
      const units = mayFloor
        ? offsets.slice(0, -1).map((start, index) => {
            const end = offsets[index + 1]!;
            const context = offsets[Math.max(0, index - CONDENSED_CONTEXT_GRAPHEMES)]!;
            return {
              text: piece.text.slice(start, end),
              start: sourceOffset + start,
              naturalPt: ctx.measureText(piece.text.slice(context, end)).width
                - (context < start ? ctx.measureText(piece.text.slice(context, start)).width : 0),
            };
          })
        : [];
      const unitPlaces = units.map((unit) => unitAdvance(unit.text, unit.start, unit.naturalPt * charScale));
      // Floored: non-cursive graphemes paint from their own origins; each run
      // of cursive graphemes stays one contextual cell (Canvas does not track
      // inside cursive scripts) advancing by its graphemes' floored sum.
      // A cursive group keeps its classification and source range; its text
      // is sliced once when the cell is emitted, so grouping stays linear.
      const pieceCells: Array<{
        start: number;
        end: number;
        advancePt: number;
        floored: boolean;
        cursive: boolean;
      }> =
        unitPlaces.some((placed) => placed.floored)
          ? []
          : (() => {
              const placed = unitAdvance(piece.text, sourceOffset, ctx.measureText(piece.text).width * charScale);
              return [{
                start: sourceOffset,
                end: sourceOffset + piece.text.length,
                advancePt: placed.advancePt,
                floored: placed.floored,
                cursive: false,
              }];
            })();
      const split = pieceCells.length === 0;
      if (split) {
        // Split cells are placed in bidi order at paint, so a cursive group
        // stays level-uniform.
        const levelChanges = bidiLevelChanges(piece.text);
        units.forEach((unit, index) => {
          const placed = unitPlaces[index]!;
          const previous = pieceCells.at(-1);
          const cursive = hasCursiveJoining(unit.text);
          if (cursive && previous?.cursive && !levelChanges(unit.start - sourceOffset)) {
            previous.end = unit.start + unit.text.length;
            previous.advancePt += placed.advancePt;
            previous.floored ||= placed.floored;
          } else {
            pieceCells.push({
              start: unit.start,
              end: unit.start + unit.text.length,
              advancePt: placed.advancePt,
              floored: placed.floored,
              cursive,
            });
          }
        });
      }
      for (const pieceCell of pieceCells) {
        const cell = {
          ...pieceCell,
          text: piece.text.slice(pieceCell.start - sourceOffset, pieceCell.end - sourceOffset),
        };
        // A split non-cursive floored cell paints one isolated cluster whose
        // ink runs past its zero allocation; directional placement anchors
        // that drawn advance.
        const naturalAdvancePt = split && cell.floored && !cell.cursive
          ? ctx.measureText(cell.text).width * charScale
          : undefined;
        const blockAxisInkBounds = plannedVerticalBlockAxisInkBounds(
          ctx, cell.text, 'sideways', drawOffsetPt, charScale, writingMode, false,
        );
        cells.push({
          range: { start: cell.start, end: cell.start + cell.text.length },
          text: cell.text,
          orientation: 'sideways',
          originPt: ax,
          advancePt: cell.advancePt,
          drawOffsetPt,
          verticalFeature: false,
          ...(blockAxisInkBounds ? { blockAxisInkBounds } : {}),
          ...(cell.floored ? { floored: true as const } : {}),
          ...(naturalAdvancePt !== undefined ? { naturalAdvancePt } : {}),
        });
        ax += cell.advancePt;
      }
      sourceOffset += piece.text.length;
      continue;
    }
    for (const ch of piece.text) {
      const cp = ch.codePointAt(0) ?? 0;
      const mode = verticalDrawMode(cp);
      const bracketCp = mode === 'rotate' ? verticalBracketFormSubstitute(cp) : null;
      const uprightFallback = mode === 'rotate'
        && bracketCp === null
        && verticalTrUprightFallback(cp);
      const routed = routedVerticalGlyphCell(ctx, ch, cp, vertCapability, growTrRotateInk);
      const placed = unitAdvance(ch, sourceOffset, routed.naturalPx * charScale);
      const advancePt = placed.advancePt;
      const flooredCell = placed.floored ? { floored: true as const } : {};
      const range = { start: sourceOffset, end: sourceOffset + ch.length };
      if (routed.vert !== null) {
        const drawOffsetPt = { xPt: 0, yPt: 0 };
        const blockAxisInkBounds = plannedVerticalBlockAxisInkBounds(
          ctx, ch, 'upright', drawOffsetPt, charScale, writingMode, true,
        );
        cells.push({
          range,
          text: ch,
          orientation: 'upright',
          originPt: ax + routed.vert.originInCellPx * charScale,
          advancePt,
          drawOffsetPt,
          verticalFeature: true,
          ...(blockAxisInkBounds ? { blockAxisInkBounds } : {}),
          ...flooredCell,
        });
      } else if (mode === 'upright' || bracketCp !== null || uprightFallback) {
        const puncCp = bracketCp !== null ? null : verticalFormSubstitute(cp);
        const drawCp = bracketCp ?? puncCp;
        const drawText = drawCp === null ? ch : String.fromCodePoint(drawCp);
        const preserveCorner = wordPreservesVerticalTuCorner(puncCp);
        const alongEm = !preserveCorner
          ? inkCenterAboveMiddlePx(ctx, drawText) / fontPt
          : 0;
        const drawOffsetPt = {
          xPt: 0,
          yPt: alongEm * fontPt,
        };
        const blockAxisInkBounds = plannedVerticalBlockAxisInkBounds(
          ctx, drawText, 'upright', drawOffsetPt, charScale, writingMode, false,
        );
        cells.push({
          range,
          text: drawText,
          orientation: 'upright',
          originPt: ax + placed.cellPt / 2,
          advancePt,
          drawOffsetPt,
          verticalFeature: false,
          ...(blockAxisInkBounds ? { blockAxisInkBounds } : {}),
          ...flooredCell,
        });
      } else {
        const drawOffsetPt = { xPt: 0, yPt: 0 };
        const blockAxisInkBounds = plannedVerticalBlockAxisInkBounds(
          ctx, ch, 'rotate', drawOffsetPt, charScale, writingMode, false,
        );
        cells.push({
          range,
          text: ch,
          orientation: 'rotate',
          originPt: ax + placed.cellPt / 2 + charScale * routed.rotateInkShiftPx,
          advancePt,
          drawOffsetPt,
          verticalFeature: false,
          ...(blockAxisInkBounds ? { blockAxisInkBounds } : {}),
          ...flooredCell,
        });
      }
      ax += advancePt;
      sourceOffset += ch.length;
    }
  }
  return cells;
}

/**
 * ECMA-376 §17.6.20 (tbRl), issues #1014/#1024 — the along-column cell delta
 * (px, before §17.3.2.43 `w:w` scaling and §17.3.2.35 pitch) beyond ordinary
 * `measureText`. Reachable Tu/Tr glyphs contribute their feature advances;
 * unreachable upright/transform glyphs contribute their manual per-glyph
 * advances, with the positive #1014 ink-overrun growth where needed. Consecutive
 * vo=R glyphs contribute one contextual sideways-piece advance, preserving the
 * same shaping and kerning as paint. Subtracting the whole-string width removes
 * only kern pairs that cross independent vertical-cell boundaries.
 * Layout folds this into the
 * segment's natural advance (`segAdvanceWidth`'s `naturalWidthPx`) so the grown
 * cell {@link drawVerticalRun} paints is matched by the measured box — measure ==
 * paint (wrapping, the next run's position, and the selection overlay all track
 * the drawn cell). The value is the per-glyph cell sum minus the whole-string
 * `measureText`: exactly 0 when neither route changes a cell AND the font applies
 * no horizontal kern to the run (the common byte-identical path), and nonzero
 * where the whole-string measure was kern-compressed but the vertical cells are
 * not (issue #1024).
 *
 * The caller must set `ctx.font` (and any kerning state) for the run before
 * calling, exactly as it does for the `measureText` that produces `naturalWidthPx`.
 *
 * @param ctx  2D context with the run's font selected.
 * @param text The run's text.
 */
export function verticalRunInkExtraPxWithCapability(
  ctx: Ctx2D,
  text: string,
  vertCapability: VertCapability,
): number {
  let routedWidth = 0;
  for (const piece of splitVerticalOrientationRuns(text)) {
    if (piece.mode === 'sideways') {
      routedWidth += ctx.measureText(piece.text).width;
      continue;
    }
    for (const ch of piece.text) {
      const cp = ch.codePointAt(0) ?? 0;
      const routed = routedVerticalGlyphCell(ctx, ch, cp, vertCapability, true);
      routedWidth += routed.naturalPx;
    }
  }
  // Canvas whole-string measurement may apply horizontal JIS punctuation kern
  // pairs (for example 。「), but vertical paint advances those independent cells.
  // Subtract the SAME whole width the caller adds so their sum is exactly the
  // routed per-glyph width at every layout call site.
  return routedWidth - ctx.measureText(text).width;
}

export function verticalRunInkExtraPx(ctx: Ctx2D, text: string): number {
  return verticalRunInkExtraPxWithCapability(
    ctx,
    text,
    (cp) => verticalVertGlyphReachable(ctx, cp),
  );
}

/**
 * Draw one run's glyphs in vertical mode. The context is assumed to already be
 * in the page's SWAPPED logical frame (the +90° page rotation is installed by
 * `renderDocumentToCanvas`), so an ordinary `fillText` advances DOWN the line.
 *
 * The run flows along logical +x (physical +y). Each glyph occupies a cell of
 * width = its horizontal advance (`ctx.measureText`) plus `letterSpacingPx`
 * (the docGrid / justification pitch the layout measured the box with), so the
 * total advance equals the run's measured width — measure == draw. Upright
 * (CJK) glyphs are counter-rotated −90° about their cell centre so they stand
 * upright; sideways (Latin/digit) pieces are painted as a single contextual
 * `fillText`, preserving the browser's shaping.
 *
 * @param ctx              2D context, already in the rotated logical page frame.
 *                         `ctx.font`/`ctx.fillStyle` are set by the caller.
 * @param text             The run's text.
 * @param x                Logical left edge of the run (px).
 * @param baseline         Logical baseline y of the line (px).
 * @param fontPx           Effective font size in px (for cell centring).
 * @param letterSpacingPx  Per-glyph extra advance: the combined docGrid cell
 *                         delta + §17.3.2.35 `w:spacing` pitch (the layout's
 *                         `segLetterSpacingPx`); 0 for the common path.
 * @param charScale        ECMA-376 §17.3.2.43 `w:w` fraction; 1 by default.
 * @param growTrRotateInk  issue #1014 — when true, a vo=Tr GEOMETRIC rotate-fallback
 *                         glyph (ー 〜 ～ “” ：) whose substitute font under-reports
 *                         its advance is sized to its along-column INK extent (and
 *                         ink-centred) so its ink cannot spill past the cell into the
 *                         next run. MUST be set ONLY where the layout advance was
 *                         grown by the SAME deficit ({@link verticalRunInkExtraPx},
 *                         gated on `LayoutTextSeg.verticalRun`) so paint == measure;
 *                         the caller passes `s.verticalRun === true`. Default false
 *                         keeps the advance-sized, advance-centred draw byte-identical
 *                         (markers and unwired vertical text boxes).
 */
export function drawVerticalRunWithCapability(
  ctx: Ctx2D,
  text: string,
  x: number,
  baseline: number,
  fontPx: number,
  letterSpacingPx: number,
  charScale = 1,
  growTrRotateInk = false,
  vertCapability: VertCapability = NO_VERT_CAPABILITY,
): void {
  const prevAlign = ctx.textAlign;
  const prevBaseline = ctx.textBaseline;
  // Cross-axis (column-thickness) distance from the alphabetic baseline to the
  // font's em-box centre — the line the UPRIGHT cells centre on. Measured once
  // per run (font-level, glyph-independent). Used to re-centre SIDEWAYS glyphs,
  // which are otherwise drawn on their baseline and so land off the centreline.
  const emBoxCenterPx = emBoxCenterAboveBaseline(ctx, text, fontPx);
  // In this rotate-layout architecture the direction-independent layout kernel
  // applies ECMA-376 §17.3.2.43 `w:w` to the line axis (`segAdvanceWidth`) even
  // for tbRl; wrapping, selection, and run boxes already depend on that advance,
  // so paint must follow measure. Sideways glyphs are rotated horizontal text,
  // making their `w:w` width axis the vertical advance axis directly. Upright
  // glyphs therefore scale the equivalent local y axis. Tate-chu-yoko is kept
  // separate: ECMA-376 §17.3.2.10 fixes its advance to one em and uses `w:w` on
  // the cross axis.
  const scaled = charScale !== 1;
  let ax = 0; // cumulative advance from run left (logical +x)
  for (const piece of splitVerticalOrientationRuns(text)) {
    if (piece.mode === 'sideways') {
      // vo=R remains horizontal text inside the rotated page frame. Paint the
      // maximal piece in one call so contextual shaping/kerning is identical to
      // the advance used by layout (and therefore to btLr's rotated-horizontal
      // frame). Letter spacing is applied inside the shaped piece, while the
      // trailing pitch remains part of the explicit cell advance as before.
      const naturalPx = ctx.measureText(piece.text).width;
      const pieceAdvance = naturalPx * charScale + textSpacingAdvance(piece.text, 0, letterSpacingPx, letterSpacingPx);
      const prevLetterSpacing = ctx.letterSpacing;
      ctx.textAlign = prevAlign;
      ctx.textBaseline = 'alphabetic';
      ctx.save();
      if (scaled) {
        ctx.translate(x + ax, 0);
        ctx.scale(charScale, 1);
        ctx.letterSpacing = `${letterSpacingPx / charScale}px`;
        ctx.fillText(piece.text, 0, baseline + emBoxCenterPx);
      } else {
        ctx.letterSpacing = `${letterSpacingPx}px`;
        ctx.fillText(piece.text, x + ax, baseline + emBoxCenterPx);
      }
      ctx.restore();
      ctx.letterSpacing = prevLetterSpacing;
      ax += pieceAdvance;
      continue;
    }
    for (const ch of piece.text) {
    const cp = ch.codePointAt(0) ?? 0;
    const mode = verticalDrawMode(cp);
    // A vo=Tr code point with a substituted Unicode vertical presentation form — the
    // brackets （）「」〈〉… and the white lenticular 〖〗 (#969) — is SUBSTITUTED and
    // drawn upright, exactly like the upright cells — UAX#50 §5 Tr means "substitute a
    // vertical glyph; rotate only as fallback". Tr code points with NO substituted form
    // (ー, quotes “”, and the colon ：/ semicolon ；whose FE13/FE14 forms are absent
    // from most render fonts) take a geometric fallback below (rotate, or — for ；—
    // upright).
    const bracketCp = mode === 'rotate' ? verticalBracketFormSubstitute(cp) : null;
    // A vo=Tr code point with NO substituted vertical form whose fallback is
    // UPRIGHT rather than the generic UAX#50 §5 ROTATE — the fullwidth semicolon
    // ；(FF1B), whose FE14 vertical form is an upright dot-over-comma, not a
    // rotation (UAX #50 / issue #969). It draws upright exactly
    // like the vo=U / vo=Tu cells; the colon ：is NOT here (its FE13 form IS a 90°
    // rotation, so it takes the rotate branch below → side-by-side dots).
    const uprightFallback = mode === 'rotate' && bracketCp === null && verticalTrUprightFallback(cp);
    // Advance/width uses the ORIGINAL code point (measure == draw, and the text
    // model / selection / find keep the original character — see the module doc).
    // #1014: a vo=Tr GEOMETRIC rotate-fallback glyph (ー 〜 ～ “” ：) is painted by a
    // plain `fillText` in the +90° page frame, so its HORIZONTAL ink maps onto the
    // along-column (advance) axis. When a substitute font UNDER-REPORTS the advance
    // (Chrome), that ink spills PAST the advance-sized cell into the next run. Size
    // the cell to the along-column INK extent instead so the ink is contained; the
    // SAME per-glyph deficit is folded into the layout advance by
    // `verticalRunInkExtraPx` (measure == draw). NO-OP unless the ink exceeds the
    // advance (every real font here reports ink ≤ advance ⇒ byte-identical), and only
    // for the geometric rotate branch (substituted/upright Tr glyphs keep their path).
    const routedCell = routedVerticalGlyphCell(ctx, ch, cp, vertCapability, growTrRotateInk);
    const vertCell = routedCell.vert;
    const cellNaturalPx = routedCell.naturalPx;
    const rotateInkShiftPx = routedCell.rotateInkShiftPx;
    const adv = cellNaturalPx * charScale + textSpacingAdvance(ch, 0, letterSpacingPx, letterSpacingPx);
    if (vertCell !== null) {
      // Preserve the feature glyph's asymmetric A/D placement by keeping its
      // origin at the nominal half-advance measured under the same composed
      // `vert` state as paint. Designed ink may poke into a neighbour cell;
      // letter spacing follows the cell and does not move the origin.
      const cx = x + ax + vertCell.originInCellPx * charScale;
      ctx.save();
      ctx.translate(cx, baseline);
      ctx.rotate(-Math.PI / 2);
      if (scaled) ctx.scale(1, charScale);
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      withVertFeature(ctx, () => ctx.fillText(ch, 0, 0));
      ctx.restore();
    } else if (mode === 'upright' || bracketCp !== null || uprightFallback) {
      // vo=U / Tu, or a substituted Tr bracket. Counter-rotate −90° about the
      // cell centre so the glyph (which the page rotation would otherwise lay on
      // its side) stands upright. For corner-hanging Tu punctuation with a Unicode
      // vertical form (、。， → U+FE10–FE12) and Tr brackets (（）「」… → U+FE35–FE44)
      // draw THAT glyph so the font supplies the vertical shape; the original
      // advance is kept. Substitution is a GLYPH-only change: the width above and
      // everything the renderer reports (selection, find) use the original `ch`.
      // ！？ are NOT substituted (see verticalFormSubstitute) — they draw upright
      // as the original fullwidth mark, which is already centred on the column.
      // A substituted Tu punctuation form (、。， → FE10–FE12) vs. everything else.
      // The Tr bracket substitute is tracked separately by `bracketCp`.
      const puncCp = bracketCp !== null ? null : verticalFormSubstitute(cp);
      const drawCp = bracketCp !== null ? bracketCp : puncCp;
      const drawStr = drawCp !== null ? String.fromCodePoint(drawCp) : ch;
      const cx = x + ax + adv / 2;
      // ALONG-COLUMN centring: an upright glyph's VERTICAL ink extent maps to the
      // column axis, so shift by its measured ink centre so the ink lands on the
      // cell centre. Per-GLYPH metric (the drawn glyph's tight ink box): for an
      // ideograph/kana it is ≈0 (cells unchanged); for a substituted vertical
      // bracket (ink hugging one cell end) it is the needed correction. An
      // unreachable Tu base glyph (including ．) uses this ordinary upright
      // fallback; its horizontal ink cannot identify a vertical alternate.
      //
      // Under `word-vertical-tu-corner-placement`, this is NOT applied to a
      // substituted Tu punctuation form (comma/full stop 、。， → FE10–FE12):
      // those glyphs are designed with their ink in the cell's upper-right
      // corner (JIS X 4051 §4.3 kutōten placement). Ink-
      // centring would force that intentional offset back to the geometric cell
      // centre, dropping the comma/full stop LOW — the reported "、。 sit too low"
      // defect (#771). Drawing them em-box-centred preserves the font's corner
      // design. The Tr brackets DO get the correction: their two halves must sit a
      // full cell apart and the font centres the em box, not the ink (#792).
      const isPunctSubstitute = wordPreservesVerticalTuCorner(puncCp);
      const alongEm =
        !isPunctSubstitute
          ? inkCenterAboveMiddlePx(ctx, drawStr) / fontPx
          : 0;
      ctx.save();
      ctx.translate(cx, baseline);
      ctx.rotate(-Math.PI / 2);
      if (scaled) ctx.scale(1, charScale);
      // In the upright local frame: `center`/`middle` puts the em box on the cell
      // centre; local +x = cross axis, local +y = along-column. Only the
      // measured along-column correction is applied to ordinary upright glyphs.
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(drawStr, 0, alongEm * fontPx);
      ctx.restore();
    } else if (mode === 'rotate') {
      // vo=Tr with NO substituted vertical form and NOT the upright-fallback
      // semicolon: ー (U+30FC), the wave dash / tilde 〜 ～, the double quotes “”,
      // and the fullwidth colon ：(FF1A). UAX#50's Tr fallback (no vertical glyph
      // available through the element/CSS route) is to ROTATE the glyph 90° CW;
      // a plain `fillText` in the +90° page frame IS that rotation, centred with
      // `center`/`middle` at the cell centre. For the colon this reproduces FE13's
      // design directly (the two vertically stacked dots become side by side);
      // for quotes, the rotation follows the font's designed vertical form.
      //
      // An unreachable `ー〜～` uses this same plain UAX #50 Tr rotation. The removed
      // #1017/#1023 mirror/shear extrapolated an inaccessible glyph design from two
      // Mincho fonts; worker/skia may therefore differ visibly from the real DOM
      // `vert` path, which is a documented limitation rather than a fabricated form.
      const cx = x + ax + adv / 2;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      // #1014: `rotateInkShiftPx` (glyph-space, non-zero ONLY when the cell was grown
      // to the ink extent above) re-centres the ink on the grown cell — a `center`
      // draw centres the glyph's ADVANCE, and an under-reported advance is off-centre
      // from the ink. It is applied separately on the advance OUTPUT axis before
      // the §17.3.2.43 `w:w` matrix. Zero leaves advance centring unchanged.
      if (scaled || rotateInkShiftPx !== 0) {
        ctx.save();
        ctx.translate(cx, baseline);
        if (rotateInkShiftPx !== 0) ctx.translate(charScale * rotateInkShiftPx, 0);
        ctx.transform(charScale, 0, 0, 1, 0, 0);
        ctx.fillText(ch, 0, 0);
        ctx.restore();
      } else {
        ctx.fillText(ch, cx, baseline);
      }
    }
    ax += adv;
    }
  }
  ctx.textAlign = prevAlign;
  ctx.textBaseline = prevBaseline;
}

export function drawVerticalRun(
  ctx: Ctx2D,
  text: string,
  x: number,
  baseline: number,
  fontPx: number,
  letterSpacingPx: number,
  charScale = 1,
  growTrRotateInk = false,
): void {
  drawVerticalRunWithCapability(
    ctx,
    text,
    x,
    baseline,
    fontPx,
    letterSpacingPx,
    charScale,
    growTrRotateInk,
    (cp) => verticalVertGlyphReachable(ctx, cp),
  );
}

/**
 * Draw one 縦中横 (tate-chū-yoko / horizontal-in-vertical) run — ECMA-376
 * §17.3.2.10 `<w:eastAsianLayout w:vert="1">`. In a vertical (tbRl) page the run
 * "keeps the text on the same line as all other text" while its characters are
 * rendered HORIZONTALLY: the whole run string is drawn UPRIGHT (counter-rotated
 * −90° to cancel the +90° page rotation, exactly like the upright CJK cells) so
 * the glyphs read left-to-right ACROSS the column, packed into ONE cell of the
 * vertical line.
 *
 * Geometry (all in the rotated logical page frame; `x` advances DOWN the column
 * = logical +x, `baseline` is the column centre-line = logical +y):
 *   - The cell spans `[x, x + cellAdvance]` along the column; the run centres on
 *     the cell centre `x + cellAdvance/2`. `cellAdvance` is one em (one cell) —
 *     the same value the layout measured (`segAdvanceWidth`), so measure==paint.
 *   - After the −90° counter-rotation the run is upright; local +x is the
 *     cross-column (the glyphs' own left→right width) and local +y is the
 *     along-column (the text's height). Drawn `center`/`middle`, so the run's
 *     em box centres on the cell centre AND on the column centre-line.
 *   - `charScale` (§17.3.2.43 `w:w`) compresses glyph width through
 *     `ctx.scale(charScale, 1)` in the upright local frame, across the column.
 *     It does not change the along-column cell height.
 *   - `vertCompress` (§17.3.2.10) compresses the run's HEIGHT to one cell so the
 *     rotated text never grows the line: if the run's natural upright height
 *     (`fontBoundingBox*`) exceeds one em, scale the along-column axis down to
 *     fit. For a single-line run (height ≈ 1 em) this is a no-op, so the common
 *     2-digit date case is unaffected; it only bites a run whose glyphs are
 *     taller than the em box.
 *
 * The whole run is one contextually-shaped `fillText`, so kerning/shaping across
 * the digits is preserved and the text model / selection keep the original
 * characters.
 *
 * @param ctx          2D context, already in the rotated logical page frame.
 *                     `ctx.font`/`ctx.fillStyle` are set by the caller.
 * @param text         The run's text (e.g. "２９").
 * @param x            Logical left edge of the cell along the column (px).
 * @param baseline     Logical column centre-line y (px).
 * @param fontPx       Effective font size in px (one em = one cell).
 * @param cellAdvance  The cell's along-column advance in px (one em; the value
 *                     the layout measured for this segment).
 * @param charScale    §17.3.2.43 `w:w` fraction (1 = 100%); compresses the
 *                     glyphs' cross-column width.
 * @param compress     §17.3.2.10 `w:vertCompress`; fit the run's height to one em.
 */
export function drawTateChuYokoRun(
  ctx: Ctx2D,
  text: string,
  x: number,
  baseline: number,
  fontPx: number,
  cellAdvance: number,
  charScale: number,
  compress: boolean,
): void {
  const prevAlign = ctx.textAlign;
  const prevBaseline = ctx.textBaseline;
  // Along-column compression factor (§17.3.2.10 vertCompress). The run is drawn
  // upright, so its along-column extent is the font's design HEIGHT
  // (fontBoundingBoxAscent + descent). When that exceeds one em and vertCompress
  // is set, scale the along-column (local y) axis so the height fits one cell.
  // Measured with a `middle` baseline (the box used below). For ordinary
  // single-line text the height is ≈1 em, so `compY` stays 1 (no-op).
  let compY = 1;
  if (compress) {
    const m = ctx.measureText(text);
    const asc = m.fontBoundingBoxAscent;
    const desc = m.fontBoundingBoxDescent;
    if (typeof asc === 'number' && typeof desc === 'number') {
      const heightPx = asc + desc;
      if (heightPx > fontPx && heightPx > 0) compY = fontPx / heightPx;
    }
  }
  const cx = x + cellAdvance / 2;
  ctx.save();
  ctx.translate(cx, baseline);
  ctx.rotate(-Math.PI / 2);
  // Upright local frame: local +x = cross-column (glyph width), local +y =
  // along-column (glyph height). `w:w` compresses width (local x); vertCompress
  // fits height (local y). center/middle centres the run's em box on the cell.
  ctx.scale(charScale, compY);
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(text, 0, 0);
  ctx.restore();
  ctx.textAlign = prevAlign;
  ctx.textBaseline = prevBaseline;
}

/**
 * Run `draw` with the context counter-rotated so a graphic that would otherwise
 * be painted lying on its side (rotated with the +90° page transform) appears
 * UPRIGHT. Used for inline / anchored images and shapes in a vertical (tbRl)
 * page: an image is not text, so it keeps its natural upright orientation even
 * though the surrounding characters advance downward.
 *
 * The box is specified by its logical top-left `(x, y)` and logical size
 * `(w, h)` — the same coordinates the horizontal draw path uses. We rotate −90°
 * about the box centre (cancelling the page rotation) and invoke `draw(dx, dy,
 * dw, dh)` with the box re-expressed in the upright local frame: the logical
 * width becomes the local HEIGHT and vice-versa, so the caller draws the image
 * at `(-h/2, -w/2, h, w)` centred on the pivot. The net effect places an upright
 * image inside the rotated page footprint.
 *
 * @param ctx  2D context already in the rotated logical page frame.
 * @param x    Logical left of the box (px).
 * @param y    Logical top of the box (px).
 * @param w    Logical width of the box (px).
 * @param h    Logical height of the box (px).
 * @param draw Callback painting the graphic at `(dx, dy, dw, dh)` in the upright
 *             local frame.
 */
export function drawUprightBox(
  ctx: Ctx2D,
  x: number,
  y: number,
  w: number,
  h: number,
  draw: (dx: number, dy: number, dw: number, dh: number) => void,
): void {
  const cx = x + w / 2;
  const cy = y + h / 2;
  ctx.save();
  ctx.translate(cx, cy);
  ctx.rotate(-Math.PI / 2);
  // In the upright local frame the logical width spans the local y-axis and the
  // logical height spans the local x-axis, so the box is (−h/2, −w/2, h, w).
  draw(-h / 2, -w / 2, h, w);
  ctx.restore();
}

/** A rectangle `{ x, y, w, h }` in one coordinate frame and linear unit. */
export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/**
 * Map a DrawingML anchor's box from PHYSICAL page space into the SWAPPED LOGICAL
 * layout frame the vertical (tbRl) renderer flows text in (ECMA-376 §17.6.20 +
 * §20.4.3.x).
 *
 * Under `word-vertical-section-physical-drawing-layer`, a
 * `<wp:positionH>` / `<wp:positionV>` anchor resolves against the physical page
 * independently of text-flow rotation, so the image stays upright at
 * physical `(physicalX, physicalY, width, height)` exactly as in
 * a horizontal document. The body text, however, is laid out in the logical frame
 * that the page transform `physical = (physicalPageWidth − logical.y, logical.x)`
 * maps to physical. Inverting that transform (`logical.x = physical.y`,
 * `logical.y = physicalPageWidth − physical.x`) projects the physical rectangle onto
 * the logical frame:
 *   - logical x-range = `[physicalY, physicalY + height]`
 *                                               (physical y ↦ logical x, downward)
 *   - logical y-range = `[physicalPageWidth − (physicalX + width),
 *                         physicalPageWidth − physicalX]`
 *                                               (physical x ↦ logical y, reversed)
 * so the logical box has `w ↔ h` swapped: logical width = physical height and
 * logical height = physical width.
 *
 * The returned box drives BOTH the float-exclusion rectangle (text wraps around
 * this logical projection, in the same frame as the flow) AND {@link drawUprightBox}
 * (which un-swaps it back to the upright physical image). Because the two derive
 * from one box, the wrap band and the painted image stay locked together
 * (packages/docx/AGENTS.md — no duplicated geometry).
 *
 * Acquisition passes canonical points; the algebra remains valid for any
 * consistent linear unit.
 */
export function physicalToLogicalAnchorBox(
  physicalX: number,
  physicalY: number,
  width: number,
  height: number,
  physicalPageWidth: number,
): Box {
  return {
    x: physicalY,
    y: physicalPageWidth - (physicalX + width),
    w: height,
    h: width,
  };
}
