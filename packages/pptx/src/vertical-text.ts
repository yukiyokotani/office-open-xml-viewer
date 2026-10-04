// ECMA-376 §20.1.10.83 ST_TextVerticalType `eaVert` ("East Asian Vertical") —
// per-glyph orientation for a text body laid out in the +90°-rotated frame.
//
// The renderer draws an eaVert body by rotating the whole context +90° clockwise
// and re-running the HORIZONTAL layout inside that rotated frame (see
// `renderTextBody`): a plain `ctx.fillText` then advances DOWN the column (logical
// +x → physical +y), which is the character progression a vertical line wants —
// but every glyph is left lying on its right side (rotated with the page). Which
// glyphs must be stood back UP is decided by the Unicode UAX#50
// Vertical_Orientation (vo) property (core `verticalOrientation`), NOT an ad-hoc
// CJK-vs-Latin guess:
//   • vo=U  (upright): CJK ideographs, kana, Hangul, fullwidth forms/digits. Drawn
//     with a local −90° counter-rotation about the cell centre, cancelling the
//     page rotation so the glyph stands upright while still advancing down.
//   • vo=Tu (transform, fallback upright): the corner-hanging comma/full stop
//     、。， are SUBSTITUTED with their U+FE10–FE12 vertical form (core
//     `verticalFormSubstitute`) and drawn upright, so the font supplies the
//     designed upper-right placement; ！？ and small kana have no form and draw
//     upright unchanged.
//   • vo=Tr (transform, fallback rotate): the fullwidth brackets （「」〈〉【】… and
//     the white lenticular brackets 〖〗 have a U+FE1x/FE3x vertical presentation
//     form (core `verticalBracketFormSubstitute`) present in the substitute fonts;
//     UAX#50 §5 makes Tr "substitute a vertical glyph, ROTATE only as fallback", so
//     we substitute and draw them upright (Word/PowerPoint-verified, #969). Tr code
//     points with no substituted form take a geometric fallback: plain ROTATE (the
//     quotes, colon, and unreachable long marks ー 〜 ～), or UPRIGHT (the semicolon
//     ；→ FE14's dot-over-comma; issue #969 follow-up, core
//     `verticalTrUprightFallback`) — FE13/FE14 are absent from most render fonts, so
//     those take the geometric path.
//   • vo=R  (rotated): Latin letters, Western digits, Latin punctuation stay
//     SIDEWAYS (rotated with the page) — the conventional "縦中横 not applied" look.
//
// This mirrors the docx tbRl vertical draw (packages/docx/src/vertical-text.ts):
// the geometry is the same rotate-layout, and both consume the SAME core UAX#50
// classifier — no new shared abstraction is introduced (issue #790). The XLSX
// stacked path is DIFFERENT geometry (no page rotation), so it keeps its own tiny
// helper; only the classifier is shared across the three formats.

import {
  graphemeClusterOffsets,
  verticalOrientation,
  verticalFormSubstitute,
  verticalBracketFormSubstitute,
  verticalTrUprightFallback,
  verticalTrLongMark,
  verticalVertGlyphReachable,
  withVertFeature,
} from '@silurus/ooxml-core';

type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
type VertCapability = (cp: number) => boolean;
const NO_VERT_CAPABILITY: VertCapability = () => false;

/**
 * Cross-axis (column-thickness) offset, in px, from the alphabetic baseline to
 * the font's EM-BOX CENTRE — `(fontBoundingBoxAscent − fontBoundingBoxDescent)/2`.
 *
 * In the rotated eaVert frame the horizontal layout hands us the run's ALPHABETIC
 * baseline; the column's cross-axis centre (where the sideways glyphs' ink already
 * sits and where the upright cells must centre) is this many px ABOVE it. Using
 * the font metric keeps upright and sideways glyphs sharing one centreline.
 *
 * Falls back to `0.38 × fontPx` when the Canvas does not report `fontBoundingBox*`.
 * Canvas can report the first available face's font box even when another
 * composite subset paints the cluster. This existing platform-box centring
 * limitation is separate from the cluster owner's PowerPoint line metrics;
 * actual ink boxes do not supply a replacement em-box centre.
 */
function emBoxCenterAboveBaselinePx(ctx: Ctx2D, sample: string, fontPx: number): number {
  const prevBaseline = ctx.textBaseline;
  ctx.textBaseline = 'alphabetic';
  const m = ctx.measureText(sample);
  ctx.textBaseline = prevBaseline;
  const asc = m.fontBoundingBoxAscent;
  const desc = m.fontBoundingBoxDescent;
  if (typeof asc === 'number' && typeof desc === 'number' && (asc !== 0 || desc !== 0)) {
    return (asc - desc) / 2;
  }
  return 0.38 * fontPx;
}

/**
 * Along-column offset, in px, from a glyph's own cell centre to its INK centre
 * when drawn UPRIGHT — `(actualBoundingBoxAscent − actualBoundingBoxDescent)/2`
 * measured with a `middle` textBaseline. For an ordinary ideograph/kana this is
 * ≈0; for a substituted vertical bracket (︵ ︶ ﹁ ﹂, whose ink hugs one cell end)
 * it is the shift that re-centres the ink so the two halves sit a full cell apart.
 *
 * Returns 0 when the Canvas does not report `actualBoundingBox*` (older engines).
 */
function inkCenterAboveMiddlePx(ctx: Ctx2D, drawStr: string): number {
  const prevAlign = ctx.textAlign;
  const prevBaseline = ctx.textBaseline;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  const m = ctx.measureText(drawStr);
  ctx.textAlign = prevAlign;
  ctx.textBaseline = prevBaseline;
  const asc = m.actualBoundingBoxAscent;
  const desc = m.actualBoundingBoxDescent;
  if (typeof asc === 'number' && typeof desc === 'number') {
    return (asc - desc) / 2;
  }
  return 0;
}

/**
 * Paint one run's glyphs for an `eaVert` text body. The context is assumed to be
 * in the page's +90°-rotated frame already (installed by `renderTextBody`), so an
 * ordinary `fillText` advances DOWN the column. Each glyph occupies a cell of
 * width = its horizontal advance (`ctx.measureText`) plus `letterSpacingPx`, so
 * the total advance equals the run's measured width (measure == draw). The glyph
 * is stood upright / substituted / rotated / left sideways per its UAX#50 class.
 *
 * Glyphs are painted one grapheme cluster at a time:
 * each upright cell needs its own counter-rotation, and a uniform per-glyph
 * advance keeps measure==draw at the segment boundary. Contextual shaping of a
 * consecutive SIDEWAYS (Latin) run — kerning / ligatures — is therefore not
 * preserved; an accepted tradeoff for vertical East-Asian text, where Latin runs
 * are short and their advances are context-free enough at these sizes.
 *
 * @param ctx              2D context, already in the rotated eaVert frame.
 *                         `ctx.font`/`ctx.fillStyle`/`ctx.strokeStyle` are the
 *                         caller's.
 * @param text             The run's text.
 * @param x                Along-column left edge of the run (px; logical +x).
 * @param baseline         Along-column ALPHABETIC baseline y of the line (px) —
 *                         the same value the horizontal draw path uses.
 * @param fontPx           Effective font size in px (for cell centring).
 * @param letterSpacingPx  Per-glyph extra advance (rPr @spc pitch); 0 for the
 *                         common path.
 * @param paint            `'fill'` or `'stroke'` (run outline, rPr > a:ln).
 */
export function drawEaVertRunWithCapability(
  ctx: Ctx2D,
  text: string,
  x: number,
  baseline: number,
  fontPx: number,
  letterSpacingPx: number,
  paint: 'fill' | 'stroke' = 'fill',
  vertCapability: VertCapability = NO_VERT_CAPABILITY,
): void {
  const prevAlign = ctx.textAlign;
  const prevBaseline = ctx.textBaseline;
  const draw = paint === 'stroke' ? ctx.strokeText.bind(ctx) : ctx.fillText.bind(ctx);
  // Cross-axis (column-thickness) centre of the column, measured once per run.
  // The alphabetic `baseline` sits this many px BELOW the em-box centre; the
  // sideways (Latin) glyphs already centre their ink there, so the upright cells
  // centre on the same line by drawing at `crossCenterY`.
  const emBoxCenterPx = emBoxCenterAboveBaselinePx(ctx, text, fontPx);
  const crossCenterY = baseline - emBoxCenterPx;
  let ax = 0; // cumulative advance from the run's left (logical +x)
  let start = 0;
  const ends = graphemeClusterOffsets(text);
  ends.push(text.length);
  for (const end of ends) {
    const ch = text.slice(start, end);
    start = end;
    if (!ch) continue;
    const cp = ch.codePointAt(0) ?? 0;
    const vo = verticalOrientation(cp);
    // Measure and paint the same complete cluster; scalar paint would let a
    // base-only subset steal ownership from the font selected for its marks.
    // Presentation forms apply only to lone scalars, preserving all modifiers.
    // Advance/width uses the ORIGINAL cluster (measure == draw; the text model,
    // selection and find keep the original character — substitution is glyph-only).
    const adv = ctx.measureText(ch).width + letterSpacingPx;
    // A vo=Tr code point with a Unicode vertical presentation form — brackets （「」…
    // and the white-lenticular 〖〗 (#969) — is substituted and drawn UPRIGHT
    // (UAX#50 §5 substitute-first). Tr code points with NO substituted form take a
    // fallback: ROTATE (ー, quotes, colon ： → FE13's side-by-side dots) or, for the
    // semicolon ；(FF1B), UPRIGHT (its FE14 form is upright dot-over-comma, not a
    // rotation). The colon/semicolon FE13/FE14 substitution was dropped in core —
    // those forms are absent from most render fonts (issue #969 follow-up).
    const single = ch === String.fromCodePoint(cp);
    const bracketCp = single && vo === 'Tr' ? verticalBracketFormSubstitute(cp) : null;
    const uprightFallback = vo === 'Tr' && bracketCp === null && verticalTrUprightFallback(cp);
    const upright = vo === 'U' || vo === 'Tu' || bracketCp !== null || uprightFallback;
    const vertGlyphSupported = verticalTrLongMark(cp) && vertCapability(cp);
    if (vertGlyphSupported) {
      const cx = x + ax + adv / 2;
      ctx.save();
      ctx.translate(cx, crossCenterY);
      ctx.rotate(-Math.PI / 2);
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      withVertFeature(ctx, () => draw(ch, 0, 0));
      ctx.restore();
    } else if (upright) {
      // vo=U / vo=Tu, or a substituted Tr bracket. Counter-rotate −90° about the
      // cell centre so the glyph (which the page rotation would lay on its side)
      // stands upright. Corner-hanging Tu punctuation (、。， → U+FE10–FE12) and Tr
      // brackets (（）「」… → U+FE35–FE44) are drawn as their vertical form so the
      // font supplies the vertical shape; ！？ and small kana draw upright unchanged.
      const puncCp = bracketCp !== null ? null : (single && vo === 'Tu' ? verticalFormSubstitute(cp) : null);
      const drawCp = bracketCp !== null ? bracketCp : puncCp;
      const drawStr = drawCp !== null ? String.fromCodePoint(drawCp) : ch;
      const cx = x + ax + adv / 2;
      // Along-column ink re-centring: an upright glyph's VERTICAL ink extent maps to
      // the column axis. ≈0 for an ideograph/kana (cells unchanged); the needed
      // correction for a substituted bracket whose ink hugs one cell end. NOT applied
      // to a substituted comma/full stop (FE10–FE12): those are DESIGNED with their
      // ink in the cell's upper-right corner, so em-box-centring preserves it.
      const isPunctSubstitute = puncCp !== null;
      const alongEm = isPunctSubstitute ? 0 : inkCenterAboveMiddlePx(ctx, drawStr) / fontPx;
      ctx.save();
      ctx.translate(cx, crossCenterY);
      ctx.rotate(-Math.PI / 2);
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      draw(drawStr, 0, alongEm * fontPx);
      ctx.restore();
    } else if (vo === 'Tr') {
      // vo=Tr with NO substituted vertical form and NOT the upright-fallback
      // semicolon: ー (U+30FC), the wave dash / tilde 〜 ～, quotes “”, and the colon
      // ：(FF1A). The Tr fallback is to ROTATE 90° CW — a plain draw in the +90° page
      // frame IS that rotation; centre it on the column. For the colon this reproduces
      // FE13's design (the vertically-stacked dots become side by side) and for the
      // quotes it matches the font's designed vertical form, both Word-verified (#969).
      //
      // Unreachable `ー〜～` deliberately use the same plain rotation. Their DOM
      // `vert` design cannot be inferred portably in worker/skia environments.
      const cx = x + ax + adv / 2;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      draw(ch, cx, crossCenterY);
    } else {
      // vo=R (Latin/digits): stay SIDEWAYS (rotated with the page). Draw at the
      // alphabetic baseline; its ink already centres on the column centreline the
      // upright cells use (crossCenterY = baseline − emBoxCenterPx), so no cross-axis
      // shift is needed — byte-identical to the horizontal path's per-glyph draw.
      ctx.textAlign = prevAlign;
      ctx.textBaseline = 'alphabetic';
      draw(ch, x + ax, baseline);
    }
    ax += adv;
  }
  ctx.textAlign = prevAlign;
  ctx.textBaseline = prevBaseline;
}

export function drawEaVertRun(
  ctx: Ctx2D,
  text: string,
  x: number,
  baseline: number,
  fontPx: number,
  letterSpacingPx: number,
  paint: 'fill' | 'stroke' = 'fill',
): void {
  drawEaVertRunWithCapability(
    ctx,
    text,
    x,
    baseline,
    fontPx,
    letterSpacingPx,
    paint,
    (cp) => verticalVertGlyphReachable(ctx, cp),
  );
}
