import { describe, expect, it } from 'vitest';
import { __test_resolveAnchorBox } from './test-support/renderer-internals.test-support.js';
import type { AnchorFloatRegistrationState } from './layout/acquisition-context.js';
import type { ImageRun } from './types.js';

// ECMA-376 §17.6.20 (tbRl) + §20.4.3.x — a DrawingML anchor on a VERTICAL page is
// resolved against the PHYSICAL (un-rotated) page, then projected into the swapped
// logical layout frame. This pins the CRITICAL fix (PR #770 review): the raw
// positionH/positionV offsets must NOT be fed unrotated into the logical frame;
// they resolve physically and map through the inverse of the section frame's
// paint matrix, so after the +90° page paint the image lands at the recorded
// physical centroid. The
// fixture below contains only the minimal redistributable geometry needed to
// preserve that behavior.

// A vertical anchor capability carries the LOGICAL (swapped) geometry PLUS `verticalPhys`
// (the physical page geometry the anchor path resolves against). Mirror what the
// measurement-state writer builds for a vertical section. resolveAnchorBox reads only
// geometry fields, so a minimal cast stand-in suffices (as in
// anchor-image-relativefrom).
const PHYS_W = 842;
const PHYS_H = 595.05;
const PHYS_ML = 56.7;
const PHYS_MT = 70.9;
const PHYS_MR = 56.7;
const PHYS_MB = 70.9;

const verticalState = {
  // Logical (swapped) geometry produced by verticalLayoutSection:
  pageWidth: PHYS_H, // logical width = physical height
  marginLeft: PHYS_MT, // logical left = physical top
  marginRight: PHYS_MB, // logical right = physical bottom
  marginTop: PHYS_MR, // logical top = physical right
  marginBottom: PHYS_ML, // logical bottom = physical left
  pageH: PHYS_W, // logical height in points = physical width
  contentX: PHYS_MT, // column band start (logical x) = physical top margin
  contentW: PHYS_H - PHYS_MT - PHYS_MB,
  verticalCJK: true,
  verticalPhys: {
    pageWidth: PHYS_W,
    pageHeight: PHYS_H,
    marginLeft: PHYS_ML,
    marginRight: PHYS_MR,
    marginTop: PHYS_MT,
    marginBottom: PHYS_MB,
    writingMode: 'vertical-rl',
  },
} as unknown as AnchorFloatRegistrationState;

const recordedAnchor: ImageRun = {
  imagePath: 'word/media/image1.png',
  mimeType: 'image/png',
  widthPt: 96.2,
  heightPt: 123.0,
  anchor: true,
  wrapMode: 'square',
  anchorXPt: 387.6,
  anchorYPt: 327.0,
  anchorXRelativeFrom: 'margin',
  anchorYRelativeFrom: 'margin',
  anchorXFromMargin: false,
  anchorYFromPara: false,
  distLeft: 9,
  distRight: 9,
  distTop: 0,
  distBottom: 0,
};

/** Reconstruct the physical centroid from a resolved LOGICAL anchor box by
 *  applying the vertical-rl page transform `physical = (cssW − logical.y,
 *  logical.x)` to the logical box centre. */
function physicalCentroid(
  box: { x: number; y: number; w: number; h: number },
  cssW: number,
): { cx: number; cy: number } {
  const logCx = box.x + box.w / 2;
  const logCy = box.y + box.h / 2;
  return { cx: cssW - logCy, cy: logCx };
}

describe('resolveAnchorBox — vertical (tbRl) physical anchor mapping (§20.4.3.x)', () => {
  it('lands an upright-section anchor at the recorded physical centroid', () => {
    const box = __test_resolveAnchorBox(recordedAnchor, verticalState, 0);
    const { cx, cy } = physicalCentroid(box, PHYS_W);
    // ±0.5px of the recorded reference centroid.
    expect(cx).toBeCloseTo(492.4, 1);
    expect(cy).toBeCloseTo(459.35, 1);
  });

  it('returns the logical-projected box (w↔h swap) so the flow wraps correctly', () => {
    const box = __test_resolveAnchorBox(recordedAnchor, verticalState, 0);
    // Physical TL is (ML+posH, MT+posV) = (444.3, 397.9): logical x = physical
    // y, logical y = page width − (physical x + width), width and height swap.
    expect(box.x).toBeCloseTo(PHYS_MT + 327.0, 4);
    expect(box.y).toBeCloseTo(PHYS_W - (PHYS_ML + 387.6 + 96.2), 4);
    expect(box.w).toBeCloseTo(123.0, 4);
    expect(box.h).toBeCloseTo(96.2, 4);
  });

  it('rotates dist* padding one quarter-turn (physical T/B ↦ logical L/R)', () => {
    const box = __test_resolveAnchorBox(recordedAnchor, verticalState, 0);
    // physical distTop/distBottom (0) ↦ logical dl/dr; physical distRight/distLeft
    // (9) ↦ logical dt/db.
    expect(box.dl).toBe(0);
    expect(box.dr).toBe(0);
    expect(box.dt).toBe(9);
    expect(box.db).toBe(9);
  });

  it('projects a vertical-lr (tbLrV) anchor through that frame\'s own paint inverse', () => {
    // tbLrV paints its logical frame through the vertical-lr reflection
    // (physical x = logical y, physical y = logical x). The physically resolved
    // box must come back to the same physical rectangle through that paint,
    // and its dist* padding must follow the same edges.
    const lrState = {
      ...verticalState,
      marginTop: PHYS_ML, // logical top = physical left
      marginBottom: PHYS_MR, // logical bottom = physical right
      verticalPhys: { ...verticalState.verticalPhys!, writingMode: 'vertical-lr' },
    } as unknown as AnchorFloatRegistrationState;
    const box = __test_resolveAnchorBox(
      { ...recordedAnchor, distLeft: 3, distRight: 9, distTop: 1, distBottom: 2 },
      lrState,
      0,
    );
    expect({ x: box.y, y: box.x, w: box.h, h: box.w }).toEqual({
      x: expect.closeTo(PHYS_ML + 387.6, 4),
      y: expect.closeTo(PHYS_MT + 327.0, 4),
      w: expect.closeTo(96.2, 4),
      h: expect.closeTo(123.0, 4),
    });
    // Logical left/right/top/bottom = physical top/bottom/left/right.
    expect([box.dl, box.dr, box.dt, box.db]).toEqual([1, 2, 3, 9]);
  });
});
