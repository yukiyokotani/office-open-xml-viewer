import { describe, it, expect } from 'vitest';
import {
  computeFloatTableBox,
  registerTableFloat,
  floatTableWrapSide,
  resolveFloatingTablePlacement,
} from './float-table-geometry.js';
import type { FloatTableBox } from './float-table-geometry.js';
import type { TblpPr } from './types.js';
import type { FloatRect } from './float-layout.js';
import type {
  FloatingTablePlacementLayout,
  FloatingTableReferenceFramesPt,
  TableLayout,
} from './layout/types.js';
import {
  beginFloatingTablePlacementTransaction,
  floatingTableRegistryDelta,
  resolveFloatingTablePlacementInTransaction,
  resolveFloatingTablePlacementsInFreshRegistry,
  validateFloatingTableRegistryDelta,
} from './layout/floating-table-transaction.js';

// Table-driven geometry assertions for floating tables (ECMA-376 §17.4.57
// `<w:tblpPr>` / §17.4.56 `<w:tblOverlap>`). The VRT does not cover the
// These synthetic cases pin the placement math and the FloatRect registration
// (xLeft/xRight/yTop/yBottom/mode/side), which resolveLineFloatWindow consumes
// to wrap the surrounding body text.
//
// Geometry is exercised directly in canonical points. A representative page:
//   pageWidth=600, margins L/R/T/B = 100/100/72/72 ⇒ content band [100,500].
//   A single column is modeled as contentX=100, contentW=400; hAnchor="text"
//   snaps to it (the #513 column-relative contract).

interface MinState {
  contentX: number;
  contentW: number;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  pageWidth: number;
  pageH: number;
  floats: FloatRect[];
  floatParaSeq: number;
}

function makeState(over: Partial<MinState> = {}): MinState {
  return {
    contentX: 100,
    contentW: 400,
    marginLeft: 100,
    marginRight: 100,
    marginTop: 72,
    marginBottom: 72,
    pageWidth: 600,
    pageH: 800,
    floats: [],
    floatParaSeq: 0,
    ...over,
  };
}

// Full TblpPr with the spec defaults; tests override only the axis under test.
// The factory sets horzAnchor:'page' (an explicit horizontal hint), so
// horzSpecified defaults to true; the "no horizontal spec at all" case is
// exercised by overriding horzSpecified:false below.
function tblp(over: Partial<TblpPr> = {}): TblpPr {
  return {
    leftFromText: 0,
    rightFromText: 0,
    topFromText: 0,
    bottomFromText: 0,
    horzAnchor: 'page',
    horzSpecified: true,
    vertAnchor: 'page',
    tblpX: 0,
    tblpY: 0,
    ...over,
  };
}

// Cast helpers: the geometry fns read only the MinState subset of BodyAcquisitionState.
const box = (tp: TblpPr, st: MinState, paraTop: number, w: number, h: number): FloatTableBox =>
  computeFloatTableBox(tp, st as never, paraTop, w, h);
const registerFloat = (
  b: FloatTableBox,
  tp: TblpPr,
  st: MinState,
  side: string,
  allowOverlap: boolean,
): void => registerTableFloat(b, tp, st as never, side, allowOverlap);
const wrapSide = (b: FloatTableBox, st: MinState): string =>
  floatTableWrapSide(b, st as never);

describe('floating table geometry (§17.4.57) — horizontal placement', () => {
  it('horzAnchor="page", tblpX=0 ⇒ table at the physical page left edge', () => {
    const st = makeState();
    const b = box(tblp({ horzAnchor: 'page', tblpX: 0 }), st, 300, 150, 60);
    expect(b.x).toBe(0); // page left + 0
    expect(b.w).toBe(150);
    expect(b.h).toBe(60);
  });

  it('horzAnchor="page", tblpX signed offset is added (signed twips round-trip)', () => {
    const st = makeState();
    expect(box(tblp({ horzAnchor: 'page', tblpX: 36 }), st, 300, 150, 60).x).toBe(36);
    expect(box(tblp({ horzAnchor: 'page', tblpX: -20 }), st, 300, 150, 60).x).toBe(-20);
  });

  it('horzAnchor="margin", tblpX anchors against the page content margin', () => {
    const st = makeState();
    const b = box(tblp({ horzAnchor: 'margin', tblpX: 10 }), st, 300, 150, 60);
    expect(b.x).toBe(100 + 10); // marginLeft + tblpX
  });

  it('horzAnchor="text", tblpX anchors against the COLUMN band (contentX, #513)', () => {
    const st = makeState({ contentX: 250, contentW: 200 }); // a right-hand column
    const b = box(tblp({ horzAnchor: 'text', tblpX: 10 }), st, 300, 80, 40);
    expect(b.x).toBe(250 + 10); // column left + tblpX
  });

  it('tblpXSpec="left" / "inside" snap to the band left, superseding tblpX', () => {
    const st = makeState();
    for (const spec of ['left', 'inside']) {
      const b = box(tblp({ horzAnchor: 'text', tblpX: 999, tblpXSpec: spec }), st, 300, 100, 40);
      expect(b.x, spec).toBe(100); // contentX, tblpX ignored
    }
  });

  it('tblpXSpec="center" centers the table in the band, superseding tblpX', () => {
    const st = makeState(); // text band [100,500], width 400
    const b = box(tblp({ horzAnchor: 'text', tblpX: 999, tblpXSpec: 'center' }), st, 300, 100, 40);
    expect(b.x).toBe(100 + (400 - 100) / 2);
  });

  it('tblpXSpec="right" / "outside" right-align the table in the band', () => {
    const st = makeState();
    for (const spec of ['right', 'outside']) {
      const b = box(tblp({ horzAnchor: 'text', tblpXSpec: spec }), st, 300, 100, 40);
      expect(b.x, spec).toBe(500 - 100); // band right − tableW
    }
  });
});

describe('floating table geometry (§17.4.57) — vertical placement', () => {
  it('vertAnchor="text", tblpY ⇒ table top at the anchor paragraph top + tblpY', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'text', tblpY: 1 }), st, 300, 150, 60);
    expect(b.y).toBe(300 + 1); // paraTop + tblpY
  });

  it('vertAnchor="margin", tblpY anchors against the top content margin', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'margin', tblpY: 5 }), st, 300, 150, 60);
    expect(b.y).toBe(72 + 5); // marginTop + tblpY
  });

  it('vertAnchor="page", tblpY anchors against the physical page top', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'page', tblpY: 5 }), st, 300, 150, 60);
    expect(b.y).toBe(0 + 5); // page top + tblpY
  });

  it('vertAnchor="text" IGNORES tblpYSpec (relative vertical not allowed) → tblpY', () => {
    const st = makeState();
    // tblpYSpec="center" would otherwise center; with vertAnchor="text" it is
    // ignored and the absolute tblpY offset is used instead (§17.4.57).
    const b = box(tblp({ vertAnchor: 'text', tblpY: 7, tblpYSpec: 'center' }), st, 300, 150, 60);
    expect(b.y).toBe(300 + 7);
  });

  it('vertAnchor="margin" + tblpYSpec="bottom" bottom-aligns on the page, superseding tblpY', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'margin', tblpY: 999, tblpYSpec: 'bottom' }), st, 300, 150, 60);
    expect(b.y).toBe((800 - 72) - 60); // (pageH − marginBottom) − tableH
  });

  it('vertAnchor="margin" + tblpYSpec="center" centers vertically, superseding tblpY', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'margin', tblpY: 999, tblpYSpec: 'center' }), st, 300, 150, 60);
    // margin band [72, 728]: start + (end − start − h)/2 = 72 + (728 − 72 − 60)/2
    expect(b.y).toBe(72 + (728 - 72 - 60) / 2);
  });

  it('vertAnchor="margin" + tblpYSpec="center": ASYMMETRIC margins center in the margin band, NOT the page', () => {
    // §22.9.2.20: tblpYSpec="center" centers in the margin BAND
    // [marginTop, pageH−marginBottom], which only equals the page centre when
    // margins are symmetric. The symmetric case above (marginTop=marginBottom=72)
    // cannot distinguish "margin band centre" from "page centre"; this asymmetric
    // case pins the band-relative behaviour.
    const st = makeState({ marginTop: 40, marginBottom: 120 }); // band [40, 680], height 640
    const b = box(tblp({ vertAnchor: 'margin', tblpY: 999, tblpYSpec: 'center' }), st, 300, 150, 60);
    // band centre = marginTop + (pageH − marginTop − marginBottom − h)/2
    //             = 40 + (800 − 40 − 120 − 60)/2 = 40 + 290 = 330
    expect(b.y).toBe(40 + (800 - 40 - 120 - 60) / 2); // 330 — NOT page centre (800−60)/2 = 370
    expect(b.y).not.toBe((800 - 60) / 2); // explicit: this is the band centre, not the page centre
  });

  // §22.9.2.20: ST_YAlign positions relative to the ANCHOR OBJECT (the vertAnchor
  // band), NOT the physical page. For vertAnchor="page" the band is [0, pageH], so
  // center/bottom must NOT carry a margin offset (the pre-band code subtracted
  // marginTop even for page-anchored tables, §17.18.100 page = page edges).
  it('vertAnchor="page" + tblpYSpec="center" centers over the FULL page (no margin offset)', () => {
    const st = makeState(); // page band [0, 800]
    const b = box(tblp({ vertAnchor: 'page', tblpY: 999, tblpYSpec: 'center' }), st, 300, 150, 60);
    expect(b.y).toBe((800 - 60) / 2); // 370 — NOT (800−60)/2 − marginTop
  });

  it('vertAnchor="page" + tblpYSpec="bottom" sits flush to the physical page bottom', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'page', tblpY: 999, tblpYSpec: 'bottom' }), st, 300, 150, 60);
    expect(b.y).toBe(800 - 60); // pageH − tableH (no marginBottom inset)
  });

  it('vertAnchor="page" + tblpYSpec="outside" sits flush to the physical page bottom', () => {
    const st = makeState();
    const b = box(tblp({ vertAnchor: 'page', tblpY: 999, tblpYSpec: 'outside' }), st, 300, 150, 60);
    expect(b.y).toBe(800 - 60);
  });

  it('vertAnchor="page" + tblpYSpec="top"/"inside" sits at the page top (band start)', () => {
    const st = makeState();
    for (const spec of ['top', 'inside']) {
      const b = box(tblp({ vertAnchor: 'page', tblpY: 999, tblpYSpec: spec }), st, 300, 150, 60);
      expect(b.y, spec).toBe(0); // page band start = 0
    }
  });
});

// Observed Office behavior, represented here with synthetic geometry:
// a vertAnchor="page" floating table whose requested tblpY would push its BOTTOM
// past the physical page edge is shifted UP so its bottom sits flush on the page
// bottom (measured top 741.9pt = 841.9 − 100 for a 100pt table), NOT left
// overflowing. computeFloatTableBox clamps it via clampAbsBoxIntoContainer.
// vertAnchor="text" is NOT clamped (its overflow is row-split by the paginator).
describe('floating table geometry (§17.4.57) — page/margin-anchored Office clamp', () => {
  it('vertAnchor="page": a box overflowing the page bottom is clamped up to pageH − boxH', () => {
    // With pageH 841.9 and a 100pt table, tblpY=775
    // would put the bottom at 875 > 841.9 ⇒ clamp to y = 841.9 − 100 = 741.9.
    const st = makeState({ pageH: 841.9 });
    const b = box(tblp({ vertAnchor: 'page', tblpY: 775 }), st, 300, 250, 100);
    expect(b.y).toBeCloseTo(841.9 - 100, 6); // 741.9 — clamped to the page bottom
  });

  it('vertAnchor="page": a box that already fits is NOT moved (clamp is idempotent)', () => {
    const st = makeState(); // pageH 800
    const b = box(tblp({ vertAnchor: 'page', tblpY: 5 }), st, 300, 150, 60);
    expect(b.y).toBe(5); // 5 + 60 = 65 ≤ 800 ⇒ unchanged
  });

  it('vertAnchor="page": a box TALLER than the page pins to the top (floor = page top)', () => {
    // boxH 900 > pageH 800: pageH − boxH = −100 would push it above the page top,
    // so the floor (containerStart = 0) wins — the box overflows the bottom instead.
    const st = makeState(); // pageH 800
    const b = box(tblp({ vertAnchor: 'page', tblpY: 50 }), st, 300, 150, 900);
    expect(b.y).toBe(0); // clamped to the page top, not −100
  });

  it('vertAnchor="margin": a box overflowing the bottom margin is clamped to (pageH − marginBottom) − boxH', () => {
    // margin band [72, 728]: tblpY=700 ⇒ y=772, bottom 872 > 728 ⇒ clamp to
    // 728 − 100 = 628 (the container end is the bottom text margin, ASSUMED — no
    // fixture pins where Word clamps a margin-anchored overflow; symmetric with page).
    const st = makeState(); // margin band [72, 728]
    const b = box(tblp({ vertAnchor: 'margin', tblpY: 700 }), st, 300, 150, 100);
    expect(b.y).toBe(728 - 100); // 628 — clamped to the margin band bottom
  });

  it('vertAnchor="text": an overflowing box is NOT clamped (paginator row-split handles it)', () => {
    // A vertAnchor="text" table rides the flow cursor; its overflow is split
    // row-by-row by canonical pagination, so the geometry must leave the box at paraTop +
    // tblpY even when that runs past the page — clamping here would corrupt the
    // per-slice band. paraTop 780, tblpY 0, tableH 100 ⇒ y stays 780 (bottom 880).
    const st = makeState(); // pageH 800
    const b = box(tblp({ vertAnchor: 'text', tblpY: 0 }), st, 780, 150, 100);
    expect(b.y).toBe(780); // NOT clamped to 700 — text-anchored is the paginator's job
  });
});

describe('floating table float registration (§17.4.57 / §17.4.56)', () => {
  it('registers a square wrap float padded by the *FromText dist values', () => {
    const st = makeState();
    const tp = tblp({ horzAnchor: 'page', leftFromText: 5, rightFromText: 9, topFromText: 3, bottomFromText: 7 });
    const b = box(tp, st, 300, 150, 60);
    registerFloat(b, tp, st, 'right', true);
    expect(st.floats).toHaveLength(1);
    const f = st.floats[0];
    expect(f.mode).toBe('square');
    expect(f.side).toBe('right');
    expect(f.imageKey).toBe(''); // non-image float (table painted directly)
    expect(f.xLeft).toBe(0 - 5); // box.x − leftFromText
    expect(f.xRight).toBe(0 + 150 + 9); // box.x + tableW + rightFromText
    expect(f.yTop).toBe(b.y - 3);
    expect(f.yBottom).toBe(b.y + 60 + 7);
    expect(f.distLeft).toBe(5);
    expect(f.distRight).toBe(9);
  });

  it('a degenerate (zero-area) box registers no float', () => {
    const st = makeState();
    const tp = tblp();
    registerFloat({ x: 0, y: 0, w: 0, h: 0 }, tp, st, 'right', true);
    expect(st.floats).toHaveLength(0);
  });

  it('overlap="never" re-seats table extents without adding text-distance padding', () => {
    // Pre-seat a blocking FLOATING TABLE [0,200]×[300,360] from another paragraph.
    const st = makeState({
      floatParaSeq: 1,
      floats: [{
        kind: 'table', tableOverlap: 'overlap', mode: 'square',
        imageKey: '', imageX: 0, imageY: 300, imageW: 200, imageH: 60,
        xLeft: 0, xRight: 200, yTop: 300, yBottom: 360, side: 'bothSides',
        distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0,
      }],
    });
    // A new floating table at page-left x=0 overlapping the blocker; never ⇒ avoid.
    const tp = tblp({
      horzAnchor: 'page', tblpX: 0, vertAnchor: 'page', tblpY: 320,
      leftFromText: 5, rightFromText: 7,
    });
    const b = box(tp, st, 0, 100, 20); // box at (0,320), overlaps the blocker
    registerFloat(b, tp, st, 'right', /* allowOverlap */ false);
    const f = st.floats[st.floats.length - 1];
    // §17.4.56 uses table extents: the table box touches x=200 while its
    // separate §17.4.57 text exclusion still begins at x=195.
    expect(f.imageX).toBe(200);
    expect(f.xLeft).toBe(195);
  });

  it('an overlap-permitted table clears a blocker-side tblOverlap=never fact', () => {
    const blocker: FloatRect = {
      kind: 'table',
      tableOverlap: 'never',
      mode: 'square',
      imageKey: '',
      imageX: 0,
      imageY: 300,
      imageW: 200,
      imageH: 60,
      xLeft: -5,
      xRight: 207,
      yTop: 296,
      yBottom: 368,
      side: 'bothSides',
      distLeft: 5,
      distRight: 7,
      distTop: 4,
      distBottom: 8,
      paraId: 1,
    };
    const st = makeState({ floatParaSeq: 1, floats: [blocker] });
    const tp = tblp({
      horzAnchor: 'page',
      tblpX: 0,
      vertAnchor: 'page',
      tblpY: 320,
      leftFromText: 3,
      rightFromText: 9,
    });
    const resolved = computeFloatTableBox(
      tp,
      st as never,
      0,
      100,
      20,
      false,
      { allowOverlap: true },
    );

    // Same paragraph ID disables the Word compatibility path. §17.4.56 still
    // clears the blocker on its raw right edge (200), not its padded edge (207).
    expect(resolved.x).toBe(200);
  });

  it('appends a pre-resolved parent exactly after its retained child entry', () => {
    const child: FloatRect = {
      kind: 'table', tableOverlap: 'never', mode: 'square', imageKey: 'nested-child',
      imageX: 20, imageY: 30, imageW: 60, imageH: 20,
      xLeft: 20, xRight: 80, yTop: 30, yBottom: 50, side: 'bothSides',
      distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
      paraId: 7,
    };
    const st = makeState({ floatParaSeq: 8, floats: [child] });
    const tp = tblp({
      leftFromText: 3, rightFromText: 5, topFromText: 7, bottomFromText: 11,
    });
    const accepted = { x: 24, y: 20, w: 140, h: 90 };

    registerTableFloat(
      accepted, tp, st as never, 'bothSides', false, /* preResolved */ true,
    );

    expect(st.floats).toHaveLength(2);
    expect(st.floatParaSeq).toBe(9);
    expect(st.floats.map((float) => float.paraId)).toEqual([7, 8]);
    expect(st.floats[1]).toMatchObject({
      imageX: accepted.x,
      imageY: accepted.y,
      imageW: accepted.w,
      imageH: accepted.h,
      xLeft: accepted.x - tp.leftFromText,
      xRight: accepted.x + accepted.w + tp.rightFromText,
      yTop: accepted.y - tp.topFromText,
      yBottom: accepted.y + accepted.h + tp.bottomFromText,
    });
  });

  // §17.4.56 scope: a floating table with <w:tblOverlap w:val="never"/> must
  // only avoid OTHER FLOATING TABLES. DrawingML anchors (§20.4.2.3) and text
  // frames keep their own placement — a never-overlap table does NOT re-seat off
  // them. These pin that scoping via FloatRect.kind.
  describe('overlap="never" scopes to other floating tables only (§17.4.56)', () => {
    // A blocker occupying [0,200]×[300,360], anchored in another paragraph,
    // parameterized by kind so we can assert table-vs-non-table behavior.
    const blocker = (kind: FloatRect['kind']): FloatRect => {
      const core = {
        mode: 'square' as const,
        imageKey: '', imageX: 0, imageY: 300, imageW: 200, imageH: 60,
      xLeft: 0, xRight: 200, yTop: 300, yBottom: 360, side: 'bothSides',
      distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0,
      };
      return kind === 'table'
        ? { ...core, kind: 'table', tableOverlap: 'overlap' }
        : { ...core, kind };
    };
    // A never-overlap floating table seated at page-left x=0, overlapping [0,200].
    const seatNeverTable = (blockers: FloatRect[]): FloatRect => {
      const st = makeState({ floatParaSeq: 1, floats: blockers });
      const tp = tblp({ horzAnchor: 'page', tblpX: 0, vertAnchor: 'page', tblpY: 320 });
      const b = box(tp, st, 0, 100, 20); // box at (0,320), overlaps [0,200]
      registerFloat(b, tp, st, 'right', /* allowOverlap (never) */ false);
      return st.floats[st.floats.length - 1];
    };

    it('does NOT re-seat off a DrawingML shape/image float (stays at x=0)', () => {
      const f = seatNeverTable([blocker('shape')]);
      expect(f.xLeft).toBe(0); // overlap permitted: §17.4.56 ignores non-tables
    });

    it('does NOT re-seat off a text-frame float (stays at x=0)', () => {
      const f = seatNeverTable([blocker('frame')]);
      expect(f.xLeft).toBe(0);
    });

    it('DOES re-seat off another floating-table float', () => {
      const f = seatNeverTable([blocker('table')]);
      expect(f.xLeft).toBeGreaterThanOrEqual(200); // avoids the other table
    });

    it('avoids only the table blocker when a shape and a table both overlap', () => {
      // shape at [0,200], table at [0,200] (same band, different paragraphs):
      // the table re-seats past the TABLE's right edge, ignoring the shape.
      const f = seatNeverTable([blocker('shape'), blocker('table')]);
      expect(f.xLeft).toBeGreaterThanOrEqual(200);
    });
  });
});

describe('floating table — omitted horizontal positioning (§17.4.57)', () => {
  // <w:tblpPr w:rightFromText="187" w:bottomFromText="72" w:vertAnchor="text"
  //           w:tblpY="1"/> + <w:tblOverlap w:val="never"/>. NO horzAnchor, NO
  // tblpX, NO tblpXSpec ⇒ horzSpecified=false. The spec-literal default would be
  // horzAnchor="page" + tblpX=0 (page edge), but Word anchors such a table at the
  // anchor paragraph's text/column LEFT (contentX), not the page edge — matching
  // the calibre PDF where the ITEM/NEEDED table aligns with the body column.
  // rightFromText 187 twip → 9.35 pt; bottomFromText 72 twip → 3.6 pt.
  const RIGHT_FROM_TEXT = 187 / 20; // 9.35 pt
  const BOTTOM_FROM_TEXT = 72 / 20; // 3.6 pt

  it('no horizontal spec ⇒ anchors at the text/column left (NOT the page edge)', () => {
    const st = makeState(); // column band [100, 500]
    const tp = tblp({
      rightFromText: RIGHT_FROM_TEXT,
      bottomFromText: BOTTOM_FROM_TEXT,
      horzAnchor: 'page', // defaulted token, but…
      horzSpecified: false, // …no horizontal hint was present in the source
      vertAnchor: 'text',
      tblpX: 0,
      tblpY: 1,
    });
    // A small ITEM/NEEDED table, say 120 wide × 50 tall, anchored at paraTop=300.
    const tableW = 120;
    const b = box(tp, st, 300, tableW, 50);
    expect(b.x).toBe(100); // text/column left (contentX), NOT page-left 0
    expect(b.y).toBe(300 + 1); // vAnchor=text ⇒ paraTop + tblpY

    // The table contributes its actual exclusion rectangle; the line-layout
    // solver chooses the widest usable free gap on either side.
    const side = wrapSide(b, st);
    expect(side).toBe('bothSides');

    registerFloat(b, tp, st, side, /* overlap="never" ⇒ */ false);
    const f = st.floats[0];
    expect(f.side).toBe('bothSides');
    expect(f.xRight).toBe(100 + tableW + RIGHT_FROM_TEXT); // right edge includes the dist padding
    expect(f.yBottom).toBe(b.y + 50 + BOTTOM_FROM_TEXT);
  });

  it('explicit horzAnchor="page" + tblpX=0 still pins to the page edge', () => {
    const st = makeState();
    // When the source DOES specify a horizontal anchor, honor the spec literally.
    const b = box(tblp({ horzAnchor: 'page', horzSpecified: true, tblpX: 0 }), st, 300, 120, 50);
    expect(b.x).toBe(0); // page left edge
  });
});

describe('floating table wrap exclusion (§17.4.57)', () => {
  const st = makeState();

  it.each([
    { x: 0, w: 120 },
    { x: 350, w: 120 },
    { x: 200, w: 200 },
  ])('offers both flanks to the shared widest-free-gap solver', ({ x, w }) => {
    expect(wrapSide({ x, y: 0, w, h: 50 }, st)).toBe('bothSides');
  });
});

function retainedFloatingPlacement(
  positioning: Partial<FloatingTablePlacementLayout['positioning']> = {},
): FloatingTablePlacementLayout {
  const child = {
    kind: 'table', id: 'nested-table',
    source: { story: 'body', storyInstance: 'body', path: [0, 0, 0] },
    flowDomainId: 'nested', ordinaryFlow: false,
    flowBounds: { xPt: 20, yPt: 10, widthPt: 80, heightPt: 40 },
    inkBounds: { xPt: 20, yPt: 10, widthPt: 80, heightPt: 40 },
    advancePt: 40, columnWidthsPt: [80], rows: [], borders: [],
  } as TableLayout;
  return Object.freeze({
    kind: 'floating-table-placement',
    occurrenceId: 'page-1:cell-0:0:nested-table',
    ownership: 'source',
    physicalPageIndex: 0,
    displayPageNumber: 1,
    hostCellId: 'cell-0',
    sourceBlockIndex: 0,
    anchorBlockIndex: 1,
    tableId: child.id,
    overlap: 'never',
    positioning: Object.freeze({
      leftFromTextPt: 1,
      rightFromTextPt: 2,
      topFromTextPt: 3,
      bottomFromTextPt: 4,
      horzAnchor: 'text',
      horzSpecified: true,
      vertAnchor: 'text',
      xPt: 5,
      yPt: 6,
      ...positioning,
    }),
    anchorBounds: Object.freeze({ xPt: 120, yPt: 250, widthPt: 200, heightPt: 20 }),
    child,
  });
}

const retainedReferenceFrames: FloatingTableReferenceFramesPt = Object.freeze({
  page: Object.freeze({ xPt: 0, yPt: 0, widthPt: 600, heightPt: 800 }),
  margin: Object.freeze({ xPt: 100, yPt: 72, widthPt: 400, heightPt: 656 }),
  text: Object.freeze({ xPt: 120, yPt: 250, widthPt: 200, heightPt: 20 }),
});

describe('retained floating table placement (§17.4.57)', () => {
  it('validates the exact probed point delta and paragraph sequence once', () => {
    const snapshot = Object.freeze({
      coordinateSpace: 'logical-page-points' as const,
      flowDomainId: 'logical-page:3',
      entries: Object.freeze([]),
      nextParagraphId: 7,
    });
    const placement = retainedFloatingPlacement({
      horzAnchor: 'page', vertAnchor: 'page', xPt: 100, yPt: 100,
    });
    const resolution = resolveFloatingTablePlacementInTransaction(
      placement,
      retainedReferenceFrames,
      beginFloatingTablePlacementTransaction(
        snapshot.entries,
        snapshot.nextParagraphId,
        snapshot.coordinateSpace,
        snapshot.flowDomainId,
      ),
    );
    const delta = floatingTableRegistryDelta(
      snapshot,
      resolution.transaction.delta,
      resolution.transaction.nextParagraphId,
    );
    const current = {
      coordinateSpace: snapshot.coordinateSpace,
      flowDomainId: snapshot.flowDomainId,
      entries: snapshot.entries,
      nextParagraphId: snapshot.nextParagraphId,
    } as const;

    expect(() => validateFloatingTableRegistryDelta(delta, current)).not.toThrow();
    expect(delta.entries[0]).toMatchObject({
      occurrenceId: placement.occurrenceId,
      overlap: 'never',
      paragraphId: 7,
      bounds: resolution.placement.bounds,
      exclusionBounds: resolution.placement.exclusionBounds,
    });
    expect(() => validateFloatingTableRegistryDelta(delta, {
      ...current,
      nextParagraphId: 8,
    })).toThrow('base/domain mismatch');
  });

  it('retains blocker-side tblOverlap=never across point-space transactions', () => {
    const placement = Object.freeze({
      ...retainedFloatingPlacement(),
      overlap: 'overlap' as const,
      occurrenceId: 'moving-overlap-table',
    });
    const blockerBounds = Object.freeze({
      xPt: 125,
      yPt: 256,
      widthPt: 80,
      heightPt: 40,
    });
    const blocker = Object.freeze({
      kind: 'table' as const,
      overlap: 'never' as const,
      occurrenceId: 'blocker-never-table',
      paragraphId: 7,
      bounds: blockerBounds,
      exclusionBounds: Object.freeze({
        xPt: 120,
        yPt: 251,
        widthPt: 90,
        heightPt: 50,
      }),
    });

    const resolution = resolveFloatingTablePlacementInTransaction(
      placement,
      retainedReferenceFrames,
      beginFloatingTablePlacementTransaction([blocker], 7),
    );

    // Same paragraph ID prevents the Word compatibility rule from masking the
    // normative blocker-side fact. Placement clears the raw edge at x=205,
    // not the padded exclusion edge at x=210.
    expect(resolution.placement.bounds.xPt).toBe(205);
    expect(resolution.transaction.delta[0]).toMatchObject({
      overlap: 'overlap',
      paragraphId: 7,
    });
  });

  it('re-resolves an upright nested float in a fresh physical-page domain', () => {
    const logicalRetained = retainedFloatingPlacement({
      horzAnchor: 'page', vertAnchor: 'margin', xPt: 10, yPt: 0,
    });
    const physicalFrames: FloatingTableReferenceFramesPt = Object.freeze({
      page: Object.freeze({ xPt: 0, yPt: 0, widthPt: 200, heightPt: 300 }),
      margin: Object.freeze({ xPt: 24, yPt: 20, widthPt: 146, heightPt: 240 }),
      text: Object.freeze({ xPt: 150, yPt: 20, widthPt: 50, heightPt: 80 }),
    });

    const physical = resolveFloatingTablePlacementsInFreshRegistry(
      [logicalRetained],
      () => physicalFrames,
      'upright-physical-page-points',
      'upright-physical-page:0',
    );

    expect(physical.transaction).toMatchObject({
      coordinateSpace: 'upright-physical-page-points',
      flowDomainId: 'upright-physical-page:0',
      nextParagraphId: 1,
    });
    expect(physical.placements[0]).toMatchObject({
      xPt: 10,
      yPt: 20,
      bounds: { xPt: 10, yPt: 20 },
    });
    expect(physical.placements[0]?.bounds).not.toBe(logicalRetained.anchorBounds);
  });

  it('resolves multiple floats in source order without mutating or duplicating the base snapshot', () => {
    const first = Object.freeze({
      ...retainedFloatingPlacement({
        horzAnchor: 'page', vertAnchor: 'page', xPt: 100, yPt: 100,
      }),
      occurrenceId: 'first',
    });
    const second = Object.freeze({
      ...retainedFloatingPlacement({
        horzAnchor: 'page', vertAnchor: 'page', xPt: 100, yPt: 100,
      }),
      occurrenceId: 'second',
    });
    const base = Object.freeze([]);
    const started = beginFloatingTablePlacementTransaction(base, 7);
    const firstResolution = resolveFloatingTablePlacementInTransaction(
      first,
      retainedReferenceFrames,
      started,
    );
    const secondResolution = resolveFloatingTablePlacementInTransaction(
      second,
      retainedReferenceFrames,
      firstResolution.transaction,
    );
    const duplicateProbe = resolveFloatingTablePlacementInTransaction(
      second,
      retainedReferenceFrames,
      secondResolution.transaction,
    );

    expect(base).toEqual([]);
    expect(secondResolution.transaction.delta.map((entry) => entry.occurrenceId)).toEqual([
      'first', 'second',
    ]);
    expect(secondResolution.placement.xPt).toBe(
      firstResolution.placement.bounds.xPt
        + firstResolution.placement.bounds.widthPt,
    );
    expect(secondResolution.placement.exclusionBounds).toMatchObject({
      xPt: secondResolution.placement.xPt - second.positioning.leftFromTextPt,
      widthPt: secondResolution.placement.bounds.widthPt
        + second.positioning.leftFromTextPt
        + second.positioning.rightFromTextPt,
    });
    expect(duplicateProbe.transaction).toBe(secondResolution.transaction);
    expect(duplicateProbe.placement.bounds).toBe(
      secondResolution.transaction.delta[1]!.bounds,
    );
    expect(secondResolution.transaction.nextParagraphId).toBe(9);
  });

  it('preserves the established page-edge slack in retained point-space placement', () => {
    const frames: FloatingTableReferenceFramesPt = Object.freeze({
      ...retainedReferenceFrames,
      page: Object.freeze({ xPt: 0, yPt: 0, widthPt: 259.75, heightPt: 800 }),
    });
    const first = Object.freeze({
      ...retainedFloatingPlacement({
        horzAnchor: 'page', vertAnchor: 'page', xPt: 100, yPt: 100,
      }),
      occurrenceId: 'page-edge-first',
    });
    const second = Object.freeze({
      ...first,
      occurrenceId: 'page-edge-second',
    });
    const firstResolution = resolveFloatingTablePlacementInTransaction(
      first,
      frames,
      beginFloatingTablePlacementTransaction([], 0),
    );
    const secondResolution = resolveFloatingTablePlacementInTransaction(
      second,
      frames,
      firstResolution.transaction,
    );

    expect(secondResolution.placement.bounds).toMatchObject({
      xPt: 180,
      yPt: 100,
    });
  });

  it('preserves the established sub-epsilon edge tolerance in retained placement', () => {
    const first = Object.freeze({
      ...retainedFloatingPlacement({
        horzAnchor: 'page', vertAnchor: 'page', xPt: 100, yPt: 100,
      }),
      occurrenceId: 'epsilon-first',
    });
    const second = Object.freeze({
      ...retainedFloatingPlacement({
        horzAnchor: 'page', vertAnchor: 'page', xPt: 179.995, yPt: 100,
      }),
      occurrenceId: 'epsilon-second',
    });
    const firstResolution = resolveFloatingTablePlacementInTransaction(
      first,
      retainedReferenceFrames,
      beginFloatingTablePlacementTransaction([], 0),
    );
    const secondResolution = resolveFloatingTablePlacementInTransaction(
      second,
      retainedReferenceFrames,
      firstResolution.transaction,
    );

    expect(secondResolution.placement.bounds).toMatchObject({
      xPt: 179.995,
      yPt: 100,
    });
  });

  it('resolves text-relative offsets and exclusion padding entirely in point space', () => {
    const placement = retainedFloatingPlacement();

    const resolved = resolveFloatingTablePlacement(placement, retainedReferenceFrames);

    expect(resolved).toMatchObject({
      kind: 'resolved-floating-table-placement',
      occurrenceId: placement.occurrenceId,
      xPt: 125,
      yPt: 256,
      bounds: { xPt: 125, yPt: 256, widthPt: 80, heightPt: 40 },
      exclusionBounds: { xPt: 124, yPt: 253, widthPt: 83, heightPt: 47 },
      overlap: 'never',
      child: placement.child,
    });
    expect(resolved.source).toBe(placement);
  });

  it('retains page and margin reference frames for aligned placement', () => {
    const placement = retainedFloatingPlacement({
      horzAnchor: 'page',
      xPt: 999,
      xAlign: 'center',
      vertAnchor: 'margin',
      yPt: 999,
      yAlign: 'bottom',
    });

    const resolved = resolveFloatingTablePlacement(placement, retainedReferenceFrames);

    expect(resolved.xPt).toBe(260);
    expect(resolved.yPt).toBe(688);
  });

  it('refuses a cell-owned alignment frame without an acquired host content band', () => {
    const placement = retainedFloatingPlacement({ widthBasis: 'host-cell-content', textAnchor: 'cell-start' });
    expect(() => resolveFloatingTablePlacement(placement, retainedReferenceFrames))
      .toThrow('Cell-owned grid frame lacks its content band');
  });

  it('uses the text frame when horizontal positioning was omitted', () => {
    const placement = retainedFloatingPlacement({
      horzAnchor: 'page',
      horzSpecified: false,
      xPt: 0,
    });

    expect(resolveFloatingTablePlacement(placement, retainedReferenceFrames).xPt).toBe(120);
  });

  it('uses grid width and table advance when row-specific origins widen flow bounds', () => {
    const placement = retainedFloatingPlacement();
    const rowExceptionEnvelope = Object.freeze({
      ...placement,
      child: Object.freeze({
        ...placement.child,
        flowBounds: Object.freeze({
          ...placement.child.flowBounds,
          widthPt: 140,
          heightPt: 90,
        }),
      }),
    });

    const resolved = resolveFloatingTablePlacement(
      rowExceptionEnvelope,
      retainedReferenceFrames,
    );

    expect(resolved.bounds.widthPt).toBe(80);
    expect(resolved.bounds.heightPt).toBe(40);
  });

  it('reuses the acquisition-time text-anchor offset after overlap resolution', () => {
    const placement = Object.freeze({
      ...retainedFloatingPlacement(),
      acquiredTextOffsetPt: Object.freeze({ xPt: 30, yPt: 40 }),
    });

    const resolved = resolveFloatingTablePlacement(placement, retainedReferenceFrames);

    expect(resolved.xPt).toBe(150);
    expect(resolved.yPt).toBe(290);
  });
});
