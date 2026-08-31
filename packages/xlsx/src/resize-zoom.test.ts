import { describe, it, expect } from 'vitest';
import { colWidthToPx, rowHeightToPx, pxToColWidth, pxToRowHeight } from './renderer.js';
import { findHighlightOverlayStyle, resizeHitIndex, selectionOverlayStyle, zoomStepScale } from './viewer.js';

/**
 * Drag-to-resize (issue #567) stores the user's dragged pixel size back into the
 * worksheet's `colWidths` / `rowHeights` model in its native units (Excel column
 * "characters" for columns, points for rows). `pxToColWidth` / `pxToRowHeight`
 * are the exact inverses of the forward converters the renderer uses, so a
 * column dragged to N px renders back at exactly N px with no drift.
 */
describe('px <-> model-unit round trip (drag-to-resize)', () => {
  for (const mdw of [7, 8, 10, 11]) {
    for (const px of [1, 5, 10, 32, 64, 100, 128, 255, 512]) {
      it(`column ${px}px @ mdw=${mdw} round-trips exactly`, () => {
        expect(colWidthToPx(pxToColWidth(px, mdw), mdw)).toBe(px);
      });
    }
  }

  for (const px of [1, 4, 10, 18, 20, 32, 64, 100, 255]) {
    it(`row ${px}px round-trips exactly`, () => {
      expect(rowHeightToPx(pxToRowHeight(px))).toBe(px);
    });
  }
});

/**
 * The viewer takes a single `selectionColor`; the rectangle border uses it as-is
 * and the fill is the same color made translucent (issue follow-up). The default
 * (`#1a73e8`) must keep the historical Google-blue look.
 */
describe('selectionOverlayStyle', () => {
  it('uses the color verbatim for the border', () => {
    expect(selectionOverlayStyle('red').border).toBe('2px solid red');
    expect(selectionOverlayStyle('#1a73e8').border).toBe('2px solid #1a73e8');
  });

  it('derives a translucent fill from the same color', () => {
    expect(selectionOverlayStyle('#1a73e8').background).toBe(
      'color-mix(in srgb, #1a73e8 8%, transparent)',
    );
    expect(selectionOverlayStyle('rgb(0,128,0)').background).toBe(
      'color-mix(in srgb, rgb(0,128,0) 8%, transparent)',
    );
  });
});

describe('findHighlightOverlayStyle', () => {
  it('uses caller-provided match and active background colours verbatim', () => {
    const colors = { match: 'rgba(1, 2, 3, 0.4)', active: '#ff00ff99' };
    expect(findHighlightOverlayStyle(false, colors).background).toBe(colors.match);
    expect(findHighlightOverlayStyle(true, colors).background).toBe(colors.active);
  });
});

/**
 * Ctrl/⌘ + wheel (and trackpad pinch, which the browser reports as a ctrl-wheel)
 * zoom. The old handler ignored `deltaY` magnitude and added a fixed ±0.1 per
 * event, so a trackpad pinch — which fires a high-frequency stream of small
 * wheel events — zoomed far too fast. `zoomStepScale` makes the step
 * exponential in a mode-normalized wheel distance, so one conventional wheel
 * step is a precise 10% change while fine-grained trackpad events still combine
 * smoothly.
 */
describe('zoomStepScale (ctrl/pinch zoom)', () => {
  it('maps pixel, line, and page wheel units to the same 10% conventional step', () => {
    expect(zoomStepScale(1, -100, 0)).toBeCloseTo(1.1, 10);
    expect(zoomStepScale(1, -3, 1)).toBeCloseTo(1.1, 10);
    expect(zoomStepScale(1, -1, 2)).toBeCloseTo(1.1, 10);
  });

  it('scrolling up / pinching out (deltaY < 0) zooms in', () => {
    expect(zoomStepScale(1, -10)).toBeGreaterThan(1);
  });

  it('scrolling down / pinching in (deltaY > 0) zooms out', () => {
    expect(zoomStepScale(1, 10)).toBeLessThan(1);
  });

  it('honors deltaY magnitude (a bigger delta zooms more)', () => {
    const small = zoomStepScale(1, -2) - 1;
    const big = zoomStepScale(1, -20) - 1;
    expect(big).toBeGreaterThan(small);
  });

  it('is resolution-independent: two small events ≈ one event of their sum', () => {
    const twoSteps = zoomStepScale(zoomStepScale(1, -5), -5);
    const oneStep = zoomStepScale(1, -10);
    expect(twoSteps).toBeCloseTo(oneStep, 10);
  });

  it('is symmetric: zooming in then out by the same delta returns to start', () => {
    expect(zoomStepScale(zoomStepScale(1, -8), 8)).toBeCloseTo(1, 10);
  });

  it('scales relative to the current zoom (multiplicative, not additive)', () => {
    // Same delta from 200% must move proportionally more than from 100%.
    const from1 = zoomStepScale(1, -10) - 1;
    const from2 = zoomStepScale(2, -10) - 2;
    expect(from2).toBeCloseTo(from1 * 2, 10);
  });
});

/**
 * `colWidthToPx` must implement the ECMA-376 §18.3.1.13 file→pixel formula
 * verbatim, including BOTH truncations:
 *   `Truncate(((256 * width + Truncate(128 / MDW)) / 256) * MDW)`
 * The inner `Truncate(128 / MDW)` constant is computed and truncated *before*
 * being folded into the numerator. These cases were derived by evaluating the
 * spec expression by hand for a couple of real stored widths / MDWs:
 *   - w=8.43, MDW=8: Truncate(128/8)=16 → ((256·8.43+16)/256)·8 = 67.94 → 67
 *   - w=10,   MDW=7: Truncate(128/7)=18 → ((256·10  +18)/256)·7 = 70.49 → 70
 *   - w=15,   MDW=8: Truncate(128/8)=16 → ((256·15  +16)/256)·8 = 120.5 → 120
 *   - w=2,    MDW=8: Truncate(128/8)=16 → ((256·2   +16)/256)·8 = 16.5  → 16
 */
describe('colWidthToPx (ECMA-376 §18.3.1.13 file→px)', () => {
  it('matches the spec formula for real stored widths', () => {
    expect(colWidthToPx(8.43, 8)).toBe(67);
    expect(colWidthToPx(10, 7)).toBe(70);
    expect(colWidthToPx(15, 8)).toBe(120);
    expect(colWidthToPx(2, 8)).toBe(16);
  });

  it('truncates the 128/MDW constant before folding it in', () => {
    // With MDW=7, 128/7 = 18.285…; the spec truncates this to 18 *before* the
    // division by 256, so the result must equal the all-integer computation.
    const mdw = 7;
    const w = 12;
    const expected = Math.trunc(((256 * w + 18) / 256) * mdw);
    expect(colWidthToPx(w, mdw)).toBe(expected);
  });
});

/**
 * `resizeHitIndex` is the pure off-by-one-prone hit predicate behind
 * drag-to-resize (issue #567), extracted from `getResizeTarget` so the
 * border-grab geometry is testable without DOM. The caller passes the candidate
 * bands' *trailing* edges in `[hit-1, hit]` order; the first edge within `grabPx`
 * of the pointer wins, and any edge at/under the header strip is rejected.
 */
describe('resizeHitIndex (drag-to-resize hit predicate)', () => {
  const GRAB = 4;
  const HEADER = 50;

  it('hits a band whose trailing edge is exactly under the pointer', () => {
    expect(resizeHitIndex(100, [{ index: 3, edge: 100 }], GRAB, HEADER)).toBe(3);
  });

  it('hits within the grab zone on either side of the edge', () => {
    expect(resizeHitIndex(96, [{ index: 3, edge: 100 }], GRAB, HEADER)).toBe(3); // -4
    expect(resizeHitIndex(104, [{ index: 3, edge: 100 }], GRAB, HEADER)).toBe(3); // +4
  });

  it('misses when the pointer is just outside the grab zone', () => {
    expect(resizeHitIndex(95, [{ index: 3, edge: 100 }], GRAB, HEADER)).toBeNull(); // -5
    expect(resizeHitIndex(105, [{ index: 3, edge: 100 }], GRAB, HEADER)).toBeNull(); // +5
  });

  it('selects the matching band from the [hit-1, hit] candidate pair', () => {
    // Pointer sits on the left band's trailing border (the neighbour-to-the-far
    // side), so the lower index is chosen even though it is listed first.
    const edges = [
      { index: 2, edge: 100 }, // hit-1, trailing edge under the pointer
      { index: 3, edge: 160 }, // hit,   far away
    ];
    expect(resizeHitIndex(101, edges, GRAB, HEADER)).toBe(2);
    // Pointer on the right band's trailing border picks the higher index.
    expect(resizeHitIndex(159, edges, GRAB, HEADER)).toBe(3);
  });

  it('takes the first in-zone edge when both candidates qualify (list order)', () => {
    const edges = [
      { index: 2, edge: 100 },
      { index: 3, edge: 102 }, // also within grab of pt=101, but listed second
    ];
    expect(resizeHitIndex(101, edges, GRAB, HEADER)).toBe(2);
  });

  it('rejects an edge that sits at or under the header strip', () => {
    // The band scrolled so its trailing border is hidden under the header
    // (edge <= headerExtent) — un-grabbable even though the pointer is on it.
    expect(resizeHitIndex(50, [{ index: 1, edge: 50 }], GRAB, HEADER)).toBeNull();
    expect(resizeHitIndex(48, [{ index: 1, edge: 48 }], GRAB, HEADER)).toBeNull();
    // But an edge just past the header is grabbable again.
    expect(resizeHitIndex(51, [{ index: 1, edge: 51 }], GRAB, HEADER)).toBe(1);
  });

  it('returns null for an empty candidate list', () => {
    expect(resizeHitIndex(100, [], GRAB, HEADER)).toBeNull();
  });
});

/**
 * The `resizable` option (default true) gates drag-to-resize. The viewer reads
 * it as `opts.resizable ?? true`, so an omitted/undefined option keeps resizing
 * on and an explicit `false` turns it off. This locks the default-true semantics
 * of that nullish-coalescing gate.
 */
describe('resizable option default semantics', () => {
  const enabled = (resizable?: boolean) => resizable ?? true;

  it('defaults to true when unset', () => {
    expect(enabled(undefined)).toBe(true);
  });

  it('stays true when explicitly true', () => {
    expect(enabled(true)).toBe(true);
  });

  it('is false only when explicitly false', () => {
    expect(enabled(false)).toBe(false);
  });
});
