import { describe, it, expect } from 'vitest';
import {
  computeUniformVisibleWindow,
  computeVisibleRange,
  computeVisibleWindow,
  createVirtualScrollGeometry,
  resolveItemStartScrollTop,
} from './virtual-scroll.js';

// Design §5.1 contract. computeVisibleRange(heights, gap, scrollTop,
// viewportHeight, overscan) → { start, end, topIndex, offsets, totalHeight }.
// offsets[i] = Σ heights[0..i-1] + i*gap; totalHeight = Σ heights + (n-1)*gap
// (gap BETWEEN items only, none after the last). start/end include overscan
// (clamped to [0,n-1]); topIndex is the first item intersecting the viewport
// top, EXCLUDING overscan.

describe('computeVisibleRange — offsets + totalHeight (gap arithmetic)', () => {
  it('computes prefix-sum offsets with a leading gap per item and gap only between items', () => {
    // 3 uniform 100-tall items, gap 10: offsets 0, 110, 220; total = 300 + 2*10 = 320.
    const r = computeVisibleRange([100, 100, 100], 10, 0, 1000, 0);
    expect(r.offsets).toEqual([0, 110, 220]);
    expect(r.totalHeight).toBe(320);
  });

  it('variable heights: offsets follow Σ heights[0..i-1] + i*gap', () => {
    // heights 50, 200, 30, gap 8 → offsets 0, 58, 266; total = 280 + 2*8 = 296.
    const r = computeVisibleRange([50, 200, 30], 8, 0, 10, 0);
    expect(r.offsets).toEqual([0, 58, 266]);
    expect(r.totalHeight).toBe(296);
  });

  it('single item has no gap in totalHeight', () => {
    const r = computeVisibleRange([120], 16, 0, 500, 0);
    expect(r.offsets).toEqual([0]);
    expect(r.totalHeight).toBe(120);
  });

  it('tolerates zero-height items', () => {
    // heights 100, 0, 100, gap 10 → offsets 0, 110, 120; total = 200 + 2*10 = 220.
    const r = computeVisibleRange([100, 0, 100], 10, 0, 1000, 0);
    expect(r.offsets).toEqual([0, 110, 120]);
    expect(r.totalHeight).toBe(220);
  });
});

describe('computeVisibleRange — visible range (uniform)', () => {
  // 10 items of 100, gap 0. Viewport 250 tall.
  const heights = Array.from({ length: 10 }, () => 100);

  it('at scrollTop 0 with no overscan mounts the items beginning within the viewport', () => {
    const r = computeVisibleRange(heights, 0, 0, 250, 0);
    // items 0 (0..100), 1 (100..200), 2 (200..300 begins at 200 < 250) begin within [0,250).
    expect(r.topIndex).toBe(0);
    expect(r.start).toBe(0);
    expect(r.end).toBe(2);
  });

  it('scrolled to 250: topIndex is the item under the viewport top', () => {
    const r = computeVisibleRange(heights, 0, 250, 250, 0);
    // top edge 250 → item 2 (200..300) intersects the top. Bottom edge 500 →
    // items whose top < 500: up to index 4 (400..500 begins at 400 < 500).
    expect(r.topIndex).toBe(2);
    expect(r.start).toBe(2);
    expect(r.end).toBe(4);
  });
});

describe('computeVisibleRange — overscan clipping at both ends', () => {
  const heights = Array.from({ length: 10 }, () => 100);

  it('clips overscan at the top (cannot go below 0)', () => {
    const r = computeVisibleRange(heights, 0, 0, 250, 2);
    expect(r.topIndex).toBe(0);
    expect(r.start).toBe(0);        // 0 - 2 clamped to 0
    expect(r.end).toBe(4);          // lastVisible 2 + overscan 2
  });

  it('clips overscan at the bottom (cannot exceed n-1)', () => {
    // scrollTop near the end so lastVisible ≈ 9; +2 overscan clamps to 9.
    const r = computeVisibleRange(heights, 0, 900, 250, 2);
    expect(r.topIndex).toBe(9);
    expect(r.end).toBe(9);          // 9 + 2 clamped to 9
    expect(r.start).toBe(7);        // 9 - 2
  });

  it('applies overscan symmetrically in the middle', () => {
    const r = computeVisibleRange(heights, 0, 300, 200, 1);
    // top 300 → topIndex 3; bottom 500 → lastVisible 4 (top 400 < 500).
    expect(r.topIndex).toBe(3);
    expect(r.start).toBe(2);        // 3 - 1
    expect(r.end).toBe(5);          // 4 + 1
  });
});

describe('computeVisibleRange — topIndex vs start distinction', () => {
  it('topIndex excludes overscan; start includes it', () => {
    const heights = Array.from({ length: 8 }, () => 50);
    const r = computeVisibleRange(heights, 0, 200, 100, 2);
    // top 200 → topIndex 4 (200..250). start = 4 - 2 = 2 (overscan). They differ.
    expect(r.topIndex).toBe(4);
    expect(r.start).toBe(2);
    expect(r.topIndex).not.toBe(r.start);
  });
});

describe('computeVisibleRange — degenerate inputs', () => {
  it('empty heights → empty mount range', () => {
    const r = computeVisibleRange([], 16, 0, 500, 1);
    expect(r).toEqual({ start: 0, end: -1, topIndex: 0, offsets: [], totalHeight: 0 });
    // start > end ⇒ a `for (i = start; i <= end; i++)` mount loop runs zero times.
    expect(r.start).toBeGreaterThan(r.end);
  });

  it('single item, scrollTop 0', () => {
    const r = computeVisibleRange([300], 16, 0, 500, 1);
    expect(r.topIndex).toBe(0);
    expect(r.start).toBe(0);
    expect(r.end).toBe(0);
    expect(r.totalHeight).toBe(300);
  });

  it('scrollTop past the end clamps topIndex/end to the last item', () => {
    const heights = [100, 100, 100];
    const r = computeVisibleRange(heights, 0, 99999, 250, 1);
    expect(r.topIndex).toBe(2);     // clamped to n-1
    expect(r.end).toBe(2);          // clamped to n-1
    expect(r.start).toBe(1);        // 2 - 1
  });

  it('negative scrollTop behaves like 0', () => {
    const heights = [100, 100, 100];
    const r = computeVisibleRange(heights, 0, -50, 150, 0);
    expect(r.topIndex).toBe(0);
    expect(r.start).toBe(0);
  });

  it('viewport top inside a gap is attributed to the PRECEDING item', () => {
    // offsets = [0, 110, 220]; scrollTop 105 is strictly inside the gap 100..110
    // between items 0 and 1. Convention (pinned): the gap is trailing padding of
    // the preceding item ⇒ topIndex 0 (mount-safe; flips to 1 exactly at 110).
    const r = computeVisibleRange([100, 100, 100], 10, 105, 50, 0);
    expect(r.topIndex).toBe(0);
    expect(r.start).toBe(0);
    expect(r.end).toBe(1); // item 1 begins (110) before the viewport bottom (155)
    // Exactly at the next item's offset the attribution flips.
    expect(computeVisibleRange([100, 100, 100], 10, 110, 50, 0).topIndex).toBe(1);
  });
});

describe('computeVisibleRange — leading/trailing padding (desk margin)', () => {
  // pad shifts offsets down by `leading` and extends totalHeight by leading+trailing.
  // offsets[i] = leading + Σ heights[0..i-1] + i*gap; total = leading + Σh + (n-1)*gap + trailing.

  it('leading pad shifts every offset down and total up by leading+trailing', () => {
    // 3 uniform 100-tall items, gap 10, pad {24, 24}:
    // offsets 24, 134, 244; total = 24 + 300 + 2*10 + 24 = 368.
    const r = computeVisibleRange([100, 100, 100], 10, 0, 1000, 0, { leading: 24, trailing: 24 });
    expect(r.offsets).toEqual([24, 134, 244]);
    expect(r.totalHeight).toBe(368);
  });

  it('omitted pad is identical to pad {0,0} (backward-compatible)', () => {
    const bare = computeVisibleRange([50, 200, 30], 8, 40, 100, 1);
    const padded = computeVisibleRange([50, 200, 30], 8, 40, 100, 1, { leading: 0, trailing: 0 });
    expect(padded).toEqual(bare);
  });

  it('viewport top inside the leading pad ⇒ topIndex 0 (existing clamp)', () => {
    // leading 50: offsets = [50, 160, 270]. scrollTop 20 is INSIDE the leading pad
    // (below every offset), so the search finds lo=0 ⇒ topIndex clamp(-1,0,n-1)=0.
    const r = computeVisibleRange([100, 100, 100], 10, 20, 80, 0, { leading: 50 });
    expect(r.topIndex).toBe(0);
    expect(r.start).toBe(0);
    // bottom edge 100 → item 0 (top 50) begins before it; item 1 (top 160) does not.
    expect(r.end).toBe(0);
  });

  it('viewport ending inside the trailing pad ⇒ end = n-1, no overrun', () => {
    // leading 0, trailing 100: offsets = [0, 110, 220], total = 300+2*10+100 = 420.
    // Scroll near the end so the viewport bottom (200+250=450) lands in the trailing
    // pad; end must clamp to the last real item (2), never past it.
    const r = computeVisibleRange([100, 100, 100], 10, 200, 250, 1, { trailing: 100 });
    // offsets [0,110,220]; scrollTop 200 → largest offset ≤ 200 is 110 (item 1).
    expect(r.topIndex).toBe(1);
    expect(r.end).toBe(2); // lastVisible (2) + overscan clamped to n-1 — no overrun
    expect(r.totalHeight).toBe(420);
  });

  it('pad + gap compose (both contribute, distinctly)', () => {
    // heights [50, 200, 30], gap 8, pad {12, 40}:
    // offsets = 12, 12+50+8=70, 12+250+16=278; total = 12 + 280 + 16 + 40 = 348.
    const r = computeVisibleRange([50, 200, 30], 8, 0, 10, 0, { leading: 12, trailing: 40 });
    expect(r.offsets).toEqual([12, 70, 278]);
    expect(r.totalHeight).toBe(348);
  });

  it('leading-only pad', () => {
    const r = computeVisibleRange([100, 100], 10, 0, 1000, 0, { leading: 30 });
    expect(r.offsets).toEqual([30, 140]);
    expect(r.totalHeight).toBe(30 + 200 + 10); // 240
  });

  it('trailing-only pad', () => {
    const r = computeVisibleRange([100, 100], 10, 0, 1000, 0, { trailing: 30 });
    expect(r.offsets).toEqual([0, 110]);
    expect(r.totalHeight).toBe(200 + 10 + 30); // 240
  });

  it('empty input ignores pad entirely (empty-doc no-op contract)', () => {
    const r = computeVisibleRange([], 16, 0, 500, 1, { leading: 24, trailing: 24 });
    expect(r).toEqual({ start: 0, end: -1, topIndex: 0, offsets: [], totalHeight: 0 });
  });
});

describe('cached virtual-scroll geometry', () => {
  it('reads variable heights once and reuses the same offsets for every scroll query', () => {
    let reads = 0;
    const heights = new Proxy(Array.from({ length: 10_000 }, () => 100), {
      get(target, property, receiver) {
        if (typeof property === 'string' && /^\d+$/.test(property)) reads++;
        return Reflect.get(target, property, receiver);
      },
    });
    const geometry = createVirtualScrollGeometry(heights, 10, {
      leading: 16,
      trailing: 16,
    });
    expect(reads).toBe(10_000);

    for (let scrollTop = 0; scrollTop < 100_000; scrollTop += 997) {
      const window = computeVisibleWindow(geometry, scrollTop, 800, 1);
      expect(window.offsets).toBe(geometry.offsets);
    }
    expect(reads).toBe(10_000);
  });

  it('computes a huge uniform window arithmetically without materializing offsets', () => {
    const window = computeUniformVisibleWindow(
      1_000_000,
      100,
      10,
      16 + 543_210 * 110,
      250,
      2,
      { leading: 16, trailing: 24 },
    );
    expect(window).toEqual({
      start: 543_208,
      end: 543_214,
      topIndex: 543_210,
      totalHeight: 16 + 1_000_000 * 100 + 999_999 * 10 + 24,
    });
  });

  it('matches the existing range semantics for uniform heights, including padding and gaps', () => {
    const heights = Array.from({ length: 20 }, () => 100);
    const expected = computeVisibleRange(
      heights,
      10,
      355,
      250,
      2,
      { leading: 24, trailing: 40 },
    );
    expect(computeUniformVisibleWindow(
      heights.length,
      100,
      10,
      355,
      250,
      2,
      { leading: 24, trailing: 40 },
    )).toEqual({
      start: expected.start,
      end: expected.end,
      topIndex: expected.topIndex,
      totalHeight: expected.totalHeight,
    });
  });
});

describe('fractional item-start navigation targets', () => {
  // heights 100.4 × 3, gap 10 → offsets [0, 110.4, 220.8]. Visible-range
  // boundaries stay exact for manual scrolling and arbitrary coordinate lookup.
  const heights = [100.4, 100.4, 100.4];

  it('keeps a viewport just before a fractional boundary on the preceding item', () => {
    expect(computeVisibleRange(heights, 10, 110, 50, 0).topIndex).toBe(0);
    expect(computeVisibleRange(heights, 10, 220, 50, 0).topIndex).toBe(1);
    expect(computeUniformVisibleWindow(3, 100.4, 10, 110, 50, 0).topIndex).toBe(0);
  });

  it('keeps the exact item boundary inclusive', () => {
    expect(computeVisibleRange(heights, 10, 110.4, 50, 0).topIndex).toBe(1);
    expect(computeUniformVisibleWindow(3, 100.4, 10, 110.4, 50, 0).topIndex).toBe(1);
  });

  it('rounds only a programmatic item-start target forward and respects maxTop', () => {
    expect(resolveItemStartScrollTop(110.4, 500)).toBe(111);
    expect(resolveItemStartScrollTop(110, 500)).toBe(110);
    expect(resolveItemStartScrollTop(110.4, 110.6)).toBe(110.6);
    expect(resolveItemStartScrollTop(110.4, 110.2)).toBe(110.2);
    expect(resolveItemStartScrollTop(-5, 500)).toBe(0);
    expect(resolveItemStartScrollTop(50, -10)).toBe(0);
  });
});
