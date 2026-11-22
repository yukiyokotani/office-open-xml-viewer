import { orientation, visiblePieces } from './run-clip.js';
import type { DocxClipRegion, DocxTextRunInfo } from './types.js';

/**
 * The overlay geometry of projected text. Every overlay of a page (find
 * highlights, comment tints in the viewer and in the comment API, the
 * visibility and extent of text-layer pieces) reads a run's geometry through
 * this module and nowhere else, so they cannot disagree about where text is
 * painted or whether it is visible.
 *
 * A run (DocxTextRunInfo) has its own frame: `x` along the run from its
 * origin, `y` across it, turned by its `transform`. In that frame:
 *
 * 1. The glyph rect of a cluster is where its paint operation draws it:
 *    along the run its glyph box (`clusterBoxes[].glyphX/glyphWidth`, which
 *    layout reads from the retained paint operation), across the run the
 *    run's font box (`highlightBounds`), or a tate-chu-yoko character's own
 *    interval across the run. It never includes tracking, justification,
 *    distribution, kashida elongation, grid padding or additional line
 *    spacing.
 * 2. The advance rect of a cluster is its advance box along the run (pen to
 *    next pen) over the run's line box. The caret, hit testing and selection
 *    extents use it.
 * 3. A run's clips (`clips`) apply on both axes. Text is visible exactly
 *    where a glyph rect has area inside every clip; an advance or a line box
 *    never makes text visible, and never hides it.
 * 4. A find highlight is the glyph rects of the matched characters cut to
 *    the clips: it lies inside the painted glyphs and the clips, and covers
 *    all of both. A comment highlight marks a range of source runs, and
 *    everything in the range belongs to it: each cluster's glyph rect
 *    widened along the run to its advance (the tracking, justification or
 *    elongation after it), and the box of each tab and inline object of the
 *    range (`inlineContent`), all cut to the clips; a cluster whose glyph
 *    the clips hide whole marks nothing. A tab takes the extent
 *    across the run of the marked text it touches. Nothing outside the
 *    range is marked.
 * 5. Rects join only where they touch or overlap along the run over the
 *    same interval across it (and, between runs, in the same orientation
 *    under the same clips). A gap is never bridged: not tracking, not text
 *    outside the marked range, not the part a clip removed. Where glyphs of
 *    different extents across the run are painted over each other, the
 *    shared area is covered once: the rects of one marking never overlap.
 *
 * Cutting a rect to clips yields rects in the run's frame. That is exact
 * for clips whose edges are parallel to the run's axes (cells, rows and text
 * boxes, turned with their text or by quarter turns); for a clip at another
 * angle each visible piece is given by its bounding rect, and the nested
 * clip containers of the viewer's layers (OverlayClipFrames) cut the drawn
 * box exactly.
 */
export interface RunRect {
  /** Along the run, from its origin. */
  readonly x: number;
  /** Across the run, from its origin. */
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/** A rect on the page, placed like DocxTextRunInfo.highlightBounds: its
 * corner, its lengths along its own axes, its orientation about the corner
 * and the clips it lies under. */
export interface OverlayBox {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly transform?: string;
  readonly clips?: readonly Readonly<DocxClipRegion>[];
}

type Run = Readonly<DocxTextRunInfo>;

/** Rects closer than this (CSS px) touch. */
const TOUCH = 1e-6;

/** The run's font box across the run, from its origin (invariant 1). */
export function runFontBox(run: Run): Readonly<{ y: number; height: number }> {
  const font = run.highlightBounds;
  if (!font) return { y: 0, height: run.h };
  const [, , c, d] = orientation(run.transform);
  return { y: c * (font.x - run.x) + d * (font.y - run.y), height: font.height };
}

/** The first cluster box that ends after `start` (boxes are in source order). */
function firstBoxAfter(boxes: NonNullable<Run['clusterBoxes']>, start: number): number {
  let low = 0;
  let high = boxes.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (boxes[middle]!.end <= start) low = middle + 1;
    else high = middle;
  }
  return low;
}

type ClusterBox = NonNullable<Run['clusterBoxes']>[number];

/** The glyph or advance rect of the share `[from, to]` of one cluster box;
 * null when it has no area. */
function clusterRect(
  run: Run,
  box: ClusterBox,
  glyph: boolean,
  across: Readonly<{ y: number; height: number }>,
  from = 0,
  to = 1,
): RunRect | null {
  const x = glyph ? box.glyphX ?? box.x : box.x;
  const width = glyph ? box.glyphWidth ?? box.width : box.width;
  // A tate-chu-yoko character has its own interval across the run, and a
  // glyph painted off the font box (a vertical-text cell) its own too.
  const y = (glyph ? box.glyphY : undefined) ?? box.y ?? across.y;
  const height = (glyph ? box.glyphHeight : undefined) ?? box.height ?? across.height;
  const part = width * (to - from);
  if (!(part > 0 && height > 0)) return null;
  return { x: x + width * (box.rtl ? 1 - to : from), y, width: part, height };
}

function clusterRects(run: Run, start: number, end: number, glyph: boolean): RunRect[] {
  const boxes = run.clusterBoxes;
  const across = glyph ? runFontBox(run) : { y: 0, height: run.h };
  if (!boxes) {
    // A run not projected from a layout: its font box (or line box) whole.
    const font = glyph ? run.highlightBounds : undefined;
    const [a, b] = orientation(run.transform);
    const x = font ? a * (font.x - run.x) + b * (font.y - run.y) : 0;
    const width = font?.width ?? run.w;
    return (end > start || run.text.length === 0) && width > 0 && across.height > 0
      ? [{ x, y: across.y, width, height: across.height }] : [];
  }
  const rects: RunRect[] = [];
  for (let index = firstBoxAfter(boxes, start); index < boxes.length && boxes[index]!.start < end; index += 1) {
    const box = boxes[index]!;
    // A range that covers part of a cluster takes that share of its box.
    const length = Math.max(1, box.end - box.start);
    const rect = clusterRect(run, box, glyph, across,
      (Math.max(start, box.start) - box.start) / length, (Math.min(end, box.end) - box.start) / length);
    if (rect) rects.push(rect);
  }
  return rects;
}

/** The glyph rects (invariant 1) of the characters `[start, end)` of `run`,
 * one per cluster, in source order; clusters that paint nothing have none. */
export function runGlyphRects(run: Run, start = 0, end = run.text.length): RunRect[] {
  return clusterRects(run, start, end, true);
}

/** The advance rects (invariant 2) of the characters `[start, end)` of `run`. */
export function runAdvanceRects(run: Run, start = 0, end = run.text.length): RunRect[] {
  return clusterRects(run, start, end, false);
}

/** The range rects (invariant 4) of `run`: each cluster's glyph rect widened
 * along the run to its advance box, in source order; a run without cluster
 * boxes has its font box. A cluster whose glyph the run's clips hide whole
 * has none: what follows a hidden glyph is not marked without it. */
export function runRangeRects(run: Run): RunRect[] {
  const boxes = run.clusterBoxes;
  if (!boxes) return clusterRects(run, 0, run.text.length, true);
  const across = runFontBox(run);
  const rects: RunRect[] = [];
  for (const box of boxes) {
    const glyph = clusterRect(run, box, true, across);
    if (glyph && !runRectVisible(run, glyph)) continue;
    const left = Math.min(box.x, glyph?.x ?? box.x);
    const right = Math.max(box.x + box.width, glyph ? glyph.x + glyph.width : box.x + box.width);
    const y = glyph?.y ?? box.glyphY ?? box.y ?? across.y;
    const height = glyph?.height ?? box.glyphHeight ?? box.height ?? across.height;
    if (right > left && height > 0) rects.push({ x: left, y, width: right - left, height });
  }
  return rects;
}

/** The page box of `rect` of `run`. */
export function placeRunRect(run: Run, rect: RunRect): OverlayBox {
  const [a, b, c, d] = orientation(run.transform);
  return {
    x: run.x + a * rect.x + c * rect.y,
    y: run.y + b * rect.x + d * rect.y,
    width: rect.width,
    height: rect.height,
    ...(run.transform ? { transform: run.transform } : {}),
    ...(run.clips ? { clips: run.clips } : {}),
  };
}

/** The parts of `rect` of `run` that its clips leave visible, in the run's
 * frame (invariant 3); empty when they hide it. */
export function clipRunRect(run: Run, rect: RunRect): RunRect[] {
  if (!(rect.width > 0 && rect.height > 0)) return [];
  if (!run.clips) return [rect];
  const [a, b, c, d] = orientation(run.transform);
  return visiblePieces(placeRunRect(run, rect), run.clips).map((piece) => {
    let left = Infinity;
    let top = Infinity;
    let right = -Infinity;
    let bottom = -Infinity;
    for (const point of piece) {
      // Orientations are rotations or reflections: the inverse is the transpose.
      const along = a * (point.x - run.x) + b * (point.y - run.y);
      const across = c * (point.x - run.x) + d * (point.y - run.y);
      left = Math.min(left, along);
      right = Math.max(right, along);
      top = Math.min(top, across);
      bottom = Math.max(bottom, across);
    }
    // The cut never grows the rect.
    left = Math.max(left, rect.x);
    top = Math.max(top, rect.y);
    right = Math.min(right, rect.x + rect.width);
    bottom = Math.min(bottom, rect.y + rect.height);
    return { x: left, y: top, width: right - left, height: bottom - top };
  }).filter((part) => part.width > 0 && part.height > 0);
}

/** Whether the clips of `run` leave any of `rect` visible (invariant 3). */
export function runRectVisible(run: Run, rect: RunRect): boolean {
  return clipRunRect(run, rect).length > 0;
}

/** Whether any glyph of the characters `[start, end)` of `run` is painted
 * inside its clips: the one visibility decision of every overlay. */
export function runTextVisible(run: Run, start = 0, end = run.text.length): boolean {
  if (!run.clips) return runGlyphRects(run, start, end).length > 0;
  return runGlyphRects(run, start, end).some((rect) => runRectVisible(run, rect));
}

/** For each cluster box of `run`, whether its glyph is painted inside the
 * run's clips (invariant 3); null for a cluster that paints nothing. Empty
 * for a run without cluster boxes. */
export function runClusterVisibility(run: Run): (boolean | null)[] {
  const across = runFontBox(run);
  return (run.clusterBoxes ?? []).map((box) => {
    const rect = clusterRect(run, box, true, across);
    return rect ? runRectVisible(run, rect) : null;
  });
}

/** `rects` of one frame joined where they touch or overlap along the run
 * over the same interval across it (invariant 5), ordered across, then
 * along. A joined rect keeps the lowest `order` of its parts. */
function joinRects<Rect extends RunRect & { readonly order?: number }>(rects: readonly Rect[]): Rect[] {
  const sorted = rects.filter((rect) => rect.width > 0 && rect.height > 0)
    .sort((left, right) => (Math.abs(left.y - right.y) <= TOUCH ? 0 : left.y - right.y)
      || (Math.abs(left.height - right.height) <= TOUCH ? 0 : left.height - right.height) || left.x - right.x);
  const merged: Rect[] = [];
  for (const rect of sorted) {
    const last = merged.at(-1);
    if (last && Math.abs(last.y - rect.y) <= TOUCH && Math.abs(last.height - rect.height) <= TOUCH
      && rect.x <= last.x + last.width + TOUCH) {
      merged[merged.length - 1] = {
        ...last,
        width: Math.max(last.width, rect.x + rect.width - last.x),
        ...(last.order !== undefined && rect.order !== undefined ? { order: Math.min(last.order, rect.order) } : {}),
      };
    } else merged.push(rect);
  }
  return withoutOverlap(merged);
}

/** `rects` with the area two of them share covered once: glyphs of
 * different extents across the run that are painted over each other
 * (condensed spacing between two sizes, a raised run, a tate-chu-yoko cell)
 * cannot join into one rect, and marking their overlap twice would darken
 * it. Each group of overlapping rects is replaced by rects of the same
 * region that do not overlap: one per stretch along the run of each band
 * across it, bands with the same stretches joined. Rects that overlap
 * nothing are returned as they are. */
function withoutOverlap<Rect extends RunRect & { readonly order?: number }>(rects: readonly Rect[]): Rect[] {
  const overlaps = (left: RunRect, right: RunRect): boolean => (
    Math.min(left.x + left.width, right.x + right.width) - Math.max(left.x, right.x) > TOUCH
    && Math.min(left.y + left.height, right.y + right.height) - Math.max(left.y, right.y) > TOUCH);
  // Groups of rects connected by overlap, found by a sweep across the run:
  // a rect is compared only with the rects still open where it starts, so
  // rects of other lines are never looked at.
  const group = rects.map((_, index) => index);
  const find = (index: number): number => {
    while (group[index] !== index) index = group[index] = group[group[index]!]!;
    return index;
  };
  let any = false;
  const byStart = rects.map((_, index) => index).sort((left, right) => rects[left]!.y - rects[right]!.y);
  let open: number[] = [];
  for (const index of byStart) {
    const rect = rects[index]!;
    open = open.filter((other) => rects[other]!.y + rects[other]!.height - rect.y > TOUCH);
    for (const other of open) {
      if (overlaps(rects[other]!, rect)) {
        group[find(index)] = find(other);
        any = true;
      }
    }
    open.push(index);
  }
  if (!any) return [...rects];
  const members = new Map<number, Rect[]>();
  rects.forEach((rect, index) => members.set(find(index), [...(members.get(find(index)) ?? []), rect]));
  const result: Rect[] = [];
  for (const parts of members.values()) {
    if (parts.length === 1) {
      result.push(parts[0]!);
      continue;
    }
    const order = parts.some((part) => part.order === undefined) ? undefined
      : Math.min(...parts.map((part) => part.order!));
    const edges = [...new Set(parts.flatMap((part) => [part.y, part.y + part.height]))].sort((a, b) => a - b);
    let open: { x: number; y: number; width: number; height: number }[] = [];
    for (let band = 0; band < edges.length - 1; band += 1) {
      const top = edges[band]!;
      const bottom = edges[band + 1]!;
      if (bottom - top <= TOUCH) continue;
      const stretches: { x: number; width: number }[] = [];
      for (const part of parts.filter((candidate) => candidate.y <= top + TOUCH && candidate.y + candidate.height >= bottom - TOUCH)
        .sort((left, right) => left.x - right.x)) {
        const last = stretches.at(-1);
        if (last && part.x <= last.x + last.width + TOUCH) last.width = Math.max(last.width, part.x + part.width - last.x);
        else stretches.push({ x: part.x, width: part.width });
      }
      // A band with the stretches of the band above continues its rects.
      const continues = open.length === stretches.length && open.every((rect, index) => (
        Math.abs(rect.y + rect.height - top) <= TOUCH && Math.abs(rect.x - stretches[index]!.x) <= TOUCH
        && Math.abs(rect.width - stretches[index]!.width) <= TOUCH));
      if (continues) for (const rect of open) rect.height = bottom - rect.y;
      else {
        open = stretches.map((stretch) => ({ ...stretch, y: top, height: bottom - top }));
        for (const rect of open) result.push(rect as Rect);
      }
    }
    if (order !== undefined) for (const rect of result) if (rect.order === undefined) (rect as { order?: number }).order = order;
  }
  return result;
}

/** `rects` of one frame joined where they touch or overlap along the run
 * over the same interval across it (invariant 5), ordered across, then
 * along; area that rects of different extents share is covered once. */
export function mergeRunRects(rects: readonly RunRect[]): RunRect[] {
  return joinRects(rects);
}

/** The rects that mark the characters `[start, end)` of `run` (invariant
 * 4): their glyph rects, cut to the run's clips and joined (invariant 5).
 * The one rect-building function of every overlay that marks text. */
export function runOverlayRects(run: Run, start = 0, end = run.text.length): RunRect[] {
  return mergeRunRects(runGlyphRects(run, start, end).flatMap((rect) => clipRunRect(run, rect)));
}

/** How many rects the tab lookup of overlayBoxes has examined so far: a
 * measure of its work for tests, which compare it across input sizes. */
let rectsExamined = 0;
export function overlayRectsExamined(): number {
  return rectsExamined;
}

/** Rects of one frame by the edges they start and end at along the run,
 * each edge's rects ordered across it. */
class EdgeIndex<Rect extends RunRect & { readonly order: number }> {
  private readonly starts = new Map<number, Rect[]>();
  private readonly ends = new Map<number, Rect[]>();

  /** Edges within TOUCH share a key or lie in neighbouring ones. */
  private static key(edge: number): number {
    return Math.round(edge / (4 * TOUCH));
  }

  add(rect: Rect): void {
    for (const [edges, edge] of [[this.starts, rect.x], [this.ends, rect.x + rect.width]] as const) {
      const key = EdgeIndex.key(edge);
      const list = edges.get(key) ?? [];
      if (list.length === 0) edges.set(key, list);
      let low = 0;
      let high = list.length;
      while (low < high) {
        const middle = (low + high) >> 1;
        if (list[middle]!.y < rect.y) low = middle + 1;
        else high = middle;
      }
      list.splice(low, 0, rect);
    }
  }

  /** The rect that ends at `left` or starts at `right` and lies across the
   * run within `[top, bottom]`: the one nearest to `order`, an earlier one
   * first. */
  touching(left: number, right: number, top: number, bottom: number, order: number): Rect | undefined {
    let best: Rect | undefined;
    for (const [edges, edge, at] of [[this.ends, left, (rect: Rect) => rect.x + rect.width],
      [this.starts, right, (rect: Rect) => rect.x]] as const) {
      for (const key of [EdgeIndex.key(edge) - 1, EdgeIndex.key(edge), EdgeIndex.key(edge) + 1]) {
        const list = edges.get(key);
        if (!list) continue;
        let low = 0;
        let high = list.length;
        while (low < high) {
          const middle = (low + high) >> 1;
          if (list[middle]!.y < top - TOUCH) low = middle + 1;
          else high = middle;
        }
        for (let index = low; index < list.length && list[index]!.y <= bottom + TOUCH; index += 1) {
          const rect = list[index]!;
          rectsExamined += 1;
          if (rect.y + rect.height > bottom + TOUCH || Math.abs(at(rect) - edge) > TOUCH) continue;
          if (!best || Math.abs(rect.order - order) < Math.abs(best.order - order)
            || (Math.abs(rect.order - order) === Math.abs(best.order - order) && rect.order < best.order)) best = rect;
        }
      }
    }
    return best;
  }
}

/** The boxes that mark a range of source runs on the page (invariant 4):
 * the range rects of each run of `runs` and the boxes of its tabs and inline
 * objects, cut to their clips and joined between runs of one orientation
 * under the same clips where they touch (invariant 5). A box stands where
 * the first of its runs does in `runs`; boxes of one run are ordered across,
 * then along it. */
export function overlayBoxes(runs: readonly Run[]): OverlayBox[] {
  type Ordered = RunRect & { readonly order: number };
  const groups = new Map<string, {
    frame: Run;
    rects: Ordered[];
    /** The rects of `rects` that mark text. */
    text: Ordered[];
    tabs: { run: Run; order: number }[];
  }>();
  const frameOf = (run: Run) => {
    const [a, b, c, d] = orientation(run.transform);
    return { along: a * run.x + b * run.y, across: c * run.x + d * run.y };
  };
  runs.forEach((run, order) => {
    const key = `${run.transform ?? ''}|${run.clips ? JSON.stringify(run.clips) : ''}`;
    let group = groups.get(key);
    if (!group) {
      // The group's frame: its runs' orientation about the page origin.
      group = { frame: { ...run, x: 0, y: 0 }, rects: [], text: [], tabs: [] };
      groups.set(key, group);
    }
    if (run.inlineContent === 'tab') {
      group.tabs.push({ run, order });
      return;
    }
    const { along, across } = frameOf(run);
    const rects = run.inlineContent === 'object'
      ? [{ x: 0, y: 0, width: run.w, height: run.h }] : runRangeRects(run);
    for (const rect of rects.flatMap((candidate) => clipRunRect(run, candidate))) {
      const placed = { x: along + rect.x, y: across + rect.y, width: rect.width, height: rect.height, order };
      group.rects.push(placed);
      if (run.inlineContent === undefined) group.text.push(placed);
    }
  });
  const boxes: { box: OverlayBox; order: number; index: number }[] = [];
  for (const group of groups.values()) {
    // A tab is marked over the extent across the run of the marked text it
    // touches, within its own line box; a tab that touches none (one beside
    // an inline object only, whose box is not a line of text) keeps its line
    // box. Rects are found by the edge they end or start at along the
    // run and by their place across it, so each tab looks at the rects of
    // its own line only. Tabs are taken in order, then the unplaced ones in
    // reverse, so a row of tabs takes its extent from the text at either
    // end in two passes.
    const edges = new EdgeIndex<Ordered>();
    for (const rect of group.text) edges.add(rect);
    const place = (tab: { run: Run; order: number }, required: boolean): boolean => {
      const { along, across } = frameOf(tab.run);
      const neighbour = edges.touching(along, along + tab.run.w, across, across + tab.run.h, tab.order);
      if (!neighbour && !required) return false;
      const rect = neighbour
        ? { x: 0, y: neighbour.y - across, width: tab.run.w, height: neighbour.height }
        : { x: 0, y: 0, width: tab.run.w, height: tab.run.h };
      for (const part of clipRunRect(tab.run, rect)) {
        const placed = { x: along + part.x, y: across + part.y, width: part.width, height: part.height, order: tab.order };
        group.rects.push(placed);
        edges.add(placed);
      }
      return true;
    };
    const waiting = group.tabs.filter((tab) => !place(tab, false));
    for (let index = waiting.length - 1; index >= 0; index -= 1) place(waiting[index]!, true);
    for (const rect of joinRects(group.rects)) {
      boxes.push({ box: placeRunRect(group.frame, rect), order: rect.order, index: boxes.length });
    }
  }
  return boxes.sort((left, right) => left.order - right.order || left.index - right.index).map(({ box }) => box);
}
