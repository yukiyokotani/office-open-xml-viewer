/**
 * IX2 pptx find-in-presentation controller.
 *
 * The pptx twin of the docx `DocxFindController`: owns per-slide run lists (the
 * `onTextRun` stream), the matches for the current query, and the active-match
 * cursor. All string/index math is core (`buildTextIndex`, `findMatches`,
 * `nextActive`/`prevActive`); this maps each hit to a `{ slide }` location.
 *
 * The viewer supplies `collectSlideRuns(slide)` — render that slide (to an
 * offscreen canvas) and return its `PptxTextRunInfo[]`. The controller caches
 * per slide until `invalidate()`. The displayed slide's runs are fed in from the
 * visible render so highlight geometry matches exactly what was drawn.
 */
import {
  buildTextIndex,
  findMatches,
  nextActive,
  prevActive,
  type FindMatch,
  type FindMatchesOptions,
  type TextMatch,
} from '@silurus/ooxml-core';
import type { PptxTextRunInfo } from './renderer';

/** Where a pptx match lives: its 0-based slide index. */
export interface PptxMatchLocation {
  slide: number;
}

interface PptxResolvedMatch {
  slide: number;
  text: string;
  slices: TextMatch['slices'];
}

function sameSearchContainer(left: PptxTextRunInfo, right: PptxTextRunInfo): boolean {
  const leftCell = left.tableCell;
  const rightCell = right.tableCell;
  if (!leftCell && !rightCell) return true;
  if (!leftCell || !rightCell) return false;
  return left.elementIndex === right.elementIndex &&
    left.origin === right.origin &&
    left.shapeId === right.shapeId &&
    leftCell.row === rightCell.row &&
    leftCell.column === rightCell.column;
}

/**
 * The shared core index joins adjacent drawing runs, which is correct within a
 * text body. Table cells are separate text containers, so reject only matches
 * whose run slices cross such a boundary. Keeping one slide-wide index avoids
 * per-cell indices and preserves the existing linear scan and run offsets.
 */
function matchStaysWithinSearchContainer(
  runs: PptxTextRunInfo[],
  slices: TextMatch['slices'],
): boolean {
  for (let index = 1; index < slices.length; index++) {
    const left = runs[slices[index - 1].runIndex];
    const right = runs[slices[index].runIndex];
    if (!left || !right || !sameSearchContainer(left, right)) return false;
  }
  return true;
}

export class PptxFindController {
  private _slideRuns = new Map<number, PptxTextRunInfo[]>();
  private _matches: PptxResolvedMatch[] = [];
  private _active = -1;
  /** Invalidates in-flight searches on clear, reload, destroy, or a newer find. */
  private _generation = 0;
  /** Advances whenever visible rendering publishes newer run geometry. */
  private _runsRevision = 0;

  constructor(
    private readonly _slideCount: () => number,
    private readonly _collectSlideRuns: (slide: number) => Promise<PptxTextRunInfo[]>,
  ) {}

  /** Drop all cached runs + matches (call on reload). */
  invalidate(): void {
    this._generation++;
    this._runsRevision++;
    this._slideRuns.clear();
    this._matches = [];
    this._active = -1;
  }

  /** The runs for a slide, if scanned (used by the highlight overlay for the
   *  displayed slide). */
  slideRuns(slide: number): PptxTextRunInfo[] | undefined {
    return this._slideRuns.get(slide);
  }

  /** Cache a slide's runs captured from the visible render. */
  setSlideRuns(slide: number, runs: PptxTextRunInfo[]): void {
    this._runsRevision++;
    this._slideRuns.set(slide, runs);
  }

  /** All match slices on a slide, tagged active — the highlight overlay input. */
  slideHighlights(slide: number): { slices: TextMatch['slices']; active: boolean }[] {
    const out: { slices: TextMatch['slices']; active: boolean }[] = [];
    for (let i = 0; i < this._matches.length; i++) {
      const m = this._matches[i];
      if (m.slide === slide) out.push({ slices: m.slices, active: i === this._active });
    }
    return out;
  }

  /** The active match's slide, or null. */
  activeSlide(): number | null {
    const m = this._matches[this._active];
    return m ? m.slide : null;
  }

  /** The public match list for the current query. */
  matches(): FindMatch<PptxMatchLocation>[] {
    return this._matches.map((m, i) => ({
      matchIndex: i,
      text: m.text,
      location: { slide: m.slide },
    }));
  }

  /** Run a fresh query across every slide, resetting the cursor. */
  async find(query: string, opts: FindMatchesOptions = {}): Promise<FindMatch<PptxMatchLocation>[]> {
    const generation = ++this._generation;
    if (query.length === 0) {
      this._runsRevision++;
      this._slideRuns.clear();
      this._matches = [];
      this._active = -1;
      return [];
    }

    const runsRevision = this._runsRevision;
    const slideRuns = new Map(this._slideRuns);
    const slides = this._slideCount();
    for (let slide = 0; slide < slides; slide++) {
      let runs = slideRuns.get(slide);
      if (!runs) {
        try {
          runs = await this._collectSlideRuns(slide);
        } catch (error) {
          if (generation !== this._generation) return [];
          throw error;
        }
        if (generation !== this._generation) return [];
        slideRuns.set(slide, runs);
      }
    }
    if (generation !== this._generation) return [];

    // A visible zoom-settle render can publish newer geometry while this scan
    // awaits an offscreen slide. Preserve those per-slide updates instead of
    // replacing them with the snapshot captured when find() began.
    const committedRuns = runsRevision === this._runsRevision
      ? slideRuns
      : new Map([...slideRuns, ...this._slideRuns]);

    // Resolve slices against the exact run lists being committed so slice
    // indices and highlight geometry remain coherent after a fresh render.
    const matches: PptxResolvedMatch[] = [];
    for (let slide = 0; slide < slides; slide++) {
      const runs = committedRuns.get(slide) ?? [];
      const index = buildTextIndex(runs);
      for (const tm of findMatches(index, query, opts)) {
        if (!matchStaysWithinSearchContainer(runs, tm.slices)) continue;
        const text = tm.slices
          .map((slice) => runs[slice.runIndex].text.slice(slice.start, slice.end))
          .join('');
        matches.push({ slide, text, slices: tm.slices });
      }
    }
    this._runsRevision++;
    this._slideRuns = committedRuns;
    this._matches = matches;
    this._active = -1;
    return this.matches();
  }

  next(): FindMatch<PptxMatchLocation> | null {
    this._active = nextActive(this._active, this._matches.length);
    return this._activePublic();
  }

  prev(): FindMatch<PptxMatchLocation> | null {
    this._active = prevActive(this._active, this._matches.length);
    return this._activePublic();
  }

  private _activePublic(): FindMatch<PptxMatchLocation> | null {
    const m = this._matches[this._active];
    if (!m) return null;
    return { matchIndex: this._active, text: m.text, location: { slide: m.slide } };
  }

}
