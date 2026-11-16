import { buildTextIndex, findMatches, foldPreservingLength, sliceRange } from '../search/text-index';
import { nextActive, prevActive } from '../search/find-cursor';
import type { FindMatchesOptions, SearchRun, TextIndex, TextMatch } from '../search/text-index';
import type { FindMatch } from '../search/find-match';
import {
  loadTermSearch,
  adoptTextHighlightGroups,
  type TextHighlightGroup,
  type TextHighlightMatch,
} from '../search/text-highlight';

interface ResolvedMatch {
  unit: number;
  text: string;
  slices: TextMatch['slices'];
  /** Text highlights only: the group's background, its paint layer (the
   *  group's position in the set), and the public `{ group, term }`. */
  color?: string;
  layer?: number;
  source?: { group: string; term: string };
}

/** What a scan's matcher reports per occurrence. */
type ScanHit = Pick<ResolvedMatch, 'slices' | 'color' | 'layer' | 'source'>;
type Matcher<Run> = (index: TextIndex, runs: Run[]) => ScanHit[];

/** One overlay box list entry: the run-slices of a match, whether it is the
 *  active one of its kind, and for a text highlight the group background and
 *  paint layer (both absent for a find match, which uses the find colours). */
export interface UnitHighlight {
  slices: TextMatch['slices'];
  active: boolean;
  color?: string;
  layer?: number;
}

/** The matches, cursor and cancellation state of one kind of search (the
 *  interactive find, or the text highlights). */
class MatchList<Location> {
  matches: ResolvedMatch[] = [];
  /** Match indexes in paint order: by layer, then document order, so a later
   *  group is painted above an earlier one wherever its matches start. */
  private _paint: number[] = [];
  active = -1;
  generation = 0;
  /** A non-empty search was requested and not cleared since. */
  armed = false;

  constructor(private readonly _location: (unit: number) => Location) {}

  reset(): void {
    this.generation++;
    this.commit([]);
    this.armed = false;
  }

  /** Install a scan's result and reset the cursor. */
  commit(matches: ResolvedMatch[]): void {
    this.matches = matches;
    // Array.prototype.sort is stable: equal layers keep document order.
    this._paint = matches.map((_, index) => index)
      .sort((a, b) => (matches[a].layer ?? 0) - (matches[b].layer ?? 0));
    this.active = -1;
  }

  /** The public form of match `index`, or `null` when out of range. */
  at(index: number): FindMatch<Location> | null {
    const match = this.matches[index];
    return match
      ? { matchIndex: index, text: match.text, location: this._location(match.unit), ...match.source }
      : null;
  }

  all(): FindMatch<Location>[] {
    return this.matches.map((_, index) => this.at(index)!);
  }

  /** Move the cursor with `nextActive` / `prevActive`. */
  step(move: (active: number, count: number) => number): FindMatch<Location> | null {
    this.active = move(this.active, this.matches.length);
    return this.at(this.active);
  }

  /** Append this list's overlay boxes of a unit to `out`, in paint order. */
  boxes(unit: number, out: UnitHighlight[]): void {
    for (const index of this._paint) {
      const match = this.matches[index];
      if (match.unit !== unit) continue;
      out.push({
        slices: match.slices, active: index === this.active, color: match.color, layer: match.layer,
      });
    }
  }
}

/** Search state for a sequence of rendered units. The adapter names the unit in
 * public locations and may reject matches spanning separate text containers.
 *
 * Two independent searches share the cached unit runs: the interactive find
 * (`find` / `next` / `prev` / `clearFind`) and the text highlights
 * (`highlight` / `highlightNext` / `highlightPrev` / `clearHighlights`). A
 * search that starts after the other has committed reuses its runs; two that
 * collect at the same time each collect the units they still miss. */
export class UnitFindController<Run extends SearchRun, Location> {
  private _runs = new Map<number, Run[]>();
  private _runsRevision = 0;
  private readonly _find: MatchList<Location>;
  private readonly _marks: MatchList<Location>;

  constructor(
    private readonly _count: () => number,
    private readonly _collectRuns: (unit: number) => Promise<Run[]>,
    location: (unit: number) => Location,
    private readonly _acceptMatch: (runs: Run[], slices: TextMatch['slices']) => boolean = () => true,
    /** Units excluded from the search (e.g. hidden slides the viewer never
     *  draws) are neither collected nor matched. Evaluated per `find`. */
    private readonly _includeUnit: (unit: number) => boolean = () => true,
    /** Loads the format's test for a layout break between two adjacent runs.
     *  Only whole-word matching needs it, so it is fetched on demand. */
    private readonly _runBreak?: () => Promise<(previous: Run, next: Run) => boolean>,
  ) {
    this._find = new MatchList(location);
    this._marks = new MatchList(location);
  }

  /** Drop both searches and the cached runs (new document, relayout). */
  invalidate(): void {
    this._find.reset();
    this._marks.reset();
    this._trimRuns();
  }

  /** Drop the interactive find; text highlights stay. */
  clearFind(): void {
    this._find.reset();
    this._trimRuns();
  }

  /** Drop the text highlights; the interactive find stays. */
  clearHighlights(): void {
    this._marks.reset();
    this._trimRuns();
  }

  /** Whether text highlights were requested and not cleared since. */
  get highlighting(): boolean { return this._marks.armed; }

  unitRuns(unit: number): Run[] | undefined { return this._runs.get(unit); }

  setUnitRuns(unit: number, runs: Run[]): void {
    this._runsRevision++;
    this._runs.set(unit, runs);
  }

  /** The overlay boxes of a unit in paint order: text highlights first (group
   *  by group in the order of the set, each in document order), then the find
   *  matches, so the find is painted on top. */
  unitHighlights(unit: number): UnitHighlight[] {
    const out: UnitHighlight[] = [];
    this._marks.boxes(unit, out);
    this._find.boxes(unit, out);
    return out;
  }

  activeUnit(): number | null { return this._find.matches[this._find.active]?.unit ?? null; }

  matches(): FindMatch<Location>[] { return this._find.all(); }

  async find(query: string, opts: FindMatchesOptions = {}): Promise<FindMatch<Location>[]> {
    // Empty search is a full clear, including cached geometry, in both formats.
    if (query.length === 0) {
      this.clearFind();
      return [];
    }
    // Whole-word matching is part of the on-demand term search; a plain find
    // starts without loading anything.
    return await this._scan(this._find, opts.wholeWord
      ? async () => (await loadTermSearch()).wholeWordMatcher(
        query, opts, foldPreservingLength, sliceRange, this._runBreak, this._acceptMatch)
      : () => (index) => findMatches(index, query, opts)) ?? [];
  }

  next(): FindMatch<Location> | null { return this._find.step(nextActive); }

  prev(): FindMatch<Location> | null { return this._find.step(prevActive); }

  /** Replace the text-highlight set and search every unit for it. A set with
   *  a repeated group id rejects and leaves the highlights and any request in
   *  flight as they are. */
  async highlight(groups: readonly TextHighlightGroup[]): Promise<TextHighlightMatch<Location>[]> {
    // Detached from the caller's objects: the scan reads the set after awaits.
    const set = adoptTextHighlightGroups(groups);
    if (set.length === 0) {
      this.clearHighlights();
      return [];
    }
    const matches = await this._scan(this._marks, async () =>
      (await loadTermSearch()).textHighlightMatcher(
        set, foldPreservingLength, sliceRange, this._runBreak, this._acceptMatch));
    return (matches ?? []) as TextHighlightMatch<Location>[];
  }

  highlightNext(): TextHighlightMatch<Location> | null {
    return this._marks.step(nextActive) as TextHighlightMatch<Location> | null;
  }

  highlightPrev(): TextHighlightMatch<Location> | null {
    return this._marks.step(prevActive) as TextHighlightMatch<Location> | null;
  }

  /** Cached geometry is refreshed only while a search is armed, so it must not
   *  outlive the last one. */
  private _trimRuns(): void {
    if (this._find.armed || this._marks.armed) return;
    this._runsRevision++;
    this._runs.clear();
  }

  /** Collect every included unit's runs, resolve the matcher `prepare` yields
   *  against them, and commit the result to `list`. Returns the committed
   *  public matches, taken at commit time so that a later request cannot
   *  change what this one reports. Returns `null` when a newer request of the
   *  same kind (or a clear) superseded this one; `list` is then left
   *  untouched, and a failure of the superseded request is not reported. */
  private async _scan(
    list: MatchList<Location>,
    prepare: () => Matcher<Run> | Promise<Matcher<Run>>,
  ): Promise<FindMatch<Location>[] | null> {
    const generation = ++list.generation;
    try {
      let matcher = prepare();
      if (matcher instanceof Promise) {
        matcher = await matcher;
        if (generation !== list.generation) return null;
      }
      // From here on mounted units keep their run geometry current.
      list.armed = true;

      const runsRevision = this._runsRevision;
      const unitRuns = new Map(this._runs);
      const units = this._count();
      const included: boolean[] = [];
      for (let unit = 0; unit < units; unit++) included.push(this._includeUnit(unit));
      for (let unit = 0; unit < units; unit++) {
        if (!included[unit]) continue;
        let runs = unitRuns.get(unit);
        if (!runs) {
          runs = await this._collectRuns(unit);
          if (generation !== list.generation) return null;
          unitRuns.set(unit, runs);
        }
      }
      if (generation !== list.generation) return null;

      // Visible renders can publish fresher geometry while an offscreen scan is
      // awaiting a unit. Resolve slices against the committed run lists.
      const committedRuns = runsRevision === this._runsRevision
        ? unitRuns
        : new Map([...unitRuns, ...this._runs]);
      const matches: ResolvedMatch[] = [];
      for (let unit = 0; unit < units; unit++) {
        if (!included[unit]) continue;
        const runs = committedRuns.get(unit) ?? [];
        for (const match of matcher(buildTextIndex(runs), runs)) {
          if (!this._acceptMatch(runs, match.slices)) continue;
          const text = match.slices.map((slice) =>
            runs[slice.runIndex].text.slice(slice.start, slice.end)).join('');
          matches.push({ ...match, unit, text });
        }
      }
      this._runsRevision++;
      this._runs = committedRuns;
      list.commit(matches);
      return list.all();
    } catch (error) {
      // A superseded request reports nothing, not even its failure (a closed
      // document's run collection, a chunk that did not load).
      if (generation !== list.generation) return null;
      throw error;
    }
  }
}
