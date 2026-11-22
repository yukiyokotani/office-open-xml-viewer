import { buildTextIndex, findMatches } from '../search/text-index';
import { nextActive, prevActive } from '../search/find-cursor';
import type { FindMatchesOptions, SearchRun, TextMatch } from '../search/text-index';
import type { FindMatch } from '../search/find-match';

interface ResolvedMatch {
  unit: number;
  text: string;
  slices: TextMatch['slices'];
}

/** Search state for a sequence of rendered units. The adapter names the unit in
 * public locations and may reject matches spanning separate text containers. */
export class UnitFindController<Run extends SearchRun, Location> {
  private _runs = new Map<number, Run[]>();
  private _matches: ResolvedMatch[] = [];
  private _active = -1;
  private _generation = 0;
  private _runsRevision = 0;

  constructor(
    private readonly _count: () => number,
    private readonly _collectRuns: (unit: number) => Promise<Run[]>,
    private readonly _location: (unit: number) => Location,
    private readonly _acceptMatch: (runs: Run[], slices: TextMatch['slices']) => boolean = () => true,
    /** Units excluded from the search (e.g. hidden slides the viewer never
     *  draws) are neither collected nor matched. Evaluated per `find`. */
    private readonly _includeUnit: (unit: number) => boolean = () => true,
  ) {}

  invalidate(): void {
    this._generation++;
    this._runsRevision++;
    this._runs.clear();
    this._matches = [];
    this._active = -1;
  }

  unitRuns(unit: number): Run[] | undefined { return this._runs.get(unit); }

  setUnitRuns(unit: number, runs: Run[]): void {
    this._runsRevision++;
    this._runs.set(unit, runs);
  }

  unitHighlights(unit: number): { slices: TextMatch['slices']; active: boolean }[] {
    const out: { slices: TextMatch['slices']; active: boolean }[] = [];
    for (let i = 0; i < this._matches.length; i++) {
      const match = this._matches[i];
      if (match.unit === unit) out.push({ slices: match.slices, active: i === this._active });
    }
    return out;
  }

  activeUnit(): number | null { return this._matches[this._active]?.unit ?? null; }

  matches(): FindMatch<Location>[] {
    return this._matches.map((match, matchIndex) => ({
      matchIndex, text: match.text, location: this._location(match.unit),
    }));
  }

  async find(query: string, opts: FindMatchesOptions = {}): Promise<FindMatch<Location>[]> {
    const generation = ++this._generation;
    // Empty search is a full clear, including cached geometry, in both formats.
    if (query.length === 0) {
      this._runsRevision++;
      this._runs.clear();
      this._matches = [];
      this._active = -1;
      return [];
    }

    const runsRevision = this._runsRevision;
    const unitRuns = new Map(this._runs);
    const units = this._count();
    const included: boolean[] = [];
    for (let unit = 0; unit < units; unit++) included.push(this._includeUnit(unit));
    for (let unit = 0; unit < units; unit++) {
      if (!included[unit]) continue;
      let runs = unitRuns.get(unit);
      if (!runs) {
        try {
          runs = await this._collectRuns(unit);
        } catch (error) {
          if (generation !== this._generation) return [];
          throw error;
        }
        if (generation !== this._generation) return [];
        unitRuns.set(unit, runs);
      }
    }
    if (generation !== this._generation) return [];

    // Visible renders can publish fresher geometry while an offscreen scan is
    // awaiting a unit. Resolve slices against the committed run lists.
    const committedRuns = runsRevision === this._runsRevision
      ? unitRuns
      : new Map([...unitRuns, ...this._runs]);
    const matches: ResolvedMatch[] = [];
    for (let unit = 0; unit < units; unit++) {
      if (!included[unit]) continue;
      const runs = committedRuns.get(unit) ?? [];
      for (const match of findMatches(buildTextIndex(runs), query, opts)) {
        if (!this._acceptMatch(runs, match.slices)) continue;
        const text = match.slices.map((slice) =>
          runs[slice.runIndex].text.slice(slice.start, slice.end)).join('');
        matches.push({ unit, text, slices: match.slices });
      }
    }
    this._runsRevision++;
    this._runs = committedRuns;
    this._matches = matches;
    this._active = -1;
    return this.matches();
  }

  next(): FindMatch<Location> | null {
    this._active = nextActive(this._active, this._matches.length);
    return this._activePublic();
  }

  prev(): FindMatch<Location> | null {
    this._active = prevActive(this._active, this._matches.length);
    return this._activePublic();
  }

  private _activePublic(): FindMatch<Location> | null {
    const match = this._matches[this._active];
    return match ? { matchIndex: this._active, text: match.text, location: this._location(match.unit) } : null;
  }
}
