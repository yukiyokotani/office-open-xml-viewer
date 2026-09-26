/**
 * IX2 docx find-in-document controller.
 *
 * Owns the search state for a {@link DocxViewer}: the per-page run lists (the
 * same `onTextRun` stream the selection / hyperlink overlay consumes), the
 * matches for the current query, and the active-match cursor `findNext` /
 * `findPrev` walk. All the string/index math is core (`buildTextIndex`,
 * `findMatches`, `nextActive`/`prevActive`); this class just threads a docx run
 * list per page through it and maps each hit to a `{ page }` location.
 *
 * The viewer supplies `collectPageRuns(page)` — render page `page` (to an
 * offscreen canvas, without disturbing the visible one) and return its
 * `DocxTextRunInfo[]`. The controller calls it once per page on `find()` and
 * caches the result until `invalidate()` (a reload / width change). The active
 * page's cached runs are what the viewer's highlight overlay is drawn from, so
 * the geometry the highlight boxes use is exactly the geometry the page was
 * drawn with.
 */
import {
  buildTextIndex,
  findMatches,
  normalizeFindQuery,
  nextActive,
  prevActive,
  type FindMatch,
  type FindMatchesOptions,
  type FindQuery,
  type TextMatch,
  withFindColor,
} from '@silurus/ooxml-core';
import type { DocxTextRunInfo } from './renderer';

/** Where a docx match lives: its 0-based page index. */
export interface DocxMatchLocation {
  page: number;
}

/** One resolved docx match: a page-located {@link FindMatch} plus the run-slices
 *  (in that page's run list) it covers, kept internally so the highlight overlay
 *  can turn them into pixel boxes. */
interface DocxResolvedMatch {
  /** 0-based page the match is on. */
  page: number;
  /** The matched text. */
  text: string;
  /** Run-slices within `pageRuns[page]`, from core `findMatches`. */
  slices: TextMatch['slices'];
  /** The colour of the term that found it, when that term set one. */
  color?: string;
}

export class DocxFindController {
  /** page index → that page's runs (from a render), or undefined = not scanned. */
  private _pageRuns = new Map<number, DocxTextRunInfo[]>();
  /** All matches for the current query, in document order (page asc, then within-page). */
  private _matches: DocxResolvedMatch[] = [];
  /** Active-match index into {@link _matches}, or -1 for none. */
  private _active = -1;
  /** Invalidates in-flight searches on clear, reload, destroy, or a newer find. */
  private _generation = 0;
  /** Advances whenever visible rendering publishes newer run geometry. */
  private _runsRevision = 0;

  constructor(
    private readonly _pageCount: () => number,
    private readonly _collectPageRuns: (page: number) => Promise<DocxTextRunInfo[]>,
  ) {}

  /** Drop all cached runs + matches (call on reload / render-width change). */
  invalidate(): void {
    this._generation++;
    this._runsRevision++;
    this._pageRuns.clear();
    this._matches = [];
    this._active = -1;
  }

  /** The runs for a page, if it has been scanned (used by the highlight overlay
   *  for the currently displayed page). */
  pageRuns(page: number): DocxTextRunInfo[] | undefined {
    return this._pageRuns.get(page);
  }

  /** Cache a page's runs captured from the visible render, so find() reuses
   *  them instead of re-rendering that page offscreen. */
  setPageRuns(page: number, runs: DocxTextRunInfo[]): void {
    this._runsRevision++;
    this._pageRuns.set(page, runs);
  }

  /** The resolved match at the given global index (for the highlight overlay). */
  private _matchAt(index: number): DocxResolvedMatch | undefined {
    return this._matches[index];
  }

  /** All match slices that fall on a given page, tagged with whether each is the
   *  active match — the exact input the highlight overlay needs. */
  pageHighlights(page: number): { slices: TextMatch['slices']; active: boolean; color?: string }[] {
    const out: { slices: TextMatch['slices']; active: boolean; color?: string }[] = [];
    for (let i = 0; i < this._matches.length; i++) {
      const m = this._matches[i];
      if (m.page === page) out.push(withFindColor({ slices: m.slices, active: i === this._active }, m.color));
    }
    return out;
  }

  /** The active match's page, or null when there is no active match. */
  activePage(): number | null {
    const m = this._matchAt(this._active);
    return m ? m.page : null;
  }

  /** The public match list for the current query. */
  matches(): FindMatch<DocxMatchLocation>[] {
    return this._matches.map((m, i) => withFindColor({
      matchIndex: i,
      text: m.text,
      location: { page: m.page },
    }, m.color));
  }

  /** Run a fresh query across every page, resetting the cursor. Returns the
   *  public match list. An empty query clears matches. */
  async find(query: FindQuery, opts: FindMatchesOptions = {}): Promise<FindMatch<DocxMatchLocation>[]> {
    const generation = ++this._generation;
    if (normalizeFindQuery(query).length === 0) {
      this._runsRevision++;
      this._pageRuns.clear();
      this._matches = [];
      this._active = -1;
      return [];
    }

    const runsRevision = this._runsRevision;
    const pageRuns = new Map(this._pageRuns);
    const pages = this._pageCount();
    for (let page = 0; page < pages; page++) {
      let runs = pageRuns.get(page);
      if (!runs) {
        try {
          runs = await this._collectPageRuns(page);
        } catch (error) {
          if (generation !== this._generation) return [];
          throw error;
        }
        if (generation !== this._generation) return [];
        pageRuns.set(page, runs);
      }
    }
    if (generation !== this._generation) return [];

    // A visible zoom-settle render can publish newer geometry while this scan
    // awaits an offscreen page. Preserve those per-page updates instead of
    // replacing them with the snapshot captured when find() began.
    const committedRuns = runsRevision === this._runsRevision
      ? pageRuns
      : new Map([...pageRuns, ...this._pageRuns]);

    // Build slices from the exact run lists being committed. Besides keeping
    // highlight coordinates current, this keeps slice indices coherent if a
    // fresh render changes run segmentation.
    const matches: DocxResolvedMatch[] = [];
    for (let page = 0; page < pages; page++) {
      const runs = committedRuns.get(page) ?? [];
      const index = buildTextIndex(runs);
      for (const tm of findMatches(index, query, opts)) {
        // The matched text as it appears in the document: read it back from the
        // run slices so it carries the document's original case (not the folded
        // query). Slices are in run order, so concatenating their run text
        // reconstructs the match.
        const text = tm.slices
          .map((s) => runs[s.runIndex].text.slice(s.start, s.end))
          .join('');
        matches.push(withFindColor({ page, text, slices: tm.slices }, tm.color));
      }
    }
    this._runsRevision++;
    this._pageRuns = committedRuns;
    this._matches = matches;
    this._active = -1;
    return this.matches();
  }

  /** Advance the active match (wrap-around) and return its public form, or null
   *  when there are no matches. */
  next(): FindMatch<DocxMatchLocation> | null {
    this._active = nextActive(this._active, this._matches.length);
    return this._activePublic();
  }

  /** Step the active match back (wrap-around). */
  prev(): FindMatch<DocxMatchLocation> | null {
    this._active = prevActive(this._active, this._matches.length);
    return this._activePublic();
  }

  private _activePublic(): FindMatch<DocxMatchLocation> | null {
    const m = this._matchAt(this._active);
    if (!m) return null;
    return withFindColor({ matchIndex: this._active, text: m.text, location: { page: m.page } }, m.color);
  }

}
