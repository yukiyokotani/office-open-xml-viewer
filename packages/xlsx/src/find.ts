/**
 * IX2 xlsx find-in-workbook controller.
 *
 * The spreadsheet twin of the docx / pptx find controllers, differing in the
 * unit searched: a *cell* is the atomic run (one display string), not a
 * glyph-run, so a match always lands wholly inside one cell and the highlight is
 * the cell rectangle rather than a sub-cell glyph range. Search runs over each
 * cell's *rendered* display text (via the viewer-supplied `cellText`, i.e. the
 * same number-format / date / rich-text flattening the grid draws) so a query
 * matches what the user sees.
 *
 * The core string math is still shared: each cell's text is a one-run
 * {@link buildTextIndex}, and {@link findMatches} counts every occurrence within
 * the cell (so "the the" registers two matches, like a browser find bar). The
 * active-match cursor is core `nextActive`/`prevActive`. A match's location is
 * `{ sheet, sheetName, ref, row, col }`.
 *
 * The controller holds two independent searches: the interactive find and the
 * text highlights (`highlight` / `highlightStep`), which keep
 * separate matches and cursors.
 */
import {
  buildTextIndex,
  findMatches,
  foldPreservingLength,
  loadTermSearch,
  nextActive,
  prevActive,
  adoptTextHighlightGroups,
  type FindMatch,
  type FindMatchesOptions,
  type TextHighlightGroup,
  type TextHighlightMatch,
} from '@silurus/ooxml-core';
import { formatA1 } from './a1.js';

/** One cell's searchable content (1-based row/col + its rendered text). */
export interface FindCell {
  row: number;
  col: number;
  text: string;
  /** Other texts the cell may display. A General number shortens to its
   *  column width at paint (ECMA-376 §18.8.30), so find also matches each
   *  shorter form; `text` is searched first, and the cell's matches come from
   *  the first of these texts that has any, so no cell is counted twice. */
  alternates?: readonly string[];
}

/** Where an xlsx match lives: the sheet, its name, and the cell (A1 + row/col). */
export interface XlsxMatchLocation {
  /** 0-based sheet index. */
  sheet: number;
  /** The sheet's display name. */
  sheetName: string;
  /** A1 cell reference, e.g. `"B7"`. */
  ref: string;
  /** 1-based row. */
  row: number;
  /** 1-based column. */
  col: number;
}

interface XlsxResolvedMatch {
  sheet: number;
  sheetName: string;
  row: number;
  col: number;
  text: string;
  /** Text highlights only: the group's background, its paint layer (the
   *  group's position in the set), and the public `{ group, term }`. */
  color?: string;
  layer?: number;
  source?: { group: string; term: string };
}

/** What a scan reports per occurrence in one cell. */
type CellHit = Pick<XlsxResolvedMatch, 'text' | 'color' | 'layer' | 'source'>;
/** Resolves one cell's display texts (the full one first) to its occurrences. */
type CellMatcher = (candidates: readonly string[]) => CellHit[];

/** The matches, cursor and cancellation state of one kind of search (the
 *  interactive find, or the text highlights). */
class MatchList {
  matches: XlsxResolvedMatch[] = [];
  active = -1;
  generation = 0;

  reset(): void {
    this.generation++;
    this.matches = [];
    this.active = -1;
  }

  /** The public form of match `index`, or `null` when out of range. */
  at(index: number): FindMatch<XlsxMatchLocation> | null {
    const m = this.matches[index];
    return m
      ? {
        matchIndex: index,
        text: m.text,
        location: { sheet: m.sheet, sheetName: m.sheetName, ref: formatA1(m.row, m.col), row: m.row, col: m.col },
        ...m.source,
      }
      : null;
  }

  all(): FindMatch<XlsxMatchLocation>[] {
    return this.matches.map((_, index) => this.at(index)!);
  }

  /** Move the cursor with `nextActive` / `prevActive`. */
  step(move: (active: number, count: number) => number): FindMatch<XlsxMatchLocation> | null {
    this.active = move(this.active, this.matches.length);
    return this.at(this.active);
  }
}

/** A cell box of the highlight overlay. */
export interface XlsxCellHighlight {
  row: number;
  col: number;
  active: boolean;
  /** Set for a text highlight: the group's CSS background. Absent for a find
   *  match, which uses the find colours. */
  color?: string;
  /** Text highlights only: how many other groups already have a box on this
   *  cell, so each further group is drawn inset and stays visible. */
  depth?: number;
}

export class XlsxFindController {
  private readonly _find = new MatchList();
  private readonly _marks = new MatchList();

  constructor(
    private readonly _sheetCount: () => number,
    private readonly _sheetName: (sheet: number) => string,
    /** Return every non-empty cell of a sheet, with its rendered display text. */
    private readonly _collectSheetCells: (sheet: number) => Promise<FindCell[]>,
  ) {}

  /** Drop both searches (call on reload). */
  invalidate(): void {
    this._find.reset();
    this._marks.reset();
  }

  /** Drop the interactive find; text highlights stay. */
  clearFind(): void { this._find.reset(); }

  /** Drop the text highlights; the interactive find stays. */
  clearHighlights(): void { this._marks.reset(); }

  /** The overlay boxes of a sheet in paint order: one per text-highlight group
   *  per cell (active when any of its occurrences is the active highlight),
   *  group by group in the order of the set, then one per find match, so the
   *  find is painted on top. */
  sheetHighlights(sheet: number): XlsxCellHighlight[] {
    const out: XlsxCellHighlight[] = [];
    const boxes = new Map<string, XlsxCellHighlight>();
    const depth = new Map<string, number>();
    // In the order of the set, so a later group is drawn inside (above) an
    // earlier one on a shared cell. Array.prototype.sort is stable: one group
    // keeps document order.
    const marks = this._marks.matches.map((m, i) => ({ m, i }))
      .filter(({ m }) => m.sheet === sheet)
      .sort((a, b) => a.m.layer! - b.m.layer!);
    for (const { m, i } of marks) {
      const cell = `${m.row},${m.col}`;
      const key = `${cell},${m.source!.group}`;
      let box = boxes.get(key);
      if (!box) {
        const level = depth.get(cell) ?? 0;
        depth.set(cell, level + 1);
        box = { row: m.row, col: m.col, active: false, color: m.color, depth: level };
        boxes.set(key, box);
        out.push(box);
      }
      if (i === this._marks.active) box.active = true;
    }
    this._find.matches.forEach((m, i) => {
      if (m.sheet === sheet) out.push({ row: m.row, col: m.col, active: i === this._find.active });
    });
    return out;
  }

  /** The active match's location, or null. */
  activeLocation(): XlsxMatchLocation | null {
    return this._find.at(this._find.active)?.location ?? null;
  }

  /** The public match list for the current query. */
  matches(): FindMatch<XlsxMatchLocation>[] { return this._find.all(); }

  /** Run a fresh query across every sheet, resetting the cursor. Matches are in
   *  document order: sheet ascending, then a sheet's cells in the order
   *  `collectSheetCells` returns them (row-major from the parser). */
  async find(query: string, opts: FindMatchesOptions = {}): Promise<FindMatch<XlsxMatchLocation>[]> {
    if (query.length === 0) {
      this._find.reset();
      return [];
    }
    // Whole-word matching is part of the on-demand term search; a plain find
    // starts without loading anything.
    return await this._scan(this._find, opts.wholeWord
      ? async () => (await loadTermSearch()).wholeWordCellMatcher(query, opts, foldPreservingLength)
      : () => (candidates) => {
        for (const candidate of candidates) {
          const index = buildTextIndex([{ text: candidate }]);
          const hits = findMatches(index, query, opts);
          // Each occurrence within the cell is one match (browser find-count
          // semantics); the located text is the matched substring in original case.
          if (hits.length > 0) {
            return hits.map((hit) => ({ text: candidate.slice(hit.slices[0].start, hit.slices[0].end) }));
          }
        }
        return [];
      }) ?? [];
  }

  next(): FindMatch<XlsxMatchLocation> | null { return this._find.step(nextActive); }

  prev(): FindMatch<XlsxMatchLocation> | null { return this._find.step(prevActive); }

  /** Replace the text-highlight set and search every sheet for it. A set with
   *  a repeated group id rejects and leaves the highlights and any request in
   *  flight as they are. */
  async highlight(groups: readonly TextHighlightGroup[]): Promise<TextHighlightMatch<XlsxMatchLocation>[]> {
    // Detached from the caller's objects: the scan reads the set after awaits.
    const set = adoptTextHighlightGroups(groups);
    if (set.length === 0) {
      this._marks.reset();
      return [];
    }
    const matches = await this._scan(this._marks, async () =>
      (await loadTermSearch()).textHighlightCellMatcher(set, foldPreservingLength));
    return (matches ?? []) as TextHighlightMatch<XlsxMatchLocation>[];
  }

  /** Move the text-highlight cursor to the next match, or with `back` to the
   *  previous one (wrap-around), and return it. */
  highlightStep(back?: boolean): TextHighlightMatch<XlsxMatchLocation> | null {
    return this._marks.step(back ? prevActive : nextActive) as TextHighlightMatch<XlsxMatchLocation> | null;
  }

  /** Read every sheet's cells, resolve the matcher `prepare` yields against
   *  each cell's display texts, and commit the result to `list`. Returns the
   *  committed public matches, taken at commit time so that a later request
   *  cannot change what this one reports. Returns `null` when a newer request
   *  of the same kind (or a clear) superseded this one; `list` is then left
   *  untouched, and a failure of the superseded request is not reported. */
  private async _scan(
    list: MatchList,
    prepare: () => CellMatcher | Promise<CellMatcher>,
  ): Promise<FindMatch<XlsxMatchLocation>[] | null> {
    const generation = ++list.generation;
    try {
      let matcher = prepare();
      if (matcher instanceof Promise) {
        matcher = await matcher;
        if (generation !== list.generation) return null;
      }
      const sheets = this._sheetCount();
      const matches: XlsxResolvedMatch[] = [];
      for (let sheet = 0; sheet < sheets; sheet++) {
        const cells = await this._collectSheetCells(sheet);
        if (generation !== list.generation) return null;
        const sheetName = this._sheetName(sheet);
        for (const cell of cells) {
          for (const hit of matcher([cell.text, ...(cell.alternates ?? [])])) {
            matches.push({ ...hit, sheet, sheetName, row: cell.row, col: cell.col });
          }
        }
      }
      list.matches = matches;
      list.active = -1;
      return list.all();
    } catch (error) {
      // A superseded request reports nothing, not even its failure (a closed
      // workbook's sheet, a chunk that did not load).
      if (generation !== list.generation) return null;
      throw error;
    }
  }
}
