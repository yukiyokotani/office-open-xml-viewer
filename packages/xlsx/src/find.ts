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
  withFindColor,
} from '@silurus/ooxml-core';
import { formatA1 } from './a1.js';

/** One cell's searchable content (1-based row/col + its rendered text). */
export interface FindCell {
  row: number;
  col: number;
  text: string;
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
  /** The colour of the term that found it, when that term set one. */
  color?: string;
}

export class XlsxFindController {
  private _matches: XlsxResolvedMatch[] = [];
  private _active = -1;
  private _generation = 0;

  constructor(
    private readonly _sheetCount: () => number,
    private readonly _sheetName: (sheet: number) => string,
    /** Return every non-empty cell of a sheet, with its rendered display text. */
    private readonly _collectSheetCells: (sheet: number) => Promise<FindCell[]>,
  ) {}

  /** Drop matches + cursor (call on reload). */
  invalidate(): void {
    this._generation++;
    this._matches = [];
    this._active = -1;
  }

  /** Matched cells on a sheet, each tagged active — the highlight overlay input. */
  sheetHighlights(sheet: number): { row: number; col: number; active: boolean; color?: string }[] {
    const out: { row: number; col: number; active: boolean; color?: string }[] = [];
    for (let i = 0; i < this._matches.length; i++) {
      const m = this._matches[i];
      if (m.sheet === sheet) out.push(withFindColor({ row: m.row, col: m.col, active: i === this._active }, m.color));
    }
    return out;
  }

  /** The active match's location, or null. */
  activeLocation(): XlsxMatchLocation | null {
    return this._locationAt(this._active);
  }

  private _locationAt(index: number): XlsxMatchLocation | null {
    const m = this._matches[index];
    if (!m) return null;
    return {
      sheet: m.sheet,
      sheetName: m.sheetName,
      ref: formatA1(m.row, m.col),
      row: m.row,
      col: m.col,
    };
  }

  /** The public match list for the current query. */
  matches(): FindMatch<XlsxMatchLocation>[] {
    return this._matches.map((m, i) => {
      const loc = this._locationAt(i) as XlsxMatchLocation;
      return withFindColor({ matchIndex: i, text: m.text, location: loc }, m.color);
    });
  }

  /** Run a fresh query across every sheet, resetting the cursor. Matches are in
   *  document order: sheet ascending, then a sheet's cells in the order
   *  `collectSheetCells` returns them (row-major from the parser). */
  async find(query: FindQuery, opts: FindMatchesOptions = {}): Promise<FindMatch<XlsxMatchLocation>[]> {
    const generation = ++this._generation;
    if (normalizeFindQuery(query).length === 0) {
      this._matches = [];
      this._active = -1;
      return [];
    }

    const sheets = this._sheetCount();
    const matches: XlsxResolvedMatch[] = [];
    for (let sheet = 0; sheet < sheets; sheet++) {
      let cells: FindCell[];
      try {
        cells = await this._collectSheetCells(sheet);
      } catch (error) {
        if (generation !== this._generation) return [];
        throw error;
      }
      if (generation !== this._generation) return [];
      const sheetName = this._sheetName(sheet);
      for (const cell of cells) {
        const index = buildTextIndex([{ text: cell.text }]);
        const hits = findMatches(index, query, opts);
        // Each occurrence within the cell is one match (browser find-count
        // semantics); the located text is the matched substring in original case.
        for (const tm of hits) {
          const slice = tm.slices[0];
          const text = cell.text.slice(slice.start, slice.end);
          matches.push(withFindColor({ sheet, sheetName, row: cell.row, col: cell.col, text }, tm.color));
        }
      }
    }
    if (generation !== this._generation) return [];
    this._matches = matches;
    this._active = -1;
    return this.matches();
  }

  next(): FindMatch<XlsxMatchLocation> | null {
    this._active = nextActive(this._active, this._matches.length);
    return this._activePublic();
  }

  prev(): FindMatch<XlsxMatchLocation> | null {
    this._active = prevActive(this._active, this._matches.length);
    return this._activePublic();
  }

  private _activePublic(): FindMatch<XlsxMatchLocation> | null {
    const loc = this._locationAt(this._active);
    if (!loc) return null;
    const m = this._matches[this._active];
    return withFindColor({ matchIndex: this._active, text: m.text, location: loc }, m.color);
  }
}
