/**
 * Format-agnostic in-document text search (IX2 — findText).
 *
 * All three formats (docx / pptx / xlsx) already emit their rendered text as an
 * ordered list of runs through `onTextRun` (the same stream IX1 turned into the
 * hyperlink-aware selection overlay). A run is the atomic drawn text segment: a
 * docx line-piece, a pptx shape run, an xlsx cell. Search is therefore the same
 * problem in every format — concatenate the run texts, match a query against the
 * joined string, and map each match back to the run(s) it fell on so the viewer
 * can draw a highlight box over the right glyphs. That reverse mapping is the
 * only subtle part (a query can straddle two runs — "Hello" split as "Hel"+"lo"),
 * and it is pure string/index math with no DOM or geometry, so it lives here in
 * `core` and is shared, not re-implemented per package.
 *
 * What stays out of core: turning a matched run-slice into a pixel rectangle
 * needs the run's font + `measureText`, which is renderer/DOM territory and
 * differs per format (docx flat spans, pptx shape-relative + rotation, xlsx cell
 * rects). Each package owns that final step; core owns the index + the match →
 * run-slice resolution.
 *
 * Normalization (IX2 default, integrator may veto): case-insensitive by default
 * (`caseSensitive: false`), matching a browser's Ctrl+F. Full-/half-width (zenkaku
 * ↔ hankaku) folding and diacritic-insensitive matching are intentionally NOT
 * done in this minimal pass — they are locale-sensitive and easy to get subtly
 * wrong; a later opt-in can add them without changing this contract. Case
 * folding is length-preserving (see {@link foldPreservingLength}): the few code
 * points whose lowercase changes UTF-16 length (e.g. U+0130 İ) match by
 * identity only, so match offsets always index the original text.
 */

/**
 * The minimal shape the index needs from a rendered run: just its text. Each
 * package passes its own richer run info (which also carries geometry) — this is
 * structurally satisfied by `DocxTextRunInfo`, `PptxTextRunInfo`,
 * `XlsxTextRunInfo`, etc., so a caller hands its `runs[]` straight in.
 */
export interface SearchRun {
  text: string;
}

/**
 * A precomputed search index over an ordered run list. `text` is the runs joined
 * in order (no separator — runs are adjacent drawn segments); `runStart[i]` is
 * the character offset in `text` where run `i` begins, so a match position in
 * `text` can be resolved back to a (run, offset-within-run) pair by binary
 * search. `folded` is the case-folded form used for matching when the search is
 * case-insensitive. INVARIANT: `folded.length === text.length` — guaranteed by
 * {@link foldPreservingLength} — so a match offset found in `folded` indexes
 * `text` directly. (A whole-string `toLowerCase()` does NOT hold this: e.g.
 * U+0130 İ lowercases to "i" + U+0307 combining dot, 1 → 2 UTF-16 code units,
 * which would shift every later match right by the delta and could even push a
 * match range past `text.length`, hanging the slice walk.)
 */
export interface TextIndex {
  /** The concatenated run texts, in run order. */
  readonly text: string;
  /** Lower-cased twin of {@link text}, for case-insensitive matching. */
  readonly folded: string;
  /** `runStart[i]` = offset in `text` where run `i` starts. Length = run count. */
  readonly runStart: readonly number[];
  /** The run count (so callers need not keep the original array alongside). */
  readonly runCount: number;
}

/**
 * The slice of one run a match covers: the run's index in the original `runs[]`
 * and the `[start, end)` character range within that run's own `text`. A match
 * that straddles N runs yields N of these (the first sliced from its start
 * offset to the run end, the last from 0 to its end offset, any middle run
 * whole). The viewer measures each slice against that run's font to get a pixel
 * rectangle.
 */
export interface MatchRunSlice {
  /** Index into the original `runs[]` handed to {@link buildTextIndex}. */
  runIndex: number;
  /** Start offset within `runs[runIndex].text` (inclusive). */
  start: number;
  /** End offset within `runs[runIndex].text` (exclusive). */
  end: number;
}

/**
 * One query match: its ordinal position among all matches (`matchIndex`, 0-based,
 * in document order) and the run-slices it covers (`slices`, one per run the
 * match touches, in run order). `slices` is never empty for a non-empty query.
 */
export interface TextMatch {
  matchIndex: number;
  slices: MatchRunSlice[];
  /** The {@link FindTerm.color} of the term that produced this match, when it
   *  set one. */
  color?: string;
}

/**
 * One search term that carries its own highlight colour. Its matches are drawn
 * in `color` instead of the viewer's `findHighlightColors.match`, so one query can
 * show several kinds of hit apart (say, search terms and flagged phrases). The
 * active match still uses `findHighlightColors.active`.
 */
export interface FindTerm {
  /** The text to find. */
  text: string;
  /** CSS background for this term's matches. Omit to use the match colour. */
  color?: string;
}

/**
 * What to search for: one term, or several searched together. Several terms
 * highlight at once — the shape a search UI produces when it hands over the
 * words (or phrases) a query matched, rather than a single find-box string. A
 * term may be a {@link FindTerm} to give its matches their own colour.
 */
export type FindQuery = string | readonly (string | FindTerm)[];

/** Options for {@link findMatches}. */
export interface FindMatchesOptions {
  /**
   * Match case exactly. Default `false` (case-insensitive, like a browser's
   * find-in-page). IX2 default — an integrator can pass `true`.
   */
  caseSensitive?: boolean;
  /**
   * Only match whole words. Default `false`. A match is rejected when a word
   * character sits on either side of it AND the match's own edge character on
   * that side is also a word character — the same boundary a regex `\b`
   * draws, so `"cat"` does not match inside `"concatenate"` while `"C++"` still
   * matches in `"C++,"`. Word characters are Unicode letters, combining marks,
   * numbers and `_`. Scripts written without spaces between words (CJK, Thai)
   * have no such boundary inside a run of text, so this option finds nothing
   * mid-sentence there; it is meant for space-delimited scripts.
   */
  wholeWord?: boolean;
}

/**
 * Reduce a {@link FindQuery} to the distinct non-empty terms to search for, as
 * {@link FindTerm}s. Only a truly empty term is dropped (no trimming — the same
 * contract as the single-string query), so a query with no terms left is the
 * "clear the find" case. When the same text appears twice, the first one (and
 * its colour) wins. Viewers use this to decide emptiness, because `['']` has a
 * non-zero `length` but searches for nothing.
 */
export function normalizeFindQuery(query: FindQuery): FindTerm[] {
  const entries = typeof query === 'string' ? [query] : query;
  const byText = new Map<string, FindTerm>();
  for (const entry of entries) {
    const term = typeof entry === 'string' ? { text: entry } : entry;
    if (term.text.length === 0 || byText.has(term.text)) continue;
    byText.set(term.text, withFindColor({ text: term.text }, term.color));
  }
  return [...byText.values()];
}

/**
 * `value` with `color` set when there is one, and no `color` key at all when
 * there is not — so a match found by a plain term looks exactly as it did
 * before terms could carry colours.
 */
export function withFindColor<T extends object>(value: T, color: string | undefined): T & { color?: string } {
  return color === undefined ? value : { ...value, color };
}

const WORD_CHARACTER = /[\p{L}\p{M}\p{N}_]/u;

function isWordCodePoint(codePoint: number | undefined): boolean {
  return codePoint !== undefined && WORD_CHARACTER.test(String.fromCodePoint(codePoint));
}

/** The code point ending just before UTF-16 offset `at`, stepping back over a
 *  surrogate pair as one character. */
function codePointBefore(s: string, at: number): number | undefined {
  if (at <= 0) return undefined;
  const low = s.charCodeAt(at - 1);
  if (low >= 0xdc00 && low <= 0xdfff && at >= 2) {
    const high = s.charCodeAt(at - 2);
    if (high >= 0xd800 && high <= 0xdbff) return s.codePointAt(at - 2);
  }
  return low;
}

/** Whether `[start, end)` of `text` stands as a whole word (see
 *  {@link FindMatchesOptions.wholeWord}). */
function isWholeWord(text: string, start: number, end: number): boolean {
  const joinsBefore = isWordCodePoint(codePointBefore(text, start)) && isWordCodePoint(text.codePointAt(start));
  const joinsAfter = isWordCodePoint(text.codePointAt(end)) && isWordCodePoint(codePointBefore(text, end));
  return !joinsBefore && !joinsAfter;
}

/** A match's `[start, end)` in the joined text, plus the index of the term that
 *  found it. */
type TermRange = [number, number, number];

/**
 * Keep a non-overlapping subset of `ranges`, preferring the longer range where
 * two collide (so a phrase term wins over one of its own words searched
 * alongside it), then the earlier one. Returned in document order. Cost is
 * linear in the total length of the candidate ranges.
 */
function dropOverlaps(ranges: TermRange[], textLength: number): TermRange[] {
  const covered = new Uint8Array(textLength);
  const byPreference = [...ranges].sort((a, b) => (b[1] - b[0]) - (a[1] - a[0]) || a[0] - b[0]);
  const kept: TermRange[] = [];
  for (const range of byPreference) {
    let free = true;
    for (let i = range[0]; i < range[1]; i++) {
      if (covered[i]) {
        free = false;
        break;
      }
    }
    if (!free) continue;
    covered.fill(1, range[0], range[1]);
    kept.push(range);
  }
  return kept.sort((a, b) => a[0] - b[0]);
}

/**
 * Case-fold `s` for case-insensitive matching WITHOUT changing its UTF-16
 * length, so every offset in the folded string indexes the original directly.
 *
 * `String.prototype.toLowerCase` is not length-preserving: a handful of code
 * points expand (U+0130 İ → "i" + U+0307, 1 → 2 code units; ligature-style
 * upper-case mappings likewise). A whole-string lowercase therefore de-syncs
 * folded offsets from text offsets and corrupts every match after the first
 * such character. Instead we fold per code point and KEEP a code point
 * unfolded whenever its lowercase form has a different code-unit length.
 *
 * Trade-off (documented contract): those rare code points only match by
 * identity, not case-insensitively (a query "i̇" will not find "İ"). That is
 * the correct side of the trade — a slightly narrower match set for a few
 * exotic characters versus wrong-glyph highlights (or a hang) for every match
 * that follows one.
 *
 * Fast path: when the whole-string lowercase already has the same length
 * (true for effectively all real documents), use it as-is; the per-code-point
 * walk only runs for strings containing a length-changing fold.
 */
function foldPreservingLength(s: string): string {
  const whole = s.toLowerCase();
  if (whole.length === s.length) return whole;
  let out = '';
  for (const ch of s) {
    // `ch` is one code point (1 or 2 code units via the string iterator).
    const lower = ch.toLowerCase();
    out += lower.length === ch.length ? lower : ch;
  }
  return out;
}

/**
 * Build a reusable {@link TextIndex} from an ordered run list. O(total chars).
 * The index is independent of the query, so a viewer builds it once per rendered
 * page/slide/sheet and reuses it across `findText` / `findNext` calls until the
 * rendered content changes.
 */
export function buildTextIndex(runs: readonly SearchRun[]): TextIndex {
  const runStart: number[] = new Array(runs.length);
  let acc = 0;
  let joined = '';
  for (let i = 0; i < runs.length; i++) {
    runStart[i] = acc;
    joined += runs[i].text;
    acc += runs[i].text.length;
  }
  return {
    text: joined,
    // Length-preserving fold: folded offsets must index `text` directly (see
    // the TextIndex invariant).
    folded: foldPreservingLength(joined),
    runStart,
    runCount: runs.length,
  };
}

/**
 * Resolve a character offset in {@link TextIndex.text} to the index of the run
 * that contains it, by binary search over `runStart`. Returns the run whose
 * `[runStart[i], runStart[i+1])` half-open range contains `pos`. `pos` is assumed
 * in `[0, text.length)`.
 */
function runAtOffset(index: TextIndex, pos: number): number {
  const { runStart } = index;
  let lo = 0;
  let hi = runStart.length - 1;
  // Invariant: runStart[lo] <= pos. Find the greatest i with runStart[i] <= pos.
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (runStart[mid] <= pos) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

/**
 * Split a `[matchStart, matchEnd)` character range in the joined text into the
 * per-run slices it covers. Empty runs (zero-length text) inside the range are
 * skipped — they contribute no glyphs to highlight. The range is assumed
 * non-empty and within `[0, text.length]`.
 */
function sliceRange(index: TextIndex, matchStart: number, matchEnd: number): MatchRunSlice[] {
  const { runStart, runCount, text } = index;
  const slices: MatchRunSlice[] = [];
  let run = runAtOffset(index, matchStart);
  let cursor = matchStart;
  // `run < runCount` is a termination guard, not a semantic bound: under the
  // TextIndex invariant (folded.length === text.length) every match range lies
  // within [0, text.length] and the walk always ends by cursor reaching
  // matchEnd inside the last run. Without the guard, a range past text.length
  // (as a length-CHANGING fold used to produce) would loop forever — cursor
  // pins at text.length while run increments unboundedly.
  while (cursor < matchEnd && run < runCount) {
    // End of the current run in joined-text coordinates.
    const runEndOffset = run + 1 < runCount ? runStart[run + 1] : text.length;
    const sliceEnd = Math.min(matchEnd, runEndOffset);
    const localStart = cursor - runStart[run];
    const localEnd = sliceEnd - runStart[run];
    if (localEnd > localStart) {
      slices.push({ runIndex: run, start: localStart, end: localEnd });
    }
    cursor = sliceEnd;
    run++;
  }
  return slices;
}

/**
 * Find every occurrence of `query` in the indexed text and return each as a
 * {@link TextMatch} carrying the run-slices it covers. Matches are
 * non-overlapping and returned in document order; scanning resumes just past each
 * match (so `"aa"` in `"aaaa"` yields two matches at 0 and 2, matching a
 * browser's find-in-page). An empty (or whitespace-trimmed-to-empty is NOT
 * applied — only truly empty) query returns `[]`.
 *
 * `query` may be several terms. Each is searched on its own and the results are
 * merged into one document-ordered list; where matches of different terms
 * overlap, the longer match is kept (see {@link dropOverlaps}), so every
 * highlighted character belongs to exactly one match. A match found by a
 * {@link FindTerm} with a `color` carries that colour.
 *
 * Pure: no DOM, no geometry. The viewer turns `slices` into pixel rectangles
 * using each run's font.
 */
export function findMatches(
  index: TextIndex,
  query: FindQuery,
  opts: FindMatchesOptions = {},
): TextMatch[] {
  const terms = normalizeFindQuery(query);
  if (terms.length === 0) return [];
  const caseSensitive = opts.caseSensitive ?? false;
  const wholeWord = opts.wholeWord ?? false;
  const haystack = caseSensitive ? index.text : index.folded;

  const ranges: TermRange[] = [];
  for (let termIndex = 0; termIndex < terms.length; termIndex++) {
    const term = terms[termIndex].text;
    // The needle must be folded with the SAME length-preserving fold as the
    // haystack: a plain toLowerCase() would expand e.g. "İ" to "i" + U+0307 and
    // never match the haystack, where "İ" was deliberately kept unfolded.
    const needle = caseSensitive ? term : foldPreservingLength(term);
    let from = 0;
    for (;;) {
      const at = haystack.indexOf(needle, from);
      if (at === -1) break;
      const end = at + needle.length;
      if (wholeWord && !isWholeWord(index.text, at, end)) {
        // A rejected occurrence must not hide one that starts inside it.
        from = at + 1;
        continue;
      }
      ranges.push([at, end, termIndex]);
      // Advance past this match so occurrences never overlap.
      from = end;
    }
  }

  // One term's matches are already disjoint and in order; only a merge of
  // several terms can collide.
  const kept = terms.length === 1 ? ranges : dropOverlaps(ranges, haystack.length);
  return kept.map(([start, end, termIndex], matchIndex) =>
    withFindColor({ matchIndex, slices: sliceRange(index, start, end) }, terms[termIndex].color));
}
