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
}

/** Options for {@link findMatches}. */
export interface FindMatchesOptions {
  /**
   * Match case exactly. Default `false` (case-insensitive, like a browser's
   * find-in-page). IX2 default — an integrator can pass `true`.
   */
  caseSensitive?: boolean;
  /**
   * Match whole words only. Default `false`. A match is kept when it starts and
   * ends at a word boundary: next to a character that is not a letter, combining
   * mark, number or `_`, at a paragraph or line change, or at the ends of the
   * text. So `cat` is not found inside `concatenate`, while `C++` is found
   * before a comma. In scripts written without spaces between words (Han,
   * Hiragana, Katakana, Thai, Lao, Khmer, Myanmar) the boundaries come from
   * `Intl.Segmenter` word segmentation and therefore depend on the browser's
   * dictionaries; where `Intl.Segmenter` is unavailable, text in those scripts
   * is matched as if `wholeWord` were `false`.
   */
  wholeWord?: boolean;
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
export function foldPreservingLength(s: string): string {
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
export function sliceRange(index: TextIndex, matchStart: number, matchEnd: number): MatchRunSlice[] {
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
 * `opts.wholeWord` is not applied here: whole-word matching is part of the
 * on-demand term search (`term-search.ts`), which the find controllers use for
 * such a query instead of this function.
 *
 * Pure: no DOM, no geometry. The viewer turns `slices` into pixel rectangles
 * using each run's font.
 */
export function findMatches(
  index: TextIndex,
  query: string,
  opts: FindMatchesOptions = {},
): TextMatch[] {
  if (query.length === 0) return [];
  const caseSensitive = opts.caseSensitive ?? false;
  const haystack = caseSensitive ? index.text : index.folded;
  // The needle must be folded with the SAME length-preserving fold as the
  // haystack: a plain toLowerCase() would expand e.g. "İ" to "i" + U+0307 and
  // never match the haystack, where "İ" was deliberately kept unfolded.
  const needle = caseSensitive ? query : foldPreservingLength(query);

  const matches: TextMatch[] = [];
  let from = 0;
  let matchIndex = 0;
  for (;;) {
    const at = haystack.indexOf(needle, from);
    if (at === -1) break;
    matches.push({ matchIndex, slices: sliceRange(index, at, at + needle.length) });
    matchIndex++;
    // Advance past this match so occurrences never overlap.
    from = at + needle.length;
  }
  return matches;
}
