/**
 * Word boundaries for whole-word matching (`wholeWord` in find and text
 * highlights).
 *
 * An offset in the searched text is a word boundary when a whole-word match may
 * start or end there. The rule, in order:
 *
 * 1. The start and the end of the text are boundaries (the text is one page,
 *    slide or cell: units are searched one by one), and so is every offset
 *    the format adapter marked as a layout break (a paragraph end, a manual
 *    line break, a tab, a different text container): the joined text has no
 *    separator character there, but the reader sees one. An automatic line
 *    wrap is not such a break.
 * 2. An offset with a non-word character on either side is a boundary. Word
 *    characters are Unicode letters, combining marks, numbers and `_` (the
 *    boundary a Unicode-aware regex `\b` draws), so `cat` is not found inside
 *    `concatenate`, while `C++` is found before a comma.
 * 3. Between two word characters there is no boundary, unless one of them
 *    belongs to a script written without spaces between words (Han, Hiragana,
 *    Katakana, Thai, Lao, Khmer, Myanmar). There the boundary comes from
 *    `Intl.Segmenter` with word granularity (UAX #29 with the engine's
 *    dictionaries), applied to the surrounding run of word characters.
 * 4. Where `Intl.Segmenter` is unavailable, every offset of rule 3 that does not
 *    precede a combining mark is a boundary: in those scripts `wholeWord` then
 *    matches like a plain substring search, rather than finding nothing.
 *
 * Rules 1 and 2 never consult the segmenter, so text in scripts that separate
 * words with spaces gives the same result in every engine.
 */

const WORD_CHARACTER = /[\p{L}\p{M}\p{N}_]/u;
const COMBINING_MARK = /\p{M}/u;
const UNSPACED_SCRIPT =
  /[\p{scx=Han}\p{scx=Hira}\p{scx=Kana}\p{scx=Thai}\p{scx=Lao}\p{scx=Khmr}\p{scx=Mymr}]/u;

/** The minimal `Intl.Segmenter` surface the boundary test uses. */
export interface WordSegmenter {
  segment(text: string): Iterable<{ index: number }>;
}

let sharedSegmenter: WordSegmenter | null | undefined;

function defaultSegmenter(): WordSegmenter | null {
  if (sharedSegmenter === undefined) {
    sharedSegmenter = null;
    try {
      const Segmenter = (Intl as { Segmenter?: new (locale?: string, options?: { granularity: 'word' }) => WordSegmenter }).Segmenter;
      if (Segmenter) sharedSegmenter = new Segmenter(undefined, { granularity: 'word' });
    } catch {
      // Stay on the documented fallback.
    }
  }
  return sharedSegmenter;
}

/** The code point ending just before UTF-16 offset `at` (`''` at the start). */
function characterBefore(text: string, at: number): string {
  if (at <= 0) return '';
  const low = text.charCodeAt(at - 1);
  if (low >= 0xdc00 && low <= 0xdfff && at >= 2) {
    const high = text.charCodeAt(at - 2);
    if (high >= 0xd800 && high <= 0xdbff) return text.slice(at - 2, at);
  }
  return text[at - 1];
}

/** The code point starting at UTF-16 offset `at` (`''` at the end). */
function characterAt(text: string, at: number): string {
  const codePoint = text.codePointAt(at);
  return codePoint === undefined ? '' : String.fromCodePoint(codePoint);
}

/**
 * Build the word-boundary test for `text` (see the module comment for the
 * rule). `breaks` holds the offsets that are boundaries by layout. `segmenter`
 * is injectable for tests; `null` selects the fallback of rule 4.
 */
export function wordBoundaryTest(
  text: string,
  breaks?: ReadonlySet<number>,
  segmenter: WordSegmenter | null = defaultSegmenter(),
): (offset: number) => boolean {
  // The word-character runs consulted so far, sorted by start and disjoint,
  // each with its segment boundaries. An offset is looked up by binary search,
  // so a run is walked and segmented once however many offsets fall into it.
  const segmented: { start: number; end: number; boundaries: Set<number> }[] = [];
  const isBreak = (offset: number): boolean => breaks?.has(offset) === true;

  const segmentBoundary = (offset: number, dictionary: WordSegmenter): boolean => {
    let low = 0;
    let high = segmented.length;
    while (low < high) {
      const middle = (low + high) >> 1;
      if (segmented[middle].end < offset) low = middle + 1;
      else high = middle;
    }
    let run = segmented[low];
    if (!run || run.start > offset) {
      let start = offset;
      while (start > 0 && !isBreak(start)) {
        const previous = characterBefore(text, start);
        if (!WORD_CHARACTER.test(previous)) break;
        start -= previous.length;
      }
      let end = offset;
      while (end < text.length && !isBreak(end)) {
        const next = characterAt(text, end);
        if (!WORD_CHARACTER.test(next)) break;
        end += next.length;
      }
      const boundaries = new Set<number>();
      for (const segment of dictionary.segment(text.slice(start, end))) {
        boundaries.add(start + segment.index);
      }
      run = { start, end, boundaries };
      segmented.splice(low, 0, run);
    }
    return run.boundaries.has(offset);
  };

  return (offset) => {
    if (offset <= 0 || offset >= text.length || isBreak(offset)) return true;
    const before = characterBefore(text, offset);
    const after = characterAt(text, offset);
    // Never inside a surrogate pair.
    if (after >= '\udc00' && after <= '\udfff' && before >= '\ud800' && before <= '\udbff') return false;
    if (!WORD_CHARACTER.test(before) || !WORD_CHARACTER.test(after)) return true;
    if (!UNSPACED_SCRIPT.test(before) && !UNSPACED_SCRIPT.test(after)) return false;
    return segmenter ? segmentBoundary(offset, segmenter) : !COMBINING_MARK.test(after);
  };
}
