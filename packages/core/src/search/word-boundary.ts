/**
 * Word boundaries for whole-word matching (`wholeWord` in find and text
 * highlights).
 *
 * An offset in the searched text is a word boundary when a whole-word match may
 * start or end there. The rule, in order:
 *
 * 1. The start and the end of the text are boundaries, and so is every offset
 *    the format adapter marked as a hard break (a paragraph end, a line change,
 *    a different text container): the joined text has no separator character
 *    there, but the reader sees one.
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
  // Segment boundaries of each word-character run consulted so far, keyed by
  // the run's start offset.
  const segmented = new Map<number, Set<number>>();
  const isBreak = (offset: number): boolean => breaks?.has(offset) === true;

  const segmentBoundary = (offset: number, dictionary: WordSegmenter): boolean => {
    let start = offset;
    while (start > 0 && !isBreak(start)) {
      const previous = characterBefore(text, start);
      if (!WORD_CHARACTER.test(previous)) break;
      start -= previous.length;
    }
    let boundaries = segmented.get(start);
    if (!boundaries) {
      let end = offset;
      while (end < text.length && !isBreak(end)) {
        const next = characterAt(text, end);
        if (!WORD_CHARACTER.test(next)) break;
        end += next.length;
      }
      boundaries = new Set<number>();
      for (const segment of dictionary.segment(text.slice(start, end))) {
        boundaries.add(start + segment.index);
      }
      segmented.set(start, boundaries);
    }
    return boundaries.has(offset);
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
