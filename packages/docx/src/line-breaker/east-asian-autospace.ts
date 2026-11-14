import { compatOf } from '../layout/compat-profile.js';
import { graphemeClusterOffsets } from '@silurus/ooxml-core';
import { wordCompressedSpaceLineFitApplies, wordEastAsianAutoSpacePt } from '../layout/line-compatibility.js';
import { RESET_SLICED_TEXT_MEASUREMENT, slicedTextMetadata } from './advance.js';
import type { LayoutSeg, LayoutTextSeg, LineLayoutEnvironment } from './model.js';

/** Side of an ECMA-376 §17.3.1.2-3 automatic-spacing boundary. */
export type EastAsianAutoSpaceClass = 'eastAsian' | 'latinLetter' | 'digit';

// §17.3.1.2 determines the regions "by the Unicode character values of the
// text content". The classes follow WORD_EAST_ASIAN_AUTO_SPACE: letters of
// the Han, Hiragana, Katakana (including half-width) and Hangul scripts
// against Latin letters (DE) and ASCII digits (DN). East Asian punctuation,
// ASCII symbols/punctuation and the full-width Latin forms are neutral; so is
// anything else (Greek, Cyrillic, U+3000, …).
const EAST_ASIAN_LETTER = /^[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]$/u;
const LETTER = /^\p{L}$/u;
const GRAPHEME_EXTENDERS = /^[\p{Grapheme_Extend}\u200D]+$/u;
const LATIN_LETTER = /^\p{Script=Latin}$/u;

export function eastAsianAutoSpaceClass(character: string): EastAsianAutoSpaceClass | undefined {
  if (character >= '0' && character <= '9' && character.length === 1) return 'digit';
  if (!LETTER.test(character)) return undefined;
  if (EAST_ASIAN_LETTER.test(character)) return 'eastAsian';
  const cp = character.codePointAt(0)!;
  // U+FF21-FF3A / U+FF41-FF5A carry Script=Latin; the rule makes them neutral.
  if (cp >= 0xff00 && cp <= 0xffef) return undefined;
  return LATIN_LETTER.test(character) ? 'latinLetter' : undefined;
}

/** Whether §17.3.1.2 (DE) or §17.3.1.3 (DN) spaces the adjacent pair. */
export function eastAsianAutoSpacePairApplies(
  left: EastAsianAutoSpaceClass | undefined,
  right: EastAsianAutoSpaceClass | undefined,
  autoSpaceDE: boolean | undefined,
  autoSpaceDN: boolean | undefined,
): boolean {
  if (left === undefined || right === undefined) return false;
  const other = left === 'eastAsian' ? right : right === 'eastAsian' ? left : undefined;
  if (other === undefined || other === 'eastAsian') return false;
  return other === 'digit' ? autoSpaceDN !== false : autoSpaceDE !== false;
}

/** Text that participates in automatic spacing as an ordinary inline cell.
 * Fixed-width fitText cells, ruby, tate-chu-yoko and complex-script text keep
 * their own geometry; each interrupts adjacency like a tab or an object. */
function participates(segment: LayoutTextSeg): boolean {
  return segment.fitTextRegionIndex === undefined && !segment.ruby && !segment.tateChuYoko
    && segment.script !== 'complexScript' && segment.optionalHyphen === undefined
    && !segment.optionalHyphenGlyph;
}

function firstScalar(text: string): string | undefined {
  for (const scalar of text) return scalar;
  return undefined;
}

/** Base character of the text's last grapheme cluster: a combining mark,
 * variation selector or voiced-sound mark belongs to its base (UAX #29). */
function lastClusterBase(text: string): string | undefined {
  if (text.length === 0) return undefined;
  const offsets = graphemeClusterOffsets(text);
  return String.fromCodePoint(text.codePointAt(offsets.at(-1) ?? 0)!);
}

/** The [start, end) piece of an ordinary text segment, with every
 * offset-bearing fact projected through the shared slicing authority. A piece
 * after the first starts at a grapheme boundary that is not a source seam. */
function segmentPiece(segment: LayoutTextSeg, start: number, end: number): LayoutTextSeg {
  return {
    ...segment,
    ...RESET_SLICED_TEXT_MEASUREMENT,
    text: segment.text.slice(start, end),
    measuredWidth: 0,
    ...slicedTextMetadata(segment, start, end),
    ...(start > 0 ? {
      joinPrev: undefined,
      hardJoinPrev: undefined,
      explicitBreakBefore: undefined,
      breakBefore: true,
    } : {}),
  };
}

/**
 * WORD_EAST_ASIAN_AUTO_SPACE (§17.3.1.2 autoSpaceDE / §17.3.1.3 autoSpaceDN):
 * record a quarter-em gap before each segment that starts on the far side of
 * an East Asian / Latin-letter or East Asian / digit boundary. The gap belongs
 * to the segment after the boundary (`autoSpaceBeforePt`); the line breaker
 * applies it only when that segment follows its partner on the same line, so
 * neither a line start nor a line end carries it.
 *
 * Script-slot shaping already separates East Asian text from ASCII letters and
 * digits; a boundary inside one segment (a Latin letter assigned to the East
 * Asian face) is split here so every boundary is a segment seam.
 */
export function assignEastAsianAutoSpace(
  segs: LayoutSeg[],
  environment: Pick<LineLayoutEnvironment, 'autoSpaceDE' | 'autoSpaceDN' | 'compat'
    | 'characterSpacingControl' | 'verticalCJK'>,
): void {
  if (environment.autoSpaceDE === false && environment.autoSpaceDN === false) return;
  // WORD_COMPRESSED_SPACE_LINE_FIT shrinks the gap with the line's spaces
  // under its document gate; its lineWrapLikeWord6 and vertical exclusions
  // are the same as for U+0020.
  const compat = compatOf(environment);
  const compressible = wordCompressedSpaceLineFitApplies(
    compat, environment.characterSpacingControl,
  ) && !compat.lineWrapLikeWord6 && environment.verticalCJK !== true;
  const applies = (left: string | undefined, right: string | undefined): boolean =>
    left !== undefined && right !== undefined && eastAsianAutoSpacePairApplies(
      eastAsianAutoSpaceClass(left), eastAsianAutoSpaceClass(right),
      environment.autoSpaceDE, environment.autoSpaceDN,
    );
  // One linear pass; the array is rewritten once only when a seam was added.
  // A cluster that is only an extender (a mark authored apart from its base,
  // possibly in the previous run) is transparent: adjacency continues from
  // the preceding base, across runs.
  let split: LayoutSeg[] | undefined;
  let before: string | undefined;
  for (let index = 0; index < segs.length; index += 1) {
    const segment = segs[index]!;
    const splits: number[] = [];
    if (!('text' in segment) || !participates(segment)) {
      before = undefined;
    } else if (segment.text.length > 0) {
      const sea = segment.seaBreaks !== undefined;
      for (const cursor of [0, ...graphemeClusterOffsets(segment.text)]) {
        const base = String.fromCodePoint(segment.text.codePointAt(cursor)!);
        if (GRAPHEME_EXTENDERS.test(base)) continue;
        if (!sea && cursor > 0 && applies(before, base)) splits.push(cursor);
        before = base;
      }
    }
    if (splits.length === 0) {
      split?.push(segment);
      continue;
    }
    split ??= segs.slice(0, index);
    const text = segment as LayoutTextSeg;
    const bounds = [0, ...splits, text.text.length];
    for (let piece = 0; piece + 1 < bounds.length; piece += 1) {
      split.push(segmentPiece(text, bounds[piece]!, bounds[piece + 1]!));
    }
  }
  if (split) {
    segs.length = 0;
    for (const segment of split) segs.push(segment);
  }
  let previous: LayoutTextSeg | undefined;
  for (const segment of segs) {
    if (!('text' in segment) || !participates(segment)) {
      previous = undefined;
      continue;
    }
    segment.autoSpaceBeforePt = undefined;
    segment.autoSpaceCompressible = undefined;
    // An empty metric host (an anchor character) has no glyph and no
    // advance; it neither forms nor interrupts the visible adjacency. A run
    // holding only grapheme extenders (a variation selector or combining
    // mark authored apart from its base) belongs to the preceding base; a
    // leading extender prefix was split off above.
    if (segment.text.length === 0 || GRAPHEME_EXTENDERS.test(segment.text)) continue;
    const left = previous ? lastClusterBase(previous.text) : undefined;
    const right = firstScalar(segment.text);
    if (previous && applies(left, right)) {
      const eastAsianSide = eastAsianAutoSpaceClass(left!) === 'eastAsian' ? previous : segment;
      segment.autoSpaceBeforePt = wordEastAsianAutoSpacePt(eastAsianSide.fontSize);
      // A positive authored character pitch keeps its expanded spacing
      // (WORD_AUTHORED_CHARACTER_SPACING_PITCH_PRIORITY) on either side.
      if (compressible && !((previous!.charSpacing ?? 0) > 0) && !((segment.charSpacing ?? 0) > 0)) {
        segment.autoSpaceCompressible = true;
      }
    }
    previous = segment;
  }
}
