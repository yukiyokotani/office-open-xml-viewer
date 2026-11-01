import { textBreakOffsetAt } from './text-break-window.js';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import type { LayoutSeg, LayoutTextSeg } from './model.js';
import { hardJoinPrefixEnd } from './advance.js';
import { hasCJKBreakOpportunity } from './text-runs.js';

export interface AtomicTextMeasurement {
  readonly baseRtl: boolean;
  /** WORD_RTL_JUSTIFIED_VISIBLE_FIT: right-to-left lines fit their visible prefix. */
  readonly lineFit?: Readonly<{ rtlVisiblePrefix: boolean }>;
  readonly reservePrefixWork: (utf16Units: number) => void;
  readonly segAdvance: (segment: LayoutTextSeg) => number;
  readonly strAdvance: (segment: LayoutTextSeg, text: string) => number;
}

/** Whether line fitting charges a line-end U+0020: right-to-left lines do
 * (library containment policy) unless WORD_RTL_JUSTIFIED_VISIBLE_FIT applies. */
export function lineEndSpaceCharged(state: Pick<AtomicTextMeasurement, 'baseRtl' | 'lineFit'>): boolean {
  return state.baseRtl && state.lineFit?.rtlVisiblePrefix !== true;
}

/** Advance of the U+0020 suffix that ends `items` across members (a run of
 * spaces may span separately formatted segments); zero-width markers are
 * looked through. The single suffix measure of every visible-prefix fit. */
export function trailingSpaceSuffixWidth(
  items: readonly LayoutSeg[],
  measurement: Pick<AtomicTextMeasurement, 'segAdvance' | 'strAdvance' | 'reservePrefixWork'>,
  /** Items before this index are known to hold `knownWidth` of the suffix,
   * running back into them only when every later item is spaces alone. */
  from = 0,
  knownWidth = 0,
): number {
  let width = 0;
  for (let index = items.length - 1; index >= from; index -= 1) {
    const item = items[index]!;
    if (!('text' in item)) return width;
    if (item.text === '') continue;
    if (!item.text.endsWith(' ')) return width;
    measurement.reservePrefixWork(item.text.length);
    const trimmed = item.text.replace(/ +$/u, '');
    width += measurement.segAdvance(item) - measurement.strAdvance(item, trimmed);
    if (trimmed.length > 0) return width;
  }
  return width + knownWidth;
}

/** Running {@link trailingSpaceSuffixWidth} of a line that grows by appends:
 * each query reads only the items appended since the last one, so a run of
 * space-only segments costs linear work. A line that was cut back or replaced
 * (a flush or a rollback) is measured afresh. */
const lineSuffixCache = new WeakMap<readonly LayoutSeg[], { length: number; last: LayoutSeg | undefined; width: number }>();
export function lineTrailingSpaceSuffixWidth(
  line: readonly LayoutSeg[],
  measurement: Pick<AtomicTextMeasurement, 'segAdvance' | 'strAdvance' | 'reservePrefixWork'>,
): number {
  const cached = lineSuffixCache.get(line);
  const valid = cached !== undefined && cached.length <= line.length
    && (cached.length === 0 || line[cached.length - 1] === cached.last);
  const width = valid
    ? trailingSpaceSuffixWidth(line, measurement, cached.length, cached.width)
    : trailingSpaceSuffixWidth(line, measurement);
  lineSuffixCache.set(line, { length: line.length, last: line[line.length - 1], width });
  return width;
}

/** §17.3.2.14 links contiguous fitText runs into one specified width.
 * Internal separators and formatting seams cannot shorten that allocation.
 * Admission and placement consume this same complete resolved region. */
export function measureFitTextUnit(
  first: LayoutTextSeg, following: Iterable<LayoutSeg>, measurement: AtomicTextMeasurement, startIndex = 0,
): number {
  let width = measurement.segAdvance(first);
  let index = 0;
  for (const member of following) {
    if (index++ < startIndex) continue;
    if (!('text' in member) || member.fitTextRegionIndex !== first.fitTextRegionIndex) break;
    width += measurement.segAdvance(member);
  }
  return width;
}

/** The source-seam group used by both gap admission and ordinary placement.
 * A hard no-break seam consumes its protected prefix; external links retain
 * their explicit hyphen/registered syntax opportunities; CJK followers glue only non-starters.
 * UAX #14 LB7 retains a space suffix across non-textual formatting seams.
 * A following fitText region is admitted separately when its head is reached;
 * at an empty line head, a preceding word can finish before that region moves
 * as its own cell. Mid-line lookahead still avoids splitting joined words.
 * Ordinary admission stops at the next legal prefix; whole-leader admission
 * may cross the leader's earlier break but still stops at its next follower
 * opportunity. Only CJK-origin whole-word admission skips all ordinary hyphens.
 * Every visited piece shares the pass work quota with exact split searches.
 * No field/link/run direction grants extra atomic ownership by itself. */
export function measureJoinedTextUnit(
  s: LayoutTextSeg, following: Iterable<LayoutSeg>, measurement: AtomicTextMeasurement,
  w = measurement.segAdvance(s), trailingSpaceW = 0, startIndex = 0, atLineStart = false, mode: 'prefix' | 'whole-word' | 'whole-leader' = 'prefix',
): Readonly<{
  width: number;
  trailingSpace: number;
  next: LayoutSeg | undefined;
  /** The joined text, piece by piece (the leader first). */
  pieces: readonly Readonly<{ segment: LayoutTextSeg; text: string }>[];
}> {
  const { segAdvance, strAdvance } = measurement;
  const preferWholeWord = mode === 'whole-word';
  const firstEnd = mode === 'prefix' && s.explicitBreaks
    ? textBreakOffsetAt(s.explicitBreaks, 0) : s.text.length;
  measurement.reservePrefixWork(firstEnd);
  const firstText = s.text.slice(0, firstEnd);
  const pieces: { segment: LayoutTextSeg; text: string }[] = [{ segment: s, text: firstText }];
  if (firstEnd < s.text.length) return {
    width: strAdvance(s, firstText), trailingSpace: 0, next: s, pieces,
  };
  let groupW = w;
  let groupTrail = s.fitTextRegionIndex === undefined ? trailingSpaceW : 0;
  let index = 0;
  let next: LayoutSeg | undefined;
  for (const f of following) {
    if (index++ < startIndex) continue;
    next = f;
    if (!('text' in f)) break;
    const wholeWordSeam = preferWholeWord && f.explicitBreakBefore && f.hyperlink?.kind !== 'external';
    if ((!f.joinPrev && !wholeWordSeam) || atLineStart && f.fitTextRegionStart) break;
    const fixedCell = f.ruby !== undefined || f.tateChuYoko === true;
    // Hard-seam acquisition scans the entire follower to find its first legal
    // grapheme. Charge that scan before allocation, not just its short result.
    if (f.hardJoinPrev) measurement.reservePrefixWork(f.text.length);
    const hardEnd = hardJoinPrefixEnd(f);
    const explicitEnd = !preferWholeWord && f.explicitBreaks ? textBreakOffsetAt(f.explicitBreaks, 0) : undefined;
    let end = fixedCell ? f.text.length : hardEnd ?? explicitEnd ?? f.text.length;
    if (!fixedCell && hardEnd === undefined && explicitEnd === undefined && hasCJKBreakOpportunity(f.text)) {
      // A CJK follower glues its leading non-starters, then resumes ordinary
      // inter-character breaks. The Latin leader never owns the whole CJK run.
      end = 0;
      for (const character of f.text) {
        if (!DEFAULT_KINSOKU_RULES.lineStartForbidden.has(character.codePointAt(0)!)) break;
        end += character.length;
      }
    }
    measurement.reservePrefixWork(end);
    const prefix = f.text.slice(0, end);
    pieces.push({ segment: f, text: prefix });
    const fw = end === f.text.length ? segAdvance(f) : strAdvance(f, prefix);
    groupW += fw;
    const trimmed = prefix.replace(/ +$/, '');
    const trailing = !fixedCell && !lineEndSpaceCharged(measurement) && prefix.endsWith(' ')
      ? fw - strAdvance(f, trimmed) : 0;
    groupTrail = trimmed.length === 0 && groupTrail > 0 ? groupTrail + trailing : trailing;
    if (end < f.text.length || !fixedCell && explicitEnd !== undefined && hardEnd === undefined) break;
    next = undefined;
  }
  return { width: groupW, trailingSpace: groupTrail, next, pieces };
}
