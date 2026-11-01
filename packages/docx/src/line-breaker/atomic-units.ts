import { textBreakOffsetAt } from './text-break-window.js';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import type { LayoutSeg, LayoutTextSeg } from './model.js';
import { hardJoinPrefixEnd } from './advance.js';
import { hasCJKBreakOpportunity } from './text-runs.js';

export interface AtomicTextMeasurement {
  readonly baseRtl: boolean;
  readonly reservePrefixWork: (utf16Units: number) => void;
  readonly segAdvance: (segment: LayoutTextSeg) => number;
  readonly strAdvance: (segment: LayoutTextSeg, text: string) => number;
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
  // A leader prefix keeps the leader's leading WORD_KERNED_SPACE_PAIRS seam
  // pair, which `w` (the whole leader as fitted) already includes.
  if (firstEnd < s.text.length) return {
    width: strAdvance(s, firstText) + (s.leadingWordBoundaryPx ?? 0), trailingSpace: 0, next: s, pieces,
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
    const trailing = !fixedCell && !measurement.baseRtl && prefix.endsWith(' ')
      ? fw - strAdvance(f, trimmed) : 0;
    groupTrail = trimmed.length === 0 && groupTrail > 0 ? groupTrail + trailing : trailing;
    if (end < f.text.length || !fixedCell && explicitEnd !== undefined && hardEnd === undefined) break;
    next = undefined;
  }
  return { width: groupW, trailingSpace: groupTrail, next, pieces };
}
