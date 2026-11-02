import { hasStrongRtl } from '@silurus/ooxml-core';
import type { NoteSourceContinuation } from './body-layout-kernel.js';
import type { LayoutParagraphBlock } from './layout-source-store.js';
import type { ParagraphLayout } from './types.js';
import { NOTE_REFRESH_PREFIX_WORK_MULTIPLIER } from './note-overflow-policy.js';

const supportedSources = new WeakMap<LayoutParagraphBlock, boolean>();

export interface NoteParagraphWindow {
  readonly paragraph: LayoutParagraphBlock;
  readonly runStart: number;
  readonly offset: number;
}

/** Recovery-only source acquisition, not an alternate line breaker. The normal
 * paragraph engine shapes geometrically growing prefixes from the unconsumed
 * source cursor. Keep an entire unfit line before the provisional final line:
 * the artificial end (including last-line justification and a truncated word)
 * can then never be committed. Capacity, not remaining paragraph length, stops
 * acquisition. Zero-height/unbreakable input is still subject to the note's
 * linear work allowance and returns its original admission error.
 *
 * Prefix trials may spend at most 4 times their largest source prefix. Four
 * is the geometric-doubling sum (<2) plus one factor of two for run metadata
 * and the lookahead frontier. If source weights do not grow that way, fail
 * rather than granting free scans. The note-wide backstop also counts every
 * prefix, whole-story acquisition, retry and convergence pass.
 */
export function acquireNoteParagraphWindow(
  paragraph: LayoutParagraphBlock,
  start: NoteSourceContinuation | undefined,
  capacityPt: number,
  maximumLines: number,
  units: (input: object) => number,
  charge: (input: object) => void,
  acquire: (window: NoteParagraphWindow) => ParagraphLayout,
): ParagraphLayout {
  // Windowing relies on local, forward text layout. Paragraph-wide shaping
  // (bidi, ruby, fitText, revisions or transformed text) has no established
  // bounded-prefix contract. Recovery may reject it with the saved error;
  // ordinary successful acquisition is unaffected.
  let supported = supportedSources.get(paragraph);
  if (supported === undefined) {
    supported = !paragraph.bidi && !paragraph.runs.some(run =>
      (run.type !== 'text' && run.type !== 'field') || run.rtl || run.cs || run.ruby
      || run.allCaps || run.smallCaps || run.fitTextVal != null || run.revision
      || run.noteRef || run.noteSeparatorCharacter || run.eastAsianVert
      || (run.type === 'text' && hasStrongRtl(run.text))
      || run.noBreakBefore || run.noBreakAfter || run.noBreakRanges?.length || run.optionalHyphen);
    supportedSources.set(paragraph, supported);
  }
  if (!supported) {
    throw new Error('Unsupported note paragraph source window');
  }
  if (start?.paragraph.boundary && !start.sourceRunCursor) {
    throw new Error('Note paragraph window requires a source-run cursor');
  }
  const boundary = start?.paragraph.boundary;
  const runStart = boundary?.segIndex ?? 0;
  const offset = boundary?.charOffset ?? 0;
  if (offset > 0 && paragraph.runs[runStart]?.type !== 'text') {
    // A PAGE result split across destinations has no stable character cut
    // when the destination changes its display length. Never guess that cut.
    throw new Error('Cannot continue inside a destination field result');
  }
  // Amortize the fixed paragraph/first-run metadata before doubling text.
  // Starting at one character would repeatedly pay a large fixed header for
  // a single long run, defeating the four-prefix work bound even on plain
  // text. This seed is structural source size, not a fitted chunk length.
  let characters = units({ ...paragraph, runs: [] })
    + units({ ...paragraph.runs[runStart], text: '' });
  let spent = 0;
  for (;;) {
    let remaining = characters;
    const runs: LayoutParagraphBlock['runs'][number][] = [];
    let index = runStart;
    let endOffset = offset;
    while (index < paragraph.runs.length && remaining > 0) {
      const run = paragraph.runs[index]!;
      if (run.type === 'text') {
        const from = index === runStart ? offset : 0;
        let length = Math.min(run.text.length - from, remaining);
        // Do not create half a UTF-16 surrogate at the provisional boundary.
        const last = run.text.charCodeAt(from + length - 1);
        if (length > 0 && from + length < run.text.length && last >= 0xd800 && last <= 0xdbff) length++;
        runs.push({ ...run, text: run.text.slice(from, from + length) });
        remaining -= Math.max(1, length);
        endOffset = from + length;
        if (endOffset < run.text.length) break;
      } else {
        runs.push(run);
        remaining--;
      }
      index++;
      endOffset = 0;
    }
    const window: NoteParagraphWindow = {
      paragraph: Object.freeze({ ...paragraph, runs: Object.freeze(runs) }), runStart, offset,
    };
    const weight = units(window.paragraph);
    if (spent + weight > NOTE_REFRESH_PREFIX_WORK_MULTIPLIER * weight) throw new Error('Note paragraph refresh prefix budget exceeded');
    charge(window.paragraph);
    spent += weight;
    const shaped = acquire(window);
    const complete = index === paragraph.runs.length && endOffset === 0;
    const lookahead = shaped.lines.at(-2);
    if (complete || shaped.lines.length > maximumLines + 1 || (lookahead && lookahead.bounds.yPt + lookahead.advancePt
      > shaped.flowBounds.yPt + capacityPt + 1e-6)) return shaped;
    characters *= 2;
  }
}

/** A window owns only a run slice. Restore original run identities before
 * partitioning so field ownership, source progress and the next destination's
 * cursor never depend on this temporary paragraph's local indices. */
export function projectNoteParagraphWindow(
  layout: ParagraphLayout,
  cursors: readonly Readonly<{ runIndex: number; offset: number }>[],
  window: NoteParagraphWindow,
): ParagraphLayout {
  return {
    ...layout, noteSourceWindow: true,
    noteLineEndBoundaries: cursors.map(cursor => ({
      segIndex: cursor.runIndex + window.runStart,
      charOffset: cursor.offset + (cursor.runIndex === 0 ? window.offset : 0),
    })),
    lines: layout.lines.map(line => ({ ...line,
      placements: line.placements.map(placement => {
        if (placement.kind !== 'text' && placement.kind !== 'resource') return placement;
        return { ...placement,
          ...(placement.sourceRunIndex === undefined ? {} : { sourceRunIndex: placement.sourceRunIndex + window.runStart }),
          ...(placement.kind === 'text' && placement.sourceRuns ? {
            sourceRuns: placement.sourceRuns.map(owner => ({
              ...owner, sourceRunIndex: owner.sourceRunIndex + window.runStart,
            })),
          } : {}),
        };
      }),
    })),
  };
}
