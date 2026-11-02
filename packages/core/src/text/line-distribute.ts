// Shared "fill the line" slack distributor for justified text — the single gap
// model behind WordprocessingML §17.18.44 (`both`/`distribute`) and DrawingML
// §20.1.10.59 (`just`/`dist`). Given a laid-out line's segments (logical order)
// and its slack, it reports — per segment — where each gap falls and the per-gap
// px, so a renderer can slice the glyph drawing and advance the pen.
//
// ── Why one kernel for both formats ─────────────────────────────────────────
// Word and PowerPoint reach the SAME gap selection: a stretch opportunity is an
// inter-word space OR an inter-CJK boundary (either side a CJK / ideographic
// glyph), evaluated across the whole line's code-point stream — boundaries are
// tested across segment edges too, so a colour change mid-phrase doesn't swallow
// a gap. They differ only in policy the CALLER owns and which this kernel does
// NOT encode:
//   • last-line policy (Word `both` / PowerPoint `just` leave the final line
//     natural; `distribute` / `dist` fill it). The caller decides whether to
//     call the kernel for a given line — the kernel just stretches whatever line
//     it is handed.
//   • the whitespace predicate. PowerPoint treats every JS `\s` char as an
//     inter-word space; Word counts only U+0020 and U+3000. Each caller injects
//     its own `isWhitespace` so the kernel reproduces that format's behaviour.
//   • the gap predicate. EXPANSION opens inter-CJK boundaries (`isGapChar` =
//     core.isCjkBreakChar); a COMPRESSION path that may shrink spaces but must
//     not overlap ideographs injects `isGapChar: () => false`.
//
// Why a per-segment / per-character model: layout merges adjacent same-style
// tokens into ONE segment, so a CJK phrase like "観察することで" is a single
// segment with no internal spaces — its inter-CJK gaps fall INSIDE the segment,
// between code points. The kernel walks the line's whole code-point stream and
// reports, per segment, the interior split offsets and a trailing-edge flag.
//
// This file is format-agnostic: no ECMA-376 §-specific policy lives here, only
// the gap geometry both renderers share. See packages/pptx/src/text-justify.ts
// and packages/docx/src/text-distribute.ts for the format adapters.

import { enumerateGaps, type GapScanState, type LineGap } from './line-gaps.js';
export { enumerateGaps, type GapScanState, type LineGap } from './line-gaps.js';

/** A laid-out segment as the distributor sees it. Only the optional text matters;
 *  an undefined `text` marks a non-text inline atom (image / math / tab) — one
 *  opaque unit bearing no stretch of its own, though a CJK neighbour can still
 *  open a gap against its edge (the atom counts as a non-CJK, non-space unit). */
export interface DistributeSeg {
  text?: string;
}

/** Per-text-segment instructions for applying the line's slack.
 *
 *  `splitBefore` lists the code-point offsets (1..len-1, counted in code points,
 *  NOT UTF-16 units) at which an INTERNAL gap falls — the glyphs before the
 *  offset are drawn, the pen advances by `perGap`, then drawing resumes.
 *  `trailingGap` marks that the boundary AFTER this segment's last code point
 *  (the inter-segment boundary) is a stretch opportunity. */
export interface SegStretch {
  /** Code-point offsets inside the segment after which to insert `perGap`. */
  splitBefore: number[];
  /** Whether to advance `perGap` after the whole segment (inter-segment gap). */
  trailingGap: boolean;
  /** px added strictly INSIDE the segment = splitBefore.length * perGap.
   *  Decorations (highlight / underline / strike / ruby centring / onTextRun
   *  width) should span measuredWidth + internalStretch. */
  internalStretch: number;
  /** Per-cut deltas when the caller supplies natural gap widths. */
  gapDeltas?: number[];
  /** Delta at the segment edge under proportional allocation. */
  trailingDelta?: number;
}

/** Result of distributing a line's slack across its gap opportunities. */
export interface DistributeResult {
  /** A caller-supplied unweighted expansion family retained its old allocation. */
  usedUnweightedExpansion?: true;
  /** px added at each gap (negative when the line is compressed). */
  perGap: number;
  /** Per-segment stretch, keyed by the segment's index in `segments`. Segments
   *  with no gap (and non-text atoms) are absent. */
  perSeg: Map<number, SegStretch>;
}

/** Tuning for {@link distributeLineSlack}; every field is optional and defaults
 *  to the WordprocessingML expansion behaviour. */
export interface DistributeOptions {
  /** Reuse the breaker's measured opportunities without rescanning text. */
  gapModel?: Readonly<{ gaps: readonly LineGap[]; state: Readonly<GapScanState> }>;
  /** Allocate proportionally to natural advances in a retained gap model. */
  proportional?: boolean;
  /** A separately measured expansion family whose allocation is unweighted.
   * Select it only for positive slack with a non-ASCII-space opportunity;
   * callers without this option retain their ordinary family. This bounds the
   * proportional observation without changing ideographic/SEA expansion. */
  unweightedExpansion?: Readonly<{ gaps: readonly LineGap[]; slack: number }>;

  /** Index of the first segment holding non-whitespace content; earlier
   *  (leading-indent / 字下げ whitespace) segments are fixed and never open a
   *  gap. Default 0 (no skip — e.g. under bidi, or for callers with no indent
   *  concept). */
  /** Require ASCII-space runs to have text cells on both sides. Tabs and
   * opaque atoms freeze adjoining spaces. Other gap families keep their policy. */
  textCellSpaceGaps?: boolean;
  /** Keep a space and its following grapheme extension atomic, including
   * across segment seams. Callers that omit this retain their gap policy. */
  atomicSpaceGaps?: boolean;
  /** Width of a gap a segment carries before its first code point (keyed by
   * the segment index including any `segmentOffset`); 0 or omitted means none.
   * It is an opportunity of the same family as the space gaps. */
  leadingGapWidth?: (segmentIndex: number) => number;
  /** Natural gap advance in the same units as slack. Supplying this selects
   * proportional allocation; omitted callers retain equal per-gap allocation. */
  gapWidth?: (segmentIndex: number, cpOffset: number, cp: number) => number;
  /** Lower bound on proportional compression, e.g. -0.25. */
  minFactor?: number;
  firstContentSi?: number;
  /** Index of the VISUALLY-final segment; it and the boundary into it open no
   *  gap. Default `segments.length - 1`. The match is EXACT, not `>=`: under
   *  bidi this is the visually-last segment's LOGICAL index, which is not the
   *  maximum si, so `>=` would wrongly suppress every logically-later segment
   *  (a pure-RTL line would skip the whole line → no justification). Pass a value
   *  ≥ segments.length (e.g. `segments.length`) to exclude NO segment: the pptx
   *  adapter does this because it draws every segment in one loop and relies on
   *  the content-span trim to suppress only the final glyph's gap. */
  lastDrawnSi?: number;
  /** Lower bound on a (negative) `perGap` when compressing (slack < 0), so a
   *  compression never eats more than a capped amount per gap. Default
   *  -Infinity (uncapped). Ignored when slack >= 0. */
  minPerGap?: number;
  /** A boundary between two non-space code points opens an inter-CJK gap when
   *  EITHER side satisfies this predicate. Default core.isCjkBreakChar (open a
   *  gap when either side is a CJK / ideographic glyph). Pass `() => false` for
   *  a compression path that must not overlap ideographs (only spaces stretch). */
  isGapChar?: (cp: number) => boolean;
  /** Classifies a code point as inter-word whitespace: it becomes ONE gap and is
   *  never treated as a CJK boundary. Default `cp === 0x20 || cp === 0x3000`
   *  (the WordprocessingML set). PowerPoint injects a wider JS-`\s` predicate.
   *  Whitespace is classified FIRST, so an ideographic space is one inter-word
   *  gap and never reaches `isGapChar`. */
  isWhitespace?: (cp: number) => boolean;
  /** WordprocessingML `thaiDistribute` (§17.18.44) / DrawingML `thaiDist`
   *  (§20.1.10.59). When true, a gap ALSO opens at every UAX#29 grapheme-cluster
   *  boundary INTERIOR to a Southeast-Asian span (Thai/Lao/Khmer), so a line of
   *  space-free SEA text is justified by widening inter-CLUSTER gaps — a combining
   *  vowel/tone mark stays glued to its base (no slack inside a cluster). Off by
   *  default: `both`/`distribute`/`just`/`dist` do NOT distribute SEA text (Word/
   *  PowerPoint leave it ragged). Only boundaries whose BOTH sides are SEA open
   *  this way; non-SEA boundaries keep the ordinary space/CJK rules. Verified
   *  against the Word-exported adjudication fixture (issue #959). */
  seaClusterGaps?: boolean;
  /** With {@link seaClusterGaps}: the SEA cluster before an inter-word space
   *  run opens its own gap too (committed with the run, so a line-end run
   *  opens none). WordprocessingML `thaiDistribute` before compatibility mode
   *  15 widens every Thai cluster and space; omitted callers keep the run as
   *  the only opportunity. */
  seaClusterSpaceGaps?: boolean;
}

/** Distribute slack over the enumerated opportunities. Existing callers use
 * equal deltas. Callers with measured gap widths use delta_i = slack*w_i/S,
 * including tiny negative slack (the legacy 0.5px floor does not apply there). */
export function distributeLineSlack<T extends DistributeSeg>(
  segments: readonly T[],
  slack: number,
  opts: DistributeOptions = {},
): DistributeResult | null {
  let proportional = opts.proportional ?? opts.gapWidth !== undefined;
  const enumeration = opts.gapModel ?? enumerateGaps(segments, opts);
  let { gaps, state } = enumeration;
  const expansion = opts.unweightedExpansion;
  const useExpansion = slack > 0 && expansion?.gaps.some(gap =>
    gap.kind === 'boundary' || gap.codePoint !== 0x20);
  if (useExpansion && expansion) {
    gaps = expansion.gaps;
    state = { ...state, naturalGapSum: gaps.length };
    slack = expansion.slack;
    proportional = false;
  }
  if (proportional ? slack === 0 : Math.abs(slack) <= 0.5) return null;
  if (!(state.naturalGapSum > 0) || gaps.length === 0) return null;
  let factor = slack / state.naturalGapSum;
  if (slack < 0) factor = Math.max(factor, proportional ? opts.minFactor ?? -Infinity : opts.minPerGap ?? -Infinity);
  const perSeg = new Map<number, SegStretch>();
  for (const gap of gaps) {
    let stretch = perSeg.get(gap.segIndex);
    if (!stretch) {
      stretch = { splitBefore: [], trailingGap: false, internalStretch: 0,
        ...(proportional ? { gapDeltas: [], trailingDelta: 0 } : {}) };
      perSeg.set(gap.segIndex, stretch);
    }
    const delta = factor * (proportional ? gap.naturalPx : 1);
    if (gap.trailing) {
      stretch.trailingGap = true;
      if (proportional) stretch.trailingDelta = delta;
    } else {
      stretch.splitBefore.push(gap.cpOffset + 1);
      stretch.internalStretch += delta;
      stretch.gapDeltas?.push(delta);
    }
  }
  return { perGap: proportional ? 0 : factor, perSeg,
    ...(useExpansion ? { usedUnweightedExpansion: true } : {}) };
}
