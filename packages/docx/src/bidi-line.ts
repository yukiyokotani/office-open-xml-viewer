import type { BidiEngine } from '@silurus/ooxml-core';
// WordprocessingML segment ordering uses the shared UAX#9 kernel with
// Word's HL1 run overrides and content-level Canvas direction.
import {
  resolveSegmentNativeLevelUnits,
  computeSegmentLineVisualOrder, computeSegmentLineSourceCellOrder, segmentLineHasRtl,
  type SegmentLineSourceCell,
  type SegmentLineVisualOrder, type SegmentLineBidiOptions,
  getDefaultBidiEngine, getBidiEngineLifecycleToken,
  certifyNativeSegmentLineOrder,
} from '@silurus/ooxml-core/internal/bidi-line';
import { wordRtlAmbiguousCharacter, wordRunDirectionSpan } from './layout/script-compatibility.js';

const options = Object.freeze({
  isTab: (segment: unknown) => 'isTab' in (segment as object),
  isRtlMarked: (segment: unknown) => (segment as { rtl?: unknown }).rtl === true,
  digitsAsAN: (segment: unknown) => (segment as { digitsAsAN?: unknown }).digitsAsAN === true,
  isRtlAmbiguous: wordRtlAmbiguousCharacter,
  // ECMA-376 §17.3.2.30 exempts numeric CS from the run RTL override.
  // Preserve the established Office punctuation policy otherwise (terminal
  // periods/prefixes and ES/ET are unchanged). Context belongs to the complete
  // canonical host, not acquired word or formatting-run boundaries.
  preserveRtlNumericCS: true,
  directionFromAnyOdd: true,
});
export type LineVisualOrder = SegmentLineVisualOrder;
export type LineOrderModel = 'line-level' | 'word';

/** Library representation boundaries shared by acquisition and final placement.
 * Leading-label isolates apply only to RTL paragraphs. Coarse distribute
 * acquisition retains its prior model. Validated final source cells can retain
 * original author scopes for positive word budgets; character/native/negative
 * and unsupported terminal or mixed-visible geometry keep the whole old path.
 * Character spread cannot place isolates, and compatibility shape text lacks
 * authoritative run-direction provenance. These keep their previous model.
 */
export function paragraphLineOrderModel(input: Readonly<{
  baseRtl: boolean; alignment?: string; runDirectionUnauthored?: boolean;
  characterSpreadBidi?: boolean;
}>): LineOrderModel {
  return input.characterSpreadBidi || input.alignment === 'distribute'
    || !input.baseRtl || input.runDirectionUnauthored ? 'line-level' : 'word';
}

export const segmentsHaveRtl = (segments: readonly unknown[]): boolean =>
  segmentLineHasRtl(segments, options);
// WORD_RUN_DIRECTION_ISOLATE: only final authored paragraphs opt in.
const authoredOptions = Object.freeze({
  ...options, isolateDirection: wordRunDirectionSpan, isolateRequiresUniformSlices: true,
});
/** Acquisition-only source units. The authored scope is assembled once from
 * the complete original source, never reconstructed from generated children. */
export const resolveNativeSourceLevelUnits = (segments: readonly unknown[], baseRtl: boolean) =>
  resolveSegmentNativeLevelUnits(segments, baseRtl, authoredOptions);

// One anchor is retained per measured slice. Intra-slice resolved-level
// boundaries are not retained, so punctuation or script changes inside one
// slice can still be misplaced (#2076); this is not a complete bidi rewrite.
/** The frozen production policy for one line-order model, shared by the
 * canonical resolver and the native certificate (and their differential test). */
export const lineOrderPolicy = (model: LineOrderModel): SegmentLineBidiOptions =>
  model === 'word' ? authoredOptions : options;

/** The order, from the start of the line, in which consecutive runs stand
 * on their own: a ruby cell's base runs, or a ruby guide's shaped runs.
 * It is the line's own resolution (UAX #9 over the characters of the joined
 * text, with the run overrides of §17.3.2.30), so European digits keep
 * their direction in a right-to-left run and run seams matter only where
 * direction does. The runs are resolved by themselves, in the paragraph's
 * base direction, whatever stands beside them on the line. */
export function runDisplayOrder(
  runs: readonly Readonly<{ text: string; rtl?: boolean; digitsAsAN?: boolean }>[],
  baseRtl: boolean,
): readonly number[] {
  if (runs.length <= 1) return runs.map((_, index) => index);
  // Only the facts of direction: a base run's cell identity must not make
  // the cell an isolate of itself.
  return computeSegmentLineVisualOrder(
    runs.map((run) => ({ text: run.text, rtl: run.rtl === true, digitsAsAN: run.digitsAsAN === true })),
    baseRtl, options,
  ).order;
}

/** The ruby cell a line item belongs to: the breaker's cell, or the cell
 * identity a placement descriptor carries (`rubyCell`). */
const rubyCellOfItem = (segment: unknown): unknown =>
  (segment as { rubyGroup?: unknown; rubyCell?: unknown }).rubyGroup
    ?? (segment as { rubyCell?: unknown }).rubyCell;

/** WORD_RUBY_ALIGN_PLACEMENT: a ruby cell is one unit of its line, so it is
 * one unit of the line's reordering too. A cell whose runs are all of one
 * direction stays together under UAX #9 as it is. A cell of mixed direction
 * would be parted when right-to-left text beside it joined its
 * right-to-left runs, so it is reordered as an isolate: an object
 * (U+FFFC) in the line, placed by the text around it, with its own runs in
 * their order inside it (runDisplayOrder, the order its widening is padded by).
 * Library policy: the specification does not say how a ruby base takes part
 * in bidirectional reordering. */
function withRubyCellIsolates(
  segments: readonly unknown[],
  baseRtl: boolean,
  compute: (segments: readonly unknown[]) => LineVisualOrder,
): LineVisualOrder {
  // Lines without ruby, nearly all of them, are resolved as they are.
  let hasCell = false;
  for (let index = 0; index < segments.length && !hasCell; index += 1) {
    hasCell = rubyCellOfItem(segments[index]) !== undefined;
  }
  if (!hasCell) return compute(segments);
  const isRtl = (segment: unknown) => (segment as { rtl?: unknown }).rtl === true;
  // Each reduced item: the original indexes it stands for.
  const groups: number[][] = [];
  let isolated = false;
  for (let index = 0; index < segments.length;) {
    const cell = rubyCellOfItem(segments[index]);
    let end = index + 1;
    while (cell !== undefined && end < segments.length && rubyCellOfItem(segments[end]) === cell) end += 1;
    const mixed = end - index > 1
      && segments.slice(index, end).some((segment) => isRtl(segment) !== isRtl(segments[index]));
    if (mixed) {
      groups.push(Array.from({ length: end - index }, (_, offset) => index + offset));
      isolated = true;
    } else {
      for (let at = index; at < end; at += 1) groups.push([at]);
    }
    index = end;
  }
  if (!isolated) return compute(segments);
  const reduced = groups.map((group) => (group.length === 1 ? segments[group[0]!] : { text: '\uFFFC' }));
  const visual = compute(reduced);
  const order: number[] = [];
  const rtl: boolean[] = new Array<boolean>(segments.length);
  visual.rtl.forEach((value, reducedIndex) => {
    const group = groups[reducedIndex]!;
    if (group.length === 1) rtl[group[0]!] = value;
    else for (const index of group) rtl[index] = isRtl(segments[index]);
  });
  for (const reducedIndex of visual.order) {
    const group = groups[reducedIndex]!;
    if (group.length === 1) order.push(group[0]!);
    else {
      const pieces = group.map((index) => segments[index] as { text: string; rtl?: boolean; digitsAsAN?: boolean });
      for (const at of runDisplayOrder(pieces, baseRtl)) order.push(group[at]!);
    }
  }
  return { order, rtl };
}

export const computeLineVisualOrder = (
  segments: readonly unknown[], baseRtl: boolean, model: LineOrderModel = 'line-level',
): LineVisualOrder => withRubyCellIsolates(segments, baseRtl, (items) =>
  computeSegmentLineVisualOrder(items, baseRtl, lineOrderPolicy(model)));

/** Final acquired source cells only; acquisition's coarse memo and decimal
 * anchors keep their established API/lifetime and native certificates. The
 * parent model is the SAME canonical resolver's original segment view, not
 * an alternate layout or paint pipeline. */
export const computeLineSourceCellOrder = (
  parents: readonly unknown[], cells: readonly SegmentLineSourceCell[],
  baseRtl: boolean, parentModel: LineOrderModel,
) => computeSegmentLineSourceCellOrder(
  parents, cells, baseRtl,
  // Proper source cells can preserve authored SPACE scope even in distribute.
  // Acquisition/decimal anchors and unsupported final lines keep their model.
  baseRtl ? authoredOptions : options,
  parentModel === 'word' ? authoredOptions : options,
);

export type LineOrderResolver = (
  segments: readonly unknown[], baseRtl: boolean, model?: LineOrderModel,
) => Readonly<{ order: readonly number[]; rtl: readonly boolean[] }>;

type OrderSegmentFacts = Readonly<{
  text: string | undefined; tab: boolean; rtl: boolean; digits: boolean;
  isolate: 'ltr' | 'rtl' | undefined;
  /** The ruby cell the segment is a piece of (withRubyCellIsolates). */
  cell: unknown;
}>;
const orderText = (segment: unknown): string | undefined => {
  const text = (segment as { text?: unknown }).text;
  return typeof text === 'string' ? text : undefined;
};
export const lineOrderEngineIdentity = (): Readonly<{
  engine: BidiEngine; token: object;
  compute: BidiEngine['computeLevels']; reorder: BidiEngine['reorderVisual'];
}> => {
  const engine = getDefaultBidiEngine();
  const compute = engine.computeLevels;
  const reorder = engine.reorderVisual;
  return { engine, compute, reorder, token: getBidiEngineLifecycleToken() } as const;
};
export const sameEngineState = (a: ReturnType<typeof lineOrderEngineIdentity>, b: ReturnType<typeof lineOrderEngineIdentity>): boolean =>
  a.engine === b.engine && a.token === b.token && a.compute === b.compute && a.reorder === b.reorder;

/** One synchronous acquisition's canonical ordering reuse. The owner drops
 * this single entry when layout returns/throws; changed contexts replace it,
 * never accumulate prefix history. Match the SAME hook facts as resolution,
 * including empty text versus objects and tab-property presence. Source
 * addresses, widths, seams and cell membership are deliberately not cached:
 * callers project their current acquisition through readonly order anew.
 * Engine lifecycle/method changes invalidate reuse, including same-object
 * setter calls; hidden semantic closure changes require that setter contract.
 * Only DOCX terminal-tab lookahead revisits the full hypothetical context in
 * convergence. XLSX/PPTX keep their existing uncached ordering consumers. */
export function createLineOrderResolver(): LineOrderResolver {
  let entry: Readonly<{
    baseRtl: boolean; model: LineOrderModel; engine: ReturnType<typeof lineOrderEngineIdentity>;
    facts: readonly OrderSegmentFacts[];
    visual: ReturnType<LineOrderResolver>;
  }> | undefined;
  return (segments, baseRtl, model = 'line-level') => {
    const active = lineOrderEngineIdentity();
    const policy = lineOrderPolicy(model);
    if (entry && entry.baseRtl === baseRtl && entry.model === model
      && sameEngineState(entry.engine, active) && entry.facts.length === segments.length
      && entry.facts.every((fact, i) => {
        const segment = segments[i];
        return fact.text === orderText(segment) && fact.tab === (policy.isTab?.(segment) === true)
          && fact.rtl === (policy.isRtlMarked?.(segment) === true)
          && fact.digits === (policy.digitsAsAN?.(segment) === true)
          && fact.isolate === policy.isolateDirection?.(segment)
          && fact.cell === rubyCellOfItem(segment);
      })) return entry.visual;
    entry = undefined;
    // These new primitive snapshots never escape this closure or get mutated.
    // Runtime freezes belong to the returned visual arrays, not private facts.
    const facts: readonly OrderSegmentFacts[] = segments.map(segment => ({
      text: orderText(segment), tab: policy.isTab?.(segment) === true,
      rtl: policy.isRtlMarked?.(segment) === true, digits: policy.digitsAsAN?.(segment) === true,
      isolate: policy.isolateDirection?.(segment),
      cell: rubyCellOfItem(segment),
    }));
    const computed = withRubyCellIsolates(segments, baseRtl, (items) =>
      certifyNativeSegmentLineOrder(items, baseRtl, policy)
        ?? computeSegmentLineVisualOrder(items, baseRtl, policy));
    // Copy before freezing: a third-party engine may own its returned array.
    const visual = Object.freeze({ order: Object.freeze([...computed.order]), rtl: Object.freeze([...computed.rtl]) });
    entry = sameEngineState(active, lineOrderEngineIdentity())
      ? Object.freeze({ baseRtl, model, engine: active, facts, visual }) : undefined;
    return visual;
  };
}

/** Physical edge a line aligns to, resolving logical start/end against base direction. */
export type AlignEdge = 'left' | 'right' | 'center' | 'justify';

/**
 * Resolve a paragraph's `w:jc` value (and base direction) to a physical edge.
 * ALL edge values are logical in WordprocessingML: `start`/`end` by definition
 * (§17.18.44), and the transitional `left`/`right` are defined as
 * "semantically equivalent to start/end" (ECMA-376 Part 4 §14.11.2) — so every
 * edge flips under an RTL base. An unset alignment defaults to the leading
 * (logical-start) edge.
 */
export function resolveAlignEdge(alignment: string | undefined, baseRtl: boolean): AlignEdge {
  switch (alignment) {
    case 'center':
      return 'center';
    case 'both':
    case 'justify':
    case 'distribute':
    // ECMA-376 §17.18.44: the three kashida settings (lowKashida / mediumKashida /
    // highKashida) and thaiDistribute are all forms of full justification between
    // both text margins — they differ only in HOW the extra space is distributed
    // (Arabic kashida elongation, Thai per-character spacing). The physical edge is
    // "justify" for all of them. NOTE: this is a MAPPING to the existing
    // inter-word/inter-character justification.
    // TODO(§17.18.44): true kashida (U+0640 tatweel) elongation — tracked in
    // https://github.com/yukiyokotani/office-open-xml-viewer/issues/724
    case 'lowKashida':
    case 'mediumKashida':
    case 'highKashida':
    case 'thaiDistribute':
      return 'justify';
    case 'end':
    case 'right':
      return baseRtl ? 'left' : 'right';
    case 'start':
    case 'left':
    case undefined:
    default:
      return baseRtl ? 'right' : 'left';
  }
}

/** ECMA-376 §17.18.44 — whether a `w:jc` value fully justifies each line by
 *  expanding inter-word (and, for distribute/thaiDistribute, inter-character)
 *  spacing. Covers `both` / `justify` / `distribute` plus the kashida and Thai
 *  variants, which this renderer maps onto the same slack-distribution kernel
 *  (see {@link resolveAlignEdge}; true kashida elongation is a follow-up). */
export function jcIsFullyJustified(alignment: string | undefined): boolean {
  switch (alignment) {
    case 'both':
    case 'justify':
    case 'distribute':
    case 'lowKashida':
    case 'mediumKashida':
    case 'highKashida':
    case 'thaiDistribute':
      return true;
    default:
      return false;
  }
}

/** ECMA-376 §17.18.44 — whether a `w:jc` value also stretches the paragraph's
 *  last line. `word-thai-distribute-cluster-policy` records that
 *  `thaiDistribute`, like `both`, leaves its final line ragged; only
 *  `distribute` stretches the last line. */
export function jcStretchesLastLine(alignment: string | undefined): boolean {
  return alignment === 'distribute';
}
