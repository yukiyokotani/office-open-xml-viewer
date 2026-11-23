import { revisionIsOmitted } from '../layout/revision-visibility.js';
import { canvasLanguageTag, type ParagraphLayoutRun, type ParagraphTextBearingRun } from '../layout/text.js';
import type { LineLayoutEnvironment } from './model.js';
import { resolveFieldText } from './text-runs.js';
import type { FieldRun } from '../types.js';
import type { RunTypographyAcquisitionInput, TypographyValueInput } from '../layout/typography-input.js';

export interface TextSequenceSource {
  readonly runIndex: number;
  /** UTF-16 offsets in the complete displayed sequence. */
  readonly start: number;
  readonly end: number;
}
export interface TextSequence {
  readonly run: Extract<ParagraphTextBearingRun, { type: 'text' }>;
  readonly sources: readonly TextSequenceSource[];
  readonly optionalHyphens: readonly Readonly<{
    offset: number;
    runIndex: number;
    run: Extract<ParagraphTextBearingRun, { type: 'text' }>;
  }>[];
}

function visibleTextRun(run: ParagraphLayoutRun, environment: LineLayoutEnvironment) {
  if (run.type !== 'text' && run.type !== 'field') return undefined;
  const r: ParagraphTextBearingRun = run;
  // These are authored units, rather than formatting-only seams: a ruby base,
  // note marker, fitText unit, noBreakHyphen owner, upright tate-chu-yoko cell
  // or explicitly advanced space.
  if (r.optionalHyphen || r.ruby || r.noteRef || r.fitTextVal != null || r.noBreakRanges?.length
    || r.noBreakBefore || r.noBreakAfter || r.eastAsianVert || r.spaceAdvancePt !== undefined
    || r.listLabel !== undefined) return undefined;
  const text = run.type === 'text' ? run.text : resolveFieldText(run as FieldRun, environment);
  const { type: _type, ...properties } = run;
  const projected = { ...properties, type: 'text' as const, text,
    isLink: 'isLink' in run ? run.isLink : false,
    hyperlink: 'hyperlink' in run ? run.hyperlink : null };
  return projected as Extract<ParagraphTextBearingRun, { type: 'text' }>;
}

type SequenceRun = Extract<ParagraphTextBearingRun, { type: 'text' }> & Readonly<{
  typographyInput?: RunTypographyAcquisitionInput;
}>;

/** Whether every scalar of `run` takes the complex-script slot. */
const forcedComplexScript = (run: Pick<SequenceRun, 'cs' | 'rtl'>): boolean => run.cs === true || run.rtl === true;

/** The effective w:val of `run`'s text: none on a forced complex-script run. */
const westernLanguageOf = (run: SequenceRun): string | undefined =>
  forcedComplexScript(run) ? undefined : canvasLanguageTag(run.lang);

/** Whether `run` names one language on every §17.3.2.20 axis (all equal,
 * including all absent): its text is in that language whichever slot a
 * scalar takes. An OpenDocument run is always uniform (its ODF script class
 * selects one language). */
const uniformLanguage = (run: SequenceRun): boolean =>
  run.lang === run.langEastAsia && run.lang === run.langBidi;

/** The language a language range records for `run`'s text. */
const rangeLanguageOf = (run: SequenceRun): string | undefined =>
  uniformLanguage(run) ? canvasLanguageTag(run.lang) : westernLanguageOf(run);

/** Which §17.3.2.20 axes a formatting key includes. */
type LanguageAxes = 'all' | 'nonWestern' | 'none';

function formattingKey(run: SequenceRun, environment: LineLayoutEnvironment, axes: LanguageAxes = 'all'): string {
  const input = run.typographyInput;
  const effective = <T>(value: TypographyValueInput<T> | undefined, fallback: T | undefined) =>
    value?.status === 'valid' && value.value !== null ? value.value : fallback;
  // Closed projection of resolved layout/paint facts. Never serialize a source
  // run wholesale: parser sidecars also carry sourceText, field instructions,
  // revision ids and lexical spellings that must not partition glyph streams.
  // Paint differences still delimit sequences; changed-format shaping is not
  // established by the identical-format invariant.
  const record = {
    fontSize: run.fontSize, fontFamily: run.fontFamily,
    fontFamilyHighAnsi: run.fontFamilyHighAnsi, fontFamilyEastAsia: run.fontFamilyEastAsia,
    fontFamilyCs: run.fontFamilyCs, fontSlots: run.fontSlots, fontHint: run.fontHint,
    fontSizeCs: run.fontSizeCs, bold: run.bold, italic: run.italic,
    boldCs: run.boldCs, italicCs: run.italicCs, rtl: run.rtl, cs: run.cs,
    // The effective w:val (canvasLanguageTag): x-none, ill-formed and absent
    // tags are one state and never delimit a sequence.
    // A run forced onto the complex-script slot (§17.3.2.7 w:cs, §17.3.2.30
    // w:rtl) never uses w:val, so it cannot delimit (e.g. Arabic joining).
    ...(axes !== 'none' ? { langBidi: run.langBidi, langEastAsia: run.langEastAsia } : {}),
    ...(axes === 'all' ? { lang: westernLanguageOf(run) } : {}),
    vertAlign: effective(input?.verticalAlign, run.vertAlign ?? undefined),
    position: effective(input?.positionPt, run.position),
    charSpacing: input?.characterSpacingPt ?? run.charSpacing,
    charScale: input?.characterScale ?? run.charScale,
    kerning: input?.kerningThresholdPt ?? run.kerning,
    snapToGrid: input?.snapToGrid ?? run.snapToGrid,
    allCaps: run.allCaps, smallCaps: run.smallCaps,
    underline: run.underline, underlineStyle: run.underlineStyle, underlineColor: run.underlineColor,
    strikethrough: run.strikethrough, doubleStrikethrough: run.doubleStrikethrough,
    color: run.color, colorAuto: run.colorAuto, background: run.background,
    highlight: run.highlight, emphasisMark: run.emphasisMark,
    // ECMA-376 §17.3.2.4 groups borders only when every attribute agrees.
    // Public borders omit theme/shadow/frame facts retained by the parser.
    border: run.border ? {
      ...run.border, val: input?.border?.val.value ?? run.border.style,
      themeColor: input?.border?.themeColor.value || undefined,
      themeTint: input?.border?.themeTint.value || undefined,
      themeShade: input?.border?.themeShade.value || undefined,
      shadow: effective(input?.border?.shadow, undefined),
      frame: effective(input?.border?.frame, undefined),
    } : undefined,
    hyperlink: run.hyperlink, hyperlinkAnchor: run.hyperlinkAnchor,
    // Markup colors/decorations depend on kind and effective author color, not
    // revision identity. Final view does not paint revision markup.
    revision: environment.showTrackedChanges === true && run.revision ? {
      kind: run.revision.kind,
      color: environment.revisionAuthorColor?.(run.revision.author) ?? '#C00000',
    } : undefined,
    textBoxLineFloor: (run as SequenceRun & { textBoxLineFloor?: boolean }).textBoxLineFloor,
    textBoxVertical: (run as SequenceRun & { textBoxVertical?: boolean }).textBoxVertical,
  };
  return JSON.stringify(record, (_key, value: unknown) =>
    value && typeof value === 'object' && !Array.isArray(value)
      ? Object.fromEntries(Object.entries(value).filter(([, v]) => v != null)
        .sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0))
      : value);
}

/** UTF-16 units of a sequence's last grapheme cluster kept as context for
 * judging the next seam ({@link SequenceTail.crossing}). */
const CLUSTER_CONTEXT_UNITS = 64;

/** The §17.3.2.20 w:val languages of a sequence whose runs disagree on it:
 * contiguous ranges of the sequence's displayed text (UTF-16 `end` offsets,
 * ascending; the last range ends with the text), every boundary a grapheme
 * cluster boundary. Only a sequence joined across a w:val-only seam inside a
 * cluster carries them; `run.lang` is then the first range's language. */
export interface SequenceLanguages {
  readonly westernLanguageRanges?: readonly Readonly<{ end: number; lang: string | undefined }>[];
  /** Every joined run names one language on all axes: a range's language
   * applies to every axis of its text, not to w:val alone. */
  readonly uniformLanguageRanges?: true;
}

let graphemeSegmenter: Intl.Segmenter | null | undefined;

/** Whether a grapheme cluster boundary between `previous` and `next` can be
 * suppressed in some context, judged from the two scalars at the seam alone.
 * Below U+0300 the only Grapheme_Cluster_Break values are Control, CR, LF
 * and Other, so only CR LF joins (UAX #29 GB3): Latin seams cost nothing.
 * Otherwise the segmenter judges `a`, the last scalar and the first scalar:
 * every rule that suppresses the boundary using context before the last
 * scalar (GB9c, GB11) needs that scalar to be Extend or ZWJ, which joins it
 * to the `a`; every other rule is decided by the pair. Three clusters thus
 * prove a boundary in any context. */
function seamMayJoin(segmenter: Intl.Segmenter, previous: string, next: string): boolean {
  const before = previous.charCodeAt(previous.length - 1);
  const after = next.charCodeAt(0);
  if (before < 0x300 && after < 0x300) return before === 0x0d && after === 0x0a;
  const trail = before >= 0xdc00 && before <= 0xdfff && previous.length > 1 ? 2 : 1;
  const lead = after >= 0xd800 && after <= 0xdbff && next.length > 1 ? 2 : 1;
  let clusters = 0;
  for (const _ of segmenter.segment(`a${previous.slice(-trail)}${next.slice(0, lead)}`)) clusters += 1;
  return clusters < 3;
}

/** The last UAX #29 grapheme cluster of one sequence's displayed text, found
 * only when a run seam that may join asks ({@link seamMayJoin}). Text
 * appended since the previous such seam is segmented once, from the start of
 * the cluster known then: a grapheme boundary depends only on the text
 * around it, so segmentation restarted at a true boundary finds the same
 * later boundaries. The work over a sequence is therefore linear: each
 * appended unit once, plus per seam a bounded context of the last cluster
 * ({@link CLUSTER_CONTEXT_UNITS}) and the next run up to the end of the
 * cluster crossing the seam. */
class SequenceTail {
  private cluster = '';
  private readonly pending: string[] = [];

  reset(): void {
    this.cluster = '';
    this.pending.length = 0;
  }

  append(display: string): void {
    if (display.length > 0) this.pending.push(display);
  }

  /** Length of the sequence's last cluster as of the latest {@link crossing}
   * that returned a positive count. */
  get clusterLength(): number { return this.cluster.length; }

  /** UTF-16 units at the start of `next` that belong to the sequence's last
   * grapheme cluster (0 when `next` starts a cluster), however long the
   * cluster. */
  crossing(next: string): number {
    graphemeSegmenter ??= typeof Intl !== 'undefined' && typeof Intl.Segmenter === 'function'
      ? new Intl.Segmenter(undefined, { granularity: 'grapheme' }) : null;
    if (!graphemeSegmenter || next.length === 0) return 0;
    const previous = this.pending.length > 0 ? this.pending[this.pending.length - 1]! : this.cluster;
    if (previous.length === 0 || !seamMayJoin(graphemeSegmenter, previous, next)) return 0;
    // The scan restarts inside the known cluster, at most its last
    // CLUSTER_CONTEXT_UNITS: a cluster that long is a run of extending
    // characters (UAX #29 GB9), whose continuation a restart inside it judges
    // alike. Each seam thus rescans a bounded context, however long the
    // cluster grows across runs.
    let cluster = this.cluster;
    if (cluster.length > CLUSTER_CONTEXT_UNITS) {
      cluster = cluster.slice(-CLUSTER_CONTEXT_UNITS);
      const first = cluster.charCodeAt(0);
      if (first >= 0xdc00 && first <= 0xdfff) cluster = cluster.slice(1);
    }
    const known = cluster + this.pending.join('');
    this.pending.length = 0;
    const seam = known.length;
    // `next` is read only up to the end of the cluster crossing the seam: a
    // bounded prefix first, doubled while that cluster reaches its end.
    let probed = 0;
    let start = 0;
    let end = 0;
    do {
      probed = Math.min(next.length, Math.max(CLUSTER_CONTEXT_UNITS, probed * 2));
      const lead = next.charCodeAt(probed - 1);
      if (probed < next.length && lead >= 0xd800 && lead <= 0xdbff) probed += 1;
      start = 0;
      end = 0;
      for (const part of graphemeSegmenter.segment(known + next.slice(0, probed))) {
        if (part.index >= seam) break;
        start = part.index;
        end = part.index + part.segment.length;
        if (end > seam) break;
      }
    } while (end === seam + probed && probed < next.length);
    this.cluster = known.slice(start);
    return end > seam ? end - seam : 0;
  }
}

/**
 * Library invariant: formatting-only source seams never enter the glyph/token
 * stream. ECMA-376 §§17.3.2.25/.19 separate the content container from its
 * resolved run formatting/kerning threshold; this normalization does not infer
 * another Office kerning switch or fitting allowance. Every layout consumer
 * uses the same concatenated sequence, so a separator pair exists exactly when
 * that separator is in the measured range, including emergency prefixes.
 *
 * Source ownership remains a separate UTF-16 map for comments, revisions and
 * field results. Tabs and real font/property changes remain semantic boundaries
 * in the ordinary segment builder. One key per source run, one join per sequence
 * and one source sweep keep normalization linear in input size.
 */
export function acquireTextSequences(
  runs: readonly ParagraphLayoutRun[],
  environment: LineLayoutEnvironment,
  displayText: (text: string, run: Extract<ParagraphTextBearingRun, { type: 'text' }>) => string,
): ReadonlyMap<number, TextSequence> {
  const sequences = new Map<number, TextSequence>();
  let start = -1;
  let key: string | undefined;
  let first: SequenceRun | undefined;
  // The last joined run and its scope: the w:val-free key of a seam is built
  // only when the languages on its two sides differ.
  let last: SequenceRun | undefined;
  let lastScope: number | string | null = null;
  const tail = new SequenceTail();
  let languages: { end: number; lang: string | undefined }[] | undefined;
  // Every joined run uniform: the ranges are each run's one language.
  let allUniform = true;
  let texts: string[] = [];
  let sources: TextSequenceSource[] = [];
  let optionalHyphens: TextSequence['optionalHyphens'][number][] = [];
  let offset = 0;
  const finish = () => {
    if (first && sources.length > 1) {
      const text = texts.join('');
      sequences.set(start, {
        run: Object.freeze({ ...first, text,
          ...(languages && languages.length > 1
            ? { westernLanguageRanges: Object.freeze(languages.map((range) => Object.freeze(range))),
              ...(allUniform ? { uniformLanguageRanges: true as const } : {}) }
            : {}),
          ...(first.typographyInput ? { typographyInput: Object.freeze({
            ...first.typographyInput, sourceText: text,
          }) } : {}),
        }),
        sources: Object.freeze(sources),
        optionalHyphens: Object.freeze(optionalHyphens),
      });
    }
    start = -1;
    first = undefined;
    last = undefined;
    languages = undefined;
    allUniform = true;
    texts = [];
    sources = [];
    optionalHyphens = [];
    offset = 0;
    tail.reset();
  };
  /** Language of the sequence's displayed unit at `position`. */
  const languageAt = (position: number): string | undefined => {
    let index = languages!.length - 1;
    while (index > 0 && languages![index - 1]!.end > position) index -= 1;
    return languages![index]!.lang;
  };
  const extendLanguage = (end: number, lang: string | undefined): void => {
    const current = languages!.at(-1)!;
    if (current.lang === lang) current.end = end;
    else languages!.push({ end, lang });
  };
  for (const [runIndex, run] of runs.entries()) {
    // ECMA-376 §17.13.5: deleted/moved-away content has no final-view glyphs or breaks.
    // It contributes neither formatting nor a shaping seam to visible neighbors;
    // markup view retains its own revision formatting and source ownership.
    const revisionKind = (run as { revision?: { kind?: string } }).revision?.kind;
    if (revisionIsOmitted(revisionKind, environment.showTrackedChanges)) continue;
    if (first && run.type === 'text' && 'optionalHyphen' in run && run.optionalHyphen) {
      // The unselected marker has no glyph and no shaping seam, even when
      // its conditional glyph uses a different font or paint style. Preserve
      // its source/format owner separately from the uninterrupted sequence.
      optionalHyphens.push(Object.freeze({ offset, runIndex, run }));
      sources.push(Object.freeze({ runIndex, start: offset, end: offset }));
      continue;
    }
    const visible = visibleTextRun(run, environment);
    const scopeKey = visible ? environment.layoutServices?.text.sourceScopeKey?.({
      text: displayText(visible.text, visible), fontSizePt: visible.fontSize,
      fonts: visible.fontSlots?.direct ?? { ascii: visible.fontFamily,
        highAnsi: visible.fontFamilyHighAnsi ?? visible.fontFamily,
        eastAsia: visible.fontFamilyEastAsia ?? visible.fontFamily,
        complexScript: visible.fontFamilyCs ?? visible.fontFamily },
      themeFonts: visible.fontSlots?.theme, themeFontPresence: visible.fontSlots?.themePresent,
      weight: (visible.rtl || visible.cs ? visible.boldCs : visible.bold) ? 700 : 400,
      style: (visible.rtl || visible.cs ? visible.italicCs : visible.italic) ? 'italic' : 'normal',
      complexScript: visible.rtl === true || visible.cs === true,
      fontHint: visible.fontHint, eastAsiaLanguage: visible.langEastAsia,
    }) : undefined;
    const scope = scopeKey === 'mixed' ? runIndex : scopeKey ?? null;
    const nextKey = visible ? JSON.stringify([formattingKey(visible, environment), scope]) : undefined;
    const display = visible ? displayText(visible.text, visible) : '';
    // §17.3.2.20 languages delimit sequences, except a seam inside one
    // grapheme cluster (UAX #29) whose sides differ only in w:val: splitting
    // there would separate a base from its variation selector, combining mark
    // or ZWJ continuation and change the cluster's glyph. Such a run joins the
    // sequence. Only the crossing cluster is affected: the units of the run
    // that complete it take the language of the cluster's first unit, and
    // every other cluster keeps the language of its own run
    // (SequenceLanguages). Source ranges stay per run. w:eastAsia and w:bidi
    // also select regional faces, so their seams keep delimiting sequences.
    let join = visible !== undefined && nextKey === key;
    let crossing = 0;
    // Judged against the previous run (whose seam this is), not the language
    // range in effect: a third run may carry the cluster's language again
    // while its seam with the second still differs in w:val.
    // A seam may also differ in every axis when both runs are uniform (one
    // language each): their whole language then transfers like w:val.
    const uniformSeam = visible !== undefined && last !== undefined && !join && uniformLanguage(visible)
      && uniformLanguage(last) && scope === lastScope
      && formattingKey(visible, environment, 'none') === formattingKey(last, environment, 'none');
    if (visible && first && last && (join
      ? rangeLanguageOf(visible) !== (languages ? languages.at(-1)!.lang : rangeLanguageOf(first))
      : uniformSeam || (scope === lastScope
        && formattingKey(visible, environment, 'nonWestern') === formattingKey(last, environment, 'nonWestern')))) {
      crossing = tail.crossing(display);
      if (crossing > 0) join = true;
    }
    if (!join) finish();
    key = nextKey;
    if (!visible) continue;
    if (!first) { start = runIndex; first = visible; }
    else if (!languages && crossing > 0) languages = [{ end: offset, lang: rangeLanguageOf(first) }];
    if (languages) {
      if (crossing > 0) extendLanguage(offset + crossing, languageAt(offset - tail.clusterLength));
      if (display.length > crossing) extendLanguage(offset + display.length, rangeLanguageOf(visible));
    }
    if (!uniformLanguage(visible)) allUniform = false;
    last = visible;
    lastScope = scope;
    texts.push(visible.text);
    tail.append(display);
    sources.push(Object.freeze({ runIndex, start: offset, end: offset + display.length }));
    offset += display.length;
  }
  finish();
  return sequences;
}
