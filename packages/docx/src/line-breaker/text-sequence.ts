import { revisionIsOmitted } from '../layout/revision-visibility.js';
import { graphemeClusterOffsets, isSymbolFontFamily } from '@silurus/ooxml-core';
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

function formattingKey(run: SequenceRun, environment: LineLayoutEnvironment): string {
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
    // No §17.3.2.20 language: a language seam is not a sequence boundary
    // (SequenceLanguages).
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

/** The languages of one stretch of a sequence's displayed text: the
 * §17.3.2.20 axes (w:val, w:eastAsia, w:bidi) and the script-class language
 * (ParagraphTextFacts.langScriptClass) of the run that owns it, as that run
 * carries them. */
export interface SequenceLanguageRange {
  /** UTF-16 end offset in the sequence's displayed text. */
  readonly end: number;
  readonly lang: string | undefined;
  readonly langEastAsia: string | undefined;
  readonly langBidi: string | undefined;
  readonly langScriptClass: string | undefined;
}

/**
 * The languages of a sequence whose runs do not all name the same ones.
 *
 * Runs that differ only in language are one sequence: a language is not
 * formatting that ends a word, a kerning pair's context or a grapheme
 * cluster. The sequence's displayed text is instead covered by contiguous
 * ranges (ascending `end`; the last ends with the text), each carrying the
 * whole language tuple of one run. Every range boundary is a grapheme cluster
 * boundary (UAX #29) of the whole displayed text: a cluster that spans runs
 * (a base and its combining mark, variation selector or ZWJ continuation, an
 * Indic conjunct) belongs to the run of its first unit, so it is shaped whole
 * in one language. The segment builder shapes each range apart, in its own
 * languages.
 *
 * Absent when every run names the same languages (always, in a document
 * without languages): the sequence run's own fields then hold for all of it.
 */
export interface SequenceLanguages {
  readonly languageRanges?: readonly SequenceLanguageRange[];
}

/** Identity of the languages `run` names, '' when it names none. The w:val
 * axis enters as its effective tag (canvasLanguageTag: x-none, ill-formed and
 * absent tags are one state), and not at all on a run forced onto the
 * complex-script slot (§17.3.2.7 w:cs, §17.3.2.30 w:rtl), which never uses
 * it. w:eastAsia and w:bidi stay as written: regional font selection and the
 * bidi digit rule read them raw. */
function languageKey(run: SequenceRun): string {
  if (run.lang === undefined && run.langEastAsia === undefined && run.langBidi === undefined
    && run.langScriptClass === undefined) return '';
  const western = run.cs === true || run.rtl === true ? undefined : canvasLanguageTag(run.lang);
  if (western === undefined && run.langEastAsia === undefined && run.langBidi === undefined
    && run.langScriptClass === undefined) return '';
  return JSON.stringify([western ?? null, run.langEastAsia ?? null, run.langBidi ?? null,
    run.langScriptClass ?? null]);
}

/** Whether a grapheme cluster boundary at `offset` of `text` can be
 * suppressed in some context, judged from the two scalars around it alone.
 * Below U+0300 the only Grapheme_Cluster_Break values are Control, CR, LF
 * and Other, so only CR LF joins (UAX #29 GB3): Latin seams cost nothing.
 * Otherwise `a`, the scalar before and the scalar after are segmented: every
 * rule that suppresses the boundary using context before the preceding
 * scalar (GB9c, GB11, GB12/GB13) needs that scalar to be Extend or ZWJ, which
 * joins it to the `a`, or a Regional_Indicator pair, which is one cluster;
 * every other rule is decided by the pair. Three clusters thus prove a
 * boundary in any context. A possible join is then decided by segmenting the
 * whole text ({@link languageRanges}), never by this test. */
function seamMayJoin(text: string, offset: number): boolean {
  const before = text.charCodeAt(offset - 1);
  const after = text.charCodeAt(offset);
  if (before < 0x300 && after < 0x300) return before === 0x0d && after === 0x0a;
  const trail = before >= 0xdc00 && before <= 0xdfff && offset > 1 ? 2 : 1;
  const lead = after >= 0xd800 && after <= 0xdbff && offset + 1 < text.length ? 2 : 1;
  return graphemeClusterOffsets(`a${text.slice(offset - trail, offset + lead)}`).length < 2;
}

/** Whether any font slot of `run` names a symbol-encoded font. */
function namesSymbolFont(run: SequenceRun): boolean {
  const slots = [run.fontSlots?.direct, run.fontSlots?.theme].flatMap((fonts) => fonts
    ? [fonts.ascii, fonts.highAnsi, fonts.eastAsia, fonts.complexScript] : []);
  return [run.fontFamily, run.fontFamilyHighAnsi, run.fontFamilyEastAsia, run.fontFamilyCs, ...slots]
    .some((family) => isSymbolFontFamily(family));
}

/** One run's stretch of a sequence: where its displayed text starts, with
 * its {@link languageKey}. Recorded only where the key changes. */
interface LanguageStretch {
  start: number;
  readonly key: string;
  readonly run: SequenceRun;
}

/**
 * The {@link SequenceLanguages} ranges of a sequence's displayed `text`,
 * given its runs' stretches in order.
 *
 * A stretch boundary inside a grapheme cluster moves forward to the end of
 * that cluster, so the cluster stays with the run of its first unit; a
 * stretch that lies wholly inside such a cluster covers nothing. The clusters
 * are those of the whole text, found by segmenting it once (and only when a
 * boundary may fall inside one, {@link seamMayJoin}): every boundary is then
 * judged in its full UAX #29 context, however long the cluster or the text.
 * The work is linear in the text.
 */
function languageRanges(
  text: string,
  stretches: LanguageStretch[],
  firstKey: string,
): readonly SequenceLanguageRange[] | undefined {
  if (stretches.some((stretch) => stretch.start > 0 && stretch.start < text.length
    && seamMayJoin(text, stretch.start))) {
    const clusters = graphemeClusterOffsets(text);
    let cluster = 0;
    for (const stretch of stretches) {
      while (cluster < clusters.length && clusters[cluster]! < stretch.start) cluster += 1;
      // The first cluster start at or after the stretch (else the text end).
      if (stretch.start > 0) stretch.start = clusters[cluster] ?? text.length;
    }
  }
  const ranges: { end: number; key: string; run: SequenceRun }[] = [];
  for (const [index, stretch] of stretches.entries()) {
    const end = stretches[index + 1]?.start ?? text.length;
    if (end <= stretch.start) continue;
    const previous = ranges.at(-1);
    if (previous?.key === stretch.key) previous.end = end;
    else ranges.push({ end, key: stretch.key, run: stretch.run });
  }
  // One range in the first run's languages is the sequence run itself.
  if (ranges.length === 0 || (ranges.length === 1 && ranges[0]!.key === firstKey)) return undefined;
  return Object.freeze(ranges.map(({ end, run }) => Object.freeze({
    end, lang: run.lang, langEastAsia: run.langEastAsia, langBidi: run.langBidi,
    langScriptClass: run.langScriptClass,
  })));
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
 *
 * §17.3.2.20 languages are not such boundaries: runs differing only in
 * language are one sequence that records which stretch is in which languages
 * ({@link SequenceLanguages}). A sequence whose runs agree records nothing,
 * so text without languages is acquired exactly as if none existed.
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
  let firstLanguages = '';
  // The runs' language stretches, recorded from the first run whose
  // languages differ from the sequence's first: a sequence of one language
  // state (every sequence of a document without languages) records nothing.
  let stretches: LanguageStretch[] | undefined;
  let texts: string[] = [];
  let displays: string[] = [];
  let sources: TextSequenceSource[] = [];
  let optionalHyphens: TextSequence['optionalHyphens'][number][] = [];
  let offset = 0;
  const finish = () => {
    if (first && sources.length > 1) {
      const text = texts.join('');
      const ranges = stretches ? languageRanges(displays.join(''), stretches, firstLanguages) : undefined;
      sequences.set(start, {
        run: Object.freeze({ ...first, text,
          ...(ranges ? { languageRanges: ranges } : {}),
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
    stretches = undefined;
    texts = [];
    displays = [];
    sources = [];
    optionalHyphens = [];
    offset = 0;
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
    const languages = visible ? languageKey(visible) : '';
    const nextKey = visible ? JSON.stringify([
      formattingKey(visible, environment), scopeKey === 'mixed' ? runIndex : scopeKey ?? null,
      // Symbol-encoded text (§17.3.3.30) is font codes, not text of a
      // language, and its conversion reads the run's own w:eastAsia (the
      // §17.3.2.26 slot of each code): such runs keep their languages as a
      // sequence boundary.
      ...(languages !== '' && namesSymbolFont(visible) ? [languages] : []),
    ]) : undefined;
    if (!visible || nextKey !== key) finish();
    key = nextKey;
    if (!visible) continue;
    const display = displayText(visible.text, visible);
    if (!first) {
      start = runIndex;
      first = visible;
      firstLanguages = languages;
    } else if (stretches) {
      if (languages !== stretches.at(-1)!.key) stretches.push({ start: offset, key: languages, run: visible });
    } else if (languages !== firstLanguages) {
      stretches = [{ start: 0, key: firstLanguages, run: first }, { start: offset, key: languages, run: visible }];
    }
    texts.push(visible.text);
    displays.push(display);
    sources.push(Object.freeze({ runIndex, start: offset, end: offset + display.length }));
    offset += display.length;
  }
  finish();
  return sequences;
}
