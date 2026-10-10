import { revisionIsOmitted } from '../layout/revision-visibility.js';
import type { ParagraphLayoutRun, ParagraphTextBearingRun } from '../layout/text.js';
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
  // note marker, fitText unit, noBreakHyphen owner or upright tate-chu-yoko cell.
  if (r.optionalHyphen || r.ruby || r.noteRef || r.fitTextVal != null || r.noBreakRanges?.length
    || r.noBreakBefore || r.noBreakAfter || r.eastAsianVert) return undefined;
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
    langBidi: run.langBidi, langEastAsia: run.langEastAsia,
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
  let texts: string[] = [];
  let sources: TextSequenceSource[] = [];
  let optionalHyphens: TextSequence['optionalHyphens'][number][] = [];
  let offset = 0;
  const finish = () => {
    if (first && sources.length > 1) {
      const text = texts.join('');
      sequences.set(start, {
        run: Object.freeze({ ...first, text,
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
    texts = [];
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
    const nextKey = visible ? JSON.stringify([
      formattingKey(visible, environment), scopeKey === 'mixed' ? runIndex : scopeKey ?? null,
    ]) : undefined;
    if (!visible || nextKey !== key) finish();
    key = nextKey;
    if (!visible) continue;
    if (!first) { start = runIndex; first = visible; }
    texts.push(visible.text);
    const display = displayText(visible.text, visible);
    sources.push(Object.freeze({ runIndex, start: offset, end: offset + display.length }));
    offset += display.length;
  }
  finish();
  return sequences;
}
