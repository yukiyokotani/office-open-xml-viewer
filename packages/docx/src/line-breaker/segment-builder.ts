import { revisionIsOmitted } from '../layout/revision-visibility.js';
import { type TextBreakWindow } from './text-break-window.js';
import type { DocxTextRun, FieldRun } from '../types';
import { acquireTextSequences } from './text-sequence.js';
import type { HyperlinkTarget, ResolvedFontMetric } from '@silurus/ooxml-core';
import {
  DEFAULT_KINSOKU_RULES,
  isUax14NoBreakPair,
  lineBreakClass,
  containsSeaScript,
  graphemeClusterOffsets,
  isSymbolFontFamily,
  symbolTextToUnicodeSegments,
} from '@silurus/ooxml-core';
import { groupFitTextRegions, type FitTextRun } from '../fit-text.js';
import { mathFallbackText } from '../layout/math-fallback-text.js';
import type { FontResolution } from '../layout/font-service.js';
import type {
  ParagraphLayoutSource,
  ParagraphLayoutRun,
  ParagraphTextBearingRun,
  TextShapeRequest,
  TextShapeSpan,
} from '../layout/text.js';
import { calcEffectiveFontPx, EAST_ASIAN_RE, assertTextShapeRunContext, independentTextShapeRequest, sliceTextShapeRequest, registeredLatinMarkGraphemeCandidate, registeredLatinSlotRunCandidate, registeredLatinSingleSeamCandidate } from '../layout/text.js';
import {
  referenceFontAverageWidthRatio,
  referenceFontLineMetrics,
} from '../reference-font-line-metrics.js';
import {
  wordKerningApplies,
  wordDocumentCharacterCompressionApplies,
  wordJapanesePunctuationRetainedExtentPt,
  wordCompressedSpaceLineFitApplies,
  wordSourceRunSpaceContinuesSequence,
  wordBalancedConsecutiveSpaceCellApplies,
  wordBalancedLinesAndCharsGridDeltaFactor,
  wordExternalLinkSyntaxBreakOffsets,
} from '../layout/line-compatibility.js';
import { type LayoutSeg, type LayoutTextSeg, type LineLayoutEnvironment } from './model.js';
import { charScaleFactor, retainHorizontalPunctuationInkClearance } from './advance.js';
import {
  COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION,
  characterSpacingControlCompresses,
  findNearbyFontSize,
  formatNoteNumber,
  hasCJKBreakOpportunity,
  isRtlBidiLang,
  resolveFieldText,
  splitByEastAsia,
  splitDigitGroups,
  splitSmallCapsCase,
  splitTextForLayout,
} from './text-runs.js';
import {
  indexedFontMetrics,
  mayUseAuthoredReferenceVerticalMetric,
  mayUseExactLocalReferenceWidthMetric,
  selectResourceAverageWidthRatio,
  selectResourceMetric,
  selectedFontLineMetric,
  type MetricTupleIndex,
} from './font-metrics.js';

/** Resolve §17.3.2.14 region geometry after raw Canvas advances are available.
 *  Region membership was fixed over tab-delimited source fragments in
 *  {@link buildSegments}; width and code-point count are deliberately derived
 *  here from the emitted segments so script/case transformations cannot create
 *  a second source of truth. */
export function resolveFitTextSegments(
  segments: LayoutTextSeg[],
  scale: number,
  measureNaturalWidthPx: (segment: LayoutTextSeg) => number,
): void {
  const regionSegments = new Map<number, LayoutTextSeg[]>();
  for (const segment of segments) {
    if (segment.fitTextRegionIndex === undefined) continue;
    const members = regionSegments.get(segment.fitTextRegionIndex) ?? [];
    members.push(segment);
    regionSegments.set(segment.fitTextRegionIndex, members);
  }

  for (const members of regionSegments.values()) {
    const first = members.find((segment) => segment.fitTextVal !== undefined);
    if (!first || first.fitTextVal === undefined) continue;

    let naturalWidthPx = 0;
    let charCount = 0;
    for (const segment of members) {
      naturalWidthPx += measureNaturalWidthPx(segment) * charScaleFactor(segment);
      charCount += [...segment.text].length;
    }

    const resolved = groupFitTextRegions(
      [
        {
          fitTextValTwips: first.fitTextVal,
          charCount,
          naturalWidthPx,
        },
      ],
      scale,
    )[0];
    if (!resolved) continue;
    members.forEach((segment, index) => {
      segment.fitTextPerGapPx = resolved.perGapPx;
      segment.fitTextTrailingPadPx =
        index === members.length - 1 ? resolved.trailingPadPx : undefined;
      segment.fitTextRegionStart = index === 0 ? true : undefined;
      segment.fitTextRegionEnd = index === members.length - 1 ? true : undefined;
    });
  }
}

export interface SegmentBuildContext {
  readonly runs: readonly ParagraphLayoutRun[];
  readonly environment: LineLayoutEnvironment;
  readonly segs: LayoutSeg[];
  readonly fitTextFragmentEntryByKey: Map<string, number>;
  readonly fitTextRegionByEntry: Map<number, number>;
  readonly selectedMetric: (
    selected: FontResolution | undefined,
    probeText: string,
  ) => ResolvedFontMetric | undefined;
  readonly selectedAverageWidth: (
    selected: FontResolution | undefined,
    probeText: string,
  ) => number | undefined;
}

/** ECMA-376 §17.3.2.5/§17.3.2.33 change displayed case, not run ownership.
 * §17.3.3.30 symbol normalization may also expand a scalar to a surrogate
 * pair. Scope and offsets use the complete display string after both changes;
 * small-caps size pieces and tabs cannot create independent Arabic proof.
 * Hidden runs are omitted by the parser (§17.3.2.41), not merged into this run.
 */
function transformedRunText(
  text: string,
  run: Extract<ParagraphLayoutSource['runs'][number], { type: 'text' | 'field' }>,
  environment: LineLayoutEnvironment,
): string {
  const r: ParagraphTextBearingRun = run;
  const display = run.allCaps || run.smallCaps ? text.toUpperCase() : text;
  const map = (text: string, family: string | null | undefined) => isSymbolFontFamily(family)
    ? symbolTextToUnicodeSegments(text, family).map((part) => part.text).join('') : text;
  if (r.rtl || r.cs) return map(display, r.fontFamilyCs ?? run.fontFamily);
  if (environment.layoutServices?.text) return map(display, run.fontFamily);
  return splitByEastAsia(display).map((part) => map(part.text,
    part.ea ? r.fontFamilyEastAsia ?? run.fontFamily : run.fontFamily)).join('');
}

export function appendTextPiece(
  state: SegmentBuildContext,
  text: string,
  base: Extract<ParagraphLayoutSource['runs'][number], { type: 'text' | 'field' }>,
  vertAlign: 'super' | 'sub' | null,
  sourceRunIndex: number,
  fullRunContext: Readonly<{ text: string; offset: number }>,
  sourceFragmentIndex?: number,
  joinPreviousRun = false,
): void {
  const {
    runs,
    environment,
    segs,
    fitTextFragmentEntryByKey,
    fitTextRegionByEntry,
    selectedMetric,
    selectedAverageWidth,
  } = state;
  const r: ParagraphTextBearingRun = base;
  const overflowPunctuationEastAsianRun = EAST_ASIAN_RE.test(text) ? true : undefined;
  const acquiredTypography = (
    r as ParagraphTextBearingRun &
      Readonly<{
        typographyInput?: import('../layout/typography-input.js').RunTypographyAcquisitionInput;
      }>
  ).typographyInput;
  // ECMA-376 §17.16.18 stores a complex field's instruction/result across
  // several physical runs. The parser rebuilds recomputed PAGE/NUMPAGES as a
  // single FieldRun, while its complete effective §17.3.2 run properties live
  // on the immutable typography acquisition sidecar. Consume those effective
  // facts exactly like an ordinary text run so the field result does not lose
  // baseline position or the core character-metric axes used by measure/paint.
  const acquiredValue = <T>(
    value: import('../layout/typography-input.js').TypographyValueInput<T> | undefined,
    fallback: T | undefined,
  ): T | undefined => (value?.status === 'valid' && value.value !== null ? value.value : fallback);
  const effectiveVertAlign =
    acquiredValue(acquiredTypography?.verticalAlign, vertAlign ?? undefined) ?? null;
  const effectivePosition = acquiredValue(acquiredTypography?.positionPt, r.position);
  const effectiveCharacterSpacing = acquiredTypography?.characterSpacingPt ?? r.charSpacing;
  // ECMA-376 §17.3.2.35 gives an authored run an explicit character pitch.
  // Word observation: a positive `w:spacing` owns that expanded pitch and suppresses
  // document-level §17.15.1.18 punctuation whitespace compression for the
  // run. Combining both adjustments collapses consecutive Japanese closing
  // punctuation even though Word preserves the authored spacing.
  const documentCharacterCompressionApplies =
    wordDocumentCharacterCompressionApplies(effectiveCharacterSpacing);
  const effectiveCharacterScale = acquiredTypography?.characterScale ?? r.charScale;
  // WORD_KERN_THRESHOLD_AUTHORITY: keep the resolved value, including zero,
  // so every measurement/paint consumer uses the same threshold decision.
  const effectiveKerningThreshold = acquiredTypography?.kerningThresholdPt ?? r.kerning;
  const effectiveSnapToGrid = acquiredTypography?.snapToGrid ?? r.snapToGrid;
  // §17.3.2.33 small caps are sized per character: lowercase LETTERS render two
  // points smaller, uppercase letters and non-alphabetic characters at the full
  // run size. `reduced` (set per case-piece in the loop below) carries that onto
  // each emitted segment; calcEffectiveFontPx shrinks only the reduced ones.
  // allCaps (§17.3.2.5) and non-caps runs are a single, non-reduced piece.

  // Ruby annotation rides with the WHOLE base text (typically 1-2 chars).
  // Splitting on word boundaries would lose the association, so attach
  // the annotation only to the first emitted segment.
  const baseRuby = r.ruby;
  const ruby = baseRuby
    ? {
        text: baseRuby.text,
        fontSizePt: baseRuby.fontSizePt,
        ...(baseRuby.hpsRaisePt != null ? { hpsRaisePt: baseRuby.hpsRaisePt } : {}),
      }
    : undefined;
  const revision = r.revision;
  const rtl = r.rtl === true ? true : undefined;
  const fitTextFragmentEntryIndex =
    sourceFragmentIndex === undefined
      ? undefined
      : fitTextFragmentEntryByKey.get(`${sourceRunIndex}:${sourceFragmentIndex}`);
  const fitTextRegionIndex =
    fitTextFragmentEntryIndex === undefined
      ? undefined
      : fitTextRegionByEntry.get(fitTextFragmentEntryIndex);

  // IX1 — resolve the run's hyperlink target ONCE (§17.16.22 external URL /
  // §17.16.23 internal anchor). An external URL (`r.hyperlink`) wins over the
  // internal `w:anchor` when both are present, matching the parser's rule. A
  // FieldRun carries neither field, so the `as DocxTextRun` guards yield
  // undefined. Purely a callback payload — it does not touch measurement.
  const hyperlink: HyperlinkTarget | undefined = r.hyperlink
    ? { kind: 'external', url: r.hyperlink }
    : r.hyperlinkAnchor
      ? { kind: 'internal', ref: r.hyperlinkAnchor }
      : undefined;

  // ECMA-376 §17.3.2.26 content classification. w:rtl/w:cs selects the cs
  // axis except for a character assigned eastAsia while rFonts@hint=eastAsia;
  // that protected span keeps the non-cs East Asian formatting axis.
  // NOTE rFonts@cs (fontFamilyCs) alone is just a font SLOT and must NOT
  // force cs: a Latin run can carry cstheme and szCs while still using w:sz.
  const forceCs = r.rtl === true || r.cs === true;

  // Complex-script (cs) formatting sources. SIZE (§17.3.2.39 szCs) and TYPEFACE
  // (§17.3.2.26 rFonts@cs) fall back to their Latin counterpart when absent —
  // the parser resolves szCs through the full style chain, mirroring a
  // directly-set `w:sz` per §17.3.2.18. But BOLD (§17.3.2.3 bCs) and ITALIC
  // (§17.3.2.17 iCs) are independent toggles: absent `bCs`/`iCs` defaults off
  // and must not inherit Latin-axis `w:b`/`w:i`, which govern only non-complex
  // content.
  const csFontSize = r.fontSizeCs ?? base.fontSize;
  const csFontFamily = r.fontFamilyCs ?? base.fontFamily;
  const highAnsiFontFamily = r.fontFamilyHighAnsi ?? base.fontFamily;
  const csBold = r.boldCs ?? false;
  const csItalic = r.italicCs ?? false;

  // ECMA-376 §17.3.2.26 eastAsia axis. Within a non-complex-script slice, CJK
  // code points take the eastAsia face while Latin/digits keep the ascii face
  // (`base.fontFamily`). Only `DocxTextRun` carries the axis; absent (field
  // runs / single-axis parser output) ⇒ fall back to ascii. Text-box runs feed
  // this same builder (via `shapeRunToDocRun`), so a text box's per-script face
  // is picked here too. Bold/italic/size are NOT axis-specific here — eastAsia
  // shares the Latin (non-cs) toggles, so only the family differs.
  const eaFontFamily = r.fontFamilyEastAsia ?? base.fontFamily;

  // `word-rtl-complex-script-european-digits-an`: use the bidi language's
  // primary subtag when present, otherwise fall back to an rtl-marked run.
  const digitsAsAN = (forceCs || Boolean(r.rtl)) && isRtlBidiLang(r.langBidi, Boolean(r.rtl));

  const emissionState: SegmentEmissionState = {
    base,
    joinPreviousRun,
    environment,
    segs,
    selectedMetric,
    selectedAverageWidth,
    r,
    effectiveVertAlign,
    effectivePosition,
    effectiveCharacterSpacing,
    documentCharacterCompressionApplies,
    effectiveCharacterScale,
    effectiveKerningThreshold,
    effectiveSnapToGrid,
    ruby,
    revision,
    fitTextFragmentEntryIndex,
    fitTextRegionIndex,
    csFontSize,
    csFontFamily,
    highAnsiFontFamily,
    csBold,
    csItalic,
    eaFontFamily,
    digitsAsAN,
    sourceRunIndex,
    sourceFragmentIndex,
    hyperlink,
    rtl,
    overflowPunctuationEastAsianRun,
    reduced: false,
    firstSeg: true,
    gluePending: false,
    scopeContext: {
      text: fullRunContext.text,
      cursor: fullRunContext.offset,
    },
  };
  // True while the next emitted segment should be GLUED to the previous one
  // (a small-caps case-piece that continues the same word). Consumed by the
  // first pushSeg of the piece so only that segment carries joinPrev.

  // Script slot for an emitted segment (§17.3.2.26): 'cs' = complex-script
  // (Arabic/Hebrew/...), 'ea' = East-Asian (CJK → eastAsia face), 'latin' =
  // Latin/digits/neutral (ascii face). Each segment stays SINGLE-FONT — one
  // family for its whole `.text` — so the measure==draw / docGrid char-grid
  // invariant holds and the draw loop needs no per-segment font switching.
  const pushSeg = (
    text: string,
    cs: boolean,
    fontFamily: string | null,
    authoritativeSpan?: TextShapeSpan,
    compressCharacterWhitespace = false,
    mappedSymbolUnicode = false,
  ): void =>
    pushSegmentPiece(
      emissionState,
      text,
      cs,
      fontFamily,
      authoritativeSpan,
      compressCharacterWhitespace,
      mappedSymbolUnicode,
    );
  const emit = (word: string, slot: 'cs' | 'ea' | 'latin') => {
    const cs = slot === 'cs';
    const fontFamily =
      slot === 'cs' ? csFontFamily : slot === 'ea' ? eaFontFamily : base.fontFamily;
    // ECMA-376 §17.3.2.26 + §17.3.3.30: a run whose rFonts axis is Symbol or
    // Wingdings stores glyphs as the FONT's own (private) code points — Word
    // commonly in the PUA (U+F020–U+F0FF). Those render as tofu in any
    // fallback face, so normalize each character to its Unicode equivalent
    // (core `symbolTextToUnicodeSegments`, the same table the list marker uses
    // via `symbolFontToUnicode`). The string is split at mapped/unmapped
    // boundaries: a MAPPED run is drawn in a generic fallback (fontFamily=null
    // → sans tail with the dingbat glyphs; keeping the symbol family would let
    // an installed Symbol/Wingdings re-interpret the Unicode code point as the
    // WRONG glyph), while an UNMAPPED run keeps the symbol family so a host
    // that ships Symbol/Wingdings still draws its native glyph. Done once at
    // build time so measure==draw (the seg.text is never transformed later).
    if (isSymbolFontFamily(fontFamily)) {
      for (const part of symbolTextToUnicodeSegments(word, fontFamily)) {
        pushSeg(part.text, cs, part.mapped ? null : fontFamily, undefined, false, part.mapped);
      }
      return;
    }
    pushSeg(word, cs, fontFamily);
  };

  // A non-complex-script slice still mixes scripts at the CJK boundary: emit
  // its maximal CJK runs on the 'ea' (eastAsia) slot and the rest on 'latin'
  // (ascii). Keeps each emitted segment single-font (so a serif ascii digit
  // sits next to a gothic eastAsia title) without changing the cs path.
  const emitNonCs = (slice: string) => {
    if (environment.layoutServices?.text) {
      // `w:sym` is parsed as a one-run private-encoding character carrying
      // its own Symbol/Wingdings family (§17.3.3.30). Keep normalization in
      // front of the service-backed script splitter as well as the legacy
      // path; otherwise a PUA code point such as Symbol F0B0 reaches Canvas
      // unchanged and renders as tofu.
      if (isSymbolFontFamily(base.fontFamily)) {
        emit(slice, 'latin');
        return;
      }
      pushSeg(slice, false, base.fontFamily);
      return;
    }
    for (const part of splitByEastAsia(slice)) emit(part.text, part.ea ? 'ea' : 'latin');
  };

  // Small caps split the run into full-size (uppercase-origin / non-cased) and
  // reduced (lowercase-origin) case-pieces; everything else is one piece. Each
  // piece is still UPPERCASED for display (allCaps or smallCaps), and `reduced`
  // drives its segments' size — see splitSmallCapsCase / calcEffectiveFontPx.
  const casePieces = base.smallCaps ? splitSmallCapsCase(text) : [{ text, reduced: false }];
  let prevPieceText = '';
  for (const piece of casePieces) {
    emissionState.reduced = piece.reduced;
    // Glue this piece's FIRST segment to the previous piece when they continue
    // the same word (the previous piece did not end at a space) — so a
    // small-caps word's full-cap initial and reduced remainder stay on one line.
    emissionState.gluePending = prevPieceText.length > 0 && !/\s$/.test(prevPieceText);
    prevPieceText = piece.text;
    const displayText = base.allCaps || base.smallCaps ? piece.text.toUpperCase() : piece.text;
    for (const word of splitTextForLayout(displayText)) {
      if (forceCs) {
        // When the run's digits are AN-classified, split a token into maximal
        // digit-groups and the surrounding separators so the per-line bidi pass
        // (which reorders at SEGMENT granularity) can place the groups in Word's
        // order — e.g. "28-02-2026" → segments [28][-][02][-][2026] reordered to
        // 2026-02-28. Canvas only reorders WITHIN a fillText using EN semantics,
        // so a single-segment date would otherwise stay 28-02-2026.
        if (digitsAsAN) {
          for (const slice of splitDigitGroups(word)) emit(slice, 'cs');
        } else {
          emit(word, 'cs');
        }
      } else {
        // Mixed Arabic+Latin word (no w:rtl / w:cs): split at script boundaries
        // so each side gets its own (cs vs Latin) size and typeface; the non-cs
        // side then sub-splits at CJK boundaries for the eastAsia face.
        // ECMA-376 §17.3.2.26 selects the cs axis only when w:cs/w:rtl
        // forces the run. Arabic/Hebrew code points in an ordinary run stay
        // on ascii/hAnsi; the text service performs the remaining grapheme-
        // safe East Asian slot split.
        emitNonCs(word);
      }
    }
  }
}

/** Resolve source-level no-break ownership and adjacent UAX14/fitText seams
 * after every display segment has been emitted. */
/** Linear projection onto immutable offset windows. A real source/font seam
 * is breakable only when the complete text proved that boundary legal. */
function projectTextBreakOffsets(group: readonly LayoutTextSeg[], offsets: readonly number[]): void {
  let cursor = 0, index = 0;
  for (const segment of group) {
    const end = cursor + segment.text.length;
    if (offsets[index] === cursor) {
      segment.joinPrev = undefined;
      segment.explicitBreakBefore = true;
      index++;
    }
    const start = index;
    while (index < offsets.length && offsets[index]! < end) index++;
    if (index > start) {
      const window: TextBreakWindow = { offsets, start, end: index, origin: cursor };
      segment.explicitBreaks = window;
    }
    cursor = end;
  }
}

export function finalizeBuiltSegments(
  runs: readonly ParagraphLayoutRun[],
  environment: LineLayoutEnvironment,
  segs: LayoutSeg[],
): void {
  // Project acquisition-owned no-break ranges through the display case
  // transform and onto the single-font layout segments produced above.
  projectNoBreakRanges(runs, segs);

  // ECMA-376 §17.3.3.18 permits ordinary U+002D breaks, including numeric
  // identifiers. This DOCX tailoring overrides LB25 only inside a word:
  // signed numbers after an opening bracket/start and LB21a's Hebrew-to-other
  // boundary retain their Unicode protection. LB9 combining bases and grapheme
  // boundaries use the complete displayed text, irrespective of font/run seams.
  // This acquires opportunities, without splitting or merging shaped segments.
  // Atomic cells and external-link syntax have their own owners below.
  const rubyOwnedRuns = new Set(segs.filter(segment => 'text' in segment && segment.ruby)
    .map(segment => segment.sourceRunIndex));
  for (let start = 0; start < segs.length;) {
    const ordinary = (s: LayoutSeg): s is LayoutTextSeg => 'text' in s
      && s.hyperlink?.kind !== 'external' && s.fitTextRegionIndex === undefined
      && !s.ruby && !s.tateChuYoko && !rubyOwnedRuns.has(s.sourceRunIndex);
    if (!ordinary(segs[start]!)) { start++; continue; }
    let end = start;
    const group: LayoutTextSeg[] = [];
    while (end < segs.length && ordinary(segs[end]!)) group.push(segs[end++] as LayoutTextSeg);
    const text = group.map(segment => segment.text).join('');
    if (text.includes('-')) {
      const boundaries = [...graphemeClusterOffsets(text), text.length];
      let boundaryIndex = 0;
      // Authored noBreakHyphen owns its whole extended grapheme. Its XML
      // range can end at HY before a following mark, so test that original
      // edge as well as the displayed cluster end below.
      const protectedOffsets = new Set<number>();
      let origin = 0;
      for (const segment of group) {
        for (const range of segment.noBreakRanges ?? []) {
          protectedOffsets.add(origin + range.start);
          protectedOffsets.add(origin + range.end);
        }
        if (segment.hardJoinPrev) protectedOffsets.add(origin);
        origin += segment.text.length;
      }
      const offsets: number[] = [];
      const prohibitedHyphenEnds = new Set<number>();
      let cursor = 0;
      let baseClass: ReturnType<typeof lineBreakClass> | undefined;
      const wordClass = (c: typeof baseClass) => c === 'AL' || c === 'HL' || c === 'NU';
      for (const scalar of text) {
        const cls = lineBreakClass(scalar.codePointAt(0)!);
        const offset = cursor + scalar.length;
        while (boundaries[boundaryIndex]! <= cursor) boundaryIndex++;
        const clusterEnd = boundaries[boundaryIndex]!;
        if (scalar === '-' && clusterEnd < text.length) {
          prohibitedHyphenEnds.add(clusterEnd);
          const rightClass = lineBreakClass(text.codePointAt(clusterEnd)!);
          if (wordClass(baseClass) && wordClass(rightClass) && !(baseClass === 'HL' && rightClass !== 'HL')
            && text.codePointAt(clusterEnd - 1) !== 0x200d
            && !protectedOffsets.has(offset) && !protectedOffsets.has(clusterEnd)) {
            offsets.push(clusterEnd);
            prohibitedHyphenEnds.delete(clusterEnd);
          }
        }
        // LB9: CM/ZWJ inherit a preceding base; the grapheme check above still
        // places any hyphen opportunity after its complete extended grapheme.
        if (cls !== 'CM' && cls !== 'ZWJ') baseClass = cls;
        cursor = offset;
      }
      projectTextBreakOffsets(group, Object.freeze(offsets));
      // A rendering seam after a rejected hyphen is not an alternate route
      // around the complete-text classifier (including Hebrew and signs).
      let seam = 0;
      for (const segment of group) {
        if (prohibitedHyphenEnds.has(seam)) segment.joinPrev = true;
        seam += segment.text.length;
      }
    }
    start = end;
  }

  // Project the registered `word-external-link-syntax-breaks` opportunities
  // across the complete semantic link and all formatting seams first, then
  // distribute them onto the existing segments.
  // Segments stay intact unless a real overflow selects one of those offsets,
  // preserving contextual shaping, decoration geometry, and paint identity on
  // lines that do not wrap.
  for (let groupStart = 0; groupStart < segs.length;) {
    const first = segs[groupStart];
    if (!('text' in first) || first.hyperlink?.kind !== 'external') {
      groupStart += 1;
      continue;
    }
    const target = first.hyperlink.url;
    let groupEnd = groupStart;
    const group: LayoutTextSeg[] = [];
    while (groupEnd < segs.length) {
      const candidate = segs[groupEnd];
      if (
        !('text' in candidate) ||
        candidate.hyperlink?.kind !== 'external' ||
        candidate.hyperlink.url !== target
      )
        break;
      group.push(candidate);
      groupEnd += 1;
    }
    const groupText = group.map((segment) => segment.text).join('');
    const protectedOffsets = new Set<number>();
    let cursor = 0;
    for (const segment of group) {
      for (const range of segment.noBreakRanges ?? []) {
        const offsets = [range.start, range.end];
        for (const offset of offsets) {
          protectedOffsets.add(cursor + offset);
        }
      }
      cursor += segment.text.length;
    }
    const legalOffsets = new Set<number>();
    for (const match of groupText.matchAll(/\S+/gu)) {
      const token = match[0];
      const tokenStart = match.index;
      const graphemeBoundaries = new Set(
        graphemeClusterOffsets(token).map((offset) => tokenStart + offset),
      );
      const tokenProtected = new Set(
        [...protectedOffsets]
          .filter((offset) => offset > tokenStart && offset <= tokenStart + token.length)
          .map((offset) => offset - tokenStart),
      );
      const tokenGraphemes = new Set([...graphemeBoundaries].map((offset) => offset - tokenStart));
      for (const offset of wordExternalLinkSyntaxBreakOffsets(
        token,
        tokenGraphemes,
        tokenProtected,
      ))
        legalOffsets.add(tokenStart + offset);
    }
    if (legalOffsets.size === 0) {
      groupStart = groupEnd;
      continue;
    }
    projectTextBreakOffsets(group, Object.freeze([...legalOffsets].sort((a, b) => a - b)));
    groupStart = groupEnd;
  }

  if (environment.balanceSingleByteDoubleByteWidth) {
    // ECMA-376 §17.15.3.3 normatively requests a 1:2 SBCS/DBCS width balance,
    // but does not define how a proportional inter-word separator becomes a
    // fixed-pitch half-width cell. The registered Word observation limits that
    // projection to two-or-more explicitly authored U+0020 spaces. One normal
    // separator remains at its natural proportional advance.
    const metricCache = new Map<string, number>();
    const adjustmentFor = (segment: LayoutTextSeg): number | undefined => {
      const service = segment.textLayoutService;
      const request = segment.textShapeRequest;
      if (!service || !request) return undefined;
      const effectiveFontSizePt = calcEffectiveFontPx(segment, 1);
      const key = [
        service.fingerprint,
        segment.fontRoute?.fingerprint ?? 'implicit-latin',
        segment.eaFloorRoute?.fingerprint ?? 'implicit-east-asia',
        effectiveFontSizePt,
        segment.bold ? 700 : 400,
        segment.italic ? 'italic' : 'normal',
        segment.kerning ?? 'none',
      ].join('|');
      const cached = metricCache.get(key);
      if (cached !== undefined) return cached;
      const naturalSpace = service.shape({
        ...independentTextShapeRequest(request, ' '),
        fontSizePt: effectiveFontSizePt,
        measure: true,
        clusterGeometry: false,
      }).advancePt;
      const ideographicCell = service.shape({
        ...independentTextShapeRequest(request, '\u4e00'),
        fontSizePt: effectiveFontSizePt,
        fontHint: 'eastAsia',
        measure: true,
        clusterGeometry: false,
      }).advancePt;
      if (
        !Number.isFinite(naturalSpace) ||
        !Number.isFinite(ideographicCell) ||
        naturalSpace < 0 ||
        ideographicCell <= 0
      )
        return undefined;
      const adjustmentPt = ideographicCell / 2 - naturalSpace;
      metricCache.set(key, adjustmentPt);
      return adjustmentPt;
    };
    let sequence: LayoutTextSeg[] = [];
    let sequenceCount = 0;
    const flushSequence = () => {
      if (wordBalancedConsecutiveSpaceCellApplies(sequenceCount)) {
        for (const segment of sequence) {
          segment.widthBalanceSpaceSequence = true;
          const adjustmentPt = adjustmentFor(segment);
          if (adjustmentPt !== undefined) {
            segment.widthBalanceSpaceAdjustmentPt = adjustmentPt;
          }
        }
      }
      sequence = [];
      sequenceCount = 0;
    };
    for (const candidate of segs) {
      if (!('text' in candidate) || candidate.script === 'complexScript') {
        flushSequence();
        continue;
      }
      const trailingSpaces = candidate.text.length - candidate.text.replace(/ +$/u, '').length;
      const spaceOnly = trailingSpaces > 0 && trailingSpaces === candidate.text.length;
      if (!spaceOnly) flushSequence();
      if (trailingSpaces > 0) {
        sequence.push(candidate);
        sequenceCount += trailingSpaces;
      } else {
        flushSequence();
      }
    }
    flushSequence();
  }

  // ── UAX#14 LB13 / ECMA-376 §17.15.1.59 (行頭禁則 — line-start-forbidden) ──────
  // A closing / mid-punctuation code point (comma, period, ; : ! ? ) ] } and
  // their CJK forms) carries NO line-break opportunity before it, so it may
  // never BEGIN a line. When such a char OPENS a segment that is glued to the
  // previous text segment — no intervening whitespace, e.g. a comma authored in
  // its own run at a formatting seam — mark it
  // `joinPrev` so the group machinery in layoutLines keeps it with the preceding
  // word and wraps "system," together instead of orphaning "," at the next
  // line's head.
  //
  // This is a UNIVERSAL Latin/Western rule (UAX#14 LB13), NOT the East-Asian
  // kinsoku feature, so it consults the application's DEFAULT forbidden table
  // UNCONDITIONALLY — independent of the document's §17.3.1.16 `w:kinsoku`
  // toggle and of any custom §17.15.1.59 `w:noLineBreaksBefore` set (which
  // REPLACES the default East-Asian table for a language and so must NOT be able
  // to drop the ASCII non-starters and re-orphan a Latin comma). The document's
  // kinsoku settings still govern the separate per-character CJK retract paths
  // (kinsokuAdjustedSplit / crossRunKinsokuRetract), which read the layout kinsoku argument.
  // The ASCII non-starters (!),.:;?]}) live in that default table (core
  // rules.ts), so one membership test covers Latin and (incidentally) CJK forms.
  for (let i = 1; i < segs.length; i++) {
    const cur = segs[i];
    if (!('text' in cur) || cur.joinPrev) continue;
    const firstCp = cur.text.codePointAt(0);
    if (firstCp === undefined || !DEFAULT_KINSOKU_RULES.lineStartForbidden.has(firstCp)) continue;
    const prev = segs[i - 1];
    // Only glue across a boundary that is NOT already a break opportunity: the
    // preceding unit must be text that does not end in whitespace (a trailing
    // space is a legal break, so the mark may legitimately start the line).
    if (!('text' in prev) || /\s$/.test(prev.text)) continue;
    cur.joinPrev = true;
  }

  // Preserve the established Word/JLReq line-end allowance for U+3000 across
  // internal script/font/width-balance shaping seams. U+3000 is BA in UAX #14,
  // so the break opportunity is after the space; splitting the space into its
  // own internal segment must not invent an opportunity before it. A real
  // U+0020 source-run boundary is delegated to
  // wordSourceRunSpaceContinuesSequence below.
  for (let i = 1; i < segs.length; i++) {
    const cur = segs[i];
    if (!('text' in cur) || cur.joinPrev || (cur.text[0] !== ' ' && cur.text[0] !== '\u3000'))
      continue;
    const prev = segs[i - 1];
    if (!('text' in prev)) continue;
    const trailingSpaceFromSameRun = cur.sourceRunIndex === prev.sourceRunIndex;
    const compatibleSourceBoundary = wordSourceRunSpaceContinuesSequence(prev.text, cur.text);
    if (!trailingSpaceFromSameRun && !compatibleSourceBoundary) continue;
    cur.joinPrev = true;
  }

  // ── UAX #14 no-break pairs (LB14/LB23/LB23a/LB24/LB25/LB28/LB30) ──
  // buildSegments intentionally splits at run / font-script boundaries, but
  // those formatting seams are not line-break opportunities. Mark the following
  // segment so layoutLines' existing atomic-group pre-flush selects the previous
  // real opportunity instead. The shared predicate is deliberately one-way:
  // false means unsupported/deferred, never "break allowed".
  for (let i = 1; i < segs.length; i++) {
    const cur = segs[i];
    if (!('text' in cur) || cur.joinPrev || cur.explicitBreakBefore || cur.text.length === 0)
      continue;
    const prev = segs[i - 1];
    if (!('text' in prev) || prev.text.length === 0) continue;

    // Whitespace is an actual wrap boundary. Check both sides because source
    // runs may start with whitespace even though ASCII spaces normally remain
    // attached to the preceding splitTextForLayout token.
    if (/\s$/u.test(prev.text) || /^\s/u.test(cur.text)) continue;

    const prevChar = [...prev.text].at(-1);
    const nextChar = [...cur.text][0];
    const prevCp = prevChar?.codePointAt(0);
    const nextCp = nextChar?.codePointAt(0);
    if (prevCp === undefined || nextCp === undefined) continue;

    // U+200B is the explicit zero-width-space opportunity from LB8 and is not
    // included in JavaScript's \s character class.
    if (prevCp === 0x200b || nextCp === 0x200b) continue;

    // SEA uses the application's dictionary tailoring, so the LB1 SA→AL default
    // must not suppress a real word boundary. CJK keeps its established
    // per-character split / kinsoku path and sparse-line safeguards.
    if (containsSeaScript(prev.text) || containsSeaScript(cur.text)) continue;
    if (hasCJKBreakOpportunity(prev.text) || hasCJKBreakOpportunity(cur.text)) continue;

    if (isUax14NoBreakPair(prevCp, nextCp)) cur.joinPrev = true;
  }

  // §17.3.2.14 fitText is a fixed-width, non-wrapping unit. Glue every segment
  // after the first in the RUN-grouped region, including script/small-caps
  // pieces emitted from the same source run.
  const seenFitTextRegions = new Set<number>();
  for (const seg of segs) {
    if (!('text' in seg) || seg.fitTextRegionIndex === undefined) continue;
    if (seenFitTextRegions.has(seg.fitTextRegionIndex)) seg.joinPrev = true;
    else {
      seg.fitTextRegionStart = true;
      seenFitTextRegions.add(seg.fitTextRegionIndex);
    }
  }

  // A formatting seam on either side of an authored optional marker is not
  // an ordinary break. Keep the word joined; the discretionary consumer
  // decides both the opportunity and the conditional glyph's measured cost.
  for (let index = 0; index < segs.length; index += 1) {
    const marker = segs[index];
    if (!('text' in marker)) continue;
    const inlineEndpoint = marker.optionalHyphenBreaks?.at(-1)?.offset === marker.text.length;
    if (!marker.optionalHyphen && !inlineEndpoint) continue;
    const previous = segs[index - 1];
    const next = segs[index + 1];
    if (marker.optionalHyphen && previous && 'text' in previous && !/\s$/u.test(previous.text)) marker.joinPrev = true;
    if (next && 'text' in next && !/^\s/u.test(next.text)) next.joinPrev = true;
  }
  // Mark each complete joined word once, avoiding a lookahead scan for every
  // ordinary word or formatting seam in the paragraph.
  for (let start = 0; start < segs.length;) {
    const first = segs[start];
    if (!('text' in first)) { start++; continue; }
    let end = start + 1;
    let optional = first.optionalHyphen !== undefined || Boolean(first.optionalHyphenBreaks?.length);
    while (end < segs.length) {
      const next = segs[end];
      if (!('text' in next) || !next.joinPrev) break;
      optional ||= next.optionalHyphen !== undefined || Boolean(next.optionalHyphenBreaks?.length);
      end++;
    }
    if (optional) for (let index = start; index < end; index++) {
      (segs[index] as LayoutTextSeg).optionalHyphenWord = true;
    }
    start = end;
  }

  retainHorizontalPunctuationInkClearance(segs);
}

export function buildSegments(
  runs: readonly ParagraphLayoutRun[],
  environment: LineLayoutEnvironment,
): LayoutSeg[] {
  const segs: LayoutSeg[] = [];
  const selectedFontMetrics =
    environment.layoutServices?.text.fontMetrics ??
    environment.layoutServices?.text.localMetrics ??
    {};
  let selectedFontMetricIndex: MetricTupleIndex | undefined;
  const metricIndex = (): MetricTupleIndex =>
    (selectedFontMetricIndex ??= indexedFontMetrics(selectedFontMetrics));
  const selectedMetric = (
    selected: FontResolution | undefined,
    probeText: string,
  ): ResolvedFontMetric | undefined => {
    if (!selected) return undefined;
    return selectResourceMetric(metricIndex(), selected, probeText);
  };
  const referenceAverageWidths = new Map<string, number | undefined>();
  const selectedAverageWidth = (
    selected: FontResolution | undefined,
    probeText: string,
  ): number | undefined => {
    if (!selected) return undefined;
    if (selected.source !== 'native') {
      const selectedRatio = selectResourceAverageWidthRatio(metricIndex(), selected, probeText);
      if (selectedRatio !== undefined) return selectedRatio;
      if (selected.source !== 'local' || !selected.resourceIdentity?.startsWith('office-local:')) {
        return undefined;
      }
    }
    if (!mayUseExactLocalReferenceWidthMetric(selected)) return undefined;
    const key = `${selected.requestedFamily}\0${selected.weight}\0${selected.style}`;
    if (!referenceAverageWidths.has(key)) {
      const admitted = referenceFontAverageWidthRatio(
        selected.requestedFamily,
        selected.weight,
        selected.style,
      );
      // Only catalog-backed tuples enter this per-build cache; arbitrary
      // authored family names cannot grow it across a long paragraph.
      if (admitted !== undefined) referenceAverageWidths.set(key, admitted);
    }
    return referenceAverageWidths.get(key);
  };
  // Group §17.3.2.14 adjacency over SOURCE RUNS before script/font, word, or
  // small-caps segmentation, but model each tab-delimited fragment as its own
  // source unit. A tab is a position-dependent advance rather than a glyph, so a
  // non-fit kernel entry at every tab boundary prevents same-id fragments from
  // linking across it. Width/count placeholders are resolved from emitted text
  // at layout scale below.
  const fitTextFragmentEntryByKey = new Map<string, number>();
  const fitTextRuns: FitTextRun[] = [];
  for (const [runIndex, run] of runs.entries()) {
    if (run.type !== 'text') {
      fitTextRuns.push({ charCount: 0, naturalWidthPx: 0 });
      continue;
    }
    const fragments = run.text.split('\t');
    for (let fragmentIndex = 0; fragmentIndex < fragments.length; fragmentIndex += 1) {
      fitTextFragmentEntryByKey.set(`${runIndex}:${fragmentIndex}`, fitTextRuns.length);
      fitTextRuns.push({
        fitTextValTwips: run.fitTextVal,
        fitTextId: run.fitTextId,
        charCount: [...fragments[fragmentIndex]].length,
        naturalWidthPx: 0,
        charScale: run.charScale,
      });
      if (fragmentIndex < fragments.length - 1) {
        fitTextRuns.push({ charCount: 0, naturalWidthPx: 0 });
      }
    }
  }
  const fitTextRegionByEntry = new Map<number, number>();
  groupFitTextRegions(fitTextRuns, 1).forEach((region, regionIndex) => {
    for (let entryIndex = region.start; entryIndex < region.end; entryIndex += 1) {
      fitTextRegionByEntry.set(entryIndex, regionIndex);
    }
  });
  const segmentBuildContext: SegmentBuildContext = {
    runs,
    environment,
    segs,
    fitTextFragmentEntryByKey,
    fitTextRegionByEntry,
    selectedMetric,
    selectedAverageWidth,
  };

  appendRunsToSegments(runs, environment, segs, segmentBuildContext, selectedMetric);

  finalizeBuiltSegments(runs, environment, segs);
  withdrawMixedSpaceEligibilityOutsideScope(environment, segs);

  return segs;
}

const AUTO_SPACE_EAST_ASIAN = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}]/u;
const AUTO_SPACE_LATIN = /\p{Script=Latin}/u;
const AUTO_SPACE_DIGIT = /[0-9]/u;
// UAX #29 GB9 extenders and ZWJ can hide a base character's adjacency. Keep
// the existing scalar checks as well: some extending characters themselves
// carry Han script membership, and their previous exclusions must remain.
// Issue #1660's compression observation excludes autospace-on adjacency and
// does not establish these extender-bearing inputs. Preserve that exclusion;
// this narrows admission without asserting an Office automatic-spacing amount.
const AUTO_SPACE_GRAPHEME_EXTEND = /[\p{Grapheme_Extend}\u200D]/u;

/**
 * Scope of WORD_COMPRESSED_SPACE_LINE_FIT (see its registered description):
 * the paragraph keeps the unchanged line breaker when
 * - §17.3.1.2-3 automatic spacing applies (enabled, with an ideograph or kana
 *   directly beside a Latin letter or ASCII digit), which the renderer does
 *   not model, or
 * - a compressible closing mark is directly followed by U+0020, where the
 *   registered rule records a full retained cell that
 *   WORD_JAPANESE_PUNCTUATION_COMPRESSION_CELL does not reproduce.
 * Preserve the scalar exclusions on the paragraph's joined text and add
 * extender-transparent base adjacency. This only withdraws additional
 * unsupported paragraphs; it cannot re-admit a previously excluded one.
 * `texts` lists the segments in order; `undefined` marks a non-text segment,
 * which ends both adjacency checks. This is not a full grapheme segmenter.
 */
export function mixedSpaceFitTextOutsideScope(
  texts: Iterable<string | undefined>,
  autoSpaceDE: boolean | undefined,
  autoSpaceDN: boolean | undefined,
): boolean {
  const pair = (left: string, right: string): boolean => {
    if (COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION.has(left) && right === ' ') return true;
    const eastAsian = AUTO_SPACE_EAST_ASIAN.test(left) ? right : AUTO_SPACE_EAST_ASIAN.test(right) ? left : undefined;
    if (eastAsian === undefined) return false;
    return (autoSpaceDE !== false && AUTO_SPACE_LATIN.test(eastAsian))
      || (autoSpaceDN !== false && AUTO_SPACE_DIGIT.test(eastAsian));
  };
  let previous: string | undefined;
  let previousBase: string | undefined;
  for (const text of texts) {
    if (text === undefined) {
      previous = undefined;
      previousBase = undefined;
      continue;
    }
    for (const character of text) {
      if (previous !== undefined && pair(previous, character)) return true;
      if (!AUTO_SPACE_GRAPHEME_EXTEND.test(character)) {
        if (previousBase !== undefined && previousBase !== previous
          && pair(previousBase, character)) return true;
        previousBase = character;
      }
      previous = character;
    }
  }
  return false;
}

function withdrawMixedSpaceEligibilityOutsideScope(
  environment: LineLayoutEnvironment,
  segs: LayoutSeg[],
): void {
  if (!segs.some((segment) => 'text' in segment && segment.mixedSpaceAverageWidthRatio !== undefined)) {
    return;
  }
  if (!mixedSpaceFitTextOutsideScope(
    segs.map((segment) => ('text' in segment ? segment.text : undefined)),
    environment.autoSpaceDE,
    environment.autoSpaceDN,
  )) return;
  for (const segment of segs) {
    if ('text' in segment) segment.mixedSpaceAverageWidthRatio = undefined;
  }
}

interface SegmentEmissionState {
  readonly base: Extract<ParagraphLayoutSource['runs'][number], { type: 'text' | 'field' }>;
  readonly joinPreviousRun: boolean;
  readonly environment: LineLayoutEnvironment;
  readonly segs: LayoutSeg[];
  readonly selectedMetric: SegmentBuildContext['selectedMetric'];
  readonly selectedAverageWidth: SegmentBuildContext['selectedAverageWidth'];
  readonly r: ParagraphTextBearingRun;
  readonly effectiveVertAlign: 'super' | 'sub' | null;
  readonly effectivePosition: number | undefined;
  readonly effectiveCharacterSpacing: number | undefined;
  readonly documentCharacterCompressionApplies: boolean;
  readonly effectiveCharacterScale: number | undefined;
  readonly effectiveKerningThreshold: number | undefined;
  readonly effectiveSnapToGrid: boolean | undefined;
  readonly ruby: { hpsRaisePt?: number | undefined; text: string; fontSizePt: number } | undefined;
  readonly revision:
    | {
        readonly kind: 'insertion' | 'deletion' | 'moveFrom' | 'moveTo' | string;
        readonly id?: string | undefined;
        readonly author?: string | undefined;
        readonly date?: string | undefined;
      }
    | undefined;
  readonly fitTextFragmentEntryIndex: number | undefined;
  readonly fitTextRegionIndex: number | undefined;
  readonly csFontSize: number;
  readonly csFontFamily: string | null;
  readonly highAnsiFontFamily: string | null;
  readonly csBold: boolean;
  readonly csItalic: boolean;
  readonly eaFontFamily: string | null;
  readonly digitsAsAN: boolean;
  readonly sourceRunIndex: number;
  readonly sourceFragmentIndex: number | undefined;
  readonly hyperlink: HyperlinkTarget | undefined;
  readonly rtl: true | undefined;
  readonly overflowPunctuationEastAsianRun: true | undefined;
  reduced: boolean;
  firstSeg: boolean;
  gluePending: boolean;
  /** The run's display text around the emitted pieces, with a cursor. The
   * text service decides a script-scoped substitute's scope over this whole
   * contiguous context, not over one word (core fontSubstituteScriptScope). */
  readonly scopeContext: { readonly text: string; cursor: number };
}

function pushSegmentPiece(
  emissionState: SegmentEmissionState,
  text: string,
  cs: boolean,
  fontFamily: string | null,
  authoritativeSpan?: TextShapeSpan,
  compressCharacterWhitespace = false,
  mappedSymbolUnicode = false,
  substituteContext?: TextShapeRequest['substituteContext'],
): void {
  const {
    base,
    joinPreviousRun,
    environment,
    segs,
    selectedMetric,
    selectedAverageWidth,
    r,
    effectiveVertAlign,
    effectivePosition,
    effectiveCharacterSpacing,
    documentCharacterCompressionApplies,
    effectiveCharacterScale,
    effectiveKerningThreshold,
    effectiveSnapToGrid,
    ruby,
    revision,
    fitTextFragmentEntryIndex,
    fitTextRegionIndex,
    csFontSize,
    csFontFamily,
    highAnsiFontFamily,
    csBold,
    csItalic,
    eaFontFamily,
    digitsAsAN,
    sourceRunIndex,
    sourceFragmentIndex,
    hyperlink,
    rtl,
    overflowPunctuationEastAsianRun,
  } = emissionState;

  // Consume an exact range of the immutable full display run. A parent
  // advances once; child spans inherit its range. Never search ahead or drop
  // context when a transform fails to project: that could invent Arabic proof.
  const scopeContext = emissionState.scopeContext;
  const retainedContext = substituteContext
    ?? Object.freeze({ text: scopeContext.text, offset: scopeContext.cursor });
  assertTextShapeRunContext({ text, substituteContext: retainedContext }, scopeContext.text);
  if (!substituteContext) scopeContext.cursor += text.length;

  if (
    environment.balanceSingleByteDoubleByteWidth &&
    !cs &&
    text.includes('\u3000') &&
    [...text].some((character) => character !== '\u3000')
  ) {
    // The registered width-balance projection treats U+3000 as a half-delta space while
    // other East-Asian glyphs receive the full delta. Split only at that
    // semantic boundary so Canvas can retain one uniform letterSpacing per
    // segment (measure == paint); the space itself has no contextual shape.
    let partOffset = 0;
    for (const part of text.split(/(\u3000+)/u).filter(Boolean)) {
      pushSegmentPiece(
        emissionState,
        part,
        cs,
        fontFamily,
        undefined,
        compressCharacterWhitespace,
        mappedSymbolUnicode,
        retainedContext
          ? { text: retainedContext.text, offset: retainedContext.offset + partOffset } : undefined,
      );
      partOffset += part.length;
    }
    return;
  }
  // ECMA-376 §17.15.1.18 / §17.18.7 — dispatch the exact
  // ST_CharacterSpacing value and split each eligible full-width character
  // so its selected face's tight ink bounds can define the removable
  // whitespace. Non-eligible text retains contextual shaping.
  if (
    !compressCharacterWhitespace &&
    documentCharacterCompressionApplies &&
    fitTextRegionIndex === undefined
  ) {
    const boundaries = [0, ...graphemeClusterOffsets(text), text.length];
    const graphemes = boundaries
      .slice(0, -1)
      .map((start, index) => text.slice(start, boundaries[index + 1]));
    if (
      graphemes.some((grapheme) =>
        characterSpacingControlCompresses(grapheme, environment.characterSpacingControl),
      )
    ) {
      pushSegmentPiece(
        emissionState, text, cs, fontFamily, undefined, true, mappedSymbolUnicode, retainedContext,
      );
      return;
    }
  }
  const bold = cs ? csBold : base.bold;
  const italic = cs ? csItalic : base.italic;
  const weight = bold ? 700 : 400;
  const style = italic ? ('italic' as const) : ('normal' as const);
  const textShapeRequest: TextShapeRequest = Object.freeze({
    text,
    ...(retainedContext ? { substituteContext: retainedContext } : {}),
    fontSizePt: cs ? csFontSize : base.fontSize,
    // A successfully decoded Symbol/Wingdings code point is Unicode text,
    // not a request for the legacy font encoding. Clear every authored
    // slot so the text service resolves a Unicode-capable generic route;
    // otherwise its §17.3.2.26 slot resolver selects Symbol again and the
    // mapped character can still be drawn with the wrong cmap.
    fonts: mappedSymbolUnicode
      ? { ascii: null, highAnsi: null, eastAsia: null, complexScript: null }
      : (r.fontSlots?.direct ?? {
          ascii: base.fontFamily,
          highAnsi: highAnsiFontFamily,
          eastAsia: eaFontFamily,
          complexScript: csFontFamily,
        }),
    themeFonts: mappedSymbolUnicode ? undefined : r.fontSlots?.theme,
    themeFontPresence: mappedSymbolUnicode ? undefined : r.fontSlots?.themePresent,
    weight,
    style,
    complexScript: cs,
    fontHint: r.fontHint,
    eastAsiaLanguage: r.langEastAsia,
    kerning:
      wordKerningApplies(cs ? csFontSize : base.fontSize, effectiveKerningThreshold, environment.compatibilityMode),
    measure: false,
  });
  const shaped = authoritativeSpan
    ? { spans: [authoritativeSpan] }
    : environment.layoutServices?.text.shape(textShapeRequest);
  const punctuationCompressions =
    compressCharacterWhitespace && documentCharacterCompressionApplies
      ? (() => {
          const boundaries = [0, ...graphemeClusterOffsets(text), text.length];
          const compressions: Array<{ end: number; adjustmentPt: number }> = [];
          for (let index = 0; index < boundaries.length - 1; index += 1) {
            const start = boundaries[index]!;
            const end = boundaries[index + 1]!;
            const compressedGrapheme = text.slice(start, end);
            if (
              !characterSpacingControlCompresses(
                compressedGrapheme,
                environment.characterSpacingControl,
              )
            )
              continue;
            const measured = environment.layoutServices?.text.shape({
              ...sliceTextShapeRequest(textShapeRequest, start, end),
              measure: true,
              clusterGeometry: false,
            });
            // ECMA-376 §17.15.1.18 defines only compression eligibility. The
            // registered Word observation retains half of the selected route's
            // ideographic cell for punctuation; this is not assumed to be half
            // of a proportional punctuation glyph's own advance. Kana in
            // `compressPunctuationAndJapaneseKana` has no observed cell floor,
            // so only its measured trailing sidebearing is removed.
            const removableUnscaledPt =
              measured?.inkBounds && measured.horizontalInkBoundsAreTight === true
                ? (() => {
                    const trailingWhitespacePt = Math.max(
                      0,
                      Math.min(measured.advancePt, measured.advancePt - measured.inkBounds.xMaxPt),
                    );
                    if (!COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION.has(compressedGrapheme)) {
                      return trailingWhitespacePt;
                    }
                    const punctuationRoute = measured.spans[0]?.fontRoute.fingerprint;
                    const ideographicCell = environment.layoutServices?.text.shape({
                      ...independentTextShapeRequest(textShapeRequest, '\u4e00'),
                      // U+3000 is semantically an ideographic space, but several
                      // proportional East Asian faces expose it to Canvas with
                      // the same narrow advance as their punctuation. The grid's
                      // full-width character cell is represented by an
                      // ideograph, not by that platform-specific space metric.
                      fontHint: 'eastAsia',
                      measure: true,
                      clusterGeometry: false,
                    });
                    const cellRoute = ideographicCell?.spans[0]?.fontRoute.fingerprint;
                    const cellAdvancePt = ideographicCell?.advancePt;
                    if (
                      !punctuationRoute ||
                      cellRoute !== punctuationRoute ||
                      cellAdvancePt === undefined ||
                      !Number.isFinite(cellAdvancePt) ||
                      cellAdvancePt <= 0
                    ) {
                      return 0;
                    }
                    const retainedExtentPt = wordJapanesePunctuationRetainedExtentPt({
                      punctuationAdvancePt: measured.advancePt,
                      punctuationInkEndPt: measured.inkBounds.xMaxPt,
                      ideographicCellAdvancePt: cellAdvancePt,
                    });
                    return Math.max(
                      0,
                      Math.min(trailingWhitespacePt, measured.advancePt - retainedExtentPt),
                    );
                  })()
                : 0;
            // §17.3.2.43 w:w scales the glyph and both of its sidebearings;
            // trim in the same post-scale coordinate space as segAdvanceWidth.
            const removablePt = removableUnscaledPt * (effectiveCharacterScale ?? 1);
            if (removablePt > 0) {
              compressions.push({ end, adjustmentPt: -removablePt });
            }
          }
          return compressions.length === 0
            ? undefined
            : Object.freeze(compressions.map((compression) => Object.freeze(compression)));
        })()
      : undefined;
  emitResolvedTextSegment(emissionState, {
    text,
    cs,
    fontFamily,
    authoritativeSpan,
    compressCharacterWhitespace,
    mappedSymbolUnicode,
    bold,
    italic,
    weight,
    style,
    textShapeRequest,
    shaped,
    punctuationCompressions,
  }); // glue applies only to a piece's FIRST segment
}

interface ResolvedSegmentFrame {
  readonly text: string;
  readonly cs: boolean;
  readonly fontFamily: string | null;
  readonly authoritativeSpan: TextShapeSpan | undefined;
  readonly compressCharacterWhitespace: boolean;
  readonly mappedSymbolUnicode: boolean;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly weight: number;
  readonly style: 'normal' | 'italic';
  readonly textShapeRequest: TextShapeRequest;
  readonly shaped: Readonly<{ spans: readonly TextShapeSpan[] }> | undefined;
  readonly punctuationCompressions: LayoutTextSeg['punctuationCompressions'];
}

function emitResolvedTextSegment(
  emissionState: SegmentEmissionState,
  frame: ResolvedSegmentFrame,
): void {
  const {
    base,
    environment,
    segs,
    selectedMetric,
    selectedAverageWidth,
    r,
    effectiveVertAlign,
    effectivePosition,
    effectiveCharacterSpacing,
    documentCharacterCompressionApplies,
    effectiveCharacterScale,
    effectiveKerningThreshold,
    effectiveSnapToGrid,
    ruby,
    revision,
    fitTextFragmentEntryIndex,
    fitTextRegionIndex,
    csFontSize,
    csFontFamily,
    highAnsiFontFamily,
    csBold,
    csItalic,
    eaFontFamily,
    digitsAsAN,
    sourceRunIndex,
    sourceFragmentIndex,
    hyperlink,
    rtl,
    overflowPunctuationEastAsianRun,
    joinPreviousRun,
  } = emissionState;
  const {
    text,
    cs,
    fontFamily,
    authoritativeSpan,
    compressCharacterWhitespace,
    mappedSymbolUnicode,
    bold,
    italic,
    weight,
    style,
    textShapeRequest: initialTextShapeRequest,
    shaped: initialShape,
    punctuationCompressions,
  } = frame;
  let textShapeRequest = initialTextShapeRequest;
  let shaped = initialShape;
  // One physical grapheme or single-seam Latin word piece retains semantic slots.
  // Both ordinary slots use the same allocation policy only in the explicit
  // absence of character grid, spacing/scaling and atomic/transformed units.
  // The text service separately proves exact registered face and cmap cover.
  const hasSlotSeam = !authoritativeSpan && shaped !== undefined && shaped.spans.length > 1;
  const joinGrapheme = hasSlotSeam && registeredLatinMarkGraphemeCandidate(text);
  const joinLatin = hasSlotSeam && registeredLatinSingleSeamCandidate(shaped!.spans)
    && textShapeRequest.kerning === true && registeredLatinSlotRunCandidate(text);
  const uniformAllocation = environment.characterGridActive === false && environment.verticalCJK !== true
    && environment.paragraphRtl === false
    && !rtl && !ruby && fitTextRegionIndex === undefined
    && !base.smallCaps && !base.allCaps && !effectiveVertAlign
    && (effectiveCharacterSpacing == null || effectiveCharacterSpacing === 0)
    && (effectiveCharacterScale == null || effectiveCharacterScale === 1)
    && !mappedSymbolUnicode && !compressCharacterWhitespace
    && (r.fontHint == null || r.fontHint === 'default');
  if (!authoritativeSpan && shaped && shaped.spans.length > 1 && (joinGrapheme || joinLatin)
    && uniformAllocation) {
    // `text` is one splitTextForLayout piece, including its pure trailing
    // U+0020 sequence. That separator is not a second WORD slot seam; the full
    // piece still shares the uniform allocation/face/cmap proof. This is never
    // a new run/line merge. Authored paint/atomic and differing policies split.
    const compoundRequest = Object.freeze({ ...textShapeRequest,
      ...(joinGrapheme ? { joinRegisteredGrapheme: true } : { joinRegisteredLatinSlots: true }),
    });
    const compound = environment.layoutServices?.text.shape(compoundRequest);
    if (compound?.spans.length === 1 && compound.spans[0]?.semanticSlotSpans) {
      shaped = compound;
      textShapeRequest = compoundRequest;
    }
  }
  // ASCII Latin base + marks followed only by U+0020, when the whole piece was
  // not adopted above (kerning off, or a peer covering the body but not the
  // separator). Prove the body alone through the existing joinRegisteredGrapheme
  // gate; the separator stays its own span. The initial push already consumed
  // this piece's scope range, so: the body is emitted directly (pushSegmentPiece
  // would rebuild its request and drop admission); the tail is pushed exactly
  // as the multi-span loop below would push it (context given, no cursor move).
  // Cross-word pair probes still require their own whole face/cmap proof.
  // Any decline falls through to that unchanged loop with the original shape.
  const markBody = !authoritativeSpan && uniformAllocation && !cs && shaped && shaped.spans.length > 1
    ? /^([A-Za-z]\p{M}+)( +)$/u.exec(text) : null;
  const tailSpan = markBody ? shaped!.spans[shaped!.spans.length - 1] : undefined;
  if (markBody && tailSpan && registeredLatinMarkGraphemeCandidate(markBody[1]!)
    && tailSpan.text === markBody[2] && tailSpan.start === markBody[1]!.length
    && tailSpan.end === text.length
    && (tailSpan.script === 'ascii' || tailSpan.script === 'highAnsi')) {
    const body = markBody[1]!;
    const bodyRequest: TextShapeRequest = Object.freeze({
      ...sliceTextShapeRequest(textShapeRequest, 0, body.length),
      joinRegisteredGrapheme: true,
      joinRegisteredLatinSlots: undefined,
    });
    const bodyShape = environment.layoutServices?.text.shape(bodyRequest);
    if (bodyShape?.spans.length === 1 && bodyShape.spans[0]?.semanticSlotSpans) {
      emitResolvedTextSegment(emissionState, {
        ...frame, text: body, textShapeRequest: bodyRequest, shaped: bodyShape,
      });
      pushSegmentPiece(
        emissionState, tailSpan.text, false,
        tailSpan.script === 'highAnsi' ? highAnsiFontFamily : base.fontFamily,
        tailSpan, false, mappedSymbolUnicode,
        sliceTextShapeRequest(textShapeRequest, tailSpan.start, tailSpan.end).substituteContext,
      );
      return;
    }
  }
  const resolvedAxisDiffers =
    shaped?.spans.some((span) => (span.script === 'complexScript') !== cs) ?? false;
  if (shaped && (shaped.spans.length > 1 || resolvedAxisDiffers)) {
    for (let spanIndex = 0; spanIndex < shaped.spans.length; spanIndex += 1) {
      const span = shaped.spans[spanIndex]!;
      const spanCs = span.script === 'complexScript';
      const spanFamily = spanCs
        ? csFontFamily
        : span.script === 'eastAsia'
          ? eaFontFamily
          : span.script === 'highAnsi'
            ? highAnsiFontFamily
            : base.fontFamily;
      const compressedSpan =
        documentCharacterCompressionApplies &&
        compressCharacterWhitespace &&
        [...span.text].some((grapheme) =>
          characterSpacingControlCompresses(grapheme, environment.characterSpacingControl),
        );
      pushSegmentPiece(
        emissionState, span.text, spanCs, spanFamily, span, compressedSpan, mappedSymbolUnicode,
        sliceTextShapeRequest(textShapeRequest, span.start, span.end).substituteContext,
      );
    }
    return;
  }
  const resolvedSpan = shaped?.spans[0];
  const localFont = resolvedSpan ? selectedMetric(resolvedSpan.font, text) : undefined;
  const eaResolution = environment.layoutServices?.text.resolve({
    fonts: textShapeRequest.fonts,
    themeFonts: textShapeRequest.themeFonts,
    themeFontPresence: textShapeRequest.themeFontPresence,
    slot: 'eastAsia',
    weight,
    style,
  });
  // A DrawingML/WPS text-body floor is reserved by its independently
  // selected eastAsia slot, even when this run paints Latin glyphs through
  // ascii. Empty coverage here means "no glyph is borrowed from this face";
  // the selected resource contributes vertical geometry only. Ordinary
  // segments still require coverage of their actual text.
  const eaFloorProbe = (r as DocxTextRun & { textBoxLineFloor?: boolean }).textBoxLineFloor
    ? ''
    : text;
  const localEaFloor = eaResolution ? selectedMetric(eaResolution, eaFloorProbe) : undefined;
  // Once a service selected a face, do not re-admit metrics by the authored
  // family name after the selected-face lookup failed (e.g. a substitute).
  // §17.3.1.33 atLeast is max(normal single-line height, authored minimum).
  // The Word OpenType projection supplies that normal box by inference;
  // exact spacing instead suppresses it.
  const naturalMetricAllowed = environment.lineSpacing?.rule !== 'exact';
  const { resourceMetric: resourceFamilyLineMetric, referenceMetric: referenceLineMetric,
    lineMetric: familyLineMetric } = selectedFontLineMetric(resolvedSpan?.font, localFont, naturalMetricAllowed);
  const { resourceMetric: resourceEaLineMetric, referenceMetric: referenceEaLineMetric,
    lineMetric: eaLineMetric } = selectedFontLineMetric(eaResolution, localEaFloor, naturalMetricAllowed);
  const resolvedEaFloorFamily =
    eaResolution?.resolvedFamily ?? localEaFloor?.family ?? eaFontFamily;
  // WORD_USE_FE_LAYOUT_INHERITED_GRID_MINIMUM was observed for an active
  // document line grid. ECMA-376 §17.6.5 disables an omitted/default grid
  // type and excludes table-cell line pitch unless §17.15.3.1
  // adjustLineHeightInTable is enabled. An inherited eastAsia font axis
  // alone must not raise a Latin-only non-grid line's normal-line floor.
  const useFeEastAsianMetric =
    environment.useFeLayout &&
    environment.lineGridActive &&
    (r.fontHint === 'eastAsia' || Boolean(resolvedEaFloorFamily?.trim()));
  const resolvedScript =
    resolvedSpan?.script ??
    authoritativeSpan?.script ??
    (cs ? 'complexScript' : EAST_ASIAN_RE.test(text) ? 'eastAsia' : 'ascii');
  const latinSpaceCompressionEligible =
    environment.characterSpacingControl === 'compressPunctuation' &&
    // MS-OE376 §2.1.472 requires full advance for fit with this
    // compatibility switch, even if display uses compression. This gate
    // covers the measured Latin fit behavior; display-space placement
    // under the switch has not yet been established.
    environment.lineWrapLikeWord6 !== true &&
    // The xAvg floor was measured for horizontal Latin words; upright
    // vertical runs and tate-chu-yoko use distinct advance allocation.
    environment.verticalCJK !== true &&
    documentCharacterCompressionApplies &&
    (resolvedScript === 'ascii' || resolvedScript === 'highAnsi') &&
    !emissionState.reduced &&
    effectiveVertAlign == null &&
    (effectiveCharacterSpacing == null || effectiveCharacterSpacing === 0) &&
    (effectiveCharacterScale == null || effectiveCharacterScale === 1) &&
    // WORD_LATIN_INTERWORD_XAVG_FLOOR retains its existing OpenType gate;
    // threshold authority must not widen this separate fit policy's scope.
    // Evidence gap: disabling implicit/zero kerning exposes justified fitting
    // losses in mode-15 zero/absent-threshold controls, and mode-14 space fitting
    // remains unresolved. Preserve this fit gate pending the separate justified
    // compression correction; do not infer an allowance from pair advances or
    // apply a font-specific scale. Positive-threshold fit contradictions likewise
    // do not establish a different shaping-table or kerning-switch rule.
    environment.enableOpenTypeFeatures !== true &&
    effectiveKerningThreshold == null;
  const latinSpaceAverageWidthRatio = latinSpaceCompressionEligible
    ? selectedAverageWidth(resolvedSpan?.font, text)
    : undefined;
  // WORD_COMPRESSED_SPACE_LINE_FIT: U+0020 on mixed East Asian / Latin lines.
  // The line breaker applies it only once its line holds East Asian text;
  // OpenType features and explicit kerning thresholds do not gate it.
  const mixedSpaceCompressionEligible =
    wordCompressedSpaceLineFitApplies(
      environment.compatibilityMode,
      environment.characterSpacingControl,
    ) &&
    environment.lineWrapLikeWord6 !== true &&
    environment.verticalCJK !== true &&
    documentCharacterCompressionApplies &&
    (resolvedScript === 'ascii' || resolvedScript === 'highAnsi') &&
    !emissionState.reduced &&
    effectiveVertAlign == null &&
    (effectiveCharacterSpacing == null || effectiveCharacterSpacing === 0) &&
    (effectiveCharacterScale == null || effectiveCharacterScale === 1);
  const mixedSpaceAverageWidthRatio = mixedSpaceCompressionEligible
    ? latinSpaceAverageWidthRatio ?? selectedAverageWidth(resolvedSpan?.font, text)
    : undefined;
  const widthBalanceGridDeltaFactor = environment.balanceSingleByteDoubleByteWidth
    ? wordBalancedLinesAndCharsGridDeltaFactor(text, resolvedScript)
    : undefined;
  segs.push({
    text,
    script: resolvedScript,
    ...(resolvedSpan?.semanticSlotSpans ? { semanticSlotSpans: resolvedSpan.semanticSlotSpans } : {}),
    ...(widthBalanceGridDeltaFactor !== undefined
      ? {
          // §17.15.3.3 defines the SBCS:DBCS width ratio as 1:2; the
          // registered Word matrix defines how that setting projects onto
          // linesAndChars charSpace. Production shaping has already split
          // the segment at §17.3.2.26 script-slot boundaries.
          widthBalanceGridDeltaFactor,
        }
      : {}),
    ...(useFeEastAsianMetric ? { metricEastAsian: true as const } : {}),
    bold,
    italic,
    underline: base.underline,
    // §17.3.2.40 underline style / colour — carried only on DocxTextRun (a
    // FieldRun draws single). Kept raw ST_Underline; the renderer normalizes
    // to DrawingML §20.1.10.82 at draw time.
    underlineStyle: r.underlineStyle,
    underlineColor: r.underlineColor,
    strikethrough: base.strikethrough,
    fontSize: cs ? csFontSize : base.fontSize,
    color: base.color,
    fontFamily: resolvedSpan?.font.resolvedFamily ?? localFont?.family ?? fontFamily,
    authoredFontFamily: resolvedSpan?.font.requestedFamily ?? fontFamily,
    fontSource: resolvedSpan?.font.source,
    authoredReferenceMetricAllowed: mayUseAuthoredReferenceVerticalMetric(resolvedSpan?.font),
    fontRoute: resolvedSpan?.fontRoute,
    resolvedLineHeightRatio: familyLineMetric?.lineHeightRatio,
    resolvedLatinGridCellAllocation: referenceLineMetric?.farEastCodePage === false,
    ...(resourceFamilyLineMetric?.lineHeightRatio != null
      ? {
          resolvedResourceVerticalMetric: true as const,
          resolvedDesignAscentRatio: resourceFamilyLineMetric.designAscentRatio,
          resolvedDesignDescentRatio: resourceFamilyLineMetric.designDescentRatio,
        }
      : {}),
    ...(referenceLineMetric
      ? {
          referenceFontVerticalMetric: true as const,
          resolvedDesignAscentRatio: referenceLineMetric.designAscentRatio,
          resolvedDesignDescentRatio: referenceLineMetric.designDescentRatio,
        }
      : {}),
    resolvedEastAsianLineHeightRatio: familyLineMetric?.eastAsianLineHeightRatio,
    ...(latinSpaceAverageWidthRatio != null && latinSpaceAverageWidthRatio > 0
      ? { latinSpaceAverageWidthRatio, latinSpaceCompressionEligible: true as const }
      : {}),
    ...(mixedSpaceAverageWidthRatio != null && mixedSpaceAverageWidthRatio > 0
      ? { mixedSpaceAverageWidthRatio }
      : {}),
    vertAlign: effectiveVertAlign,
    measuredWidth: 0,
    textLayoutService: environment.layoutServices?.text,
    textShapeRequest,
    ...(resolvedSpan?.substituteScope !== undefined
      ? { substituteScope: resolvedSpan.substituteScope } : {}),
    breakBefore: resolvedSpan?.breakBefore ?? authoritativeSpan?.breakBefore ?? true,
    smallCaps: emissionState.reduced,
    joinPrev:
      (emissionState.firstSeg && (r.noBreakBefore === true || joinPreviousRun)) ||
      emissionState.gluePending ||
      authoritativeSpan?.breakBefore === false
        ? true
        : undefined,
    hardJoinPrev:
      emissionState.firstSeg && (r.noBreakBefore === true || joinPreviousRun) ? true : undefined,
    doubleStrikethrough: base.doubleStrikethrough ?? false,
    highlight: base.highlight ?? null,
    // §17.3.2.12 w:em — carried on both DocxTextRun and FieldRun (a field's
    // resolved/fallback text stamps the mark the same as a plain run).
    emphasisMark: base.emphasisMark,
    background: base.background ?? null,
    colorAuto: r.colorAuto ?? false,
    border: r.border ?? null,
    ruby: emissionState.firstSeg ? ruby : undefined,
    revision,
    ...(revision && environment.showTrackedChanges === true
      ? {
          trackChangesMarkup: {
            kind: revision.kind,
            authorColor: environment.revisionAuthorColor?.(revision.author) ?? '#C00000',
          },
        }
      : {}),
    rtl,
    digitsAsAN: digitsAsAN ? true : undefined,
    // §17.3.2.26 declared eastAsia axis — used by text-box line floors and
    // the compatibility-owned useFELayout body metric path.
    eaFloorFamily: resolvedEaFloorFamily,
    eaFloorRoute: eaResolution?.route,
    resolvedEaFloorLineHeightRatio: eaLineMetric?.lineHeightRatio,
    resolvedEaFloorEastAsianLineHeightRatio: eaLineMetric?.eastAsianLineHeightRatio,
    textBoxLineFloor: (r as DocxTextRun & { textBoxLineFloor?: boolean }).textBoxLineFloor,
    textBoxVertical: (r as DocxTextRun & { textBoxVertical?: boolean }).textBoxVertical,
    // IX1 — resolved hyperlink target of the originating run, for the
    // text-layer clickable overlay and URL-aware line-break opportunities.
    // It does not change glyph measurement or drawing.
    hyperlink,
    snapToCharacterGrid: effectiveSnapToGrid !== false,
    // WD4 — run character metrics (§17.3.2.35 spacing / §17.3.2.43 w /
    // §17.3.2.24 position / §17.3.2.19 kern). Uniform across the run, so
    // every emitted segment carries the same values; the measure and paint
    // passes apply them identically (measure==paint).
    charSpacing: effectiveCharacterSpacing,
    punctuationCompressions,
    eastAsiaLanguage: r.langEastAsia,
    overflowPunctuationEastAsianRun,
    overflowPunctuationBidiLanguage: r.langBidi,
    charScale: effectiveCharacterScale,
    fitTextVal: fitTextRegionIndex === undefined ? undefined : r.fitTextVal,
    fitTextId: fitTextRegionIndex === undefined ? undefined : r.fitTextId,
    fitTextRegionIndex,
    fitTextRunIndex: fitTextRegionIndex === undefined ? undefined : fitTextFragmentEntryIndex,
    position: effectivePosition,
    positionExtendsLineBox: environment.positionExtendsLineBox !== false,
    kerning: effectiveKerningThreshold,
    // ECMA-376 §17.3.2.10 eastAsianLayout — 縦中横 is meaningful ONLY in a
    // vertical (tbRl) page, so fold the vertical gate in HERE at build time
    // (buildSegments receives it through LineLayoutEnvironment). Measure/paint then read a single
    // pre-gated flag. `vertCompress` rides only when `vert` is set (spec: it
    // is ignored otherwise).
    tateChuYoko: environment.verticalCJK && r.eastAsianVert === true ? true : undefined,
    tateChuYokoCompress:
      environment.verticalCJK && r.eastAsianVert === true && r.eastAsianVertCompress === true
        ? true
        : undefined,
    // #1014 — an upright-vertical (tbRl) per-glyph segment (NOT a 縦中横 cell,
    // which is one drawTateChuYokoRun cell). Marks the segment for the vo=Tr
    // rotate-fallback ink-extent advance correction in the measure passes.
    verticalRun: environment.verticalCJK && r.eastAsianVert !== true ? true : undefined,
  });
  emissionState.firstSeg = false;
  emissionState.gluePending = false;
}

function appendRunsToSegments(
  runs: readonly ParagraphLayoutRun[],
  environment: LineLayoutEnvironment,
  segs: LayoutSeg[],
  segmentBuildContext: SegmentBuildContext,
  selectedMetric: SegmentBuildContext['selectedMetric'],
): void {
  // A native reserved-separator paragraph holds only text-free metric
  // participants; coalescing them would drop a participant's own CHPX.
  const sequences = runs.some((run) => run.type === 'text'
    && (run as ParagraphTextBearingRun).noteSeparatorCharacter !== undefined)
    ? new Map<number, never>()
    : acquireTextSequences(runs, environment, (text, run) => transformedRunText(text, run, environment));
  let sequenceEnd = -1;
  let joinNextVisibleText = false;
  for (const [runIndex, sourceRun] of runs.entries()) {
    if (runIndex <= sequenceEnd) continue;
    const sequence = sequences.get(runIndex);
    const run = sequence?.run ?? sourceRun;
    if (sequence) sequenceEnd = sequence.sources.at(-1)!.runIndex;
    // ECMA-376 §17.13.5 final view (the default): deleted (`w:del`,
    // §17.13.5.14) and moved-away (`w:moveFrom`, §17.13.5.22) content is not
    // part of the document's final state, so no segment is produced and line
    // breaking/pagination see the accepted document state. The markup view
    // (`showTrackedChanges`) keeps every revision run visible so it can be
    // decorated. Insertions/moveTo render in both views, and revision metadata
    // remains available through the parsed model for consumer-owned review UI.
    const runRevisionKind = (run as { revision?: { kind?: string } }).revision?.kind;
    if (revisionIsOmitted(runRevisionKind, environment.showTrackedChanges)) {
      continue;
    }
    const joinFromPreviousNoBreakHyphen = joinNextVisibleText;
    joinNextVisibleText =
      run.type === 'text' && (run as ParagraphTextBearingRun).noBreakAfter === true;
    const emittedStart = segs.length;
    if (run.type === 'text') {
      const t = run as unknown as DocxTextRun & { type: 'text' };
      if ((run as ParagraphTextBearingRun).optionalHyphen === true) {
        // §17.3.3.29: acquire the conditional hyphen through the ordinary
        // font/shape/paint route, retaining its own run rather than borrowing
        // either neighbor. Its empty source marker contributes no geometry
        // unless the breaker chooses this authored boundary.
        appendTextPiece(segmentBuildContext, '-', t, t.vertAlign ?? null, runIndex,
          { text: '-', offset: 0 });
        for (let index = emittedStart; index < segs.length; index += 1) {
          const glyph = segs[index];
          if (!('text' in glyph)) throw new Error('An optional hyphen lost its text authority');
          segs[index] = { ...glyph, text: '', metricOnly: true,
            optionalHyphen: Object.freeze({ ...glyph }), measuredWidth: 0,
            sourceRunIndex: runIndex };
        }
        continue;
      }
      if ((run as ParagraphTextBearingRun).noteSeparatorCharacter !== undefined) {
        // MS-DOC 2.3.3 reserved separator character: the U+0003/U+0004 rule
        // control or its story's content paragraph mark. Neither has a glyph
        // here; the rule ink is retained separately. As for a suppressed note
        // mark, a bounded Latin probe resolves the run's own four font slots
        // and selected face through the ordinary text service, and only its
        // vertical metrics remain (zero advance, no ink). Control/mark code
        // points are not East Asian, so an East Asian slot alone is not used.
        // The run context is the transformed display probe (caps/small caps
        // or symbol mapping), exactly as for any other text piece.
        const probe = transformedRunText('x', t, environment);
        appendTextPiece(segmentBuildContext, 'x', t, t.vertAlign ?? null, runIndex,
          { text: probe, offset: 0 });
        for (let index = emittedStart; index < segs.length; index += 1) {
          const segment = segs[index];
          if (!('text' in segment)) throw new Error('A separator metric probe lost its text authority');
          segment.text = '';
          segment.metricOnly = true;
          // Keep the probe for vertical metrics only (pass-operations,
          // paragraph sourceMetrics); display and width stay empty.
          segment.metricProbeText = probe;
          segment.sourceRunIndex = runIndex;
          if (segment.textShapeRequest) {
            segment.textShapeRequest = Object.freeze(independentTextShapeRequest(segment.textShapeRequest, ''));
          }
        }
        continue;
      }
      // ECMA-376 §17.11: substitute a footnote/endnote reference marker's glyph
      // with the note's resolved sequential number. The body `*Reference` run
      // (§17.11.14 footnoteReference / §17.11.7 endnoteReference) carries the
      // id; the in-note `*Ref` placeholder (§17.11.13 footnoteRef / §17.11.6
      // endnoteRef) carries an empty id, so we fall back to the note number
      // currently being drawn. Numbering and formatting are independent: the
      // mark takes its run's effective §17.3.2.42 w:vertAlign (direct §17.3.2.28
      // rPr or style), and no superscript is synthesized when it is absent.
      const noteText = t.noteRef
        ? t.noteRef.id
          ? environment.noteNumbers?.get(`${t.noteRef.kind}:${t.noteRef.id}`)
          : environment.noteReferenceNumber
        : undefined;
      if (t.noteRef) {
        // CT_FtnEdnRef/@customMarkFollows suppresses the automatic glyph, not
        // the note relationship. Keep an immutable zero-width host so a note
        // remains attached to the physical line/page of this reference.
        // Number 0 is the acquisition map's custom-note sentinel; the note's
        // own automatic *Ref placeholder is suppressed by the same contract.
        if (t.noteRef.customMarkFollows === true || noteText === 0) {
          // As for an empty/anchor-only mark, a bounded Latin probe resolves
          // the four font slots and selected-face metrics through the ordinary
          // text service. Discard its ink/text, never its font authority.
          appendTextPiece(segmentBuildContext, 'x', t, t.vertAlign ?? null, runIndex,
            { text: 'x', offset: 0 });
          for (let index = emittedStart; index < segs.length; index += 1) {
            const segment = segs[index];
            if (!('text' in segment)) throw new Error('A note metric probe lost its text authority');
            segment.text = '';
            segment.metricOnly = true;
            segment.sourceRunIndex = runIndex;
            if (segment.textShapeRequest) {
              segment.textShapeRequest = Object.freeze(independentTextShapeRequest(segment.textShapeRequest, ''));
            }
          }
          continue;
        }
        const label =
          noteText != null
            ? formatNoteNumber(
                noteText,
                t.noteRef.kind === 'footnote'
                  ? environment.noteNumbering?.footnote
                  : t.noteRef.kind === 'endnote'
                    ? environment.noteNumbering?.endnote
                    : undefined,
              )
            : t.text || '';
        if (label.length > 0) {
          appendTextPiece(
            segmentBuildContext,
            label,
            t,
            t.vertAlign ?? null,
            runIndex,
            { text: transformedRunText(label, t, environment), offset: 0 },
            0,
            joinFromPreviousNoBreakHyphen,
          );
        }
        for (let index = emittedStart; index < segs.length; index += 1) {
          segs[index].sourceRunIndex = runIndex;
        }
        continue;
      }
      // Split on tab chars so tab alignment can be resolved during layout.
      const parts = t.text.split('\t');
      const fullDisplayText = transformedRunText(t.text, t, environment);
      let displayOffset = 0;
      for (let i = 0; i < parts.length; i++) {
        if (parts[i].length > 0) {
          appendTextPiece(
            segmentBuildContext,
            parts[i],
            t,
            t.vertAlign,
            runIndex,
            { text: fullDisplayText, offset: displayOffset },
            i,
            i === 0 && joinFromPreviousNoBreakHyphen,
          );
        }
        displayOffset += transformedRunText(parts[i], t, environment).length + 1;
        if (i < parts.length - 1) {
          segs.push({
            isTab: true,
            fontSize: t.fontSize,
            measuredWidth: 0,
            bold: t.bold,
            italic: t.italic,
            sourceRunIndex: runIndex,
            ...(sequence ? { sourceTextOffset: displayOffset - 1 } : {}),
          });
        }
      }
    } else if (run.type === 'image') {
      const img = run;
      segs.push({
        imagePath: img.imagePath,
        mimeType: img.mimeType,
        ...(img.anchor ? {} : { inlinePicture: true as const }),
        widthPt: img.widthPt,
        heightPt: img.heightPt,
        rotation: img.rotation,
        flipH: img.flipH,
        flipV: img.flipV,
        anchor: img.anchor ?? false,
        anchorXPt: img.anchorXPt ?? 0,
        anchorYPt: img.anchorYPt ?? 0,
        anchorXFromMargin: img.anchorXFromMargin ?? false,
        anchorYFromPara: img.anchorYFromPara ?? false,
        colorReplaceFrom: img.colorReplaceFrom,
        duotone: img.duotone,
        alpha: img.alpha,
        srcRect: img.srcRect ?? undefined,
        measuredWidth: 0,
      });
    } else if (run.type === 'chart') {
      // ECMA-376 §21.2 chart. Flow it as a picture box of the `<wp:extent>`
      // natural size: the same LayoutImageSeg shape (empty `imagePath`/
      // `mimeType` sentinels so `'imagePath' in seg` routes it through the image
      // measurement/split path) with only a chart resource marker; the model
      // payload remains owned by the paint resource registry.
      //
      // A `<wp:anchor>` (floating) chart (§20.4.2.3) carries `anchor: true` and
      // its parsed page-offset fields, exactly like an anchor ImageRun: the
      // measure pass zeroes an anchor seg's width (it is not part of the inline
      // flow) and anchor acquisition retains it at the resolved absolute box.
      const chartRun = run;
      segs.push({
        imagePath: '',
        mimeType: '',
        widthPt: chartRun.widthPt,
        heightPt: chartRun.heightPt,
        anchor: chartRun.anchor ?? false,
        anchorXPt: chartRun.anchorXPt ?? 0,
        anchorYPt: chartRun.anchorYPt ?? 0,
        anchorXFromMargin: chartRun.anchorXFromMargin ?? false,
        anchorYFromPara: chartRun.anchorYFromPara ?? false,
        chart: true,
        chartResourceKey: (chartRun as Partial<import('../layout/text.js').ParagraphChartRun>)
          .resourceKey,
        measuredWidth: 0,
      });
    } else if (run.type === 'shape' && run.inline === true) {
      // `wp:inline` hosts arbitrary DrawingML, including WPS shapes (§20.4.2.8).
      // Reserve its extent in the same line-breaking path as an inline picture;
      // paragraph acquisition replaces the sentinel with a retained drawing
      // placement at the resolved pen position.
      segs.push({
        imagePath: '',
        mimeType: '',
        widthPt: run.widthPt,
        heightPt: run.heightPt,
        anchor: false,
        anchorXPt: 0,
        anchorYPt: 0,
        anchorXFromMargin: false,
        anchorYFromPara: false,
        inlineShape: true,
        measuredWidth: 0,
      });
    } else if (run.type === 'unavailableDrawing') {
      const acquiredAnchor =
        'anchorAcquisitionInput' in run ? run.anchorAcquisitionInput : undefined;
      segs.push({
        imagePath: '',
        mimeType: '',
        widthPt: run.widthPt,
        heightPt: run.heightPt,
        anchor: acquiredAnchor !== undefined,
        anchorXPt: 0,
        anchorYPt: 0,
        anchorXFromMargin: false,
        anchorYFromPara: false,
        unavailableResourceKind: run.resourceKind,
        measuredWidth: 0,
      });
    } else if (run.type === 'break') {
      if (run.breakType === 'line') {
        // Determine font size for the line break height from surrounding text runs
        const fontSize = findNearbyFontSize(runs, runs.indexOf(run));
        segs.push({ lineBreak: true, fontSize, measuredWidth: 0 });
      }
      // page/column breaks handled at the document level (splitPages)
    } else if (run.type === 'field') {
      const f = run as unknown as FieldRun & { type: 'field' };
      const text = resolveFieldText(f, environment);
      if (text) {
        appendTextPiece(
          segmentBuildContext,
          text,
          f,
          f.vertAlign,
          runIndex,
          { text: transformedRunText(text, f, environment), offset: 0 },
          undefined,
          joinFromPreviousNoBreakHyphen,
        );
      }
    } else if (run.type === 'math') {
      // The parser resolves the paragraph font size; fall back to a nearby run only
      // if it is somehow absent.
      const fontSize = run.fontSize || findNearbyFontSize(runs, runs.indexOf(run));
      const resourceKey = 'resourceKey' in run ? run.resourceKey : undefined;
      if (environment.layoutServices && !resourceKey) {
        throw new Error('Service-backed math layout requires a normalized structural resource key');
      }
      const mathMetadata = resourceKey
        ? environment.layoutServices?.math.resolve(resourceKey)
        : undefined;
      segs.push({
        math: true,
        mathResourceKey: resourceKey ?? '',
        mathMetadata,
        display: run.display,
        fontSize,
        color: null,
        fallbackText: 'fallbackText' in run ? run.fallbackText : mathFallbackText(run.nodes),
        measuredWidth: 0,
        mathAscent: 0,
        mathDescent: 0,
        jc: run.jc,
      });
    } else if (run.type === 'ptab') {
      // ECMA-376 §17.3.3.23 absolute-position tab. Emit a tab segment carrying the
      // ptab descriptor; layoutLines resolves it to an absolute X (independent of
      // the paragraph's tab stops) and fills the gap with the run's leader.
      segs.push({
        isTab: true,
        fontSize: run.fontSize || findNearbyFontSize(runs, runs.indexOf(run)),
        measuredWidth: 0,
        leader: run.leader,
        ptab: { alignment: run.alignment, relativeTo: run.relativeTo },
      });
    } else if (run.type === 'anchorHost') {
      const eastAsian = run.fontFamilyEastAsia != null;
      const bold = run.bold ?? false;
      const italic = run.italic ?? false;
      const authoredFamily = run.fontFamilyEastAsia ?? run.fontFamily ?? null;
      const weight = bold ? 700 : 400;
      const style = italic ? ('italic' as const) : ('normal' as const);
      // An anchor host has no glyph of its own. Resolve a script-matched probe
      // as for an empty paragraph mark, then require selected-face identity and
      // cmap coverage. Empty text would vacuously admit any subset resource.
      const probeText = eastAsian ? 'あ' : 'x';
      const selected = environment.layoutServices?.text.resolve({
        text: probeText,
        fonts: {
          ascii: run.fontFamily,
          highAnsi: run.fontFamily,
          eastAsia: run.fontFamilyEastAsia,
          complexScript: run.fontFamily,
        },
        slot: eastAsian ? 'eastAsia' : 'ascii',
        weight,
        style,
      });
      const localFont = selectedMetric(selected, probeText);
      // The mark hosting a floating anchor participates in the same natural
      // atLeast line box as a visible text line.
      const naturalMetricAllowed = environment.lineSpacing?.rule !== 'exact';
      const resourceMetric =
        naturalMetricAllowed || localFont?.designAscentRatio == null ? localFont : undefined;
      const referenceMetric =
        naturalMetricAllowed && !resourceMetric && mayUseAuthoredReferenceVerticalMetric(selected)
          ? referenceFontLineMetrics(selected.requestedFamily, weight, style)
          : undefined;
      const familyLineMetric = resourceMetric ?? referenceMetric;
      segs.push({
        text: '',
        metricOnly: true,
        ...(eastAsian ? { metricEastAsian: true as const } : {}),
        bold,
        italic,
        underline: false,
        strikethrough: false,
        fontSize: run.fontSize,
        color: null,
        fontFamily: selected?.resolvedFamily ?? authoredFamily,
        authoredFontFamily: selected?.requestedFamily ?? authoredFamily,
        fontSource: selected?.source,
        authoredReferenceMetricAllowed: mayUseAuthoredReferenceVerticalMetric(selected),
        fontRoute: selected?.route,
        resolvedLineHeightRatio: familyLineMetric?.lineHeightRatio,
        ...(resourceMetric?.lineHeightRatio != null
          ? {
              resolvedResourceVerticalMetric: true as const,
              resolvedDesignAscentRatio: resourceMetric.designAscentRatio,
              resolvedDesignDescentRatio: resourceMetric.designDescentRatio,
            }
          : {}),
        ...(referenceMetric
          ? {
              referenceFontVerticalMetric: true as const,
              resolvedDesignAscentRatio: referenceMetric.designAscentRatio,
              resolvedDesignDescentRatio: referenceMetric.designDescentRatio,
            }
          : {}),
        resolvedEastAsianLineHeightRatio: familyLineMetric?.eastAsianLineHeightRatio,
        vertAlign: null,
        measuredWidth: 0,
        eaFloorFamily: eastAsian ? (selected?.resolvedFamily ?? authoredFamily) : null,
        resolvedEaFloorLineHeightRatio: eastAsian ? familyLineMetric?.lineHeightRatio : undefined,
        resolvedEaFloorEastAsianLineHeightRatio: eastAsian
          ? familyLineMetric?.eastAsianLineHeightRatio
          : undefined,
        snapToCharacterGrid: false,
      });
    }
    let sequenceOffset = 0;
    let markerIndex = 0;
    for (let index = emittedStart; index < segs.length; index += 1) {
      const segment = segs[index];
      segment.sourceRunIndex = runIndex;
      if (sequence) {
        segment.sourceTextSequence = sequence.sources;
        // Authored optional controls need the displayed source sweep even
        // when no shaping service is installed. Sequences without controls
        // retain their existing anchor projection unchanged.
        segment.sourceTextOffset = sequence.optionalHyphens.length > 0 ? sequenceOffset
          : 'text' in segment ? segment.textShapeRequest?.substituteContext?.offset ?? 0
          : segment.sourceTextOffset ?? 0;
        const end = sequenceOffset + ('text' in segment ? segment.text.length : 'isTab' in segment ? 1 : 0);
        if ('text' in segment) {
          const opportunities: NonNullable<LayoutTextSeg['optionalHyphenBreaks']>[number][] = [];
          while (markerIndex < sequence.optionalHyphens.length
            && sequence.optionalHyphens[markerIndex]!.offset <= end) {
            const marker = sequence.optionalHyphens[markerIndex++]!;
            const glyphs: LayoutSeg[] = [];
            appendTextPiece({ ...segmentBuildContext, segs: glyphs }, '-', marker.run,
              marker.run.vertAlign ?? null, marker.runIndex, { text: '-', offset: 0 });
            if (glyphs.length !== 1 || !('text' in glyphs[0]!)) {
              throw new Error('An optional hyphen lost its single conditional glyph');
            }
            const glyph = glyphs[0] as LayoutTextSeg;
            glyph.sourceRunIndex = marker.runIndex;
            opportunities.push(Object.freeze({ offset: marker.offset - sequenceOffset,
              glyph: Object.freeze(glyph) }));
          }
          if (opportunities.length) segment.optionalHyphenBreaks = Object.freeze(opportunities);
        }
        sequenceOffset = end;
      }
    }
  }
}

function projectNoBreakRanges(runs: readonly ParagraphLayoutRun[], segs: LayoutSeg[]): void {
  for (const [runIndex, run] of runs.entries()) {
    if (run.type !== 'text') continue;
    const textRun = run as Extract<ParagraphTextBearingRun, { type: 'text' }>;
    const sourceRanges = textRun.noBreakRanges;
    if (!sourceRanges || sourceRanges.length === 0) continue;
    const displayedRanges = sourceRanges.map((range) => {
      const transformOffset = (offset: number) => {
        const prefix = textRun.text.slice(0, offset);
        return textRun.allCaps || textRun.smallCaps ? prefix.toUpperCase().length : prefix.length;
      };
      return { start: transformOffset(range.start), end: transformOffset(range.end) };
    });
    let displayedCursor = 0;
    for (const candidate of segs) {
      if (candidate.sourceRunIndex !== runIndex) continue;
      if (!('text' in candidate)) {
        if ('isTab' in candidate) displayedCursor += 1;
        continue;
      }
      const segmentEnd = displayedCursor + candidate.text.length;
      if (
        displayedCursor > 0 &&
        displayedRanges.some(
          (range) => range.start === displayedCursor || range.end === displayedCursor,
        )
      ) {
        candidate.joinPrev = true;
        candidate.hardJoinPrev = true;
      }
      const local = displayedRanges
        .filter((range) => range.start >= displayedCursor && range.end <= segmentEnd)
        .map((range) =>
          Object.freeze({
            start: range.start - displayedCursor,
            end: range.end - displayedCursor,
          }),
        );
      if (local.length > 0) candidate.noBreakRanges = Object.freeze(local);
      displayedCursor = segmentEnd;
    }
  }
}
