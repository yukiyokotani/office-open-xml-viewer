import { isCjkBreakChar } from '../text/cjk-ranges.js';
import { DEFAULT_KINSOKU_RULES, kinsokuAdjustedSplit } from '../text/kinsoku/index.js';
import { isUax14NoBreakPair, lineBreakClass } from '../text/line-break.js';
import { containsSeaScript, graphemeClusterOffsets, seaMixedBreakOffsets } from '../text/sea-break.js';
import { resolveDrawingMlTabWidths, type DrawingMlTabStop } from './tab.js';

export type DrawingMlInputRun<T> =
  | { type: 'text'; text: string; style: T }
  | { type: 'break'; style?: T }
  | { type: 'object'; width: number; style: T; payload?: unknown; display?: boolean };

export type DrawingMlLineSegment<T> =
  | { type: 'text'; text: string; style: T; width: number }
  | { type: 'tab'; style: T; width: number }
  | { type: 'object'; style: T; width: number; payload?: unknown };

export interface DrawingMlBrokenLine<T> {
  segments: DrawingMlLineSegment<T>[];
  width: number;
  /** The line ended at an authored `<a:br>` or LF, not a soft wrap. */
  endsWithBreak?: boolean;
  /**
   * Indexes of input text runs that paint nothing (empty, or only trimmed or
   * wrapped spaces) but sit on this line. Their run properties (for example
   * the font size) still take part in the line's metrics.
   */
  hiddenRuns?: number[];
  /**
   * For an empty line opened by a line feed inside a text run, that run's
   * input index. PowerPoint sizes such a line by the run (controls L07, L08).
   */
  lineFeedRun?: number;
  /** Index of the authored a:br that closes this line, for its rPr metrics. */
  endBreakRun?: number;
}

export interface DrawingMlBreakOptions<T> {
  /** Inner text width after body insets and paragraph margins, in canvas px. */
  maxWidth: number;
  /** Signed first-line indent; continuation lines use the full `maxWidth`. */
  firstLineIndent?: number;
  /** Resolve font and other run properties in the format adapter. */
  measureText(text: string, style: T): number;
  /** Adjacent authored runs may shape together only when all paint metadata agrees. */
  sameStyle?(left: T, right: T): boolean;
  /** Advance at a style seam within one authored run (for `rPr@spc`). */
  boundaryAdvance?(left: T, right: T): number;
  tabStops?: readonly DrawingMlTabStop[];
  /** `a:pPr@defTabSz` in canvas px; Office's ordinary fallback is 1 inch. */
  defaultTabSize?: number;
  /** Pen position relative to the leading inset; includes paragraph margin. */
  tabStartPen?(lineIndex: number): number;
  /**
   * `a:pPr@eaLnBrk` (ECMA-376 §21.1.2.2.7, default true). PowerPoint controls
   * E00–E12 show that `eaLnBrk="0"` does not keep an East Asian word whole:
   * ideographs still break per character. It lifts the East Asian line-start
   * and line-end rules instead: a CJK bracket may end a line (E01, E03) and
   * a break is allowed after an opening bracket before Latin text (E08, E10,
   * E12), where `eaLnBrk="1"` keeps the bracket with that text (E09).
   */
  eastAsianLineBreak?: boolean;
  /**
   * Negative tracking can make a longer prefix narrower. Every prefix is then
   * a fit candidate; widths are accumulated from segment heads and
   * per-grapheme advances measured with a bounded left shaping context.
   */
  nonMonotoneMeasure?: boolean;
}

type Atom<T> =
  | { type: 'text'; text: string; style: T; run: number; start: number; end: number }
  | { type: 'tab'; style: T; run: number }
  | { type: 'object'; style: T; run: number; width: number; payload?: unknown };

/** Internal source provenance for format-specific line metrics. It never
 * participates in measurement or visual coalescing. Only emitted segments
 * retain ranges; prefix-fit probes allocate no provenance. Offsets are UTF-16
 * within the input run's display text, before trimming and wrapping. */
export type DrawingMlSourceRange = Readonly<{ run: number; start: number; end: number }>;
const sourceRanges = new WeakMap<object, readonly DrawingMlSourceRange[]>();
export function drawingMlSegmentSourceRanges(segment: object): readonly DrawingMlSourceRange[] {
  return sourceRanges.get(segment) ?? [];
}

const latin = /^\p{Script_Extensions=Latin}$/u;
const isSpace = <T>(atom: Atom<T>): boolean => atom.type === 'text' && atom.text === ' ';
const isCjk = <T>(atom: Atom<T>): boolean =>
  atom.type === 'text' && isCjkBreakChar(atom.text.codePointAt(0) ?? 0);

/**
 * Shared DrawingML soft-wrap phase. ECMA-376 CT_TextBody/CT_TextParagraphProperties
 * defines `wrap`, insets, indent and tab stops, but leaves the choice among
 * Unicode soft opportunities to the host. The Office controls C00–C11 establish
 * the boundaries below at matched geometry in PowerPoint and Excel:
 * authored run seams are not breaks, ordinary spaces separate words, a Latin
 * compound hyphen and CJK/Latin seam are breaks, NBSP is not ordinary space,
 * and an overwide unbreakable word falls back to grapheme boundaries. Both
 * hosts put a tab on its next visual line when its following word will not fit.
 * C08 also shows that a CJK punctuation run seam can remain a break boundary;
 * the kinsoku adjustment therefore stays within an authored run here.
 *
 * The input adapter owns fonts and measurement; this phase never changes a
 * run's style or fabricates a font. The output is logical order. Alignment,
 * bidi, line metrics and paint are later phases.
 */
export function breakDrawingMlText<T>(
  runs: readonly DrawingMlInputRun<T>[],
  options: DrawingMlBreakOptions<T>,
): DrawingMlBrokenLine<T>[] {
  const sameStyle = options.sameStyle ?? ((a: T, b: T) => a === b);
  const regions: Atom<T>[][] = [[]];
  /** Region in which each input run starts. */
  const runRegion: number[] = [];
  /** Input run whose line feed opened each region, else -1. */
  const regionLineFeedRun: number[] = [-1];
  const regionEndBreakRun: number[] = [];
  let runIndex = 0;
  for (const run of runs) {
    runRegion.push(regions.length - 1);
    if (run.type === 'break') {
      regionEndBreakRun[regions.length - 1] = runIndex;
      regions.push([]);
      regionLineFeedRun.push(-1);
      runIndex++;
      continue;
    }
    if (run.type === 'object') {
      // ECMA-376 §22.1 m:oMathPara is display math. It occupies its own visual
      // line: break after preceding text and before following text. An authored
      // hard break already supplies the first edge.
      if (run.display && regions.at(-1)!.length > 0) {
        regions.push([]);
        regionLineFeedRun.push(-1);
      }
      regions.at(-1)!.push({ type: 'object', style: run.style, run: runIndex, width: run.width, payload: run.payload });
      if (run.display) {
        regions.push([]);
        regionLineFeedRun.push(-1);
      }
      runIndex++;
      continue;
    }
    const offsets = [0, ...graphemeClusterOffsets(run.text), run.text.length];
    for (let i = 0; i + 1 < offsets.length; i++) {
      const text = run.text.slice(offsets[i], offsets[i + 1]);
      if (text === '\n') {
        regions.push([]);
        regionLineFeedRun.push(runIndex);
      }
      else if (text === '\t') regions.at(-1)!.push({ type: 'tab', style: run.style, run: runIndex });
      else if (text !== '\r') regions.at(-1)!.push({ type: 'text', text, style: run.style, run: runIndex, start: offsets[i], end: offsets[i + 1] });
    }
    runIndex++;
  }

  const lines: DrawingMlBrokenLine<T>[] = [];
  const regionFirstLine: number[] = [];
  const regionAtomLine: Int32Array[] = [];
  const emitted = new Uint8Array(runs.length);
  for (let regionIndex = 0; regionIndex < regions.length; regionIndex++) {
    let atoms = regions[regionIndex];
    regionFirstLine.push(lines.length);
    // Office C05/C06: paragraph-terminal U+0020 spaces do not create a
    // continuation line, even across authored runs. An authored <a:br> is a
    // line ending inside the paragraph: preserve its preceding spaces because
    // they still advance a centred/right-aligned line (observed in an Office
    // PDF with a space before <a:br>).
    // NBSP stays in either case.
    if (regionIndex === regions.length - 1) {
      while (atoms.length > 0 && isSpace(atoms.at(-1)!)) atoms.pop();
    }

    // Font/paint units and authored style seams may divide one Unicode
    // grapheme. Recover boundaries from contiguous text, retaining each unit's
    // measurement/paint metadata. Regional-indicator pairing can also move a
    // boundary into an existing run-local atom: split that atom at the recovered
    // boundary rather than keeping several paragraph graphemes inseparable.
    // This is the library's grapheme-safe emergency-wrap contract, not an
    // additional Office font-slot rule. Tabs, objects and authored breaks
    // interrupt text. Work/storage stay linear; regions without adjacent
    // text-run seams need no second segmentation.
    let paragraphBounds: number[] | undefined;
    if (atoms.some((atom, i) => i > 0 && atom.type === 'text'
      && atoms[i - 1].type === 'text' && atom.run !== atoms[i - 1].run)) {
      const text = atoms.map((atom) => atom.type === 'text' ? atom.text : '\n').join('');
      paragraphBounds = graphemeClusterOffsets(text);
      paragraphBounds.push(text.length);
      const splitAtoms: Atom<T>[] = [];
      let offset = 0;
      let bound = 0;
      for (const atom of atoms) {
        const stop = offset + (atom.type === 'text' ? atom.text.length : 1);
        while (paragraphBounds[bound] <= offset) bound++;
        let from = offset;
        while (paragraphBounds[bound] < stop) {
          const to = paragraphBounds[bound++];
          if (atom.type === 'text') splitAtoms.push({ ...atom, text: atom.text.slice(from - offset, to - offset), start: atom.start + from - offset, end: atom.start + to - offset });
          from = to;
        }
        splitAtoms.push(atom.type === 'text' && from > offset
          ? { ...atom, text: atom.text.slice(from - offset), start: atom.start + from - offset } : atom);
        offset = stop;
      }
      atoms = regions[regionIndex] = splitAtoms;
    }
    const end = atoms.length;
    const graphemeBoundary = new Uint8Array(end + 1).fill(1);
    if (paragraphBounds) {
      let offset = 0;
      let bound = 0;
      for (let i = 1; i <= end; i++) {
        const atom = atoms[i - 1];
        offset += atom.type === 'text' ? atom.text.length : 1;
        while (paragraphBounds[bound] < offset) bound++;
        graphemeBoundary[i] = paragraphBounds[bound] === offset ? 1 : 0;
      }
    }
    const graphemeStops = [0];
    for (let i = 1; i <= end; i++) if (graphemeBoundary[i]) graphemeStops.push(i);
    const seaBreaks = new Set<number>();
    if (atoms.some((atom) => atom.type === 'text' && containsSeaScript(atom.text))) {
      let full = '';
      const atomOffsets = [0];
      for (const atom of atoms) {
        full += atom.type === 'text' ? atom.text : '\ufffc';
        atomOffsets.push(full.length);
      }
      const offsets = new Set(seaMixedBreakOffsets(full, { cjk: true }));
      for (let i = 1; i < atomOffsets.length; i++) if (offsets.has(atomOffsets[i])) seaBreaks.add(i);
    }

    const eastAsianRules = options.eastAsianLineBreak !== false;
    // Line-break-prohibited adjacency that eaLnBrk="0" keeps (not a CJK rule).
    const gluedCodePoint = (cp: number): boolean => cp === 0xa0 || cp === 0x202f || cp === 0x2060 || cp === 0xfeff;

    const mayBreakAt = (index: number): boolean => {
      if (index <= 0 || index >= end || !graphemeBoundary[index]) return false;
      const prev = atoms[index - 1];
      const next = atoms[index];
      // A tab is a break-before opportunity, not break-after: C11 carries
      // the tab to the next line and seats its first glyph at the 1-inch stop.
      if (next.type === 'tab') return true;
      if (prev.type === 'tab') return false;
      if (isSpace(prev) || isSpace(next)) return true;
      if (seaBreaks.has(index)) return true;
      if (prev.type === 'object' || next.type === 'object') return true;
      const prevCp = [...prev.text].at(-1)?.codePointAt(0);
      const nextCp = next.text.codePointAt(0);
      if (prevCp === undefined || nextCp === undefined) return false;
      if (prevCp === 0x200b) return true;
      // Observed PowerPoint table controls T00–T08: NBSP binds the adjacent
      // words. At a narrow width Office moves the whole "and NBSP Partner"
      // group after the preceding ordinary space; once that group fits, it
      // stays on the first line. The matching Arial controls bound the rule
      // independently of Segoe UI's host-only font metrics.
      if (!eastAsianRules && (isCjk(prev) || isCjk(next))
          && !gluedCodePoint(prevCp) && !gluedCodePoint(nextCp)) return true;
      if (isUax14NoBreakPair(prevCp, nextCp)) return false;
      if (isCjk(prev) || isCjk(next)) return true;
      if (index > 1 && (lineBreakClass(prevCp) === 'HY' || lineBreakClass(prevCp) === 'HH')) {
        const before = atoms[index - 2];
        if (before.type === 'text' && latin.test([...before.text].at(-1) ?? '')
            && latin.test([...next.text][0] ?? '')) return true;
      }
      return false;
    };

    // A continuation of an overwide word can begin at every grapheme. Cache
    // the paragraph-local predicate once; searching the whole remaining word
    // on every continuation made a narrow box quadratic in word length.
    const breakAt = new Uint8Array(end + 1);
    for (let i = 1; i < end; i++) breakAt[i] = mayBreakAt(i) ? 1 : 0;

    // Most DrawingML paragraphs use one resolved paint style and no tabs or
    // objects. Build UTF-16 offsets once so every fit probe measures exactly
    // the same substring that makeSegments would coalesce, without rebuilding
    // the segment array for each candidate prefix. This also preserves shaping
    // across authored run seams when their resolved styles agree.
    const plainStyle = atoms[0]?.style;
    const plainText = end > 0 && atoms.every((atom) => atom.type === 'text'
      && sameStyle(plainStyle, atom.style))
      ? atoms.map((atom) => atom.type === 'text' ? atom.text : '').join('') : null;
    const plainOffsets = plainText === null ? null : new Int32Array(end + 1);
    if (plainOffsets) {
      for (let i = 0; i < end; i++) {
        const atom = atoms[i];
        plainOffsets[i + 1] = plainOffsets[i] + (atom.type === 'text' ? atom.text.length : 0);
      }
    }

    let explicitTabTail: Uint8Array | null = null;
    let explicitTabTailStart = 0;
    const isSingleExplicitTabCell = (index: number): boolean => {
      const tab = atoms[index];
      const first = atoms[index + 1];
      if (tab?.type !== 'tab' || !options.tabStops?.length || first?.type !== 'text'
          || !/^[\p{Script_Extensions=Latin}\p{Number}]$/u.test(first.text)) return false;
      // Cache the unchanged suffix predicate; repeated tab-cell probes must
      // neither copy nor rescan the remaining paragraph on each wrapped line.
      // Line starts only advance, so no later query can reach an already
      // consumed line. Start storage and construction at the current line.
      if (explicitTabTail === null) {
        explicitTabTailStart = start;
        explicitTabTail = new Uint8Array(end - explicitTabTailStart + 1);
        explicitTabTail[end - explicitTabTailStart] = 1;
        for (let i = end - 1; i >= explicitTabTailStart; i--) {
          const atom = atoms[i];
          const offset = i - explicitTabTailStart;
          explicitTabTail[offset] = explicitTabTail[offset + 1] && atom.type === 'text'
            && !isSpace(atom) && !isCjk(atom) ? 1 : 0;
        }
      }
      return explicitTabTail[index + 1 - explicitTabTailStart] === 1;
    };

    let kinsokuText: { chars: string[]; offsets: Int32Array } | null = null;
    let kinsokuTextStart = 0;
    const kinsokuCodePoints = (): { chars: string[]; offsets: Int32Array } => {
      if (kinsokuText === null) {
        // Flatten the unconsumed suffix once, starting at the current line.
        // Line starts only advance: consumed atoms cannot affect kinsoku's
        // line-local retraction floor and must not be scanned or retained.
        // Suffix-relative offsets preserve the code-point split/retraction
        // semantics for multi-code-point graphemes and omitted non-text atoms.
        kinsokuTextStart = start;
        const chars: string[] = [];
        const offsets = new Int32Array(end - kinsokuTextStart + 1);
        for (let i = kinsokuTextStart; i < end; i++) {
          const atom = atoms[i];
          if (atom.type === 'text') for (const ch of atom.text) chars.push(ch);
          offsets[i + 1 - kinsokuTextStart] = chars.length;
        }
        kinsokuText = { chars, offsets };
      }
      return kinsokuText;
    };

    const appendAtom = (segments: DrawingMlLineSegment<T>[], atom: Atom<T>): void => {
      if (atom.type === 'text') {
        const last = segments.at(-1);
        if (last?.type === 'text' && sameStyle(last.style, atom.style)) {
          last.text += atom.text;
          return;
        }
        segments.push({ type: 'text', text: atom.text, style: atom.style, width: 0 });
      } else if (atom.type === 'tab') {
        segments.push({ type: 'tab', style: atom.style, width: 0 });
      } else {
        segments.push({ type: 'object', style: atom.style, width: atom.width, payload: atom.payload });
      }
    };

    const measureSegments = (segments: DrawingMlLineSegment<T>[], lineIndex: number): number => {
      let noStopGap = 0;
      const items = segments.map((seg, i) => {
        if (seg.type === 'text') {
          seg.width = options.measureText(seg.text, seg.style);
          const previous = segments[i - 1];
          if (previous?.type === 'text') {
            seg.width += options.boundaryAdvance?.(previous.style, seg.style) ?? 0;
          }
        }
        if (seg.type === 'tab') {
          // The non-monotone scan reuses this prefix array after measuring its
          // previous length. A tab's earlier resolved gap is not an input to
          // the next candidate's tab-stop resolution.
          seg.width = 0;
          noStopGap = options.measureText(' ', seg.style);
        }
        return { isTab: seg.type === 'tab', width: seg.width };
      });
      if (items.some((item) => item.isTab)) {
        const widths = resolveDrawingMlTabWidths(
          items,
          options.tabStops ?? [],
          options.tabStartPen?.(lineIndex) ?? (lineIndex === 0 ? options.firstLineIndent ?? 0 : 0),
          Infinity,
          noStopGap,
          options.defaultTabSize ?? 0,
        );
        for (let i = 0; i < segments.length; i++) segments[i].width = widths[i];
      }
      return segments.reduce((sum, seg) => sum + seg.width, 0);
    };

    const makeSegments = (start: number, stop: number, lineIndex: number, retainSources = false): DrawingMlBrokenLine<T> => {
      const retain = (segments: DrawingMlLineSegment<T>[]) => {
        if (!retainSources) return;
        let segmentIndex = 0, length = 0;
        let ranges: DrawingMlSourceRange[] = [];
        for (let i = start; i < stop; i++) {
          const atom = atoms[i], segment = segments[segmentIndex];
          if (atom.type !== 'text') { segmentIndex++; continue; }
          const previous = ranges.at(-1);
          if (previous?.run === atom.run && previous.end === atom.start) ranges[ranges.length - 1] = { ...previous, end: atom.end };
          else ranges.push({ run: atom.run, start: atom.start, end: atom.end });
          length += atom.text.length;
          if (segment.type === 'text' && length === segment.text.length) {
            sourceRanges.set(segment, ranges); ranges = []; length = 0; segmentIndex++;
          }
        }
      };
      if (plainText !== null && plainOffsets && plainStyle !== undefined) {
        const text = plainText.slice(plainOffsets[start], plainOffsets[stop]);
        const width = options.measureText(text, plainStyle);
        const segments: DrawingMlLineSegment<T>[] = [{ type: 'text', text, style: atoms[start].style, width }];
        retain(segments);
        return { segments, width };
      }
      const segments: DrawingMlLineSegment<T>[] = [];
      for (let i = start; i < stop; i++) appendAtom(segments, atoms[i]);
      retain(segments);
      return { segments, width: measureSegments(segments, lineIndex) };
    };

    // Negative tracking (`rPr@spc < 0`) can make a longer prefix narrower, so
    // the fit inspects every candidate prefix and keeps the greatest fitting
    // one. Measuring each candidate afresh made a narrow box cubic in run
    // length. The paragraph is measured once instead: each text segment's
    // first SHAPING_CONTEXT + 1 graphemes as one string, then each later
    // grapheme's advance after the preceding SHAPING_CONTEXT graphemes of the
    // same segment. Tracking is a per-grapheme addend, so the accumulated
    // width equals the measured string's width whenever shaping context
    // (kerning pairs, joining) stays within that window. Segment seams, tabs
    // and objects contribute exactly as measureSegments does.
    const SHAPING_CONTEXT = 8;
    const BLOCK = 32;
    interface NonMonotoneModel {
      /** Start of the coalesced text segment containing each text atom. */
      segStart: Int32Array;
      segEnd: Int32Array;
      /** Segment text width through atom t while t is within its head. */
      head: Float64Array;
      /** Advance of atom t after its SHAPING_CONTEXT predecessors. */
      tail: Float64Array;
      tailSum: Float64Array;
      tailBlockMin: Float64Array;
      /** Line width of atoms [0, i) with segments coalesced from atom 0. */
      natural: Float64Array;
      naturalBlockMin: Float64Array;
      /** min(natural[i..end]): a lower bound for tabbed prefixes (tabs never narrow). */
      naturalSuffixMin: Float64Array;
      /** min(tailSum[i..segment end]) for a stop i inside a text segment. */
      tailSuffixMin: Float64Array;
      tabsBefore: Int32Array;
      /** Index into gapValues of each tab atom's no-stop gap (`measureText(' ')`). */
      tabGap: Int32Array;
      gapValues: number[];
    }
    let model: NonMonotoneModel | null = null;
    // nonMonotoneMeasure is fixed for this call: this model is built on the
    // first finite-budget line, before any atoms can be consumed. Its segment,
    // tab and block indexes are then reused; only bounded line heads are lazy.
    const atomText = (from: number, to: number): string => {
      let text = '';
      for (let i = from; i < to; i++) {
        const atom = atoms[i];
        if (atom.type === 'text') text += atom.text;
      }
      return text;
    };
    const blockMins = (values: Float64Array): Float64Array => {
      const mins = new Float64Array(Math.ceil(values.length / BLOCK)).fill(Infinity);
      for (let i = 0; i < values.length; i++) {
        const b = Math.floor(i / BLOCK);
        if (values[i] < mins[b]) mins[b] = values[i];
      }
      return mins;
    };
    const buildModel = (): NonMonotoneModel => {
      const segStart = new Int32Array(end).fill(-1);
      const segEnd = new Int32Array(end).fill(-1);
      const head = new Float64Array(end).fill(NaN);
      const tail = new Float64Array(end);
      const tailSum = new Float64Array(end + 1);
      const natural = new Float64Array(end + 1);
      const tabsBefore = new Int32Array(end + 1);
      let first = -1;
      let style: T | undefined;
      let boundary: number | null = null;
      let base = 0;
      let text = 0;
      const close = (stop: number): void => {
        for (let t = first; first >= 0 && t < stop; t++) segEnd[t] = stop;
      };
      for (let t = 0; t < end; t++) {
        const atom = atoms[t];
        tabsBefore[t + 1] = tabsBefore[t] + (atom.type === 'tab' ? 1 : 0);
        if (!(atom.type === 'text' && first >= 0 && sameStyle(style as T, atom.style))) {
          close(t);
          base = natural[t];
          if (atom.type === 'text') {
            boundary = first >= 0 ? options.boundaryAdvance?.(style as T, atom.style) ?? 0 : null;
            first = t;
            style = atom.style;
          } else {
            first = -1;
            style = undefined;
          }
        }
        if (atom.type === 'text') {
          const segmentStyle = style as T;
          segStart[t] = first;
          if (t - first <= SHAPING_CONTEXT) {
            text = options.measureText(atomText(first, t + 1), segmentStyle);
            head[t] = text;
          } else {
            const context = atomText(t - SHAPING_CONTEXT, t);
            tail[t] = options.measureText(context + atom.text, segmentStyle)
              - options.measureText(context, segmentStyle);
            text += tail[t];
          }
          natural[t + 1] = base + (boundary === null ? text : text + boundary);
        } else {
          natural[t + 1] = base + (atom.type === 'object' ? atom.width : 0);
        }
        tailSum[t + 1] = tailSum[t] + tail[t];
      }
      close(end);
      const naturalSuffixMin = new Float64Array(end + 1);
      naturalSuffixMin[end] = natural[end];
      for (let i = end - 1; i >= 0; i--) naturalSuffixMin[i] = Math.min(natural[i], naturalSuffixMin[i + 1]);
      const tailSuffixMin = new Float64Array(end + 1).fill(Infinity);
      for (let i = end; i >= 1; i--) {
        if (atoms[i - 1].type !== 'text') continue;
        tailSuffixMin[i] = i === segEnd[i - 1] ? tailSum[i] : Math.min(tailSum[i], tailSuffixMin[i + 1]);
      }
      // A tab's no-stop gap only applies without a default grid. The last tab
      // of a prefix supplies the gap for every tab in it (measureSegments).
      const tabGap = new Int32Array(end);
      const gapValues: number[] = [];
      for (let t = 0; t < end; t++) {
        const atom = atoms[t];
        if (atom.type !== 'tab' || (options.defaultTabSize ?? 0) > 0) continue;
        const gap = options.measureText(' ', atom.style);
        let index = gapValues.indexOf(gap);
        if (index < 0) index = gapValues.push(gap) - 1;
        tabGap[t] = index;
      }
      if (gapValues.length === 0) gapValues.push(0);
      return {
        segStart, segEnd, head, tail, tailSum,
        tailBlockMin: blockMins(tailSum),
        natural,
        naturalBlockMin: blockMins(natural),
        naturalSuffixMin,
        tailSuffixMin,
        tabsBefore,
        tabGap,
        gapValues,
      };
    };

    /** Greatest i in (from, to] with offset + values[i] <= budget, else -1. */
    const lastFitting = (
      values: Float64Array, mins: Float64Array, offset: number, from: number, to: number, budget: number,
    ): number => {
      let i = to;
      while (i > from) {
        const lo = Math.floor(i / BLOCK) * BLOCK;
        if (lo > from && offset + mins[lo / BLOCK] > budget) {
          i = lo - 1;
          continue;
        }
        if (graphemeBoundary[i] && offset + values[i] <= budget) return i;
        i--;
      }
      return -1;
    };

    /** Text width of atoms [lineStart, stop) within one coalesced segment. */
    const lineHeadWidths = new Map<number, number>();
    const firstSegmentWidth = (m: NonMonotoneModel, lineStart: number, stop: number): number => {
      if (lineStart === m.segStart[lineStart]) {
        const last = Math.min(stop, lineStart + SHAPING_CONTEXT + 1) - 1;
        return m.head[last] + (m.tailSum[stop] - m.tailSum[last + 1]);
      }
      const headStop = Math.min(stop, lineStart + SHAPING_CONTEXT + 1);
      let width = lineHeadWidths.get(headStop);
      if (width === undefined) {
        width = options.measureText(atomText(lineStart, headStop), atoms[lineStart].style);
        lineHeadWidths.set(headStop, width);
      }
      return width + (m.tailSum[stop] - m.tailSum[headStop]);
    };

    const nonMonotoneFit = (lineStart: number, lineIndex: number, budget: number): number => {
      const m = model ??= buildModel();
      lineHeadWidths.clear();
      if (m.tabsBefore[end] !== m.tabsBefore[lineStart]) {
        return tabbedNonMonotoneFit(m, lineStart, lineIndex, budget);
      }
      const first = atoms[lineStart];
      const firstEnd = first.type === 'text' ? m.segEnd[lineStart] : lineStart + 1;
      const firstWidth = first.type === 'text' ? firstSegmentWidth(m, lineStart, firstEnd)
        : first.type === 'object' ? first.width : 0;
      if (firstEnd < end) {
        // The next segment's seam advance is relative to this line's first style.
        const next = atoms[firstEnd];
        const seam = first.type === 'text' && next.type === 'text' && options.boundaryAdvance
          ? options.boundaryAdvance(first.style, next.style)
            - options.boundaryAdvance(atoms[m.segStart[lineStart]].style, next.style)
          : 0;
        const found = lastFitting(m.natural, m.naturalBlockMin,
          firstWidth + seam - m.natural[firstEnd], firstEnd, end, budget);
        if (found >= 0) return found;
      }
      if (graphemeBoundary[firstEnd] && firstWidth <= budget) return firstEnd;
      if (first.type !== 'text') return lineStart;
      const headStop = lineStart + SHAPING_CONTEXT + 1;
      if (firstEnd - 1 > headStop) {
        const offset = firstSegmentWidth(m, lineStart, headStop) - m.tailSum[headStop];
        const found = lastFitting(m.tailSum, m.tailBlockMin, offset, headStop, firstEnd - 1, budget);
        if (found >= 0) return found;
      }
      for (let i = Math.min(firstEnd - 1, headStop); i > lineStart; i--) {
        if (graphemeBoundary[i] && firstSegmentWidth(m, lineStart, i) <= budget) return i;
      }
      return lineStart;
    };

    // Tabs make each prefix's width depend on its stop resolution. A stop
    // depends only on the pen before its tab, and only the last tab of a
    // prefix still depends on the (open) cell after it, so the resolved pen is
    // carried along the scan instead of re-resolving every candidate. One
    // state is kept per distinct no-stop gap, because the prefix's last tab
    // supplies that gap to all of its tabs. Tabs never narrow a line, so the
    // scan stops once the tab-free width of every later prefix overflows.
    const tabbedNonMonotoneFit = (
      m: NonMonotoneModel, lineStart: number, lineIndex: number, budget: number,
    ): number => {
      const startPen = options.tabStartPen?.(lineIndex) ?? (lineIndex === 0 ? options.firstLineIndent ?? 0 : 0);
      const stops = options.tabStops ?? [];
      const defaultTab = options.defaultTabSize ?? 0;
      const gaps = m.gapValues;
      const states = gaps.length;
      const closed = new Float64Array(states);
      const pen = new Float64Array(states).fill(startPen);
      const tabPen = new Float64Array(states);
      const stopPos = new Float64Array(states);
      const fraction = new Float64Array(states);
      const gapTab = new Uint8Array(states);
      const cell = new Float64Array(states);
      let hasTab = false;
      let lastGap = 0;
      let open = 0;
      let openIsItem = false;

      const first = atoms[lineStart];
      const firstEnd = first.type === 'text' ? m.segEnd[lineStart] : lineStart + 1;
      const firstWidth = first.type === 'text' ? firstSegmentWidth(m, lineStart, firstEnd)
        : first.type === 'object' ? first.width : 0;
      const next = atoms[firstEnd];
      const seam = firstEnd < end && first.type === 'text' && next.type === 'text' && options.boundaryAdvance
        ? options.boundaryAdvance(first.style, next.style)
          - options.boundaryAdvance(atoms[m.segStart[lineStart]].style, next.style)
        : 0;
      const restOffset = firstWidth + seam - m.natural[firstEnd];
      const headStop = lineStart + SHAPING_CONTEXT + 1;
      const tailOffset = first.type === 'text' && firstEnd > headStop
        ? firstSegmentWidth(m, lineStart, headStop) - m.tailSum[headStop] : 0;
      const prunable = gaps.every((gap) => gap >= 0);
      const slack = 1e-9 * Math.max(1, Math.abs(budget));
      const noLaterFit = (i: number): boolean => {
        if (!prunable || i < headStop && i < firstEnd) return false;
        let bound = Infinity;
        if (i < firstEnd) bound = tailOffset + m.tailSuffixMin[i + 1];
        if (firstEnd < end) bound = Math.min(bound, restOffset + m.naturalSuffixMin[Math.max(i + 1, firstEnd + 1)]);
        return bound > budget + slack;
      };

      let fit = lineStart;
      let segFirst = -1;
      let boundary: number | null = null;
      for (let i = lineStart + 1; i <= end; i++) {
        const index = i - 1;
        const atom = atoms[index];
        if (atom.type === 'text' && segFirst >= 0 && m.segStart[index] === m.segStart[index - 1]) {
          const text = firstSegmentWidthFrom(m, segFirst, lineStart, i);
          open = boundary === null ? text : text + boundary;
        } else {
          if (openIsItem) {
            for (let g = 0; g < states; g++) {
              if (hasTab) cell[g] += open;
              else { closed[g] += open; pen[g] += open; }
            }
          }
          if (atom.type === 'tab') {
            for (let g = 0; g < states; g++) {
              if (hasTab) {
                // The previous tab's cell is now complete.
                const followingWidth = cell[g];
                let target = tabPen[g] + gaps[g];
                if (!gapTab[g]) {
                  target = stopPos[g] - followingWidth * fraction[g];
                  if (target < tabPen[g]) target = tabPen[g];
                }
                closed[g] += target - tabPen[g] + followingWidth;
                pen[g] = target + followingWidth;
              }
              tabPen[g] = pen[g];
              let stop: DrawingMlTabStop | null = null;
              for (const candidate of stops) {
                if (candidate.pos > pen[g] && (stop === null || candidate.pos < stop.pos)) stop = candidate;
              }
              if (stop === null && defaultTab > 0) {
                stop = { pos: (Math.floor(pen[g] / defaultTab) + 1) * defaultTab, algn: 'l' };
              }
              gapTab[g] = stop === null ? 1 : 0;
              stopPos[g] = stop?.pos ?? 0;
              fraction[g] = stop?.algn === 'ctr' ? 0.5 : stop?.algn === 'r' || stop?.algn === 'dec' ? 1 : 0;
              cell[g] = 0;
            }
            hasTab = true;
            lastGap = m.tabGap[index];
            segFirst = -1;
            open = 0;
            openIsItem = false;
          } else if (atom.type === 'text') {
            const previous = segFirst >= 0 ? atoms[segFirst] : undefined;
            boundary = previous?.type === 'text' ? options.boundaryAdvance?.(previous.style, atom.style) ?? 0 : null;
            segFirst = index;
            const text = firstSegmentWidthFrom(m, segFirst, lineStart, i);
            open = boundary === null ? text : text + boundary;
            openIsItem = true;
          } else {
            segFirst = -1;
            open = atom.width;
            openIsItem = true;
          }
        }
        let width: number;
        if (!hasTab) {
          width = closed[lastGap] + open;
        } else {
          const followingWidth = cell[lastGap] + open;
          let tabWidth = gaps[lastGap];
          if (!gapTab[lastGap]) {
            let target = stopPos[lastGap] - followingWidth * fraction[lastGap];
            if (target < tabPen[lastGap]) target = tabPen[lastGap];
            tabWidth = target - tabPen[lastGap];
          }
          width = closed[lastGap] + tabWidth + followingWidth;
        }
        if (graphemeBoundary[i] && width <= budget) {
          if (i === end) return end;
          fit = i;
        }
        if (i < end && noLaterFit(i)) break;
      }
      return fit;
    };
    /** Width of segment text [segFirst, stop), where segFirst is lineStart or a natural start. */
    const firstSegmentWidthFrom = (
      m: NonMonotoneModel, segFirst: number, lineStart: number, stop: number,
    ): number => segFirst === lineStart
      ? firstSegmentWidth(m, lineStart, stop)
      : firstSegmentWidth(m, segFirst, stop);

    const atomLine = new Int32Array(end).fill(-1);
    regionAtomLine.push(atomLine);
    const markLine = (from: number, to: number): void => {
      for (let i = from; i < to; i++) {
        atomLine[i] = lines.length - 1;
        emitted[atoms[i].run] = 1;
      }
    };
    let start = 0;
    if (end === 0) {
      lines.push({
        segments: [], width: 0,
        ...(regionIndex + 1 < regions.length ? { endsWithBreak: true } : {}),
        ...(regionLineFeedRun[regionIndex] >= 0 ? { lineFeedRun: regionLineFeedRun[regionIndex] } : {}),
        ...(regionEndBreakRun[regionIndex] !== undefined ? { endBreakRun: regionEndBreakRun[regionIndex] } : {}),
      });
      continue;
    }
    let stopIndex = 0;
    while (start < end) {
      while (graphemeStops[stopIndex] < start) stopIndex++;
      const lineIndex = lines.length;
      const budget = options.maxWidth - (lineIndex === 0 ? options.firstLineIndent ?? 0 : 0);
      const widthAt = (stop: number): number =>
        plainText !== null && plainOffsets && plainStyle !== undefined
          ? options.measureText(plainText.slice(plainOffsets[start], plainOffsets[stop]), plainStyle)
          : makeSegments(start, stop, lineIndex).width;
      let fit = end;
      if (Number.isFinite(budget) && options.nonMonotoneMeasure) {
        // A later prefix can fit after an earlier overflow, so every candidate
        // is still checked and the greatest fitting one wins (the whole line
        // when it fits). Prefix widths are accumulated, not re-measured.
        fit = nonMonotoneFit(start, lineIndex, budget);
      } else if (Number.isFinite(budget)) {
        // Exponential search only measures prefixes close to the fit boundary.
        // In a one-grapheme box this avoids measuring the whole remaining word
        // on every visual line. The existing binary search still chooses the
        // greatest fitting prefix when advances are monotone.
        let lo = stopIndex;
        let hi = graphemeStops.length - 1;
        let step = 1;
        while (lo < graphemeStops.length - 1) {
          const probe = Math.min(graphemeStops.length - 1, stopIndex + step);
          if (widthAt(graphemeStops[probe]) <= budget) {
            lo = probe;
            if (probe === graphemeStops.length - 1) break;
            step *= 2;
          } else {
            hi = probe - 1;
            break;
          }
        }
        while (lo < hi) {
          const mid = Math.ceil((lo + hi) / 2);
          if (widthAt(graphemeStops[mid]) <= budget) lo = mid;
          else hi = mid - 1;
        }
        fit = graphemeStops[lo];
      }
      if (fit >= end) {
        const line = makeSegments(start, end, lineIndex, true);
        if (regionIndex + 1 < regions.length) line.endsWithBreak = true;
        lines.push(line);
        markLine(start, end);
        break;
      }

      let split = 0;
      for (let i = start + 1; i <= fit; i++) if (breakAt[i]) split = i;
      if (split > start && isSingleExplicitTabCell(split)) {
        // Office keeps the tab and following Latin cell on the authored line
        // for an explicit stop, even when the cell passes the box edge. The
        // control C11 uses the default grid and wraps before its tab instead.
        split = end;
      }
      if (isSingleExplicitTabCell(start)) {
        // An authored stop makes the next Latin cell one visual unit even if
        // that cell contains a hyphen or another soft opportunity. PowerPoint
        // keeps this cell together in its explicit-stop paragraphs; C11 has
        // only the default grid and still breaks within the following word.
        split = start + 2;
        while (split < end && !isSpace(atoms[split]) && atoms[split].type !== 'tab') split++;
      }
      if (split === 0) {
        if (atoms[start].type === 'tab' && fit <= start + 1) {
          // A stop beyond the box still carries its next cell. A CJK cell
          // contributes at least one glyph; an indivisible Latin cell stays
          // together. The painter can clamp an explicit off-box stop.
          split = start + 1;
          if (split < end && atoms[split].type === 'text') {
            split++;
            while (split < end && !breakAt[split]) split++;
          }
        } else {
          // Overwide word: grapheme-safe emergency break. PowerPoint control
          // E06 also splits a Latin word that spans a font seam (C01 splits
          // one with identical styles), so no style seam keeps it whole.
          split = Math.max(graphemeStops[stopIndex + 1], fit);
        }
      }

      while (split < end && !graphemeBoundary[split]) split++;

      // Kinsoku (§17.15.1.58–.60) adjusts an in-run CJK boundary. The Office
      // C08 control is a counterexample at an authored run seam, so leave that
      // boundary intact instead of inferring a cross-run retraction rule.
      // eaLnBrk="0" lifts the East Asian line-start/line-end rules (E01).
      if (eastAsianRules && split < end && atoms[split - 1].run === atoms[split].run
          && (isCjk(atoms[split - 1]) || isCjk(atoms[split]))) {
        const { chars, offsets } = kinsokuCodePoints();
        const codeStart = offsets[start - kinsokuTextStart];
        const codeSplit = offsets[split - kinsokuTextStart];
        if (codeSplit - codeStart > 1 && codeSplit < chars.length) {
          const adjusted = kinsokuAdjustedSplit(chars, codeSplit, DEFAULT_KINSOKU_RULES, codeStart + 1);
          if (adjusted < codeSplit) {
            // Kinsoku returns a code-point position; atoms may hold a whole
            // grapheme or a styled piece. Map back before enforcing the
            // paragraph grapheme boundary, keeping the cached suffix linear.
            let candidate = split;
            while (candidate > start && offsets[candidate - kinsokuTextStart] > adjusted) candidate--;
            while (candidate > start && !graphemeBoundary[candidate]) candidate--;
            if (candidate > start) split = candidate;
          }

        }
      }

      // Keep authored spaces on the closed line. The controlled PDFs establish
      // the visible word break, but cannot identify the advance of invisible
      // trailing spaces; retaining them preserves that unresolved paint detail.
      lines.push(makeSegments(start, split, lineIndex, true));
      markLine(start, split);
      start = split;
      while (start < end && isSpace(atoms[start]) && graphemeBoundary[start + 1]) start++;
    }
  }

  for (let region = 0; region < regionEndBreakRun.length; region++) {
    const run = regionEndBreakRun[region];
    if (run === undefined) continue;
    const first = regionFirstLine[region];
    const after = region + 1 < regionFirstLine.length ? regionFirstLine[region + 1] : lines.length;
    if (after > first) lines[after - 1].endBreakRun = run;
  }

  // A text run that paints nothing sits on the line holding the preceding
  // emitted atom of its region, else on the region's first line.
  for (let index = 0; index < runs.length; index++) {
    if (runs[index].type !== 'text' || emitted[index]) continue;
    const region = runRegion[index];
    const atoms = regions[region];
    const atomLine = regionAtomLine[region];
    // Atoms are in run order: find the first atom of a later run, then the
    // nearest emitted atom before it.
    let lo = 0;
    let hi = atoms.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (atoms[mid].run < index) lo = mid + 1;
      else hi = mid;
    }
    let line = regionFirstLine[region];
    for (let i = lo - 1; i >= 0; i--) {
      if (atomLine[i] >= 0) { line = atomLine[i]; break; }
    }
    (lines[line].hiddenRuns ??= []).push(index);
  }
  return lines;
}
