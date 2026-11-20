import type { MatchRunSlice, TextSelectionContextOptions } from '@silurus/ooxml-core';
import type { PptxTextRunInfo } from './renderer.js';
import type { Slide, SlideElement, TextBody } from './types.js';
import { preparedPowerPointText } from './font-display-text.js';
import { readPptxTextSelectionContext } from './selection-context.js';

/** Fixed-version slide-owned body occurrence; table cells have separate identities. */
export interface PptxTextBodyIdentity {
  readonly slideIndex: number;
  readonly slidePart: string;
  readonly shapeId: string;
  readonly cell?: Readonly<{ row: number; column: number }>;
}

/** UTF-16 source slice, with one LF for each authored break and paragraph boundary. */
export interface PptxTextSource {
  readonly body: PptxTextBodyIdentity;
  readonly context: Readonly<{ text: string; nonpainted?: readonly Readonly<{ start: number; end: number }>[] }>;
  readonly start: number;
  readonly end: number;
}

/** Shared body contexts keep structured-clone worker transfers linear in source size. */
export type PptxTextSourceCache = Map<string, Pick<PptxTextSource, 'body' | 'context'> | undefined> & {
  uniqueShapeIds?: ReadonlySet<string>;
};

/** Asset-version-bound anchor. Callers must bind it to the exact asset version. */
export interface PptxTextAnchor {
  readonly type: 'pptx-text';
  readonly version: 1;
  readonly body: PptxTextBodyIdentity;
  readonly start: number;
  readonly end: number;
  readonly exact: string;
  readonly prefix: string;
  readonly suffix: string;
  readonly bodyLength: number;
}

/** All-or-nothing source anchoring failure; resolution never searches for replacement text. */
export type PptxTextAnchorFailure = {
  readonly isSuccess: false;
  readonly reason: 'invalid-range' | 'unsupported' | 'cross-body' | 'source-not-found' | 'text-mismatch' | 'ambiguous' | 'incomplete';
};

/** Library source policy for DrawingML a:p/a:r/a:br (ECMA-376 §21.1.2).
 * Fields, equations and substituted glyphs cannot promise exact source projection. */
export function pptxTextBodyText(body: TextBody): string | undefined {
  const paragraphs: string[] = [];
  for (const paragraph of body.paragraphs) {
    let text = '';
    for (const run of paragraph.runs) {
      if (run.type === 'break') text += '\n';
      else if (run.type === 'text' && !run.fieldType && preparedPowerPointText(run) === run.text &&
        !/[\uf020-\uf0ff]/u.test(run.text)) text += run.text;
      else return undefined;
    }
    paragraphs.push(text);
  }
  return paragraphs.join('\n');
}

/** Attach provenance only to unique slide-owned bodies and exact native slices. */
export function sourcePptxTextRun(slide: Slide, element: SlideElement, run: PptxTextRunInfo, cache: PptxTextSourceCache): PptxTextSource | undefined {
  if (run.origin !== 'slide' || run.elementIndex === undefined ||
    slide.elementSources?.[run.elementIndex]?.origin !== 'slide' || !slide.partName || !element.id ||
    run.sourceStart === undefined || run.sourceEnd === undefined) return undefined;
  if (!cache.uniqueShapeIds) {
    const ids = new Set<string>();
    const duplicates = new Set<string>();
    for (const [index, candidate] of slide.elements.entries()) {
      if (!candidate.id || slide.elementSources?.[index]?.origin !== 'slide') continue;
      if (ids.has(candidate.id)) duplicates.add(candidate.id);
      ids.add(candidate.id);
    }
    for (const id of duplicates) ids.delete(id);
    cache.uniqueShapeIds = ids;
  }
  const key = `${run.elementIndex}:${run.tableCell?.row ?? ''}:${run.tableCell?.column ?? ''}`;
  if (!cache.has(key)) {
    const body = element.type === 'shape' ? element.textBody : element.type === 'table' && run.tableCell
      ? element.rows[run.tableCell.row]?.cells[run.tableCell.column]?.textBody : undefined;
    const text = body ? pptxTextBodyText(body) : undefined;
    const isUnique = cache.uniqueShapeIds.has(element.id);
    cache.set(key, text === undefined || !isUnique ? undefined : {
      body: { slideIndex: slide.index, slidePart: slide.partName, shapeId: element.id,
        ...(run.tableCell ? { cell: { ...run.tableCell } } : {}) },
      context: { text, nonpainted: run.sourceNonpainted },
    });
  }
  const source = cache.get(key);
  if (!source || source.context.text.slice(run.sourceStart, run.sourceEnd) !== run.text ||
    run.sourceEnd - run.sourceStart !== run.text.length) return undefined;
  return { ...source, start: run.sourceStart, end: run.sourceEnd };
}

/** Compare source occurrences, never rendered-run indexes or text. */
function isSameBody(left: PptxTextBodyIdentity, right: PptxTextBodyIdentity): boolean {
  return left.slideIndex === right.slideIndex && left.slidePart === right.slidePart &&
    left.shapeId === right.shapeId && left.cell?.row === right.cell?.row && left.cell?.column === right.cell?.column;
}

/** Reject UTF-16 offsets inside surrogate pairs or grapheme clusters. */
function isBoundary(text: string, offset: number): boolean {
  if (!Number.isSafeInteger(offset) || offset < 0 || offset > text.length) return false;
  if (offset === 0 || offset === text.length) return true;
  for (const { index } of new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text)) {
    if (index >= offset) return index === offset;
  }
  return false;
}

/** Validate stored source identity without coercion. */
function isBody(body: PptxTextBodyIdentity): boolean {
  return !!body && Number.isSafeInteger(body.slideIndex) && body.slideIndex >= 0 &&
    typeof body.slidePart === 'string' && /^ppt\/slides\/[^/]+\.xml$/u.test(body.slidePart) &&
    typeof body.shapeId === 'string' && body.shapeId.length > 0 &&
    (body.cell === undefined || (!!body.cell && Number.isSafeInteger(body.cell.row) && body.cell.row >= 0 &&
      Number.isSafeInteger(body.cell.column) && body.cell.column >= 0));
}

/** Only native layout omissions, TABs and authored line breaks may lack paint callbacks. */
function isNonpaintedGap(context: PptxTextSource['context'], start: number, end: number): boolean {
  const gap = context.text.slice(start, end);
  return /^[\n\t]*$/u.test(gap) || (/^[\n\t ]*$/u.test(gap) &&
    !!context.nonpainted?.some((range) => range.start <= start && range.end >= end));
}

/** Resolve against a complete current native run snapshot, independent of wrapping and geometry. */
export function resolvePptxTextAnchor(anchor: PptxTextAnchor, runs: readonly PptxTextRunInfo[]):
  { readonly isSuccess: true; readonly slices: MatchRunSlice[] } | PptxTextAnchorFailure {
  if (!anchor || anchor.type !== 'pptx-text' || anchor.version !== 1 || !isBody(anchor.body) ||
    typeof anchor.exact !== 'string' || !anchor.exact || typeof anchor.prefix !== 'string' ||
    typeof anchor.suffix !== 'string' || !Number.isSafeInteger(anchor.bodyLength) ||
    !Number.isSafeInteger(anchor.start) || !Number.isSafeInteger(anchor.end) ||
    anchor.start < 0 || anchor.end <= anchor.start || anchor.end > anchor.bodyLength ||
    anchor.exact.length !== anchor.end - anchor.start || anchor.prefix.length !== Math.min(32, anchor.start) ||
    anchor.suffix.length !== Math.min(32, anchor.bodyLength - anchor.end)) return { isSuccess: false, reason: 'invalid-range' };
  const located = runs.flatMap((run, runIndex) => run.source && isSameBody(anchor.body, run.source.body)
    ? [{ run, runIndex, source: run.source }] : []);
  if (!located.length) return { isSuccess: false, reason: 'source-not-found' };
  const text = located[0].source.context.text;
  if (!isBoundary(text, anchor.start) || !isBoundary(text, anchor.end) ||
    text.length !== anchor.bodyLength || text.slice(anchor.start, anchor.end) !== anchor.exact ||
    text.slice(Math.max(0, anchor.start - 32), anchor.start) !== anchor.prefix ||
    text.slice(anchor.end, anchor.end + 32) !== anchor.suffix ||
    located.some(({ run, source }) => !Number.isSafeInteger(source.start) || !Number.isSafeInteger(source.end) ||
      source.start < 0 || source.end > text.length || source.context.text !== text ||
      source.end - source.start !== run.text.length || text.slice(source.start, source.end) !== run.text)) {
    return { isSuccess: false, reason: 'text-mismatch' };
  }
  const slices: MatchRunSlice[] = [];
  let cursor = anchor.start;
  for (const { runIndex, source } of located.sort((a, b) => a.source.start - b.source.start)) {
    const start = Math.max(anchor.start, source.start);
    const end = Math.min(anchor.end, source.end);
    if (end <= start) continue;
    if (start < cursor) return { isSuccess: false, reason: 'ambiguous' };
    if (!isNonpaintedGap(located[0].source.context, cursor, start)) return { isSuccess: false, reason: 'incomplete' };
    slices.push({ runIndex, start: start - source.start, end: end - source.start });
    cursor = end;
  }
  if (!slices.length || !isNonpaintedGap(located[0].source.context, cursor, anchor.end)) return { isSuccess: false, reason: 'incomplete' };
  return { isSuccess: true, slices };
}

/** Convert exact snapshot slices into one body anchor, rejecting omitted or unsupported slices. */
export function createPptxTextAnchor(runs: readonly PptxTextRunInfo[], slices: readonly MatchRunSlice[]):
  { readonly isSuccess: true; readonly anchor: PptxTextAnchor } | PptxTextAnchorFailure {
  if (!slices.length) return { isSuccess: false, reason: 'invalid-range' };
  const source = runs[slices[0].runIndex]?.source;
  if (!source) return { isSuccess: false, reason: 'unsupported' };
  let start = Infinity;
  let end = -1;
  for (const slice of slices) {
    const run = runs[slice.runIndex];
    if (!run?.source) return { isSuccess: false, reason: 'unsupported' };
    if (!isSameBody(source.body, run.source.body)) return { isSuccess: false, reason: 'cross-body' };
    if (!Number.isSafeInteger(slice.runIndex) || !isBoundary(run.text, slice.start) ||
      !isBoundary(run.text, slice.end) || slice.start >= slice.end) return { isSuccess: false, reason: 'invalid-range' };
    start = Math.min(start, run.source.start + slice.start);
    end = Math.max(end, run.source.start + slice.end);
  }
  const anchor: PptxTextAnchor = {
    type: 'pptx-text', version: 1, body: { ...source.body, ...(source.body.cell ? { cell: { ...source.body.cell } } : {}) },
    start, end, exact: source.context.text.slice(start, end), bodyLength: source.context.text.length,
    prefix: source.context.text.slice(Math.max(0, start - 32), start), suffix: source.context.text.slice(end, end + 32),
  };
  const resolved = resolvePptxTextAnchor(anchor, runs);
  if (!resolved.isSuccess) return resolved;
  if (resolved.slices.length !== slices.length || resolved.slices.some((slice) => !slices.some((candidate) =>
    candidate.runIndex === slice.runIndex && candidate.start === slice.start && candidate.end === slice.end))) {
    return { isSuccess: false, reason: 'incomplete' };
  }
  return { isSuccess: true, anchor };
}

/** Native text-layer registry; discarded spans cannot retain a presentation. */
export const pptxTextSpanRuns = new WeakMap<HTMLElement, PptxTextRunInfo>();

/** Read one untruncated native DOM range, rejecting cross-viewer and unsupported content. */
export function readPptxTextAnchor(root: HTMLElement, selection: Selection | null, options: TextSelectionContextOptions = {}):
  { readonly isSuccess: true; readonly anchor: PptxTextAnchor } | PptxTextAnchorFailure {
  if (!selection || selection.rangeCount !== 1 || selection.isCollapsed) return { isSuccess: false, reason: 'invalid-range' };
  const context = readPptxTextSelectionContext(root, selection, options);
  const range = selection.getRangeAt(0);
  if (!root.contains(range.startContainer) || !root.contains(range.endContainer)) return { isSuccess: false, reason: 'cross-body' };
  if (!context || context.truncated) return { isSuccess: false, reason: 'incomplete' };
  const spans = [...root.querySelectorAll<HTMLElement>('[data-ooxml-selection-run="pptx"]')].filter((span) => range.intersectsNode(span));
  const runs: PptxTextRunInfo[] = [];
  const slices: MatchRunSlice[] = [];
  let selected = '';
  for (const span of spans) {
    const run = pptxTextSpanRuns.get(span);
    if (!run || span.textContent !== run.text) return { isSuccess: false, reason: 'unsupported' };
    const before = root.ownerDocument.createRange();
    before.selectNodeContents(span);
    if (span.contains(range.startContainer)) before.setEnd(range.startContainer, range.startOffset);
    const start = span.contains(range.startContainer) ? before.toString().length : 0;
    before.selectNodeContents(span);
    if (span.contains(range.endContainer)) before.setEnd(range.endContainer, range.endOffset);
    const end = span.contains(range.endContainer) ? before.toString().length : run.text.length;
    if (end <= start) continue;
    slices.push({ runIndex: runs.length, start, end });
    runs.push(run);
    selected += run.text.slice(start, end);
  }
  if (selected !== range.toString()) return { isSuccess: false, reason: 'incomplete' };
  return createPptxTextAnchor(runs, slices);
}
