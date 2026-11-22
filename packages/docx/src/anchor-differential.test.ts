import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { textRunsForPage } from './text-run-projection.js';
import { buildDocxHighlightLayer } from './find-highlight-layer.js';
import { resolveCommentAnchorRuns } from './comments.js';
import { visualBoundaryXPt, visualClusterStartXPt } from './layout/visual-cluster-geometry.js';
import type { AnchorAcquisitionInput } from './layout/anchor-input.js';
import type { ParagraphLayout } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, DocxTextRunInfo } from './types.js';

/**
 * WORD_ANCHOR_HOST_NO_LINE_METRIC, differentially: a floating drawing anchored
 * at any source boundary of a paragraph changes nothing but the drawing. For
 * a generated set of paragraphs, every sampled boundary receives an anchor
 * and the result is compared with the same paragraph without it:
 *
 * - pagination, blocks, line boxes and grid cells;
 * - cluster positions, paint operations and decorations;
 * - selection rectangles and their source runs, a find highlight for every
 *   single character (the per-character extent a caret offset also uses),
 *   and comment anchors over every run and at every run boundary;
 * - the host of the drawing, which stands at the visual position of that
 *   boundary in the anchor-free layout.
 */

type Run = DocParagraph['runs'][number];
type RunSpec = readonly [text: string, extra?: object];

/** Half-em advances, or `ems` of the em for the characters it lists. */
function measureContext(ems: Readonly<Record<string, number>> = {}): CanvasRenderingContext2D {
  let font = '10px serif';
  const size = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (value: string) => ({
      width: [...value].reduce((sum, character) => sum + (ems[character] ?? 0.5), 0) * size(),
      fontBoundingBoxAscent: size() * 0.8, fontBoundingBoxDescent: size() * 0.2,
      actualBoundingBoxAscent: size() * 0.8, actualBoundingBoxDescent: size() * 0.2,
    } as TextMetrics),
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, textAlign: 'left', direction: 'ltr',
  } as unknown as CanvasRenderingContext2D;
}

const textRun = (value: string, extra: object = {}): Run => ({
  type: 'text', text: value, bold: false, italic: false, underline: false, strikethrough: false,
  fontSize: 12, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null, vertAlign: null,
  hyperlink: null, ...extra,
}) as unknown as Run;

const missingEdges = { topPt: null, topStatus: 'missing', rightPt: null, rightStatus: 'missing',
  bottomPt: null, bottomStatus: 'missing', leftPt: null, leftStatus: 'missing' };
type AnchorKind = 'page' | 'character';
/** A wrapNone drawing: positioned from the page, or from its anchoring
 * character and line with zero offsets. */
const acquisition = (kind: AnchorKind): AnchorAcquisitionInput => ({
  occurrenceId: 'subject',
  simplePosition: { enabled: false, status: 'valid', xPt: 0, xStatus: 'valid', yPt: 0, yStatus: 'valid' },
  horizontal: { relativeFrom: kind, relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: kind === 'page' ? 150 : 0 } },
  vertical: { relativeFrom: kind === 'page' ? 'paragraph' : 'line', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: 0 } },
  extent: { widthPt: 5, heightPt: 7, widthStatus: 'valid', heightStatus: 'valid' },
  parentEffectExtent: missingEdges, anchorDistances: missingEdges,
  relativeSize: { horizontal: null, vertical: null },
  wrap: { kind: 'none', authoredKinds: ['wrapNone'], side: null, distances: missingEdges, effectExtent: null, polygon: null },
  behavior: { behindDoc: false, behindDocStatus: 'valid', relativeHeight: 1, relativeHeightStatus: 'valid',
    locked: false, lockedStatus: 'valid', allowOverlap: true, allowOverlapStatus: 'valid',
    layoutInCell: true, layoutInCellStatus: 'valid' },
  group: null,
}) as unknown as AnchorAcquisitionInput;
const anchorRuns = (kind: AnchorKind): Run[] => [
  { type: 'anchorHost', fontSize: 40, bold: true, __anchorOccurrenceId: 'subject' } as unknown as Run,
  { type: 'image', imagePath: 'word/media/p.png', mimeType: 'image/png', widthPt: 5, heightPt: 7, anchor: true,
    anchorXPt: 0, anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: false,
    __anchorAcquisition: acquisition(kind) } as unknown as Run,
];

const paragraph = (runs: readonly Run[], options: object = {}): BodyElement => ({
  type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0,
  spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [], defaultFontSize: 12,
  defaultFontFamily: 'NotInMetrics', runs: [...runs], ...options,
}) as unknown as BodyElement;
const table = (content: BodyElement[]): BodyElement => ({
  type: 'table', colWidths: [120], borders: { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null },
  cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'fixed',
  rows: [{ rowHeight: null, rowHeightRule: 'auto', isHeader: false, cells: [{ content, colSpan: 1, vMerge: null,
    borders: { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null },
    background: null, vAlign: 'top', widthPt: 120 }] }],
}) as unknown as BodyElement;

// ── The generated set ───────────────────────────────────────────────────────

const rtl = { rtl: true };
const kashidaText = '\u0628\u062a\u062a\u0628\u062a \u0628\u062a\u062a\u0628\u062a \u0628\u062a\u062a\u0628\u062a';
const scripts: readonly (readonly [name: string, runs: readonly RunSpec[], options?: object])[] = [
  ['left-to-right', [['ab cd '], ['efgh ij']]],
  ['two formats', [['ab '], ['cd ', { bold: true }], ['ef']]],
  ['right-to-left', [['אב גד ', rtl], ['הו', rtl]]],
  ['right-to-left with digits', [['אב123גד 45', rtl]]],
  ['mixed directions', [['ab אבג cd']]],
  ['Arabic joining', [['لام لا با', rtl]]],
  ['East Asian with automatic spacing', [['漢字ab字12字']], { autoSpaceDE: true, autoSpaceDN: true }],
  ['East Asian punctuation', [['字字、「字字」字。字字字']]],
  ['Thai', [['ภาษาไทย คำ']]],
  ['ligatures', [['office fi ffl']]],
  ['tabs', [['a\tbc\td']], { tabStops: [{ pos: 30, alignment: 'left', leader: 'none' }] }],
  ['fitted text', [['abcd', { fitTextVal: 960, fitTextId: 1 }], [' ef']]],
  ['soft hyphen and no-break space', [['co­op x yz']]],
  // One run holding both directions, in a right-to-left paragraph.
  ['digits inside one right-to-left run', [['\u05d0\u05d1123\u05d2\u05d3', rtl]], { bidi: true }],
  ['a Latin word inside one right-to-left run', [['\u05d0\u05d1 abc \u05d2\u05d3', rtl]], { bidi: true }],
  ['digits inside one unflagged run', [['\u05d0\u05d1123\u05d2\u05d3']], { bidi: true }],
  // Justification that elongates or distributes (its own alignment).
  ['Arabic under highKashida', [[kashidaText, rtl]], { bidi: true, alignment: 'highKashida' }],
  ['Arabic under lowKashida', [[kashidaText, rtl]], { bidi: true, alignment: 'lowKashida' }],
  ['Arabic under mediumKashida', [[kashidaText, rtl]], { bidi: true, alignment: 'mediumKashida' }],
  ['distributed text', [['ab cd ef gh ij kl']], { alignment: 'distribute' }],
  ['Thai distributed', [['\u0e20\u0e32\u0e29\u0e32\u0e44\u0e17\u0e22 \u0e04\u0e33 \u0e44\u0e17\u0e22']], { alignment: 'thaiDistribute' }],
];
const naturalScripts = new Set(['left-to-right', 'two formats', 'right-to-left', 'right-to-left with digits',
  'mixed directions', 'Arabic joining', 'East Asian punctuation', 'ligatures', 'digits inside one right-to-left run',
  'a Latin word inside one right-to-left run', 'digits inside one unflagged run', 'fitted text',
  'East Asian with automatic spacing', 'Arabic under highKashida', 'Arabic under lowKashida',
  'Arabic under mediumKashida', 'distributed text']);
const alignments = ['left', 'both'] as const;
const underlines: readonly object[] = [{}, { underline: true }, { underline: true, underlineStyle: 'dash' }];
const modes = [12, 14, 15] as const;
const grids: readonly object[] = [{}, { docGridType: 'lines', docGridLinePitch: 20 }];
const containers = ['body', 'body near the page end', 'cell', 'cell of a row split across pages', 'header', 'footnote'] as const;
const directions: readonly (string | null)[] = [null, 'tbRlV'];
const widths = [200, 100] as const;
const lineSpacings: readonly (object | null)[] = [null, { rule: 'auto', value: 1.5, explicit: true },
  { rule: 'exact', value: 14, explicit: true }, { rule: 'atLeast', value: 20, explicit: true }];
/** Room left under the filler of 'body near the page end': from less than
 * one line to two. */
const fillers = [100, 112, 122, 130] as const;

interface Case {
  readonly name: string;
  readonly specs: readonly RunSpec[];
  readonly options: object;
  readonly settings: object;
  readonly section: object;
  readonly container: typeof containers[number];
  readonly vertical: boolean;
  readonly filler: number;
  /** Horizontal body text with one cluster per character: page coordinates
   * are the paragraph's own. */
  readonly naturalSpacing: boolean;
}

/** A deterministic sample of the axes: every script meets every value of
 * every other axis at least once, without the full product. */
function cases(perScript: number): Case[] {
  const result: Case[] = [];
  let state = 12345;
  const next = (count: number): number => {
    state = (Math.imul(state, 1103515245) + 12345) & 0x7fffffff;
    return (state >>> 8) % count;
  };
  scripts.forEach(([scriptName, specs, scriptOptions], scriptIndex) => {
    for (let sample = 0; sample < perScript; sample += 1) {
      // The first samples walk each axis in turn; the rest are drawn.
      const pick = <T,>(values: readonly T[], axis: number): T =>
        values[sample < 6 ? (sample + scriptIndex * (axis + 1) + axis) % values.length : next(values.length)]!;
      const alignment = pick(alignments, 0);
      const underline = pick(underlines, 1);
      const mode = pick(modes, 2);
      const grid = pick(grids, 3);
      const useFELayout = pick([false, true], 4);
      const container = pick(containers, 5);
      const direction = pick(directions, 6);
      const width = pick(widths, 7);
      const lineSpacing = pick(lineSpacings, 8);
      const filler = pick(fillers, 9);
      result.push({
        name: [scriptName, alignment, JSON.stringify(underline), `mode ${mode}`, JSON.stringify(grid),
          useFELayout ? 'useFELayout' : 'no useFELayout', container, direction ?? 'horizontal', `page ${width}`,
          `spacing ${JSON.stringify(lineSpacing)}`, `filler ${filler}`].join(', '),
        specs: specs.map(([value, extra]) => [value, { ...extra, ...underline }] as const),
        options: { alignment, lineSpacing, ...scriptOptions },
        filler,
        settings: { compatibilityMode: mode, ...(useFELayout ? { useFELayout: true } : {}) },
        section: { ...grid, ...(direction ? { textDirection: direction, pageWidth: 200, pageHeight: width + 60 }
          : { pageWidth: width, pageHeight: 200 }) },
        container,
        vertical: direction !== null,
        naturalSpacing: direction === null && container.startsWith('body') && naturalScripts.has(scriptName),
      });
    }
  });
  return result;
}

/** Source boundaries of the subject text: every code unit that starts a
 * grapheme, and the end; long texts keep the ends and every second one. */
function boundaries(specs: readonly RunSpec[]): number[] {
  const value = specs.map(([part]) => part).join('');
  const all: number[] = [];
  for (let index = 0; index <= value.length; index += 1) {
    if (index < value.length && /\p{M}/u.test(value[index]!)) continue;
    all.push(index);
  }
  return all.length <= 12 ? all : all.filter((_, index) => index % 2 === 0 || index === all.length - 1);
}

/** The subject runs split at `offset`, with or without the anchor there. */
function subjectRuns(specs: readonly RunSpec[], offset: number, kind: AnchorKind | null): { runs: Run[]; anchorAt: number } {
  const runs: Run[] = [];
  let cursor = 0;
  let anchorAt = -1;
  const anchor = (): void => {
    anchorAt = runs.length;
    if (kind) runs.push(...anchorRuns(kind));
  };
  for (const [value, extra] of specs) {
    const local = offset - cursor;
    if (local === 0 && anchorAt < 0) anchor();
    if (local > 0 && local < value.length) {
      runs.push(textRun(value.slice(0, local), extra));
      anchor();
      // §17.3.2.14 groups source-adjacent runs only, so a float splits a
      // fitText region; the plain counterpart splits it with a new id.
      const fit = extra as { fitTextVal?: number; fitTextId?: number } | undefined;
      runs.push(textRun(value.slice(local), fit?.fitTextVal === undefined ? extra
        : { ...fit, fitTextId: (fit.fitTextId ?? 0) + 1000 }));
    } else {
      runs.push(textRun(value, extra));
    }
    cursor += value.length;
  }
  if (anchorAt < 0) anchor();
  return { runs, anchorAt };
}

function model(testCase: Case, subject: readonly Run[]): { model: DocxDocumentModel; prefixRuns: number } {
  const after = paragraph([textRun('zz')]);
  let body: BodyElement[];
  let header: BodyElement[] | null = null;
  let footnotes: unknown[] = [];
  let prefixRuns = 0;
  const subjectParagraph = paragraph(subject, testCase.options);
  switch (testCase.container) {
    case 'body': body = [subjectParagraph, after]; break;
    case 'body near the page end':
      body = [paragraph([textRun('x')], { lineSpacing: { rule: 'exact', value: testCase.filler, explicit: true } }), subjectParagraph, after];
      break;
    case 'cell': body = [table([subjectParagraph, paragraph([textRun('y')])]), after]; break;
    case 'cell of a row split across pages':
      // The row is taller than the page: its lines are placed page by page,
      // and the subject stands on either side of the split.
      body = [table([
        ...Array.from({ length: Math.round(testCase.filler / 12) }, () => paragraph([textRun('y')])),
        subjectParagraph, paragraph([textRun('y')]), paragraph([textRun('y')]),
      ]), after];
      break;
    case 'header': body = [after]; header = [subjectParagraph]; break;
    case 'footnote': {
      prefixRuns = 1;
      body = [paragraph([textRun('B'), textRun('1', { noteRef: { kind: 'footnote', id: '1' } })]), after];
      footnotes = [{ id: '1', content: [paragraph([textRun('', { noteRef: { kind: 'footnote', id: '' } }), ...subject],
        testCase.options)] }];
      break;
    }
  }
  return {
    prefixRuns,
    model: {
      section: { pageWidth: 200, pageHeight: 200, marginTop: 40, marginRight: 20, marginBottom: 20, marginLeft: 20,
        headerDistance: 5, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
        sectionStart: 'nextPage', columns: null, ...testCase.section },
      body, settings: testCase.settings,
      headers: { default: header ? { body: header } : null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes, endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel,
  };
}

// ── What is compared ────────────────────────────────────────────────────────

const round = (value: number): number => Math.round(value * 1e6) / 1e6;
/** Keys that name the drawing and its bookkeeping, or source positions the
 * anchoring character shifts by one unit. */
const dropped = new Set(['paintOrder', 'capabilities', 'segIndex', 'drawings', 'textBoxes', 'anchorCollisions', 'anchorFrames',
  'cellContainmentBounds', 'resources', 'events', 'id', 'source', 'sourceRunIndex', 'paragraphId']);

/** Everything a page lays out and paints, without the floating drawing, its
 * host and source addresses. Ranges keep their lengths. */
function comparable(value: unknown): unknown {
  if (typeof value === 'number') return round(value);
  if (value === null || typeof value !== 'object') return value;
  if (Array.isArray(value)) {
    return value
      .filter((entry) => !(entry && typeof entry === 'object'
        && ((entry as { kind?: string }).kind === 'anchor-host'
          || ((entry as { kind?: string }).kind === 'drawing' && 'drawingId' in entry))))
      .map(comparable);
  }
  const result: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(value)) {
    if (dropped.has(key) || entry === undefined) continue;
    if (key === 'range' && entry && typeof entry === 'object' && 'start' in entry && 'end' in entry) {
      // A line's own range counts the anchoring character; its content's do not.
      if ('placements' in value) continue;
      result.rangeLength = (entry as { end: number }).end - (entry as { start: number }).start;
      continue;
    }
    result[key] = comparable(entry);
  }
  return result;
}

const fakeElement = () => ({ style: { cssText: '' }, children: [] as { style: { cssText: string } }[], innerHTML: '',
  appendChild(child: { style: { cssText: string } }) { this.children.push(child); } });
/** UTF-16 units handed to the find overlay's measurer. */
let measuredUnits = 0;
const measureForFont = (font: string) => (value: string): number => {
  measuredUnits += value.length;
  return [...value].length * Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10') * 0.5;
};

/** The highlight of every single character of every run. */
function findHighlights(runs: DocxTextRunInfo[]): string[] {
  const scope = globalThis as { document?: unknown };
  const previous = scope.document;
  scope.document = { createElement: fakeElement };
  try {
    const layer = fakeElement();
    const matches = runs.flatMap((run, runIndex) => Array.from({ length: run.text.length }, (_, start) => ({
      active: false, slices: [{ runIndex, start, end: start + 1 }],
    })));
    buildDocxHighlightLayer(layer as never, runs, matches, 1000, 1000, measureForFont);
    return layer.children.map((child) => child.style.cssText);
  } finally {
    scope.document = previous;
  }
}

/** Each character's highlight that is not the glyph box of a cluster showing
 * that character: where paint draws the glyph, from its pen (the left edge,
 * or the leading right edge of a right-to-left cluster), at its own drawn
 * advance (the retained glyph advance, or a floored glyph's drawn one). */
function misplacedHighlights(runs: DocxTextRunInfo[], laidOut: readonly ParagraphLayout[]): string[] {
  const painted = new Map<string, Array<readonly [number, number]>>();
  for (const subject of laidOut) {
    for (const line of subject.lines) {
      for (const placement of line.placements) {
        if (placement.kind !== 'text') continue;
        placement.clusters.forEach((cluster, index) => {
          const shown = placement.text.slice(cluster.range.start - placement.range.start, cluster.range.end - placement.range.start);
          const floored = placement.paintOps.find((operation) => operation.range.start === cluster.range.start
            && operation.naturalAdvancePt !== undefined)?.naturalAdvancePt;
          const glyph = cluster.glyphAdvancePt ?? floored ?? cluster.advancePt;
          painted.set(shown, [...(painted.get(shown) ?? []),
            [visualClusterStartXPt(placement, index), glyph], [visualBoundaryXPt(placement, index) - glyph, glyph]]);
        });
      }
    }
  }
  const scope = globalThis as { document?: unknown };
  const previous = scope.document;
  scope.document = { createElement: fakeElement };
  const misplaced: string[] = [];
  try {
    runs.forEach((run, runIndex) => {
      for (let start = 0; start < run.text.length; start += 1) {
        const layer = fakeElement();
        buildDocxHighlightLayer(layer as never, runs, [{ active: false, slices: [{ runIndex, start, end: start + 1 }] }],
          100, 100, measureForFont);
        const css = layer.children[0]?.style.cssText ?? '';
        const left = /left:([-\d.e]+)%/.exec(css)?.[1];
        const width = /width:([-\d.e]+)%/.exec(css)?.[1];
        if (left === undefined || width === undefined) continue;
        if (!(painted.get(run.text[start]!) ?? []).some(([xPt, widthPt]) =>
          Math.abs(xPt - Number(left)) < 1e-6 && Math.abs(widthPt - Number(width)) < 1e-6)) {
          misplaced.push(`${JSON.stringify(run.text)}[${start}] at ${left} width ${width}`);
        }
      }
    });
  } finally {
    scope.document = previous;
  }
  return misplaced;
}

function paragraphs(node: unknown, found: ParagraphLayout[] = []): ParagraphLayout[] {
  if (!node || typeof node !== 'object') return found;
  if (Array.isArray(node)) { node.forEach((entry) => paragraphs(entry, found)); return found; }
  const record = node as Record<string, unknown>;
  if (record.kind === 'paragraph' && Array.isArray(record.lines)) found.push(record as unknown as ParagraphLayout);
  for (const [key, entry] of Object.entries(record)) if (key !== 'paintOrder') paragraphs(entry, found);
  return found;
}

function observe(testCase: Case, subject: readonly Run[], anchorAt: number, anchored: boolean) {
  const measuredBefore = measuredUnits;
  const built = model(testCase, subject);
  const layout = layoutDocument(built.model, createLayoutServices(built.model, { measureContext: measureContext() }),
    { currentDateMs: 0 });
  void anchorAt;
  void anchored;
  const selection: unknown[] = [];
  const finds: string[][] = [];
  const comments: unknown[] = [];
  layout.pages.forEach((_, pageIndex) => {
    const runs = textRunsForPage(layout, pageIndex, { scale: 1 });
    // Source runs by their order in their paragraph: the anchoring runs
    // renumber the runs after them and own no text.
    const story = (run: Readonly<DocxTextRunInfo>): string => JSON.stringify(run.source);
    const indices = new Map<string, number[]>();
    for (const run of runs) {
      if (run.sourceRunIndex === undefined) continue;
      const list = indices.get(story(run)) ?? [];
      if (!list.includes(run.sourceRunIndex)) list.push(run.sourceRunIndex);
      indices.set(story(run), list.sort((a, b) => a - b));
    }
    const geometry = (run: Readonly<DocxTextRunInfo>) => {
      const { source: _source, sourceRunIndex, paragraphId: _paragraphId, ...rest } = run as DocxTextRunInfo & { paragraphId?: unknown };
      return { ...comparable(rest) as object, story: run.source?.story,
        sourceRun: sourceRunIndex === undefined ? null : indices.get(story(run))!.indexOf(sourceRunIndex) };
    };
    selection.push(runs.map(geometry));
    finds.push(findHighlights(runs));
    // A comment over each source run, and a point comment at every run
    // position of the paragraph (the anchoring runs' own included: a point
    // there resolves as at the text that follows).
    for (const [key, list] of indices) {
      const sample = runs.find((run) => story(run) === key)!;
      const resolve = (start: number, end: number) => JSON.stringify(resolveCommentAnchorRuns({
        commentId: 'c', source: sample.source, startRunIndex: start, endRunIndex: end,
        reference: { source: sample.source, runIndex: start, affinity: 'following' },
      } as never, runs).map(geometry));
      const points: string[] = [];
      for (let index = list[0]!; index <= list.at(-1)! + 1; index += 1) {
        const resolved = resolve(index, index);
        if (points.at(-1) !== resolved) points.push(resolved);
      }
      comments.push({ ranges: list.map((index) => resolve(index, index + 1)), points });
    }
  });
  const laidOut = paragraphs(layout.pages.map((page) => page.layers));
  const misplaced = testCase.naturalSpacing && layout.pages.length === 1
    ? misplacedHighlights(textRunsForPage(layout, 0, { scale: 1 }), laidOut) : [];
  // Selection rectangles of one line cover disjoint intervals.
  const overlapping: string[] = [];
  if (testCase.naturalSpacing) {
    layout.pages.forEach((_, pageIndex) => {
      const runs = textRunsForPage(layout, pageIndex, { scale: 1 }).filter((run) => run.w > 0);
      runs.forEach((left, index) => runs.slice(index + 1).forEach((right) => {
        if (Math.abs(left.y - right.y) < 1e-6 && Math.min(left.x + left.w, right.x + right.w) - Math.max(left.x, right.x) > 1e-6) {
          overlapping.push(`${JSON.stringify(left.text)} [${left.x}, ${left.x + left.w}) and ${JSON.stringify(right.text)} [${right.x}, ${right.x + right.w})`);
        }
      }));
    });
  }
  const drawings = layout.pages.flatMap((page) => page.layers.paintOrder.flatMap((entry) => entry.kind === 'drawing'
    ? [[round(entry.node.inkBounds.xPt), round(entry.node.inkBounds.yPt)]] : []));
  return {
    drawings, misplaced, overlapping,
    // Find reads the runs' retained cluster boxes: it measures no text.
    measured: measuredUnits - measuredBefore,
    pages: layout.pages.map((page) => comparable(page.layers)),
    selection, finds, comments,
    paragraphs: laidOut,
  };
}

/** Where the source boundary `offset` (in the paragraph's text) is painted
 * in an anchor-free paragraph: the leading edge of the content that starts
 * there, or the trailing edge of the text that ends there. `precedingPt` is
 * the visual start of the character before the boundary on that line. */
function boundaryPosition(subject: ParagraphLayout, offset: number): { xPt: number; line: number; precedingPt: number | null } | null {
  let found: { xPt: number; line: number } | null = null;
  let ending: { xPt: number; line: number } | null = null;
  for (const [lineIndex, line] of subject.lines.entries()) {
    for (const placement of line.placements) {
      if (placement.kind === 'text') {
        const cluster = placement.clusters.findIndex((entry) => entry.range.start === offset);
        if (cluster >= 0) found ??= { xPt: visualBoundaryXPt(placement, cluster), line: lineIndex };
        else if (placement.range.start < offset && offset < placement.range.end) return null;
        else if (placement.range.end === offset && placement.clusters.length > 0) {
          ending = { xPt: visualBoundaryXPt(placement, placement.clusters.length), line: lineIndex };
        }
      } else if (placement.kind === 'tab' && placement.range.start === offset) {
        found ??= { xPt: placement.bounds!.xPt, line: lineIndex };
      }
    }
  }
  // Only the end of the paragraph's content is decided by what precedes it.
  const last = subject.lines.at(-1)?.placements.filter((placement) => placement.kind === 'text' || placement.kind === 'tab')
    .reduce((end, placement) => Math.max(end, placement.range.end), 0) ?? 0;
  const position = found ?? (offset >= last ? ending : null);
  if (!position) return null;
  let preceding: { start: number; xPt: number } | null = null;
  for (const placement of subject.lines[position.line]!.placements) {
    if (placement.kind !== 'text') continue;
    placement.clusters.forEach((cluster, index) => {
      if (cluster.range.start < offset && (!preceding || cluster.range.start > preceding.start)) {
        preceding = { start: cluster.range.start, xPt: visualClusterStartXPt(placement, index) };
      }
    });
  }
  return { ...position, precedingPt: (preceding as { xPt: number } | null)?.xPt ?? null };
}

/** The first place two compared values differ, for a readable failure. */
function firstDifference(actual: unknown, expected: unknown, path = ''): string | null {
  if (Object.is(actual, expected)) return null;
  if (actual && expected && typeof actual === 'object' && typeof expected === 'object') {
    const keys = new Set([...Object.keys(actual), ...Object.keys(expected)]);
    for (const key of keys) {
      const found = firstDifference((actual as Record<string, unknown>)[key], (expected as Record<string, unknown>)[key],
        `${path}/${key}`);
      if (found) return found;
    }
    return null;
  }
  return `${path}: with the anchor ${JSON.stringify(actual)?.slice(0, 120)}, without ${JSON.stringify(expected)?.slice(0, 120)}`;
}

const sampled = cases(Number(process.env.ANCHOR_DIFFERENTIAL_SAMPLES ?? 6));

describe('a floating anchor at any source boundary changes nothing but the drawing', () => {
  for (const testCase of sampled) {
    it(testCase.name, () => {
      const offsets = boundaries(testCase.specs);
      offsets.forEach((offset, order) => {
        const kind: AnchorKind = order % 2 === 0 ? 'character' : 'page';
        const plainRuns = subjectRuns(testCase.specs, offset, null);
        const anchoredRuns = subjectRuns(testCase.specs, offset, kind);
        const plain = observe(testCase, plainRuns.runs, plainRuns.anchorAt, false);
        const anchored = observe(testCase, anchoredRuns.runs, anchoredRuns.anchorAt, true);
        const label = `${kind}-relative anchor at offset ${offset}`;
        for (const facet of ['pages', 'selection', 'finds', 'comments'] as const) {
          const difference = firstDifference(anchored[facet], plain[facet]);
          if (difference) throw new Error(`${label}: ${facet}${difference}`);
        }

        // A highlight lies over the character it marks, with or without the anchor.
        expect(plain.misplaced, `${label}: find highlights without the anchor`).toEqual([]);
        expect(anchored.misplaced, `${label}: find highlights`).toEqual([]);
        expect([plain.measured, anchored.measured], `${label}: text measured by find`).toEqual([0, 0]);
        expect(plain.overlapping, `${label}: selection rectangles without the anchor`).toEqual([]);
        expect(anchored.overlapping, `${label}: selection rectangles`).toEqual([]);

        // The host: on the line and at the visual position of its boundary.
        const hostParagraph = anchored.paragraphs.findIndex((candidate) =>
          candidate.lines.some((line) => line.placements.some((placement) => placement.kind === 'anchor-host')));
        expect(hostParagraph, label).toBeGreaterThanOrEqual(0);
        const lines = anchored.paragraphs[hostParagraph]!.lines;
        const hostLine = lines.findIndex((line) => line.placements.some((placement) => placement.kind === 'anchor-host'));
        const host = lines[hostLine]!.placements.find((placement) => placement.kind === 'anchor-host')!;
        const prefix = testCase.container === 'footnote' ? 1 : 0;
        const expected = boundaryPosition(plain.paragraphs[hostParagraph]!, prefix + offset);
        if (expected) {
          expect([hostLine, round(host.bounds!.xPt)], `${label}: host position`).toEqual([expected.line, round(expected.xPt)]);
          // The host carries where its source-preceding character starts,
          // relative to itself: still true of the finished page, whatever
          // repositioned or copied the line's placements after planning.
          const recorded = (host as { precedingCharacterOffsetPt?: number | null }).precedingCharacterOffsetPt;
          expect(recorded, `${label}: preceding character record`).not.toBeUndefined();
          expect(recorded == null ? null : round(host.bounds!.xPt + recorded), `${label}: preceding character`)
            .toBe(expected.precedingPt === null ? null : round(expected.precedingPt));
          // The drawing itself, positioned from its character and line: at
          // the boundary on a horizontal line; on a line of a vertical
          // section's body, at the start of the character before it (the
          // line start when none). A cell lays its lines out in its own
          // horizontal frame, in a row split across pages as well, where the
          // host placements are repositioned after planning.
          if (kind === 'character' && (testCase.container.startsWith('body') || testCase.container.startsWith('cell'))) {
            const line = lines[hostLine]!.bounds;
            expect(anchored.drawings, `${label}: drawing position`).toEqual([[
              round(testCase.vertical && testCase.container.startsWith('body') ? expected.precedingPt ?? line.xPt : expected.xPt),
              round(line.yPt)]]);
          }
        }
      });
    });
  }
});

describe('a character-relative anchor in kashida-justified text', () => {
  it('stands where the elongated text paints its boundary', () => {
    // An oracle independent of the cluster table: the painted text itself.
    // Every painted character, inserted tatweels included, is 6pt wide, and
    // a right-to-left piece paints its text from its right edge.
    // The second text holds an authored tatweel: painted, it cannot be told
    // from an inserted one, so a boundary beside it may stand on either side
    // of the tatweels inserted there (the plan decides; see the direct test
    // of the cluster table).
    const base = cases(1).find((candidate) => candidate.name.startsWith('Arabic under highKashida'))!;
    const tatweel = '\u0640';
    for (const value of [kashidaText, kashidaText.replace('\u062a\u062a', `\u062a${tatweel}\u062a`)]) {
      const narrow = { ...base, specs: [[value, rtl]] as const, container: 'body' as const, vertical: false,
        section: { pageWidth: 130, pageHeight: 200 } };
      let elongated = 0;
      for (const offset of boundaries(narrow.specs)) {
        const { runs } = subjectRuns(narrow.specs, offset, 'character');
        const built = model(narrow, runs);
        const layout = layoutDocument(built.model, createLayoutServices(built.model, { measureContext: measureContext() }),
          { currentDateMs: 0 });
        const subject = paragraphs(layout.pages.map((page) => page.layers))[0]!;
        for (const line of subject.lines) {
          const host = line.placements.find((placement) => placement.kind === 'anchor-host');
          if (!host) continue;
          const piece = line.placements.find((placement) => placement.kind === 'text'
            && placement.range.start <= offset && offset < placement.range.end);
          if (piece?.kind !== 'text' || piece.paintOps[0]?.sourceMapping !== 'kashida') continue;
          elongated += 1;
          const painted = [...piece.paintOps[0].text];
          const source = [...piece.text];
          const target = offset - piece.range.start;
          // Painted characters before the boundary: the source characters
          // before it, then the tatweels inserted before its character.
          let consumed = 0;
          let early = 0;
          while (consumed < target) {
            if (painted[early] === source[consumed]) consumed += 1;
            early += 1;
          }
          let late = early;
          while (late < painted.length && painted[late] !== source[target]) late += 1;
          const run = (text: readonly string[], from: number): number => {
            let count = 0;
            while (text[from + count] === tatweel) count += 1;
            return count;
          };
          // Beside an authored tatweel the inserted ones are not identified.
          if (source[target] === tatweel) late = early + run(painted, early) - run(source, target);
          const rightPt = piece.bounds.xPt + piece.bounds.widthPt;
          const besideAuthored = source[target] === tatweel || source[target - 1] === tatweel;
          const xPt = round(host.bounds!.xPt);
          if (besideAuthored) {
            expect(xPt, `anchor at offset ${offset}`).toBeGreaterThanOrEqual(round(rightPt - late * 6));
            expect(xPt, `anchor at offset ${offset}`).toBeLessThanOrEqual(round(rightPt - early * 6));
            expect((rightPt - xPt) / 6 % 1, `anchor at offset ${offset}`).toBeCloseTo(0, 6);
          } else {
            expect(xPt, `anchor at offset ${offset}`).toBe(round(rightPt - late * 6));
          }
        }
      }
      expect(elongated).toBeGreaterThan(0);
    }
  });
});

describe('find highlights and hosts in pieces whose paint differs from their source order', () => {
  const lay = (runs: readonly Run[], options: object, pageWidth = 200, ems: Readonly<Record<string, number>> = {}) => {
    const built = model({ ...cases(1)[0]!, container: 'body', vertical: false, options,
      settings: { compatibilityMode: 15 }, section: { pageWidth, pageHeight: 200 } }, runs);
    const layout = layoutDocument(built.model, createLayoutServices(built.model, { measureContext: measureContext(ems) }),
      { currentDateMs: 0 });
    const subject = paragraphs(layout.pages.map((page) => page.layers))[0]!;
    return { layout, subject, placements: subject.lines.flatMap((line) => line.placements) };
  };
  /** Left edge and width of the highlight of each match of `needle`, by run text. */
  const highlights = (layout: Parameters<typeof textRunsForPage>[0], needle: string): [string, number, number][] => {
    const runs = textRunsForPage(layout, 0, { scale: 1 });
    const scope = globalThis as { document?: unknown };
    const previous = scope.document;
    scope.document = { createElement: fakeElement };
    try {
      return runs.flatMap((run, runIndex) => [...run.text].flatMap((character, start) => {
        if (character !== needle) return [];
        const layer = fakeElement();
        buildDocxHighlightLayer(layer as never, runs, [{ active: false, slices: [{ runIndex, start, end: start + 1 }] }],
          100, 100, measureForFont);
        return layer.children.map((child) =>
          [run.text, round(Number(/left:([-\d.e]+)%/.exec(child.style.cssText)![1])),
            round(Number(/width:([-\d.e]+)%/.exec(child.style.cssText)![1]))] as [string, number, number]);
      }));
    } finally {
      scope.document = previous;
    }
  };

  it('keeps a host before a fitted right-to-left run at the fitted edge of the run', () => {
    // Four 6pt letters fitted to 48pt: the run starts at its right edge,
    // and the tracking after its last letter is outside the fitted width.
    const fitted = textRun('\u05d0\u05d1\u05d2\u05d3', { rtl: true, fitTextVal: 960, fitTextId: 1 });
    const { placements } = lay([...anchorRuns('character'), fitted], { bidi: true });
    const piece = placements.find((placement) => placement.kind === 'text')!;
    const host = placements.find((placement) => placement.kind === 'anchor-host')!;
    expect(piece.bounds!.widthPt).toBe(48);
    expect(round(host.bounds!.xPt)).toBe(round(piece.bounds!.xPt + 48));
  });

  it('keeps a host before a compressed fitted right-to-left run at the run\'s right edge', () => {
    // Four 6pt letters fitted to 12pt: each advances 2pt, the final glyph
    // paints past its pitch to the fitted edge, and the run starts at its
    // right edge, where the highlight of its first letter ends as well.
    const fitted = textRun('\u05d0\u05d1\u05d2\u05d3', { rtl: true, fitTextVal: 240, fitTextId: 1 });
    const { layout, placements } = lay([...anchorRuns('character'), fitted], { bidi: true });
    const piece = placements.find((placement) => placement.kind === 'text')!;
    const host = placements.find((placement) => placement.kind === 'anchor-host')!;
    expect(piece.bounds!.widthPt).toBe(12);
    expect(round(host.bounds!.xPt)).toBe(round(piece.bounds!.xPt + 12));
    const [[, x, width]] = highlights(layout, '\u05d0') as [[string, number, number]];
    expect(round(x + width)).toBe(round(piece.bounds!.xPt + 12));
  });

  it('highlights a condensed letter whose advance floors to nothing where it is painted', () => {
    // 'abc' at 12pt condensed by 8pt: b's advance floors to zero while its
    // own operation still draws it at its 6pt natural advance.
    const { layout, placements } = lay([textRun('abc', { charSpacing: -8 })], {});
    const piece = placements.find((placement) => placement.kind === 'text')!;
    expect(piece.kind === 'text' && piece.paintOps.some((operation) => operation.naturalAdvancePt !== undefined)).toBe(true);
    const found = highlights(layout, 'b');
    expect(found).toHaveLength(1);
    expect(found[0]![2]).toBeGreaterThan(0);
  });

  it('highlights a floored condensed letter over its painted glyph only', () => {
    // 'mim' at 12pt (6/3/6pt natural) condensed by 4pt: i floors to no
    // advance and is drawn at its 3pt natural advance; its box is final.
    const { layout } = lay([textRun('mim', { charSpacing: -4 })], {}, 200, { i: 0.25 });
    const found = highlights(layout, 'i');
    expect(found).toHaveLength(1);
    expect(found[0]![2]).toBe(3);
  });

  it('keeps fitText regions on either side of a float independent', () => {
    // ab and cd share fitText id 1 and 24pt, with a float in its own run
    // between them: they are not source-adjacent, so each fills 24pt.
    const fit = { fitTextVal: 480, fitTextId: 1 };
    const { placements } = lay([textRun('ab', fit), ...anchorRuns('page'), textRun('cd', fit)], {});
    const texts = placements.filter((placement) => placement.kind === 'text');
    expect(texts.map((placement) => [placement.kind === 'text' ? placement.text : '', placement.bounds!.widthPt]))
      .toEqual([['ab', 24], ['cd', 24]]);
  });

  it('places a character-relative float after a compressed fitted run at the run\'s end', () => {
    // abcd (6pt letters) fitted to 12pt: each advances 3pt and the final glyph
    // paints to the run's end, where the float's anchoring character stands.
    const fitted = textRun('abcd', { fitTextVal: 240, fitTextId: 1 });
    const { placements } = lay([fitted, ...anchorRuns('character')], {});
    const piece = placements.find((placement) => placement.kind === 'text')!;
    const host = placements.find((placement) => placement.kind === 'anchor-host')!;
    expect(piece.bounds!.widthPt).toBe(12);
    expect(round(host.bounds!.xPt)).toBe(round(piece.bounds!.xPt + 12));
  });

  it('highlights a Latin letter inside one mixed-direction piece where it is painted', () => {
    // alef a - b bet, two same-format runs around a float. Painted left to
    // right: bet, a, hyphen, b, alef (6pt each); "b" is the fourth.
    const rtlRun = (value: string) => textRun(value, { rtl: true });
    const { layout, placements } = lay([rtlRun('\u05d0a'), ...anchorRuns('page'), rtlRun('-b\u05d1')], { bidi: true });
    const texts = placements.filter((placement) => placement.kind === 'text');
    expect(texts.map((placement) => placement.kind === 'text' ? placement.text : '')).toEqual(['\u05d0a-b\u05d1']);
    const left = texts[0]!.bounds!.xPt;
    expect(highlights(layout, 'b').map(([, x]) => x)).toEqual([round(left + 18)]);
    expect(highlights(layout, '-').map(([, x]) => x)).toEqual([round(left + 12)]);
    expect(highlights(layout, 'a').map(([, x]) => x)).toEqual([round(left + 6)]);
    expect(highlights(layout, '\u05d0').map(([, x]) => x)).toEqual([round(left + 24)]);
  });

  it('highlights letters of kashida-justified text where the elongated text paints them', () => {
    // A 90pt band, a float after the first two letters. Each highlight of
    // beh is checked against the painted text (6pt per painted character,
    // inserted tatweels included, from the piece's right edge): it ends at
    // the letter's right edge, its pen, and covers the letter's glyph.
    const rtlRun = (value: string) => textRun(value, { rtl: true });
    const { layout, subject } = lay([rtlRun(kashidaText.slice(0, 2)), ...anchorRuns('page'), rtlRun(kashidaText.slice(2))],
      { bidi: true, alignment: 'highKashida' }, 130);
    const expected: number[] = [];
    let elongated = 0;
    for (const line of subject.lines) {
      for (const piece of line.placements) {
        if (piece.kind !== 'text') continue;
        const painted = [...(piece.paintOps[0]?.text ?? piece.text)];
        if (piece.paintOps[0]?.sourceMapping === 'kashida') elongated += 1;
        const source = [...piece.text];
        let prefix = 0;
        source.forEach((character, index) => {
          // Tatweels inserted before a character precede it.
          while (prefix < painted.length && painted[prefix] !== character) prefix += 1;
          if (character === '\u0628') expected.push(round(piece.bounds.xPt + piece.bounds.widthPt - prefix * 6));
          prefix += 1;
          void index;
        });
      }
    }
    expect(elongated).toBeGreaterThan(0);
    const found = highlights(layout, '\u0628');
    expect(found.map(([, x, width]) => round(x + width)).sort((a, b) => a - b)).toEqual(expected.sort((a, b) => a - b));
    // The highlight is the letter's glyph only: the elongation added after
    // it belongs to its advance, not its glyph.
    expect(found.every(([, , width]) => width === 6)).toBe(true);
  });

  it('measures no text to highlight the last letter of a long right-to-left run', () => {
    const measured = (length: number): number => {
      const { layout } = lay([textRun('\u05d0'.repeat(length - 1) + '\u05d1', { rtl: true })], { bidi: true }, 4000);
      const before = measuredUnits;
      expect(highlights(layout, '\u05d1')).toHaveLength(1);
      return measuredUnits - before;
    };
    expect([measured(50), measured(200)]).toEqual([0, 0]);
  });
});

describe('a float-only mark that ends a section before a continuous break', () => {
  /** The top of the next section's text and the drawings painted, for a
   * section-ending mark holding `mark` with `spaceBeforePt` before it. */
  const observe = (mark: readonly Run[], spaceBeforePt: number, markOptions: object = {}) => {
    const section = { pageWidth: 200, pageHeight: 400, marginTop: 40, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 5, footerDistance: 0, titlePage: false, evenAndOddHeaders: false, columns: null };
    const built = {
      section: { ...section, sectionStart: 'continuous' },
      body: [
        paragraph([textRun('A')]),
        paragraph(mark, { spaceBefore: spaceBeforePt, ...markOptions }),
        { type: 'sectionBreak', kind: 'continuous', geom: { ...section, sectionStart: 'nextPage' } },
        paragraph([textRun('B')], { spaceBefore: 10 }),
      ],
      settings: { compatibilityMode: 15 },
      headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(built, createLayoutServices(built, { measureContext: measureContext() }), { currentDateMs: 0 });
    const runs = textRunsForPage(layout, 0, { scale: 1 });
    return {
      top: round(runs.find((run) => run.text === 'B')!.y),
      drawings: layout.pages.flatMap((page) => page.layers.paintOrder.filter((entry) => entry.kind === 'drawing')).length,
      borders: paragraphs(layout.pages.map((page) => page.layers)).flatMap((laidOut) => laidOut.borders).length,
      shaded: paragraphs(layout.pages.map((page) => page.layers)).filter((laidOut) => laidOut.shading !== undefined).length,
    };
  };

  it('keeps its space before suppressed, as its empty counterpart does', () => {
    const empty = observe([], 20);
    expect(observe(anchorRuns('page'), 20)).toEqual({ ...empty, drawings: 1 });
  });

  it('collapses, as its empty counterpart does, and still paints its drawing', () => {
    const empty = observe([], 0);
    expect(observe(anchorRuns('page'), 0)).toEqual({ ...empty, drawings: 1 });
  });

  it('paints no border or shading of its own when it collapses, only its drawing', () => {
    // A 1pt bottom border and a fill on the collapsed mark.
    const decorated = { borders: { bottom: { style: 'single', color: '000000', width: 1, space: 1 } },
      shading: 'FFFF00' };
    const empty = observe([], 0, decorated);
    expect(empty.borders).toBe(0);
    expect(observe(anchorRuns('page'), 0, decorated)).toEqual({ ...empty, drawings: 1 });
  });
});

describe('a hidden mark that anchors only floating drawings', () => {
  /** The top of the following text and the drawings painted. */
  const observe = (mark: readonly Run[], pageBreakBefore = false) => {
    const built = model({ ...cases(1)[0]!, container: 'body', vertical: false, options: {},
      settings: { compatibilityMode: 15 }, section: { pageWidth: 200, pageHeight: 400 } }, []).model;
    const body = [
      paragraph([textRun('A')]),
      paragraph(mark, { markVanish: true, ...(pageBreakBefore ? { pageBreakBefore: true } : {}) }),
      paragraph([textRun('B')]),
    ];
    const document = { ...built, body } as unknown as DocxDocumentModel;
    const layout = layoutDocument(document, createLayoutServices(document, { measureContext: measureContext() }), { currentDateMs: 0 });
    const runs = layout.pages.flatMap((_, index) => textRunsForPage(layout, index, { scale: 1 }).map((run) => ({ run, index })));
    const b = runs.find(({ run }) => run.text === 'B')!;
    return {
      page: b.index,
      top: round(b.run.y),
      drawings: layout.pages.flatMap((page) => page.layers.paintOrder.filter((entry) => entry.kind === 'drawing')).length,
    };
  };

  it('joins the next paragraph as its empty counterpart does and still paints its drawing', () => {
    const empty = observe([]);
    expect(observe(anchorRuns('page'))).toEqual({ ...empty, drawings: 1 });
  });

  it('keeps its drawing when its page break applies', () => {
    const empty = observe([], true);
    expect(observe(anchorRuns('page'), true)).toEqual({ ...empty, drawings: 1 });
  });
});

describe('floating anchors on many lines of one paragraph', () => {
  it('are paginated in work proportional to their number', () => {
    // One float per line and no text, lines ended by manual breaks: a scan
    // of the paragraph's drawings from every line reads n * n entries.
    const visited = (count: number): number => {
      const runs: Run[] = [];
      for (let index = 0; index < count; index += 1) {
        runs.push({ type: 'anchorHost', fontSize: 2, __anchorOccurrenceId: `p${index}` } as unknown as Run);
        runs.push({ type: 'image', imagePath: 'word/media/p.png', mimeType: 'image/png', widthPt: 3, heightPt: 2,
          anchor: true, anchorXPt: 0, anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: false,
          __anchorAcquisition: { ...acquisition('page'), occurrenceId: `p${index}` } } as unknown as Run);
        runs.push({ type: 'break', breakType: 'line' } as unknown as Run);
      }
      const built = model({ ...cases(1)[0]!, container: 'body', options: {},
        settings: { compatibilityMode: 14, useFELayout: true },
        section: { pageWidth: 200, pageHeight: 20000, docGridType: 'lines', docGridLinePitch: 20 } }, runs);
      let entries = 0;
      const names = ['filter', 'find', 'findIndex', 'slice', 'flatMap', 'some', 'every', 'map', Symbol.iterator] as const;
      const prototype = Array.prototype as unknown as Record<string | symbol, (...args: unknown[]) => unknown>;
      const originals = names.map((name) => prototype[name]!);
      names.forEach((name, index) => {
        prototype[name] = function (this: unknown[], ...args: unknown[]) {
          entries += this.length;
          return originals[index]!.apply(this, args);
        };
      });
      try {
        const layout = layoutDocument(built.model, createLayoutServices(built.model, { measureContext: measureContext() }),
          { currentDateMs: 0 });
        expect(layout.pages.flatMap((page) => page.layers.paintOrder.filter((entry) => entry.kind === 'drawing')))
          .toHaveLength(count);
      } finally {
        names.forEach((name, index) => { prototype[name] = originals[index]!; });
      }
      return entries;
    };
    expect(visited(800)).toBeLessThan(visited(200) * 5);
  });
});
