import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import type { DocxCompatSettingsInput } from './layout/compat-profile.js';
import { withCompat } from './test-support/compat-environment.test-support.js';
import {
  measureParagraph,
  type MeasuredParagraph,
  type ParagraphMeasurementEnvironment,
  type TextMeasurer,
} from './paragraph-measure.js';
import type { ParagraphLayoutContext } from './layout-context.js';
import type { DocParagraph } from './types.js';

type Run = DocParagraph['runs'][number];

function makeContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  const size = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => ({
      width: [...text].length * size() * 0.5,
      fontBoundingBoxAscent: size() * 0.8, fontBoundingBoxDescent: size() * 0.2,
      actualBoundingBoxAscent: size() * 0.8, actualBoundingBoxDescent: size() * 0.2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}
const measurer: TextMeasurer = { context: makeContext(), fontFamilyClasses: {} };

const text = (value: string, fontSize = 10, extra: object = {}): Run => ({
  type: 'text', text: value, bold: false, italic: false, underline: false, strikethrough: false,
  fontSize, color: null, fontFamily: null, isLink: false, background: null, vertAlign: null,
  hyperlink: null, ...extra,
}) as unknown as Run;
const deletion = { revision: { kind: 'deletion', author: 'A' } };
const picture = (extra: object = {}): Run => ({
  type: 'image', imagePath: 'p.png', mimeType: 'image/png', widthPt: 16, heightPt: 16, anchor: false,
  anchorXPt: 0, anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: false, ...extra,
}) as unknown as Run;
const inlineObject = {
  inlineObjectOrigin: 'drawingml',
  inlineObjectShapeInput: { fonts: { ascii: 'Arial', highAnsi: 'Arial', eastAsia: 'Arial' }, fontSizePt: 10,
    weight: 400, style: 'normal', complexScript: false },
};

/** Every input shape of a floating drawing the layout receives. */
const host = (extra: object = {}): Run => ({ type: 'anchorHost', fontSize: 40, fontFamily: 'Arial',
  fontFamilyEastAsia: 'Yu Mincho', bold: true, italic: false, ...extra }) as unknown as Run;
const acquisition = { anchorAcquisitionInput: { occurrenceId: 'a' } };
const anchorShapes: readonly (readonly [string, readonly Run[]])[] = [
  ['parser host and picture payload', [host(), picture({ anchor: true, ...acquisition })]],
  ['parser host and shape payload', [host(), { type: 'shape', widthPt: 40, heightPt: 20, ...acquisition } as unknown as Run]],
  ['parser host and unavailable placeholder', [host(), { type: 'unavailableDrawing', resourceKind: 'image',
    widthPt: 40, heightPt: 20, ...acquisition } as unknown as Run]],
  ['public anchored picture', [picture({ anchor: true, widthPt: 40, heightPt: 60 })]],
  ['public anchored chart', [{ type: 'chart', chart: {}, widthPt: 40, heightPt: 60, anchor: true } as unknown as Run]],
  ['public floating shape', [{ type: 'shape', widthPt: 40, heightPt: 60 } as unknown as Run]],
  ['public floating group', [{ type: 'shape', widthPt: 40, heightPt: 60, children: [] } as unknown as Run]],
];

const characterGrid = { characterGrid: { active: true, kind: 'linesAndChars', pitchPt: 12, deltaPt: 2 } } as Partial<ParagraphLayoutContext>;
const bases: readonly (readonly [string, readonly Run[], Partial<DocParagraph>?, Partial<ParagraphMeasurementEnvironment>?,
  Partial<ParagraphLayoutContext>?])[] = [
  ['empty', []],
  ['text', [text('Latin words fill')]],
  ['wrapping text of two sizes', [text('Latin words fill the band and wrap onto later lines '), text('larger words follow here', 18)]],
  ['inline picture', [picture(inlineObject)], { defaultFontSize: 40 }, { inlineObjectRunLeadingAdmitted: true }],
  ['inline picture beside text', [text('x '), picture(inlineObject)], {}, { inlineObjectRunLeadingAdmitted: true }],
  ['ruby', [text('漢字', 10, { ruby: { text: 'かんじ', fontSizePt: 5, fontFamily: null, bold: false, italic: false,
    hpsRaisePt: 9, align: 'center' } }), text('の文')]],
  ['raised text', [text('raised', 10, { position: 6 })]],
  ['text ending at a line break', [text('before'), { type: 'break', breakType: 'line' } as unknown as Run]],
  ['tracked deletion beside text', [text('kept'), text('gone', 30, deletion)]],
  ['tracked deletion alone', [text('gone', 30, deletion)]],
  ['deleted ruby alone', [text('漢', 30, { ...deletion, ruby: { text: 'かん', fontSizePt: 15, hpsRaisePt: 20 } })]],
  ['soft hyphen alone', [text('\u00AD', 30)]],
  ['optional hyphen marker alone', [text('', 30, { optionalHyphen: true })]],
  ['empty text run alone', [text('', 30)]],
  ['space alone', [text(' ', 30)]],
  ['one word in two runs', [text('Latinwordsfillthebandand'), text('wraponlywhereawordends here')]],
  ['justified text', [text('Latin words fill the band '), text('and wrap onto later lines here')], { alignment: 'both' },
    {}, { isJustified: true }],
  ['text on a character grid', [text('Latin '), text('words 漢字 here')], {}, {}, characterGrid],
  ['text and a tab', [text('a'), { type: 'tab' } as unknown as Run, text('b')]],
  ['math alone', [{ type: 'math', nodes: [{ kind: 'run', text: 'x' }], display: false, fontSize: 10 } as unknown as Run]],
  ['numbered text', [text('item')], { numbering: { text: '1.', numId: 1, level: 0, fontSize: 10 } as never }],
  ['numbered empty item', [], { numbering: { text: '1.', numId: 1, level: 0, fontSize: 10 } as never }],
];

const spacings = [null, { rule: 'exact', value: 14 }, { rule: 'atLeast', value: 14 }, { rule: 'auto', value: 1.5 }] as const;

function measure(runs: readonly Run[], overrides: Partial<DocParagraph>, context: Partial<ParagraphLayoutContext>,
  environment: Partial<ParagraphMeasurementEnvironment> & DocxCompatSettingsInput): MeasuredParagraph {
  const spacing = (context.lineSpacing ?? null) as DocParagraph['lineSpacing'];
  return measureParagraph(
    { alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0,
      lineSpacing: spacing, numbering: null, tabStops: [], defaultFontSize: 10, runs: [...runs], ...overrides } as DocParagraph,
    { lineGrid: { active: false, pitchPt: null }, characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
      rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true }, physicalIndentLeftPt: 0,
      physicalIndentRightPt: 0, firstIndentPt: 0, lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0, baseRtl: false,
      isJustified: false, stretchLastLine: false, tabStops: [],
      hasRuby: runs.some((run) => Boolean((run as { ruby?: unknown }).ruby)),
      hasEastAsianText: runs.some((run) => /[぀-鿿]/u.test((run as { text?: string }).text ?? '')),
      kinsoku: DEFAULT_KINSOKU_RULES, defaultTabPt: 36, ...context } as ParagraphLayoutContext,
    { startYPt: 0, paragraphXPt: 0, availableWidthPt: 150, maximumYPt: 10000 } as never,
    measurer,
    { pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb', documentHasEastAsianText: false,
      ...withCompat(environment) } as ParagraphMeasurementEnvironment,
  );
}

/** Everything that sizes or places a line, without the anchors themselves. */
function metrics(measured: MeasuredParagraph) {
  const { lines, markAnchorLine: _anchors, placement: _placement, ...paragraphFacts } = measured;
  return {
    ...paragraphFacts,
    lines: lines.filter((line) => !line.emptyMarkLine).map((line) => ({
      topYPt: line.topYPt, advancePt: line.advancePt,
      ascent: line.layout.ascent, descent: line.layout.descent, height: line.layout.height,
      intendedSingle: line.layout.intendedSingle, gridCountSingle: line.layout.gridCountSingle,
      visibleAscent: line.layout.visibleAscent, visibleDescent: line.layout.visibleDescent,
      eastAsian: line.layout.eastAsian, hasRuby: line.layout.hasRuby,
      uniformPositionAuto: line.layout.uniformPositionAuto,
      inlinePictureTextSingle: line.layout.inlinePictureTextSingle,
      xOffset: line.layout.xOffset, availWidth: line.layout.availWidth, endsWithBreak: line.layout.endsWithBreak,
      // The anchoring character separates two shaping sequences, so a text
      // may arrive in more pieces: its content, width and position may not
      // change.
      text: line.layout.segments.map((segment) => 'text' in segment ? segment.text
        : 'imagePath' in segment ? (segment.anchor ? '' : '\uFFFC') : 'isTab' in segment ? '\t' : '\uFFFC').join(''),
      width: line.layout.segments.reduce((sum, segment) => sum + (segment.measuredWidth ?? 0), 0),
      positions: [...new Set(line.layout.segments.flatMap((segment) => 'text' in segment && segment.text
        ? [segment.lineRelativePosition ?? 0] : []))],
    })),
  };
}

describe('WORD_ANCHOR_HOST_NO_LINE_METRIC: a floating anchor changes no line metric', () => {
  for (const [baseName, base, overrides = {}, env = {}, baseContext = {}] of bases) {
    for (const spacing of spacings) {
      for (const grid of [false, true]) {
        for (const compatibilityMode of [14, 15]) {
          const context = { ...baseContext, lineSpacing: spacing as ParagraphLayoutContext['lineSpacing'],
            ...(grid ? { lineGrid: { active: true, pitchPt: 18 } } : {}) };
          const settings = { ...env, compatibilityMode, ...(grid ? { useFeLayout: true } : {}) };
          it(`${baseName}, ${spacing ? `${spacing.rule} ${spacing.value}` : 'no'} spacing, ${grid ? 'grid' : 'no grid'}, mode ${compatibilityMode}`, () => {
            for (const showTrackedChanges of [false, true]) {
              const view = { ...settings, showTrackedChanges };
              const expected = metrics(measure(base, overrides, context, view));
              for (const [shapeName, anchor] of anchorShapes) {
                for (const [where, runs] of [
                  ['before', [...anchor, ...base]],
                  ['inside', [...base.slice(0, 1), ...anchor, ...base.slice(1)]],
                  ['after', [...base, ...anchor]],
                  ['twice', [...anchor, ...base, ...anchor]],
                ] as const) {
                  expect(metrics(measure(runs, overrides, context, view)),
                    `${shapeName} ${where}, ${showTrackedChanges ? 'markup' : 'final'} view`).toEqual(expected);
                }
              }
            }
          });
        }
      }
    }
  }
});

// ── Whole documents: pagination, headers and table cells ───────────────────

import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { AnchorAcquisitionInput } from './layout/anchor-input.js';
import type { PaintNode } from './layout/types.js';
import type { BodyElement, DocxDocumentModel } from './types.js';
import { attachAnchorPlacements } from './line-breaker/model.js';
import { planLine } from './layout/paragraph.js';

function documentContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  const size = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (value: string) => ({
      width: [...value].length * size() * 0.5,
      fontBoundingBoxAscent: size() * 0.8, fontBoundingBoxDescent: size() * 0.2,
      actualBoundingBoxAscent: size() * 0.8, actualBoundingBoxDescent: size() * 0.2,
    } as TextMetrics),
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, textAlign: 'left', direction: 'ltr',
  } as unknown as CanvasRenderingContext2D;
}

const missingEdges = { topPt: null, topStatus: 'missing', rightPt: null, rightStatus: 'missing',
  bottomPt: null, bottomStatus: 'missing', leftPt: null, leftStatus: 'missing' };
/** A wrapNone drawing positioned from the page: it excludes no text. */
const pageAnchor = (occurrenceId: string): AnchorAcquisitionInput => ({
  occurrenceId,
  simplePosition: { enabled: false, status: 'valid', xPt: 0, xStatus: 'valid', yPt: 0, yStatus: 'valid' },
  horizontal: { relativeFrom: 'page', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: 150 } },
  vertical: { relativeFrom: 'paragraph', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: 0 } },
  extent: { widthPt: 30, heightPt: 50, widthStatus: 'valid', heightStatus: 'valid' },
  parentEffectExtent: missingEdges, anchorDistances: missingEdges,
  relativeSize: { horizontal: null, vertical: null },
  wrap: { kind: 'none', authoredKinds: ['wrapNone'], side: null, distances: missingEdges, effectExtent: null, polygon: null },
  behavior: { behindDoc: false, behindDocStatus: 'valid', relativeHeight: 1, relativeHeightStatus: 'valid',
    locked: false, lockedStatus: 'valid', allowOverlap: true, allowOverlapStatus: 'valid',
    layoutInCell: true, layoutInCellStatus: 'valid' },
  group: null,
}) as unknown as AnchorAcquisitionInput;

const publicPlacement = { anchorXPt: 150, anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: true,
  anchorXRelativeFrom: 'page', anchorYRelativeFrom: 'paragraph', wrapMode: 'none' };
const documentAnchorShapes: readonly (readonly [string, () => Run[]])[] = [
  ['parser host and picture payload', () => [
    { type: 'anchorHost', fontSize: 40, bold: true, __anchorOccurrenceId: 'p' } as unknown as Run,
    picture({ anchor: true, widthPt: 30, heightPt: 50, imagePath: 'word/media/p.png', __anchorAcquisition: pageAnchor('p') }),
  ]],
  ['parser host and unavailable placeholder', () => [
    { type: 'anchorHost', fontSize: 40, __anchorOccurrenceId: 'u' } as unknown as Run,
    { type: 'unavailableDrawing', resourceKind: 'image', widthPt: 30, heightPt: 50,
      __anchorAcquisition: pageAnchor('u') } as unknown as Run,
  ]],
  ['public anchored picture', () => [picture({ anchor: true, widthPt: 30, heightPt: 50,
    imagePath: 'word/media/p.png', ...publicPlacement })]],
  ['public anchored chart', () => [{ type: 'chart', chart: { plotArea: {} }, widthPt: 30, heightPt: 50,
    anchor: true, ...publicPlacement } as unknown as Run]],
  ['public floating shape', () => [{ type: 'shape', widthPt: 30, heightPt: 50, presetGeometry: 'rect',
    subpaths: [], fill: null, stroke: null, rotation: 0, flipH: false, flipV: false, ...publicPlacement } as unknown as Run]],
  ['public floating group', () => [0, 20].map((offset) => ({ type: 'shape', widthPt: 10, heightPt: 50,
    presetGeometry: 'rect', subpaths: [], fill: null, stroke: null, rotation: 0, flipH: false, flipV: false,
    ...publicPlacement, anchorXPt: 150 + offset }) as unknown as Run)],
];

type ParagraphOptions = Partial<DocParagraph>;
const bodyParagraph = (runs: readonly Run[], options: ParagraphOptions = {}): BodyElement => ({
  type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0,
  spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [], defaultFontSize: 12,
  defaultFontFamily: 'NotInMetrics', runs: [...runs], ...options,
}) as unknown as BodyElement;
const word = (value: string, fontSize = 12, extra: object = {}): Run =>
  text(value, fontSize, { fontFamily: 'NotInMetrics', ...extra });

function documentModel(body: BodyElement[], header: BodyElement[] | null, settings: object): DocxDocumentModel {
  return {
    section: { pageWidth: 200, pageHeight: 200, marginTop: 40, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 5, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
      sectionStart: 'nextPage', columns: null },
    body, settings,
    headers: { default: header ? { body: header } : null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [], endnotes: [], fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

const tableOf = (content: BodyElement[], cell: object = {}): BodyElement => ({
  type: 'table', colWidths: [160], borders: { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null },
  cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'fixed',
  rows: [{ rowHeight: null, rowHeightRule: 'auto', isHeader: false, cells: [{ content, colSpan: 1, vMerge: null,
    borders: { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null },
    background: null, vAlign: 'top', widthPt: 160, ...cell }] }],
}) as unknown as BodyElement;

const round = (value: number): number => Math.round(value * 1e6) / 1e6;

/** Every placed block of a page layer with the geometry of its lines and of
 * each glyph cluster, without the floating drawings and their hosts. */
function geometry(node: PaintNode): unknown {
  if (node.kind === 'paragraph') {
    return {
      kind: node.kind, y: round(node.flowBounds.yPt), height: round(node.flowBounds.heightPt),
      mark: node.paragraphMark ? [round(node.paragraphMark.bounds.yPt), round(node.paragraphMark.bounds.heightPt)] : null,
      lines: node.lines.flatMap((line) => {
        const content = line.placements.filter((placement) => placement.kind !== 'anchor-host'
          && !(placement.kind === 'drawing' && placement.advancePt === 0));
        // The mark line that carries an anchor-only paragraph's anchors.
        if (content.length === 0 && line.markLine) return [];
        return [{
          y: round(line.bounds.yPt), advance: round(line.advancePt), baseline: round(line.baselinePt),
          glyphs: content.flatMap((placement) => placement.kind === 'text'
            ? placement.clusters.map((cluster) => round(placement.origin.xPt + cluster.offset.xPt))
            : [round(placement.bounds?.xPt ?? 0)]),
          text: content.map((placement) => placement.kind === 'text' ? placement.text : `<${placement.kind}>`).join(''),
        }];
      }),
    };
  }
  if (node.kind === 'table') {
    return {
      kind: node.kind, y: round(node.flowBounds.yPt), height: round(node.flowBounds.heightPt),
      rows: node.rows.map((row) => ({ height: round(row.heightPt), cells: row.cells.map((cell) =>
        cell.blocks.map((block) => ({ offset: round(block.offsetPt), layout: geometry(block.layout as PaintNode) }))) })),
    };
  }
  return { kind: node.kind };
}

function documentGeometry(model: DocxDocumentModel, showTrackedChanges: boolean) {
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: documentContext() }),
    { currentDateMs: 0, showTrackedChanges });
  return {
    geometry: layout.pages.map((page) => ({
      header: page.layers.header.map(geometry),
      body: page.layers.body.map(geometry),
    })),
    drawings: layout.pages.reduce((count, page) =>
      count + page.layers.paintOrder.filter((entry) => entry.kind === 'drawing').length, 0),
  };
}

describe('WORD_ANCHOR_HOST_NO_LINE_METRIC: a floating anchor changes no pagination', () => {
  const numbering = { numId: 1, level: 0, format: 'decimal', text: '1.', indentLeft: 36, tab: 36, suff: 'tab',
    fontFamily: 'NotInMetrics', fontFacts: { fontSize: 12, fontFamily: 'NotInMetrics' } };
  const subjects: readonly (readonly [string, readonly Run[], ParagraphOptions?])[] = [
    ['empty paragraph', []],
    ['empty paragraph kept with the next', [], { keepNext: true }],
    ['numbered empty item', [], { numbering, indentLeft: 36, indentFirst: -36 } as ParagraphOptions],
    ['text', [word('Latin words fill the band and wrap onto later lines')]],
    ['justified text', [word('Latin words fill the band and wrap '), word('onto later lines here')], { alignment: 'both' }],
    ['numbered text', [word('item')], { numbering, indentLeft: 36, indentFirst: -36 } as ParagraphOptions],
    ['inline picture', [picture({ imagePath: 'word/media/i.png', heightPt: 16 })], { defaultFontSize: 30 }],
    ['ruby', [word('漢字', 12, { ruby: { text: 'かんじ', fontSizePt: 6, fontFamily: null, bold: false, italic: false,
      hpsRaisePt: 11, align: 'center' } }), word('の文')]],
    ['tracked deletion', [word('gone', 30, deletion)]],
    ['tracked deletion beside text', [word('kept '), word('gone', 30, deletion)]],
    ['text ending at a line break', [word('before'), { type: 'break', breakType: 'line' } as unknown as Run]],
  ];
  const spacings: readonly (readonly [string, ParagraphOptions])[] = [
    ['automatic', {}],
    ['exact 14', { lineSpacing: { rule: 'exact', value: 14, explicit: true } } as ParagraphOptions],
    ['atLeast 14', { lineSpacing: { rule: 'atLeast', value: 14, explicit: true } } as ParagraphOptions],
    ['auto 1.5', { lineSpacing: { rule: 'auto', value: 1.5, explicit: true } } as ParagraphOptions],
  ];
  // The subject starts at the page top and with 12.5, 11.5 and 9.5pt left on
  // the page: a 12pt text line fits only the first, an empty 12pt mark (page
  // fit by its baseline) all but the last.
  const filler = (heightPt: number): BodyElement => bodyParagraph([word('x')],
    { lineSpacing: { rule: 'exact', value: heightPt, explicit: true } } as ParagraphOptions);
  const stories: readonly (readonly [string, (subject: BodyElement) => readonly [BodyElement[], BodyElement[] | null]])[] = [
    ['body, page top', (subject) => [[subject, bodyParagraph([word('After')])], null]],
    ['body, 12.5pt left', (subject) => [[filler(127.5), subject, bodyParagraph([word('After')])], null]],
    ['body, 11.5pt left', (subject) => [[filler(128.5), subject, bodyParagraph([word('After')])], null]],
    ['body, 9.5pt left', (subject) => [[filler(130.5), subject, bodyParagraph([word('After')])], null]],
    ['header', (subject) => [[bodyParagraph([word('Body')])], [subject, bodyParagraph([word('Header')])]]],
    ['table cell', (subject) => [[tableOf([bodyParagraph([word('Cell')]), subject, bodyParagraph([word('End')])]),
      bodyParagraph([word('After')])], null]],
    ['last paragraph of a table cell', (subject) => [[tableOf([bodyParagraph([word('Cell')]), subject]),
      bodyParagraph([word('After')])], null]],
    // §17.4.21: the end-of-cell mark of a hideMark cell owns no row height.
    ['sole paragraph of a hideMark table cell', (subject) => [[tableOf([subject], { hideMark: true }),
      bodyParagraph([word('After')])], null]],
  ];
  const settingsMatrix: readonly (readonly [string, object, object])[] = [
    ['mode 15', { compatibilityMode: 15 }, {}],
    ['mode 14', { compatibilityMode: 14 }, {}],
    ['mode 14 on a lines grid', { compatibilityMode: 14, useFeLayout: true }, { docGridType: 'lines', docGridLinePitch: 18 }],
  ];
  // Each input shape is laid out once per case, before or after the content
  // in turn, so every shape meets every position across the matrix (the
  // position inside the content is the measurement matrix's above).
  let turn = 0;
  for (const [subjectName, runs, options = {}] of subjects) {
    const tracked = runs.some((run) => (run as { revision?: unknown }).revision !== undefined);
    for (const [storyName, story] of stories) {
      it(`${subjectName} in ${storyName}`, () => {
        for (const [spacingName, spacing] of spacings) {
          for (const [settingsName, settings, grid] of settingsMatrix) {
            for (const showTrackedChanges of tracked ? [false, true] : [false]) {
              const lay = (subjectRuns: readonly Run[]) => {
                const [body, header] = story(bodyParagraph(subjectRuns, { ...options, ...spacing }));
                const model = documentModel(body, header, settings);
                Object.assign(model.section, grid);
                return documentGeometry(model, showTrackedChanges);
              };
              const expected = lay(runs);
              for (const [shapeName, anchor] of documentAnchorShapes) {
                turn += 1;
                for (const [where, anchored] of [turn % 2 === 0
                  ? ['before', [...anchor(), ...runs]] as const : ['after', [...runs, ...anchor()]] as const]) {
                  const label = `${shapeName} ${where}; ${spacingName} spacing, ${settingsName}, `
                    + `${showTrackedChanges ? 'markup' : 'final'} view`;
                  const laid = lay(anchored);
                  expect(laid.geometry, label).toEqual(expected.geometry);
                  // The drawings themselves are placed, never omitted.
                  expect(laid.drawings, label).toBeGreaterThan(expected.drawings);
                }
              }
            }
          }
        }
      });
    }
  }
});

describe('a hideMark cell whose final paragraph holds only floating anchors', () => {
  const numbering = { numId: 1, level: 0, format: 'decimal', text: '1.', indentLeft: 36, tab: 36, suff: 'tab',
    fontFamily: 'NotInMetrics', fontFacts: { fontSize: 12, fontFamily: 'NotInMetrics' } };
  const rowHeight = (runs: readonly Run[], options: ParagraphOptions = {}, showTrackedChanges = false): number => {
    const model = documentModel([tableOf([bodyParagraph(runs, options)], { hideMark: true }),
      bodyParagraph([word('After')])], null, { compatibilityMode: 15 });
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: documentContext() }),
      { currentDateMs: 0, showTrackedChanges });
    const table = layout.pages[0]!.layers.body.find((node) => node.kind === 'table');
    if (table?.kind !== 'table') throw new Error('Expected a table');
    return table.rows[0]!.heightPt;
  };
  const anchor = documentAnchorShapes[0]![1];
  const numbered = { numbering, indentLeft: 36, indentFirst: -36 } as ParagraphOptions;

  it('keeps the flow height of a painted list marker', () => {
    // The 12pt "1." is ink: its row may not collapse under the rows below.
    expect(rowHeight([], numbered)).toBeCloseTo(12, 6);
    expect(rowHeight(anchor(), numbered)).toBe(rowHeight([], numbered));
    expect(rowHeight(anchor())).toBe(rowHeight([]));
    expect(rowHeight([])).toBeLessThan(12);
  });

  it('classifies the mark in the active revision view', () => {
    const deleted = word('gone', 12, deletion);
    // The final view omits the deletion: still nothing but the mark.
    expect(rowHeight([...anchor(), deleted])).toBe(rowHeight([]));
    expect(rowHeight([deleted])).toBe(rowHeight([]));
    // The markup view shows it: the row holds its line, with and without the anchor.
    expect(rowHeight([deleted], {}, true)).toBeCloseTo(12, 6);
    expect(rowHeight([...anchor(), deleted], {}, true)).toBe(rowHeight([deleted], {}, true));
  });
});

describe('floating anchors are placed in linear work', () => {
  const counted = <T extends object>(items: T[], reads: { count: number }): T[] => new Proxy(items, {
    get(target, key, receiver) {
      if (typeof key === 'string' && /^\d+$/.test(key)) reads.count += 1;
      return Reflect.get(target, key, receiver);
    },
  });

  it('attaches hosts to the finished lines without rescanning the hosts that follow', () => {
    const reads = (hostCount: number): number => {
      const hostSegment = { text: '', metricOnly: true, anchorHostWithoutMetric: true, measuredWidth: 0 };
      const content = { text: 'x', measuredWidth: 5, src: { segIndex: hostCount, charOffset: 0 } };
      const count = { count: 0 };
      const segs = counted([...Array.from({ length: hostCount }, () => ({ ...hostSegment })), content] as never[], count);
      const [line] = attachAnchorPlacements([{ segments: [content], height: 10, ascent: 8, descent: 2,
        intendedSingle: 0, gridCountSingle: 10, xOffset: 0, availWidth: 100,
        consumedEnd: { segIndex: hostCount + 1, charOffset: 0 } } as never], segs, undefined);
      expect(line!.segments).toHaveLength(hostCount + 1);
      expect(line!.segments.at(-1)).toBe(content);
      return count.count;
    };
    // A rescan per host reads about n * n / 2 segments.
    expect(reads(400)).toBeLessThan(reads(100) * 5);
  });

  it('plans a line holding many hosts in one sweep over its content', () => {
    const reads = (count: number): number => {
      const rangeReads = { count: 0 };
      // Zero-advance content (metric hosts that are no floating anchors),
      // each followed by a floating anchor host at the same offset.
      const content = Array.from({ length: count }, (_, index) => new Proxy({ kind: 'anchor-host',
        measuredWidthPt: 0, range: { start: index, end: index } }, {
        get(target, key, receiver) {
          if (key === 'range') rangeReads.count += 1;
          return Reflect.get(target, key, receiver);
        },
      }));
      const hosts = Array.from({ length: count }, (_, index) => ({ kind: 'anchor-host', anchorPlacement: true,
        measuredWidthPt: 0, range: { start: count, end: count }, anchorOccurrenceId: `a${index}` }));
      const planned = planLine({ paragraphXPt: 0, availableWidthPt: 100, alignment: 'left', baseRtl: false,
        isFirstLine: true, isLastLine: true, stretchLastLine: false,
        line: { range: { start: 0, end: count }, topPt: 0, baselinePt: 8, advancePt: 10, xOffsetPt: 0,
          availableWidthPt: 100, endsWithBreak: false, segments: [...content, ...hosts] as never } });
      expect(planned.placements).toHaveLength(2 * count);
      // Every floating host follows all content, in source order.
      expect(planned.placements.slice(count).map((placement) =>
        (placement as { anchorOccurrenceId?: string }).anchorOccurrenceId)).toEqual(hosts.map((host) => host.anchorOccurrenceId));
      return rangeReads.count;
    };
    // A search per host reads every content range again for each host.
    expect(reads(400)).toBeLessThan(reads(100) * 5);
  });
});
