import { describe, expect, it } from 'vitest';
import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService } from './layout/text.js';
import { createLayoutServices } from './layout-runtime.js';
import { createLayoutServicesRuntimeView } from './layout/runtime-state.js';
import { paragraphAcquisitionInput } from './parser-model.js';
import { acquireParagraphResult, type ParagraphAcquisitionOptions } from './layout/paragraph.js';
import { createParagraphAcquisitionCacheServicesView } from './layout/runtime-state.js';
import { layoutDocument } from './document-layout.js';
import { paintLayoutPage } from './paint/canvas-page.js';
import { buildSegments, segAdvanceWidth } from './line-layout.js';
import { measureParagraphIntrinsicWidths } from './layout/intrinsic-width.js';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import type { ParagraphLayoutContext } from './layout-context.js';
import type { DocParagraph, DocxDocumentModel, DocTable } from './types.js';
import { makeRecordingCanvas } from './test-support/recording-canvas.test-support.js';

type Variant = 'fixed' | 'proportional' | 'native-pair' | 'missing-proof';
function service(variant: Variant = 'fixed', measured?: string[], proofs?: string[]) {
  return createTextLayoutService({
    fonts: createFontResolver(['Pair Face', 'Following Face'].map(family => ({
      requestedFamily: family, resolvedFamily: `${family} Loaded`, source: 'local',
      resourceIdentity: `office-local:${family}`, weight: 400, style: 'normal',
      coverageProbeFamily: `${family} Zero`,
    }))),
    measurer: {
      fingerprint: `closing-pair:${variant}`,
      ...(variant === 'missing-proof' ? {} : { ownsGlyphs: () => { proofs?.push('ownership'); return true; }, sameGlyphRaster: () => { proofs?.push('raster'); return true; } }),
      measure(r) {
        measured?.push(r.text);
        let advancePt = [...r.text].reduce((sum, c) => sum + r.fontSizePt
          * (variant === 'proportional' && c === '。' ? 0.6 : 1), 0);
        if (variant === 'native-pair') advancePt -= (r.text.match(/。」/gu)?.length ?? 0) * r.fontSizePt / 2;
        return { advancePt, ascentPt: r.fontSizePt * 0.8, descentPt: r.fontSizePt * 0.2,
          inkBounds: { xMinPt: 0, xMaxPt: advancePt / 2, ascentPt: r.fontSizePt * 0.8, descentPt: r.fontSizePt * 0.2 },
          horizontalInkBoundsAreTight: true };
      },
    },
  });
}
function run(text: string, extra: Record<string, unknown> = {}) {
  return { type: 'text', text, fontSize: 12, fontFamily: 'Pair Face', fontFamilyEastAsia: 'Pair Face',
    bold: false, italic: false, underline: false, strikethrough: false, ...extra } as DocParagraph['runs'][number];
}
function paragraph(runs: DocParagraph['runs']): DocParagraph {
  return { alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs, defaultFontSize: 12, defaultFontFamily: 'Pair Face', widowControl: false, kinsoku: false } as DocParagraph;
}
function model(p: DocParagraph): DocxDocumentModel {
  return { body: [{ type: 'paragraph', ...p }], settings: { compatibilityMode: 15 },
    section: { pageWidth: 468, pageHeight: 600, marginLeft: 0, marginRight: 0, marginTop: 0, marginBottom: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false },
    headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null } } as DocxDocumentModel;
}
const intrinsicContext = {
  lineGrid: { active: false, pitchPt: null }, characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
  rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true }, physicalIndentLeftPt: 0, physicalIndentRightPt: 0,
  firstIndentPt: 0, lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0, baseRtl: false, isJustified: false,
  stretchLastLine: false, tabStops: [], hasRuby: false, hasEastAsianText: true, kinsoku: DEFAULT_KINSOKU_RULES, defaultTabPt: 36,
} as ParagraphLayoutContext;

describe('first-owner closing-pair allocation', () => {
  it('shares acquisition advance through intrinsic width, a carried line and published paint', async () => {
    const p = paragraph([run('あ'.repeat(39) + 'り。」い')]);
    const doc = model(p); const { canvas, fills } = makeRecordingCanvas();
    const services = createLayoutServicesRuntimeView(createLayoutServices(doc, { measureContext: canvas.getContext('2d') }), { text: service() });
    const layout = layoutDocument(doc, services);
    const published = layout.pages[0].layers.body.find(node => node.kind === 'paragraph');
    if (!published || published.kind !== 'paragraph') throw Error('Missing paragraph');
    expect(published.lines.map(l => l.placements.filter(p => p.kind === 'text').map(p => p.text).join('')))
      .toEqual(['あ'.repeat(39), 'り。」い']);
    const tail = published.lines[1].placements.filter(p => p.kind === 'text');
    const clusters = tail.flatMap(p => p.clusters);
    expect(clusters.map(c => c.advancePt)).toEqual([12, 6, 12, 12]);
    expect(tail.flatMap(p => p.paintOps).map(op => op.text).join('')).toBe('り。」い');
    await paintLayoutPage(layout, 0, canvas, { scale: 1, dpr: 1 });
    const period = fills.find(f => f.text === '。'); const close = fills.find(f => f.text.startsWith('」'));
    expect(period).toBeDefined(); expect(close).toBeDefined();
    expect((close?.x ?? 0) - (period?.x ?? 0)).toBe(6);
    const widths = measureParagraphIntrinsicWidths(p, intrinsicContext, 10000,
      { context: canvas.getContext('2d') as CanvasRenderingContext2D, fontFamilyClasses: {} },
      { pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb', documentHasEastAsianText: true,
        compatibilityMode: 15, closingPairRootSource: true, layoutServices: services });
    expect(widths.maxWidthPt).toBe(510);
  });


  it('keeps root eligibility owned through exact and reusable line cache hits', () => {
    const p = paragraph([run('。」')]), doc = model(p);
    const { canvas } = makeRecordingCanvas();
    const services = createParagraphAcquisitionCacheServicesView(createLayoutServicesRuntimeView(
      createLayoutServices(doc, { measureContext: canvas.getContext('2d') }), { text: service() }));
    const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
    const input = paragraphAcquisitionInput(p, source);
    const options: ParagraphAcquisitionOptions = {
      id: 'root', source, flowDomainId: 'body', ordinaryFlow: true, context: intrinsicContext,
      placement: { startYPt: 0, paragraphXPt: 0, availableWidthPt: 468, maximumYPt: 600, suppressSpaceBefore: true },
      measurer: { context: canvas.getContext('2d') as CanvasRenderingContext2D, fontFamilyClasses: {} },
      environment: { pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb', documentHasEastAsianText: true,
        compatibilityMode: 15, closingPairRootSource: true, layoutServices: services },
      exclusions: [], paragraphBorderEdges: { top: 'top', bottom: 'bottom' }, trailingExtentPt: 0,
    };
    const width = (o: ParagraphAcquisitionOptions) => acquireParagraphResult(input, o).layout.lines[0]!.placements
      .filter(piece => piece.kind === 'text').reduce((sum, piece) => sum + piece.advancePt, 0);
    expect(width(options)).toBe(18);
    expect(width({ ...options, environment: { ...options.environment, closingPairRootSource: undefined } })).toBe(24);
    expect(width({ ...options, id: 'nested', source: { story: 'header', storyInstance: 'header:1', path: [0] },
      placement: { ...options.placement, startYPt: 20 } })).toBe(24);
    expect(width({ ...options, id: 'rebase', placement: { ...options.placement, startYPt: 30 } })).toBe(18);
  });


  it('preserves a Kashida full cell while projecting an observed left boundary', () => {
    const source = 'あ。」い';
    const low = { ...paragraph([run(source)]), alignment: 'lowKashida' } as DocParagraph;
    const left = paragraph([run(source)]);
    const doc = model(low);
    doc.body.push({ type: 'paragraph', ...left });
    doc.section.pageWidth = 48;
    const { canvas } = makeRecordingCanvas();
    const services = createLayoutServicesRuntimeView(createLayoutServices(doc, { measureContext: canvas.getContext('2d') }), { text: service() });
    const published = layoutDocument(doc, services).pages.flatMap(page => page.layers.body)
      .filter(node => node.kind === 'paragraph');
    expect(published).toHaveLength(2);
    const advances = published.map(p => p.lines.flatMap(line => line.placements)
      .filter(piece => piece.kind === 'text').flatMap(piece => piece.clusters)
      .map(cluster => [cluster.range.start, cluster.range.end, cluster.advancePt]));
    expect(advances).toEqual([
      [[0, 1, 12], [1, 2, 12], [2, 3, 12], [3, 4, 12]],
      [[0, 1, 12], [1, 2, 6], [2, 3, 12], [3, 4, 12]],
    ]);
  });

  it('projects all seven measured closing boundaries and preserves an opening counter', () => {
    const source = 'あ。」い、」う」」え」。お」、か」「き）（く「」';
    const doc = model(paragraph([run(source)]));
    const { canvas } = makeRecordingCanvas();
    const services = createLayoutServicesRuntimeView(createLayoutServices(doc, { measureContext: canvas.getContext('2d') }), { text: service() });
    const published = layoutDocument(doc, services).pages[0].layers.body.find(n => n.kind === 'paragraph');
    if (!published || published.kind !== 'paragraph') throw Error('Missing paragraph');
    const clusters = published.lines.flatMap(line => line.placements)
      .filter(piece => piece.kind === 'text').flatMap(piece => piece.clusters);
    // Each source-owned boundary is independent: losing a closing projection
    // cannot be hidden by accidentally projecting the opening at source22.
    const firstOwners = [1, 4, 7, 10, 13, 16, 19];
    expect(clusters.map(cluster => [cluster.range.start, cluster.range.end, cluster.advancePt]))
      .toEqual(Array.from({ length: 24 }, (_, start) => [start, start + 1, firstOwners.includes(start) ? 6 : 12]));
  });

  it('preserves identical-format source ownership across a split pair', () => {
    const doc = model(paragraph([run('あ。'), run('」い')]));
    const { canvas } = makeRecordingCanvas();
    const services = createLayoutServicesRuntimeView(createLayoutServices(doc, { measureContext: canvas.getContext('2d') }), { text: service() });
    const published = layoutDocument(doc, services).pages[0].layers.body.find(n => n.kind === 'paragraph');
    if (!published || published.kind !== 'paragraph') throw Error('Missing paragraph');
    const pieces = published.lines.flatMap(l => l.placements).filter(p => p.kind === 'text');
    expect(pieces.map(p => p.text).join('')).toBe('あ。」い');
    expect(pieces.reduce((sum, p) => sum + p.advancePt, 0)).toBe(42);
    expect(pieces.flatMap(p => p.clusters).map(c => [c.range.start, c.range.end, c.advancePt]))
      .toEqual([[0, 1, 12], [1, 2, 6], [2, 3, 12], [3, 4, 12]]);
    expect(pieces.flatMap(p => p.sourceRuns ?? []).map(r => [r.sourceRunIndex, r.range.start, r.range.end]))
      .toEqual([[0, 0, 2], [1, 2, 4]]);
  });

  it('keeps AutoFit cell intrinsic width natural inside a body-owned table', () => {
    const p = paragraph([run('。」')]);
    const borders = { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null };
    const table = { type: 'table', layout: 'autofit', colWidths: [1], widthPt: 1, borders, jc: 'left',
      cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
      rows: [{ rowHeight: null, rowHeightRule: 'auto', isHeader: false,
        cells: [{ content: [{ type: 'paragraph', ...p }], colSpan: 1, vMerge: null, borders,
          background: null, vAlign: 'top', widthPt: null, widthPct: null,
          marginTop: 0, marginBottom: 0, marginLeft: 0, marginRight: 0 }] }] } as unknown as DocTable & { type: 'table' };
    const doc = { ...model(p), body: [table] };
    const { canvas } = makeRecordingCanvas();
    const services = createLayoutServicesRuntimeView(createLayoutServices(doc, { measureContext: canvas.getContext('2d') }), { text: service() });
    const published = layoutDocument(doc, services).pages[0].layers.body.find(n => n.kind === 'table');
    if (!published || published.kind !== 'table') throw Error('Missing table');
    expect(published.columnWidthsPt).toEqual([24]);
  });

  it('does not project across a field hidden inside an identical-format sequence', () => {
    const text = service();
    const field = { ...run('unused'), type: 'field', fieldType: 'other', instruction: 'UNKNOWN', fallbackText: '」' } as DocParagraph['runs'][number];
    const segments = buildSegments([run('。'), field], { pageIndex: 0, totalPages: 1, compatibilityMode: 15,
      closingPairAdvanceAdmitted: true, layoutServices: { text } as never });
    expect(segments.map(s => 'text' in s ? s.text : '').join('')).toBe('。」');
    expect(segments.reduce((sum, seg) => sum + ('text' in seg ? segAdvanceWidth(seg,
      text.shape({ ...seg.textShapeRequest!, measure: true, clusterGeometry: false }).advancePt, undefined, 1) : 0), 0)).toBe(24);
  });

  it('preflights huge unselected font metadata before new native or proof work', () => {
    const measured: string[] = [], proofs: string[] = [], text = service('fixed', measured, proofs);
    const segments = buildSegments([run('。」', { fontSlots: { direct: { ascii: 'unused'.repeat(200000), eastAsia: 'Pair Face' }, theme: {}, themePresent: {} } })],
      { pageIndex: 0, totalPages: 1, compatibilityMode: 15, closingPairAdvanceAdmitted: true, layoutServices: { text } as never });
    expect(measured).toEqual([]); expect(proofs).toEqual([]);
    expect(segments.reduce((sum, seg) => sum + ('text' in seg ? segAdvanceWidth(seg, seg.text.length * 12, undefined, 1) : 0), 0)).toBe(24);
  });


  it('does not spend optional native work on an unavailable selected route', () => {
    const measured: string[] = [], proofs: string[] = [], text = service('fixed', measured, proofs);
    const segments = buildSegments([run('。」', { fontFamilyEastAsia: 'Unavailable Face' })],
      { pageIndex: 0, totalPages: 1, compatibilityMode: 15, closingPairAdvanceAdmitted: true, layoutServices: { text } as never });
    expect(measured).toEqual([]); expect(proofs).toEqual([]);
    expect(segments.reduce((sum, seg) => sum + ('text' in seg ? segAdvanceWidth(seg, seg.text.length * 12, undefined, 1) : 0), 0)).toBe(24);
  });

  it('uses the first owner across a size and family seam, preserving source text', () => {
    const text = service();
    const segments = buildSegments([run('あ。', { fontSize: 20 }), run('」い', { fontSize: 8, fontFamily: 'Following Face', fontFamilyEastAsia: 'Following Face' })],
      { pageIndex: 0, totalPages: 1, compatibilityMode: 15, closingPairAdvanceAdmitted: true, layoutServices: { text } as never });
    const ctx = makeRecordingCanvas().canvas.getContext('2d') as CanvasRenderingContext2D;
    expect(segments.map(s => 'text' in s ? s.text : '').join('')).toBe('あ。」い');
    expect(segments.reduce((sum, s) => sum + ('text' in s ? segAdvanceWidth(s, text.shape({ ...s.textShapeRequest!, measure: true, clusterGeometry: false }).advancePt, undefined, 1) : 0), 0)).toBe(46);
  });


  it('declines a dense over-budget acquisition atomically before optional native work', () => {
    const measured: string[] = [];
    const text = service('fixed', measured);
    const source = '。」'.repeat(4096);
    const segments = buildSegments([run(source)], { pageIndex: 0, totalPages: 1, compatibilityMode: 15,
      closingPairAdvanceAdmitted: true, layoutServices: { text } as never });
    expect(measured).toEqual([]);
    expect(segments.reduce((sum, s) => sum + ('text' in s
      ? segAdvanceWidth(s, s.text.length * 12, undefined, 1) : 0), 0)).toBe(source.length * 12);
  });

  it('bounds the direct contextual batch endpoint work independently of pair proof metadata', () => {
    const measured: string[] = [], text = service('fixed', measured);
    const request = { text: '。」'.repeat(4096), fontSizePt: 12, fonts: { eastAsia: 'Pair Face' } };
    const result = text.acquireContextualAdvances?.([{ request, additionalWorkUtf16: 0,
      ranges: Array.from({ length: 4096 }, (_, i) => ({ start: 2 * i, end: 2 * i + 1 })) }]);
    expect(result).toBeUndefined();
    expect(measured).toEqual([]);
  });

  it('retains exact full-shape contextual rounding after another selected span', () => {
    const text = createTextLayoutService({ fonts: createFontResolver([]), measurer: {
      fingerprint: 'multispan-rounding', measure(r) {
        return { advancePt: r.text === 'X' ? 1e16 : r.text.length * 1.5, ascentPt: 1, descentPt: 0 };
      },
    } });
    const request = { text: 'X。」', fontSizePt: 12, fonts: { ascii: 'Latin', eastAsia: 'CJK' } };
    const full = text.shape({ ...request, measure: true, clusterGeometry: true });
    const sparse = text.acquireContextualAdvances?.([{ request, ranges: [{ start: 1, end: 2 }] }]);
    expect(full.spans).toHaveLength(2);
    expect(sparse?.[0].clusters).toEqual(full.clusters?.filter(c => c.range.start === 1));
    expect(sparse?.[0].clusters[0].advancePt).toBe(2);
  });

  it.each(['english', 'proportional', 'native-pair', 'missing-proof', 'prior-CSC', 'tab', 'unknown-service'] as const)('preserves the independent %s refusal', kind => {
    const text = service(kind === 'proportional' || kind === 'native-pair' || kind === 'missing-proof' ? kind : 'fixed');
    const runs = [run(kind === 'tab' ? '。\t」' : '。」', kind === 'english' ? { langEastAsia: 'en-US' } : {})];
    const env = { pageIndex: 0, totalPages: 1, compatibilityMode: 15, closingPairAdvanceAdmitted: true,
      ...(kind === 'prior-CSC' ? { characterSpacingControl: 'compressPunctuation' } : {}), layoutServices: { text: kind === 'unknown-service' ? { ...text, acquireContextualAdvances: undefined } : text } as never };
    const withPair = buildSegments(runs, env);
    const withoutPair = buildSegments(runs, { ...env, closingPairAdvanceAdmitted: false });
    const ctx = makeRecordingCanvas().canvas.getContext('2d') as CanvasRenderingContext2D;
    expect(withPair.map(s => 'text' in s ? segAdvanceWidth(s, text.shape({ ...s.textShapeRequest!, measure: true, clusterGeometry: false }).advancePt, undefined, 1) : 0))
      .toEqual(withoutPair.map(s => 'text' in s ? segAdvanceWidth(s, text.shape({ ...s.textShapeRequest!, measure: true, clusterGeometry: false }).advancePt, undefined, 1) : 0));
  });
});
