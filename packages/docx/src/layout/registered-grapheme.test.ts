import { describe, expect, it } from 'vitest';
import { createFontResolver, type FontInventoryFace } from './font-service.js';
import { createTextLayoutService, independentTextShapeRequest, replaceTextShapeRequest,
  sliceTextShapeRequest, type TextShapeRequest } from './text.js';
import type { ResolvedFontMetric } from '@silurus/ooxml-core';
import { createLayoutServices } from '../layout-runtime.js';
import { buildSegments, layoutLines, type LineLayoutEnvironment } from '../line-layout.js';
import type { DocRun, DocxDocumentModel } from '../types.js';
import { layoutDocument } from '../document-layout.js';
import { sourceOwnedTextPlacements } from './text-source-ownership.js';
import type { TextPlacement } from './types.js';
import { resolveDocumentLayoutSettings, resolveParagraphLayoutContext, resolveSectionLayoutContext } from '../layout-context.js';
import { measureParagraphIntrinsicWidths } from './intrinsic-width.js';

const identity = 'embedded:test-immutable-resource';
function fixture(ranges: ResolvedFontMetric['unicodeRanges'] | null = [[0x54, 0x54], [0x301, 0x301]],
  metricOverrides: Partial<ResolvedFontMetric> = {}, faceOverrides: Partial<FontInventoryFace> = {}) {
  return createTextLayoutService({
    fonts: createFontResolver([Object.fromEntries(Object.entries({
      requestedFamily: 'Test Face', resolvedFamily: 'Test Resource', resourceIdentity: identity,
      source: 'embedded', weight: 400, style: 'normal', ...faceOverrides,
    }).filter(([, value]) => value !== undefined)) as unknown as FontInventoryFace]),
    fontMetrics: { resource: { family: 'Test Resource', sourceIdentity: identity,
      weight: 400, style: 'normal', ...(ranges ? { unicodeRanges: ranges } : {}), ...metricOverrides } },
    measurer: { fingerprint: 'distinct-mark-anchors', measure: ({ text }) => ({
      advancePt: text === '\u0301' ? 0 : 6, ascentPt: 9, descentPt: 0,
      inkBounds: { xMinPt: 0, xMaxPt: text === 'T\u0301' ? 6 : 7, ascentPt: 9, descentPt: 0 },
      horizontalInkBoundsAreTight: true,
    }) },
  });
}
const request: TextShapeRequest = { text: 'T\u0301', fontSizePt: 10,
  fonts: { ascii: 'Test Face', highAnsi: 'Test Face' }, joinRegisteredGrapheme: true };

describe('registered single-grapheme physical shaping', () => {
  it('keeps semantic slots but measures one physical glyph unit and one cluster', () => {
    const result = fixture().shape(request);
    expect(result.spans).toHaveLength(1);
    expect(result.inkBounds?.xMaxPt).toBe(6);
    expect(result.spans[0].semanticSlotSpans?.map(s => [s.start, s.end, s.script])).toEqual([
      [0, 1, 'ascii'], [1, 2, 'highAnsi'],
    ]);
    expect(result.clusters).toEqual([{ range: { start: 0, end: 2 }, offsetPt: 0, advancePt: 6 }]);
    expect(Object.isFrozen(result.spans[0].semanticSlotSpans)).toBe(true);
  });
  it.each([null, [[0x54, 0x54]], []] as (ResolvedFontMetric['unicodeRanges'] | null)[])(
    'rejects absent or incomplete explicit cmap %j', ranges => {
      expect(fixture(ranges).shape(request).spans).toHaveLength(2);
    });
  it('does not extend joining into another grapheme', () => {
    expect(fixture([[0, 0xffff]]).shape({ ...request, text: 'T\u0301é' }).spans).toHaveLength(2);
  });
  it('does not collide with the previous split cache in either order or measure mode', () => {
    for (const order of [false, true]) {
      const service = fixture();
      for (const join of [order, !order, order]) {
        expect(service.shape({ ...request, measure: false, joinRegisteredGrapheme: join }).spans)
          .toHaveLength(join ? 1 : 2);
        expect(service.shape({ ...request, measure: true, joinRegisteredGrapheme: join }).spans)
          .toHaveLength(join ? 1 : 2);
      }
    }
  });
  it('clears admission for partial slices, independent probes and transformed text', () => {
    expect(sliceTextShapeRequest(request, 0, 2).joinRegisteredGrapheme).toBe(true);
    expect(sliceTextShapeRequest(request, 0, 1).joinRegisteredGrapheme).not.toBe(true);
    expect(independentTextShapeRequest(request, request.text).joinRegisteredGrapheme).not.toBe(true);
    expect(replaceTextShapeRequest(request, request.text).joinRegisteredGrapheme).not.toBe(true);
  });
});

const documentModel = { section: { pageWidth: 612, pageHeight: 792, marginTop: 72,
  marginRight: 72, marginBottom: 72, marginLeft: 72, headerDistance: 36, footerDistance: 36 },
  body: [], headers: { default: null, first: null, even: null },
  footers: { default: null, first: null, even: null } } as unknown as DocxDocumentModel;
const context = { font: '', fontKerning: 'none', letterSpacing: '0px', measureText: () => ({
  width: 99, actualBoundingBoxAscent: 9, actualBoundingBoxDescent: 0,
  fontBoundingBoxAscent: 9, fontBoundingBoxDescent: 0,
}) } as unknown as CanvasRenderingContext2D;
function run(text = 'T\u0301', extra = {}): DocRun {
  return { type: 'text', text, fontFamily: 'Test Face', fontFamilyHighAnsi: 'Test Face',
    fontSize: 10, bold: false, italic: false, underline: false, strikethrough: false,
    color: null, isLink: false, background: null, vertAlign: null, hyperlink: null,
    ...extra } as DocRun;
}
function environment(extra: Partial<LineLayoutEnvironment> = {}, doc = documentModel): LineLayoutEnvironment {
  const base = createLayoutServices(doc, { measureContext: context });
  return { pageIndex: 0, totalPages: 1, characterGridActive: false, paragraphRtl: false,
    layoutServices: Object.freeze({ ...base, text: fixture() }), ...extra };
}
describe('registered grapheme allocation admission', () => {
  it('builds and measures one segment with original slot facts', () => {
    const segs = buildSegments([run()], environment());
    expect(segs).toHaveLength(1);
    expect(segs[0]).toMatchObject({ text: 'T\u0301', textShapeRequest: { joinRegisteredGrapheme: true } });
    const lines = layoutLines(context, segs, 100, 0, 1);
    expect(lines[0].segments).toHaveLength(1);
    expect(lines[0].segments[0].measuredWidth).toBe(6);
  });
  it.each([
    [{ paragraphRtl: true }, {}], [{ paragraphRtl: undefined }, {}],
    [{ characterGridActive: true }, {}], [{ characterGridActive: undefined }, {}],
    [{ verticalCJK: true }, {}], [{}, { charSpacing: 1 }], [{}, { charScale: 0.8 }],
    [{}, { fitTextVal: 20, fitTextId: 1 }], [{}, { smallCaps: true }],
    [{}, { allCaps: true }], [{}, { rtl: true }], [{}, { vertAlign: 'super' }],
    [{}, { ruby: { text: 'ruby', fontSizePt: 5 } }],
  ])('preserves slot separation with other allocation/atomic policies %j %j', (env, format) => {
    const segs = buildSegments([run('T\u0301', format)], environment(env));
    expect(segs.filter(s => 'text' in s).every(s => !('semanticSlotSpans' in s))).toBe(true);
  });
  it('retains authored paint boundaries even when they bisect a grapheme', () => {
    const segs = buildSegments([run('T'), run('\u0301', { color: '#ff0000' })], environment());
    expect(segs).toHaveLength(2);
    expect(segs.every(s => !('semanticSlotSpans' in s))).toBe(true);
  });
  it('joins a formatting-only source seam without losing source ownership', () => {
    const segs = buildSegments([run('T'), run('\u0301')], environment());
    expect(segs).toHaveLength(1);
    expect(segs[0].sourceTextSequence).toHaveLength(2);
  });
});

it('retains one paint operation and shared cluster geometry across source owners', () => {
  const doc = { ...documentModel, body: [{ type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0,
    lineSpacing: null, numbering: null, tabStops: [], runs: [run('T'), run('\u0301')],
  }] } as DocxDocumentModel;
  const layout = layoutDocument(doc, environment({}, doc).layoutServices!, { currentDateMs: 0 });
  const seen = new WeakSet<object>();
  const placements: TextPlacement[] = [];
  function walk(value: unknown): void {
    if (!value || typeof value !== 'object' || seen.has(value)) return;
    seen.add(value);
    if ('kind' in value && value.kind === 'text' && 'paintOps' in value) {
      placements.push(value as TextPlacement); return;
    }
    for (const child of Object.values(value)) walk(child);
  }
  walk(layout.pages);
  const placement = placements.find(p => p.text === 'T\u0301')!;
  expect(placement).toBeDefined();
  expect(placement.paintOps).toHaveLength(1);
  expect(placement.paintOps[0].text).toBe('T\u0301');
  expect(placement.clusters).toHaveLength(1);
  expect(placement.sourceRuns).toHaveLength(2);
  const owners = sourceOwnedTextPlacements(placement);
  expect(owners.map(owner => owner.text)).toEqual(['T', '\u0301']);
  expect(owners.map(owner => owner.semanticSlotSpans?.map(slot => [slot.start, slot.end, slot.script])))
    .toEqual([[[0, 1, 'ascii']], [[0, 1, 'highAnsi']]]);
  expect(owners[0].bounds).toEqual(owners[1].bounds);
  expect(structuredClone(placement)).toEqual(placement);
});
it('retains justified separator context without splitting the compound grapheme', () => {
  // Synthetic shaper: letters 6pt, an attached mark 0pt, a detached (string-
  // initial) mark 2pt, and kerning removes 1pt only from the " T" pair.
  const paintedOrigins = (text: string, markCovered: boolean) => {
    const service = createTextLayoutService({
      fonts: createFontResolver([{ requestedFamily: 'Test Face', resolvedFamily: 'Test Resource',
        resourceIdentity: identity, source: 'embedded', weight: 400, style: 'normal' }]),
      fontMetrics: { resource: { family: 'Test Resource', sourceIdentity: identity, weight: 400,
        style: 'normal', unicodeRanges: markCovered ? [[0x20, 0x7a], [0x301, 0x301]] : [[0x20, 0x7a]] } },
      measurer: { fingerprint: 'detached-mark-spacing', measure: ({ text, kerning }) => ({
        advancePt: [...text].reduce((sum, ch, index) =>
          sum + (ch === '́' ? (index === 0 ? 2 : 0) : 6), 0)
          - (kerning && text.includes(' T') ? 1 : 0),
        ascentPt: 9, descentPt: 0,
      }) },
    });
    // Mode-15 `both` is the only production consumer of the boundary repair.
    const doc = { ...documentModel, settings: { compatibilityMode: 15 }, body: [{ type: 'paragraph',
      alignment: 'both', indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0,
      spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
      runs: [run(text, { kerning: 1 })] }] } as unknown as DocxDocumentModel;
    const services = Object.freeze({ ...createLayoutServices(doc, { measureContext: context }), text: service });
    const placements: TextPlacement[] = [];
    const seen = new WeakSet<object>();
    (function walk(value: unknown): void {
      if (!value || typeof value !== 'object' || seen.has(value)) return;
      seen.add(value);
      if ('kind' in value && value.kind === 'text' && 'paintOps' in value) {
        placements.push(value as TextPlacement); return;
      }
      for (const child of Object.values(value)) walk(child);
    })(layoutDocument(doc, services, { currentDateMs: 0 }).pages);
    const x = (p: TextPlacement) => p.origin.xPt + p.paintOps[0].offset.xPt;
    return placements.map(p => [p.text, x(p) - x(placements[0]), Boolean(p.semanticSlotSpans)]);
  };
  // The compound remains one paint shape. The bounded registered Latin
  // probe revalidates the whole pair as one face, preserving the -1pt context
  // without introducing the detached mark's 2pt advance.
  expect(paintedOrigins('a T́', true)).toEqual([['a ', 0, false], ['T́', 11, true]]);
  // Ordinary pairs keep the native repair, including when the proof is
  // withheld and the same text keeps its per-slot spans.
  expect(paintedOrigins('a T', true)).toEqual([['a ', 0, false], ['T', 11, false]]);
  expect(paintedOrigins('a T́', false)).toEqual(
    [['a ', 0, false], ['T', 11, false], ['́', 17, false]]);
});
it('rejects native faces even when the resolver memoizes the same object', () => {
  const service = createTextLayoutService({ fonts: createFontResolver([]),
    measurer: { fingerprint: 'native', measure: () => ({ advancePt: 1, ascentPt: 1, descentPt: 0 }) } });
  expect(service.shape(request).spans).toHaveLength(2);
});
it('rejects distinct registered resources even when their concrete Canvas route is identical', () => {
  const service = createTextLayoutService({
    fonts: createFontResolver(['A', 'B'].map(name => ({ requestedFamily: name,
      resolvedFamily: 'Shared Route', source: 'embedded', resourceIdentity: `embedded:${name}`,
      weight: 400, style: 'normal' }))),
    fontMetrics: Object.fromEntries(['A', 'B'].map(name => [name, { family: 'Shared Route',
      sourceIdentity: `embedded:${name}`, weight: 400, style: 'normal', unicodeRanges: [[0, 0xffff]],
    }])),
    measurer: { fingerprint: 'same-route', measure: () => ({ advancePt: 6, ascentPt: 9, descentPt: 0 }) },
  });
  const result = service.shape({ ...request, fonts: { ascii: 'A', highAnsi: 'B' } });
  expect(result.spans[0].fontRoute).toEqual(result.spans[1].fontRoute);
  expect(result.spans.map(span => span.font.resourceIdentity)).toEqual(['embedded:A', 'embedded:B']);
  expect(result.spans).toHaveLength(2);
});

it.each([{ sourceIdentity: undefined }, { sourceIdentity: 'embedded:other-resource' },
  { weight: 700 }, { style: 'italic' as const }, { family: 'Other Resource' }])(
  'rejects a metric without the exact selected resource tuple %j', metric => {
    expect(fixture(undefined, metric).shape(request).spans).toHaveLength(2);
  });
it('requires every resource sharing the selected Canvas face to cover all or none of the grapheme', () => {
  const spans = (peer: Partial<ResolvedFontMetric>) => createTextLayoutService({
    fonts: createFontResolver([{ requestedFamily: 'Alias A', resolvedFamily: 'Shared Face',
      source: 'local', resourceIdentity: 'provided-sfnt:a', weight: 400, style: 'normal' }]),
    fontMetrics: {
      'alias a': { family: 'Shared Face', requestedFamily: 'Alias A', sourceIdentity: 'provided-sfnt:a',
        weight: 400, style: 'normal', unicodeRanges: [[0x54, 0x54], [0x301, 0x301]] },
      // CSS family names match case-insensitively; weight/style default to 400/normal.
      'alias b': { family: 'shared face', requestedFamily: 'Alias B', sourceIdentity: 'provided-sfnt:b', ...peer },
    },
    measurer: { fingerprint: 'shared-face', measure: () => ({ advancePt: 6, ascentPt: 9, descentPt: 0 }) },
  }).shape({ ...request, fonts: { ascii: 'Alias A', highAnsi: 'Alias A' } }).spans.length;
  // Canvas may select either face, painting the base and mark from different resources.
  expect(spans({ unicodeRanges: [[0x54, 0x54]] })).toBe(2);
  expect(spans({})).toBe(2);
  expect(spans({ unicodeRanges: [[0x54, 0x54], [0x301, 0x301]] })).toBe(1);
  expect(spans({ unicodeRanges: [[0x41, 0x41]] })).toBe(1);
});
it('requires nonmissing resource identity on registered faces', () => {
  expect(fixture(undefined, {}, { resourceIdentity: undefined }).shape(request).spans).toHaveLength(2);
});
it.each(['css', 'google', 'substitute'] as const)('does not admit unproven registration source %s', source => {
  expect(fixture(undefined, {}, { source }).shape(request).spans).toHaveLength(2);
});

it('keeps character-grid allocation even when its line-grid pitch is disabled', () => {
  const paragraph = { type: 'paragraph' as const, alignment: 'left', indentLeft: 0,
    indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0,
    lineSpacing: null, numbering: null, tabStops: [], runs: [run()],
  } as import('../types.js').DocParagraph;
  const doc = { ...documentModel, section: { ...documentModel.section,
    docGridType: 'linesAndChars', docGridCharSpace: 4096, docGridLinePitch: 0,
  }, body: [paragraph] } as DocxDocumentModel;
  const settings = resolveDocumentLayoutSettings(doc);
  const paragraphContext = resolveParagraphLayoutContext(settings,
    resolveSectionLayoutContext(settings, doc.section), { story: 'body', containers: [], lineNumberingEligible: true }, paragraph);
  expect(paragraphContext.lineGrid.active).toBe(false);
  expect(paragraphContext.characterGrid.active).toBe(true);
  const calls: TextShapeRequest[] = [];
  const text = fixture();
  const services = createLayoutServices(doc, { measureContext: context });
  const recorded = Object.freeze({ ...services, text: Object.freeze({ ...text, shape(request: TextShapeRequest) {
    calls.push(request); return text.shape(request);
  } }) });
  layoutDocument(doc, recorded, { currentDateMs: 0 });
  expect(calls.some(call => call.joinRegisteredGrapheme === true)).toBe(false);
  calls.length = 0;
  measureParagraphIntrinsicWidths(paragraph, paragraphContext, 100,
    { context, fontFamilyClasses: {} }, { pageIndex: 0, totalPages: 1,
      pageWritingMode: 'horizontal-tb', documentHasEastAsianText: false, layoutServices: recorded });
  expect(calls.some(call => call.joinRegisteredGrapheme === true)).toBe(false);
  calls.length = 0;
  measureParagraphIntrinsicWidths(paragraph, { ...paragraphContext,
    characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 } }, 100,
    { context, fontFamilyClasses: {} }, { pageIndex: 0, totalPages: 1,
      pageWritingMode: 'horizontal-tb', documentHasEastAsianText: false, layoutServices: recorded });
  expect(calls.some(call => call.joinRegisteredGrapheme === true)).toBe(true);
});
it.each(['א\u0301', 'س\u0301', '.\u0301', '1\u0301', 'é\u0301', 'T\u200d\u0301'])(
  'retains existing shaping outside the proven strong Latin base class %s', text => {
    expect(fixture([[0, 0xffff]]).shape({ ...request, text }).spans
      .every(span => !span.semanticSlotSpans)).toBe(true);
  });
