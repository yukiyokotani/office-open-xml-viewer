import { expect, it } from 'vitest';
import { createFontResolver } from './font-service.js';
import { createTextLayoutService, independentTextShapeRequest, replaceTextShapeRequest,
  sliceTextShapeRequest, sliceSemanticSlotSpans } from './text.js';
import { buildSegments, layoutLines, type LineLayoutEnvironment, type LayoutTextSeg } from '../line-layout.js';
import { LineMeasurementAdapter } from '../line-breaker/measurement-adapter.js';
import { slicedTextMetadata } from '../line-breaker/advance.js';
import type { DocRun, DocxDocumentModel, DocParagraph } from '../types.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import { sourceOwnedTextPlacements } from './text-source-ownership.js';
import type { TextPlacement } from './types.js';
import { measureParagraphIntrinsicWidths } from './intrinsic-width.js';
import { resolveDocumentLayoutSettings, resolveParagraphLayoutContext, resolveSectionLayoutContext } from '../layout-context.js';
import { textRunGeometryForPage } from './text-index.js';
import { textRunsForPage } from '../text-run-projection.js';
import { buildTextIndex, findMatches } from '@silurus/ooxml-core';

// Independent glyph-provider response: six-point ordinary glyphs, attached
// marks zero, detached marks two; one kern pair and contextual ligatures.
const nativeAdvance = (text: string, kerning = true): number =>
  [...text].reduce((n, c, i) => n + (c === '\u0301' ? (i === 0 ? 2 : 0) : 6), 0)
  - (kerning ? (text.match(/Té/g)?.length ?? 0) : 0)
  - (kerning ? (text.match(/ [Té]/g)?.length ?? 0) : 0)
  - (text.includes('ffí') ? 7 : text.includes('fí') ? 3 : 0);
const identity = 'embedded:public-latin-context';
function fixture(peer?: readonly (readonly [number, number])[] | null, onMeasure?: (text: string) => void) {
  return createTextLayoutService({
    fonts: createFontResolver([{ requestedFamily: 'Test Face', resolvedFamily: 'Test Resource',
      resourceIdentity: identity, source: 'embedded', weight: 400, style: 'normal' }]),
    fontMetrics: { one: { family: 'Test Resource', sourceIdentity: identity, weight: 400,
      style: 'normal', unicodeRanges: [[0x20, 0x7a], [0xe9, 0xef], [0x301, 0x301]] },
      ...(peer === undefined ? {} : { peer: { family: 'Test Resource', sourceIdentity: 'peer',
        weight: 400, style: 'normal' as const, ...(peer ? { unicodeRanges: peer } : {}) } }) },
    measurer: { fingerprint: 'independent-context-response', measure: ({ text, kerning }) => {
      onMeasure?.(text);
      return { advancePt: nativeAdvance(text, kerning), ascentPt: 9, descentPt: 0 };
    } },
  });
}
const run = (text: string, extra = {}): DocRun => ({ type: 'text', text, fontFamily: 'Test Face',
  fontFamilyHighAnsi: 'Test Face', fontSize: 10, bold: false, italic: false, underline: false,
  strikethrough: false, color: null, isLink: false, background: null, vertAlign: null,
  hyperlink: null, kerning: 1, ...extra } as DocRun);
const environment = (service = fixture(), extra = {}): LineLayoutEnvironment => ({
  pageIndex: 0, totalPages: 1, characterGridActive: false, paragraphRtl: false,
  layoutServices: { text: service } as LineLayoutEnvironment['layoutServices'], ...extra,
});
const context = { font: '', fontKerning: 'none', letterSpacing: '0px',
  measureText: (text: string) => ({ width: nativeAdvance(text, false),
    actualBoundingBoxAscent: 9, actualBoundingBoxDescent: 0,
    fontBoundingBoxAscent: 9, fontBoundingBoxDescent: 0 }),
} as unknown as CanvasRenderingContext2D;
const segments = (text: string, env = environment(), extra = {}): LayoutTextSeg[] =>
  buildSegments([run(text, extra)], env).filter((s): s is LayoutTextSeg => 'text' in s);

it.each(['Té', 'fí', 'ffí'])('acquires %s as one physical shape, preserving semantic slots', text => {
  const s = segments(text);
  expect(s).toHaveLength(1);
  const adapter = new LineMeasurementAdapter(context, 1, () => '10px Test Resource');
  expect(adapter.measureSegment(s[0]).width).toBe(nativeAdvance(text));
  expect(s[0].semanticSlotSpans?.map(v => [v.start, v.end, v.script])).toEqual([
    [0, text.length - 1, 'ascii'], [text.length - 1, text.length, 'highAnsi'],
  ]);
});
it('retains compound context at the ordinary word separator', () => {
  const s = segments('a T\u0301');
  const adapter = new LineMeasurementAdapter(context, 1, () => '10px Test Resource');
  const width = s.reduce((n, v, i) => n + adapter.measureSegment(v).width
    + adapter.wordBoundaryAdvance(s[i - 1], v), 0);
  expect(width).toBe(nativeAdvance('a T\u0301'));
});
it.each([null, 20])('keeps the established slot split when kerning threshold is %s', kerning => {
  expect(segments('Té', environment(), { kerning })).toHaveLength(2);
});
it.each([null, [[0x54, 0x54]] as const])('declines an unknown or partial peer %j', peer => {
  expect(segments('Té', environment(fixture(peer)))).toHaveLength(2);
});
it.each([{ charSpacing: 1 }, { charScale: 0.8 }, { smallCaps: true },
  { rtl: true }, { vertAlign: 'super' }, { fontFamilyHighAnsi: 'Different Face' }])(
  'keeps real allocation or font boundaries %j', extra => {
    const s = segments('Té', environment(), extra);
    expect(s.length).toBeGreaterThan(0);
    expect(s.every(v => v.semanticSlotSpans === undefined
      && v.textShapeRequest?.joinRegisteredLatinSlots !== true)).toBe(true);
  });
it('keeps an authored paint seam while absorbing an identical-format source seam', () => {
  const acquire = (right: DocRun) => buildSegments([run('T'), right], environment())
    .filter((s): s is LayoutTextSeg => 'text' in s);
  expect(acquire(run('é', { color: '#ff0000' }))).toHaveLength(2);
  const joined = acquire(run('é'));
  expect(joined).toHaveLength(1);
  expect(joined[0].sourceTextSequence?.map(v => [v.start, v.end])).toEqual([[0, 1], [1, 2]]);
});
it('reproves retained fragments and clears independent or transformed admission', () => {
  const service = fixture();
  const s = segments('TTéé', environment(service))[0];
  expect(s.text).toBe('TTéé');
  const metadata = slicedTextMetadata(s, 1, 3);
  expect(sliceSemanticSlotSpans(metadata.semanticSlotSpans, metadata.semanticSlotRange!.start,
    metadata.semanticSlotRange!.end)?.map(v => [v.start, v.end, v.script])).toEqual([
    [0, 1, 'ascii'], [1, 2, 'highAnsi'],
  ]);
  const request = s.textShapeRequest!;
  expect(service.shape({ ...sliceTextShapeRequest(request, 1, 3), measure: true }).advancePt).toBe(11);
  expect(service.shape(independentTextShapeRequest(request, 'Té')).spans).toHaveLength(2);
  expect(service.shape(replaceTextShapeRequest(request, 'Té')).spans).toHaveLength(2);
});
it('keeps admitted and ordinary shapes under distinct cache provenance', () => {
  // Build the request separately: segment construction may itself populate
  // the service cache, so each ordering needs its own cold service below.
  const admitted = segments('Té', environment(fixture()))[0].textShapeRequest!;
  expect(admitted.joinRegisteredLatinSlots).toBe(true);
  // All request facts stay identical except the admission proof flag, so
  // removing that flag from the cache key would alias these two shapes.
  const ordinary = { ...admitted, joinRegisteredLatinSlots: undefined };
  const ordinaryFirst = fixture();
  expect(ordinaryFirst.shape(ordinary).spans).toHaveLength(2);
  expect(ordinaryFirst.shape(admitted).spans).toHaveLength(1);
  const admittedFirst = fixture();
  expect(admittedFirst.shape(admitted).spans).toHaveLength(1);
  expect(admittedFirst.shape(ordinary).spans).toHaveLength(2);
});

function document(runs: DocRun[], alignment: DocParagraph['alignment'] = 'left'): DocxDocumentModel {
  return { section: { pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72,
    marginBottom: 72, marginLeft: 72, headerDistance: 36, footerDistance: 36 },
    settings: { compatibilityMode: 15 }, body: [{ type: 'paragraph', alignment,
      indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0,
      lineSpacing: null, numbering: null, tabStops: [], runs }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}
function textPlacements(doc: DocxDocumentModel): TextPlacement[] {
  const services = Object.freeze({ ...createLayoutServices(doc, { measureContext: context }), text: fixture() });
  const result: TextPlacement[] = [];
  const seen = new WeakSet<object>();
  function walk(value: unknown): void {
    if (!value || typeof value !== 'object' || seen.has(value)) return;
    seen.add(value);
    if ('kind' in value && value.kind === 'text' && 'paintOps' in value) {
      result.push(value as TextPlacement); return;
    }
    for (const child of Object.values(value)) walk(child);
  }
  walk(layoutDocument(doc, services, { currentDateMs: 0 }).pages);
  return result;
}
it('keeps registered contextual Latin shaping across an unselected same-format optional owner', () => {
  const doc = document([run('ff'), run('', { __optionalHyphen: true }), run('í')]);
  const services = Object.freeze({ ...createLayoutServices(doc, { measureContext: context }), text: fixture() });
  const layout = layoutDocument(doc, services, { currentDateMs: 0 });
  const paragraph = layout.pages[0]!.layers.body.find(node => node.kind === 'paragraph');
  if (!paragraph || paragraph.kind !== 'paragraph') throw new Error('Expected acquired paragraph');
  const placements = paragraph.lines.flatMap(line => line.placements)
    .filter((p): p is TextPlacement => p.kind === 'text');
  // The independent provider substitutes ffí as one physical unit (18−7).
  // An empty authored source owner must not split that contextual operation.
  expect(placements).toHaveLength(1);
  expect(placements[0].paintOps.map(op => op.text)).toEqual(['ffí']);
  expect(placements[0].advancePt).toBe(11);
  expect(placements[0].semanticSlotSpans?.map(s => [s.start, s.end, s.script]))
    .toEqual([[0, 2, 'ascii'], [2, 3, 'highAnsi']]);
  const owned = textRunGeometryForPage(layout, 0).map(g => g.placement);
  expect(owned.map(p => [p.sourceRunIndex, p.text, p.range]))
    .toEqual([[0, 'ff', { start: 0, end: 2 }], [1, '', { start: 2, end: 2 }],
      [2, 'í', { start: 2, end: 3 }]]);
  expect(owned[1].advancePt).toBe(0);
  const index = buildTextIndex(textRunsForPage(layout, 0, { scale: 1 }));
  expect(index.text).toBe('ffí');
  expect(findMatches(index, 'ffí')).toHaveLength(1);
});
it('retains registered slot windows and zero logical length when an optional glyph is selected', () => {
  const doc = document([run('TTé'), run('', { __optionalHyphen: true }), run('éé')]);
  doc.section.marginRight = 612 - 72 - 23;
  const services = Object.freeze({ ...createLayoutServices(doc, { measureContext: context }), text: fixture() });
  const layout = layoutDocument(doc, services, { currentDateMs: 0 });
  const paragraph = layout.pages[0]!.layers.body.find(node => node.kind === 'paragraph');
  if (!paragraph || paragraph.kind !== 'paragraph') throw new Error('Expected acquired paragraph');
  const owned = textRunGeometryForPage(layout, 0).map(g => g.placement);
  // Whole word is 29 points; contextual TTé is 17 and the owned glyph is 6.
  expect(paragraph.lines).toHaveLength(2);
  expect(owned.map(p => p.text)).toEqual(['TTé', '-', 'éé']);
  expect(owned.map(p => p.advancePt)).toEqual([17, 6, 12]);
  expect(owned[0].semanticSlotSpans?.map(s => [s.start, s.end, s.script]))
    .toEqual([[0, 2, 'ascii'], [2, 3, 'highAnsi']]);
  expect(owned[1]).toMatchObject({ optionalHyphenGlyph: true, sourceRunIndex: 1,
    range: { start: 3, end: 3 }, clusters: [{ range: { start: 3, end: 3 }, advancePt: 6 }] });
  expect(owned[2]).toMatchObject({ sourceRunIndex: 2, range: { start: 3, end: 5 } });
  expect(owned[2].semanticSlotSpans?.map(s => [s.start, s.end, s.script]))
    .toEqual([[0, 2, 'highAnsi']]);
  expect(paragraph.lines.map(line => line.range)).toEqual([{ start: 0, end: 3 }, { start: 3, end: 5 }]);
  expect(owned[2].clusters.map(c => c.range)).toEqual([{ start: 3, end: 4 }, { start: 4, end: 5 }]);
  const index = buildTextIndex(textRunsForPage(layout, 0, { scale: 1 }));
  expect(index.text).toBe('TTééé');
  expect(findMatches(index, 'TTééé')).toHaveLength(1);
  expect(findMatches(index, 'TTé-éé')).toHaveLength(0);
});
it('materializes retained slot windows on sliced placements', () => {
  const text = 'é' + 'T'.repeat(200);
  const placements = textPlacements(document([run(text)]));
  expect(placements.length).toBeGreaterThan(1);
  expect(placements.map(p => p.text).join('')).toBe(text);
  for (const p of placements) {
    const slots = p.semanticSlotSpans!;
    expect(slots[0].start).toBe(0);
    expect(slots.at(-1)!.end).toBe(p.text.length);
    expect(slots.every((v, i) => i === 0 || slots[i - 1].end === v.start)).toBe(true);
    expect(sourceOwnedTextPlacements(p).map(o => o.text).join('')).toBe(p.text);
  }
  expect(placements[0].semanticSlotSpans!.map(v => v.script)).toEqual(['highAnsi', 'ascii']);
  expect(placements.slice(1).every(p => p.semanticSlotSpans!.length === 1
    && p.semanticSlotSpans![0].script === 'ascii')).toBe(true);
});
it('paints a contextual substitution as one string and projects original owners separately', () => {
  const placements = textPlacements(document([run('ff'), run('í')]));
  expect(placements).toHaveLength(1);
  expect(structuredClone(placements)).toEqual(placements);
  expect(placements[0].paintOps.map(op => op.text)).toEqual(['ffí']);
  expect(placements[0].bounds.widthPt).toBe(nativeAdvance('ffí'));
  const owners = sourceOwnedTextPlacements(placements[0]);
  expect(owners.map(p => p.text)).toEqual(['ff', 'í']);
  expect(owners.map(p => p.semanticSlotSpans?.map(v => [v.start, v.end, v.script])))
    .toEqual([[[0, 2, 'ascii']], [[0, 1, 'highAnsi']]]);
});
it('breaks long joined text with shared slot metadata and remeasures each fragment', () => {
  const s = segments('T'.repeat(24) + 'é');
  const lines = layoutLines(context, s, 12, 0, 1);
  const fragments = lines.flatMap(line => line.segments)
    .filter((v): v is LayoutTextSeg => 'text' in v);
  expect(fragments.map(v => v.text).join('')).toBe('T'.repeat(24) + 'é');
  expect(fragments.every(v => v.measuredWidth <= 12)).toBe(true);
  for (const f of fragments) {
    expect(f.measuredWidth).toBe(nativeAdvance(f.text));
    const slots = sliceSemanticSlotSpans(f.semanticSlotSpans,
      f.semanticSlotRange?.start ?? 0, f.semanticSlotRange?.end ?? f.text.length)!;
    expect(slots[0].start).toBe(0);
    expect(slots.at(-1)!.end).toBe(f.text.length);
    expect(slots.every((v, i) => i === 0 || slots[i - 1].end === v.start)).toBe(true);
  }
});
it('preserves known all-or-none peers and keeps paragraph grid admission closed', () => {
  for (const peer of [[], [[0x20, 0x7a], [0xe9, 0xef], [0x301, 0x301]]] as const)
    expect(segments('Té', environment(fixture(peer)))).toHaveLength(1);
  for (const state of [true, undefined])
    expect(segments('Té', environment(fixture(), { characterGridActive: state })))
      .toHaveLength(2);
  expect(segments('Té', environment(fixture(), { paragraphRtl: true }))).toHaveLength(2);
});
it('acquires native max-content context while retaining its separate word units', () => {
  const doc = document([run('a Té')], 'both');
  const paragraph = doc.body[0] as DocParagraph;
  const settings = resolveDocumentLayoutSettings(doc);
  const paragraphContext = resolveParagraphLayoutContext(settings,
    resolveSectionLayoutContext(settings, doc.section),
    { story: 'body', containers: [], lineNumberingEligible: true }, paragraph);
  const services = Object.freeze({ ...createLayoutServices(doc, { measureContext: context }), text: fixture() });
  const intrinsic = measureParagraphIntrinsicWidths(paragraph, paragraphContext, 100,
    { context, fontFamilyClasses: {} }, { pageIndex: 0, totalPages: 1,
      pageWritingMode: 'horizontal-tb', documentHasEastAsianText: false,
      compatibilityMode: 15, layoutServices: services });
  expect(intrinsic.maxWidthPt).toBe(nativeAdvance('a Té'));
  const placements = textPlacements(doc);
  expect(placements.map(p => p.paintOps.map(op => op.text))).toEqual([['a '], ['Té']]);
});
it('commits separator context on the intrinsic pass without a soft wrap', () => {
  const lines = layoutLines(context, segments('a Té'), 1, 0, 1, [], undefined,
    {}, 0, undefined, undefined, undefined, undefined, false, false, false,
    undefined, 'intrinsic');
  expect(lines).toHaveLength(1);
  const text = lines[0].segments.filter((s): s is LayoutTextSeg => 'text' in s);
  expect(text.map(s => s.text)).toEqual(['a ', 'Té']);
  expect(text.reduce((n, s) => n + s.measuredWidth, 0)).toBe(nativeAdvance('a Té'));
});

it('keeps the internal slot-window projection bounded for resolved alternating metadata', () => {
  const text = 'Té'.repeat(1024);
  const service = fixture();
  const base = segments('Té', environment(service))[0];
  const s = { ...base, text, textShapeRequest: { ...base.textShapeRequest!, text,
    joinRegisteredLatinSlots: undefined, substituteContext: { text, offset: 0 } },
    semanticSlotSpans: service.shape({ text, fontSizePt: 10,
      fonts: { ascii: 'Test Face', highAnsi: 'Test Face' }, kerning: true,
      measure: false, clusterGeometry: false }).spans };
  // These are real resolved slot facts, not a claim that a multi-seam word
  // is admitted as one physical shape. Exercise the internal window alone.
  let reads = 0;
  const spans = new Proxy(s.semanticSlotSpans!, { get(target, key, receiver) {
    if (typeof key === 'string' && /^\d+$/.test(key)) reads += 1;
    return Reflect.get(target, key, receiver);
  } });
  const source = { ...s, semanticSlotSpans: spans };
  for (let offset = 0; offset < source.text.length; offset++) {
    const suffix = slicedTextMetadata(source, offset, source.text.length);
    expect(suffix.semanticSlotRange).toEqual({ start: offset, end: source.text.length });
  }
  // Binary positioning plus O(1) window creation, rather than repeatedly
  // scanning/copying every remaining slot of each emergency suffix.
  expect(reads).toBeLessThanOrEqual(source.text.length * (Math.ceil(Math.log2(spans.length)) + 2));
});

it('keeps multi-seam words on their existing slot path without repeated full-tail acquisition', () => {
  for (const n of [32, 64, 128, 256]) {
    let measuredUnits = 0;
    const text = 'Té'.repeat(n);
    const service = fixture(undefined, t => { measuredUnits += t.length; });
    const s = segments(text, environment(service));
    expect(s).toHaveLength(2 * n);
    expect(s.every(v => v.semanticSlotSpans === undefined)).toBe(true);
    const lines = layoutLines(context, s, 12, 0, 1);
    expect(lines.flatMap(l => l.segments).filter((v): v is LayoutTextSeg => 'text' in v)
      .map(v => v.text).join('')).toBe(text);
    // Exactly the two distinct one-scalar shapes are measured and cached.
    // This retains the baseline placement/measurement path, rather than
    // repeatedly native-shaping each remaining joined word tail.
    expect(lines).toHaveLength(1);
    expect(measuredUnits).toBe(2);
  }
});
it('revalidates the single-seam limit on a cross-word context probe', () => {
  const s = segments('a éT');
  expect(s.at(-1)!.semanticSlotSpans).toBeDefined();
  const adapter = new LineMeasurementAdapter(context, 1, () => '10px Test Resource');
  // ascii left + highAnsi/ascii right has two seams. The individual right
  // token remains a valid unit; separator context is independently declined.
  expect(adapter.wordBoundaryAdvance(s[0], s[1])).toBe(0);
});

it.each(['Té ', 'fí ', 'ffí ', 'Tó  '])('retains the one-word seam in %j with its authored trailing spaces', text => {
  // Tó uses the same ordinary slot/cmap proof as Té.
  const service = createTextLayoutService({
    fonts: createFontResolver([{ requestedFamily: 'Test Face', resolvedFamily: 'Test Resource',
      resourceIdentity: identity, source: 'embedded', weight: 400, style: 'normal' }]),
    fontMetrics: { one: { family: 'Test Resource', sourceIdentity: identity, weight: 400,
      style: 'normal', unicodeRanges: [[0x20, 0x7a], [0xe9, 0xf3], [0x301, 0x301]] } },
    measurer: { fingerprint: 'independent-context-with-trailing-space', measure: ({ text, kerning }) => ({
      advancePt: nativeAdvance(text, kerning), ascentPt: 9, descentPt: 0,
    }) },
  });
  const s = segments(text, environment(service));
  expect(s).toHaveLength(1);
  expect(s[0].semanticSlotSpans?.map(v => v.script)).toEqual(['ascii', 'highAnsi', 'ascii']);
  const adapter = new LineMeasurementAdapter(context, 1, () => '10px Test Resource');
  expect(adapter.measureSegment(s[0]).width).toBe(nativeAdvance(text));
});
it('keeps a contextual substitution and its trailing separator in one retained paint unit', () => {
  const placements = textPlacements(document([run('ff'), run('í ')]));
  expect(placements).toHaveLength(1);
  expect(structuredClone(placements)).toEqual(placements);
  expect(placements[0].paintOps.map(op => op.text)).toEqual(['ffí ']);
  expect(placements[0].bounds.widthPt).toBe(nativeAdvance('ffí '));
  const owners = sourceOwnedTextPlacements(placements[0]);
  expect(owners.map(p => p.text)).toEqual(['ff', 'í ']);
  expect(owners.map(p => p.semanticSlotSpans?.map(v => [v.start, v.end, v.script])))
    .toEqual([[[0, 2, 'ascii']], [[0, 1, 'highAnsi'], [1, 2, 'ascii']]]);
});
it('keeps alternating multi-seam words with trailing spaces on their baseline path', () => {
  for (const n of [32, 64, 128, 256]) {
    let measuredUnits = 0;
    const text = 'Té'.repeat(n) + ' ';
    const service = fixture(undefined, t => { measuredUnits += t.length; });
    const s = segments(text, environment(service));
    expect(s).toHaveLength(2 * n + 1);
    expect(s.every(v => v.semanticSlotSpans === undefined)).toBe(true);
    const lines = layoutLines(context, s, 12, 0, 1);
    expect(lines.flatMap(l => l.segments).filter((v): v is LayoutTextSeg => 'text' in v)
      .map(v => v.text).join('')).toBe(text);
    expect(measuredUnits).toBe(3);
  }
});

it('retains full-unit cmap proof when the optional tail is admitted', () => {
  // This peer covers the entire word body but not its authored SPACE.
  // Full-unit all-or-none proof must still decline the physical join.
  const bodyOnly = [[0x54, 0x54], [0xe9, 0xe9]] as const;
  expect(segments('Té ', environment(fixture(bodyOnly)))).toHaveLength(3);
  expect(segments('Té ', environment(fixture(null)))).toHaveLength(3);
  expect(segments('Té ', environment(fixture([])))).toHaveLength(1);
});
it('keeps ordinary separator context after a joined emergency suffix becomes one live slot', () => {
  const s = segments('éTTTTT T');
  const sourceLeft = s.find(v => v.text.endsWith(' '))!;
  const right = s.at(-1)!;
  const start = sourceLeft.text.length - 3;
  const left = { ...sourceLeft, text: sourceLeft.text.slice(start),
    ...slicedTextMetadata(sourceLeft, start, sourceLeft.text.length) };
  expect(left.text).toBe('TT ');
  expect(left.script).toBe('ascii');
  expect(left.semanticSlotSpans).toBeDefined(); // retained original ownership
  const adapter = new LineMeasurementAdapter(context, 1, () => '10px Test Resource');
  expect(adapter.wordBoundaryAdvance(left, right)).toBe(-1);
  // A single live slot restores only the preexisting ordinary same-route guard,
  // rather than admitting every one-span service result as a physical join.
  expect(adapter.wordBoundaryAdvance(left, { ...right, fontRoute: {
    ...right.fontRoute!, fingerprint: right.fontRoute!.fingerprint + ':foreign',
  } })).toBe(0);
  const lines = layoutLines(context, s, 25, 0, 1, [], undefined, {}, 0,
    undefined, undefined, undefined, undefined, false, true, false,
    undefined, 'bounded', undefined, false, true);
  const second = lines[1].segments.filter((v): v is LayoutTextSeg => 'text' in v);
  expect(second.map(v => v.text)).toEqual(['TT ', 'T']);
  expect(second.reduce((n, v) => n + v.measuredWidth, 0)).toBe(23);
  expect(second.at(-1)!.leadingWordBoundaryPx).toBe(-1);
});
