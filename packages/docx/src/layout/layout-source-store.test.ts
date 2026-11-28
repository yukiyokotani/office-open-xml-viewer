import { describe, expect, it } from 'vitest';
import type { BodyElement, DocParagraph, DocxDocumentModel } from '../types.js';
import {
  layoutSourceModelAdapter,
  layoutSourceModelAdapterFromOwnedModel,
  layoutSourceStore,
} from '../layout-source-model-adapter.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import { layoutSourceStoreOf } from './runtime-state.js';
import { createLayoutBlockRepository, sealLayoutSourceStore, type LayoutSourceStore } from './layout-source-store.js';

function documentFacts(source: LayoutSourceStore) {
  return Object.freeze({
    ...source.documentLayoutSettings,
    kinsoku: Object.freeze({
      enabled: source.documentLayoutSettings.kinsoku.enabled,
      lineStartForbidden: Object.freeze([...source.documentLayoutSettings.kinsoku.lineStartForbidden]),
      lineEndForbidden: Object.freeze([...source.documentLayoutSettings.kinsoku.lineEndForbidden]),
    }),
  });
}

function documentWithUnavailableDrawing(): DocxDocumentModel {
  const paragraph = {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [{
      type: 'unavailableDrawing',
      resourceKind: 'image',
      widthPt: 24,
      heightPt: 12,
    }],
  } as unknown as DocParagraph;
  return {
    section: {
      pageWidth: 612,
      pageHeight: 792,
      marginTop: 72,
      marginRight: 72,
      marginBottom: 72,
      marginLeft: 72,
      headerDistance: 36,
      footerDistance: 36,
    },
    body: [paragraph],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

/** Reserved separator wire as a producer emits it. */
function reservedSeparatorDocument(noteId = '1'): DocxDocumentModel {
  const run = (fontSize: number) => ({
    text: '', bold: false, italic: false, underline: false, strikethrough: false, fontSize,
    color: null, fontFamily: 'Control Face', isLink: false, background: null, vertAlign: null, hyperlink: null,
  });
  const paragraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0,
    lineSpacing: null, numbering: null, tabStops: [], runs: [], defaultFontSize: 9,
    paragraphMarkFontFacts: { fontSize: 9 },
  };
  const rule = (mark: 'short' | 'full') => ({
    rule: { mark, control: { run: run(14) } },
    paragraph: { paragraph, contentMark: { run: run(9) } },
  });
  const document = documentWithUnavailableDrawing();
  document.footnotes = [{ id: noteId, content: [] }];
  return { ...document, __noteLayoutSettings: {
    footnoteSeparator: 'short', footnoteContinuationSeparator: 'full',
    reservedSeparators: { footnote: {
      separator: rule('short'),
      continuationSeparator: rule('full'),
      continuationNotice: {},
    } },
  } } as unknown as DocxDocumentModel;
}

describe('LayoutSourceStore', () => {
  it('caches one sealed normalized source for raw and normalized model identities', () => {
    const raw = documentWithUnavailableDrawing();
    const adapted = layoutSourceModelAdapter(raw);
    const first = adapted.source;

    expect(Object.isFrozen(first)).toBe(true);
    expect('document' in first).toBe(false);
    expect(first).toBe(layoutSourceStore(raw));
    expect(first).toBe(layoutSourceStore(adapted.document));
    expect(adapted.document).not.toBe(raw);
  });

  it('retains parser-private acquisition facts on the normalized document identity', () => {
    const adapted = layoutSourceModelAdapter(documentWithUnavailableDrawing());
    const store = adapted.source;
    const paragraph = store.blocks.resolve({ story: 'body', storyInstance: 'body', path: [0] });
    expect(paragraph.type).toBe('paragraph');
    if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph fixture');
    const acquired = store.acquisition.acquisitionInputs.paragraphAcquisitionInput(
      paragraph,
      { story: 'body', storyInstance: 'body', path: [0] },
    );

    expect((adapted.document.body[0] as DocParagraph).runs).toEqual([]);
    expect(paragraph.runs).toBe(acquired.runs);
    expect(acquired.runs).toMatchObject([{
      type: 'unavailableDrawing',
      resourceKind: 'image',
      widthPt: 24,
      heightPt: 12,
    }]);
    expect(store.bodyLayoutInput.sequence[0]).toMatchObject({
      kind: 'body-block',
      block: { kind: 'paragraph', inkless: false },
    });
    expect(() => store.acquisition.publicAnchorBridge(
      { story: 'body', storyInstance: 'body', path: [99] }, 0,
    )).toThrow(/unknown paragraph acquisition source/i);
    expect(() => store.acquisition.publicAnchorBridge(
      { story: 'body', storyInstance: 'body', path: [0] }, 99,
    )).toThrow(/unknown paragraph anchor bridge index/i);
  });

  it('retains unavailable drawing geometry inside a complete text-box story', () => {
    const model = documentWithUnavailableDrawing();
    const nested = structuredClone(model.body[0]) as Extract<BodyElement, { type: 'paragraph' }>;
    const outer = structuredClone(model.body[0]) as Extract<BodyElement, { type: 'paragraph' }>;
    outer.runs = [{
      type: 'shape', widthPt: 120, heightPt: 60, zOrder: 0,
      subpaths: [], fill: null, stroke: null,
      textBoxContent: [nested],
    }] as unknown as DocParagraph['runs'];
    model.body = [outer];

    const adapted = layoutSourceModelAdapter(model);
    const root = adapted.source.blocks.storyRoot({
      story: 'textbox', storyInstance: 'body:body:0.0', path: [],
    });
    expect(root).toHaveLength(1);
    expect(root[0]?.type).toBe('paragraph');
    if (root[0]?.type !== 'paragraph') throw new Error('Expected text-box paragraph fixture');
    expect(root[0].runs).toMatchObject([{
      type: 'unavailableDrawing',
      resourceKind: 'image',
      widthPt: 24,
      heightPt: 12,
    }]);
    expect((adapted.document.body[0] as DocParagraph).runs[0]).not.toHaveProperty(
      'textBoxContent.0.runs.0.type',
      'unavailableDrawing',
    );
  });

  it('keeps the builder-owned stream adapter semantically equal to the model adapter', () => {
    const raw = documentWithUnavailableDrawing();
    const nestedParagraph = raw.body[0]!;
    raw.body = [{
      type: 'table', colWidths: [100], layout: 'fixed', alignment: 'left', indent: 0,
      bidiVisual: false, cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      cellMarginRight: 0, rows: [{ isHeader: false, cantSplit: false, gridBefore: 0,
        gridAfter: 0, rowHeight: null, rowHeightRule: 'auto', cells: [{ colSpan: 1,
          rowSpan: 1, vAlign: 'top', marginTop: null, marginBottom: null,
          marginLeft: null, marginRight: null, content: [nestedParagraph] }] }],
    } as unknown as DocxDocumentModel['body'][number]];
    raw.headers.default = { body: [structuredClone(nestedParagraph)] };
    raw.footnotes = [{ id: '1', content: [structuredClone(nestedParagraph)] }];
    const baseline = layoutSourceModelAdapter(structuredClone(raw));
    const owned = structuredClone(raw);
    const ownedBody = owned.body;
    const ownedNestedParagraph = (owned.body[0] as Extract<DocxDocumentModel['body'][number], { type: 'table' }>)
      .rows[0]!.cells[0]!.content[0]!;
    const streamed = layoutSourceModelAdapterFromOwnedModel(
      structuredClone(raw),
      owned,
    );

    expect(streamed.document).toEqual(baseline.document);
    expect(streamed.source.bodyLayoutInput).toEqual(baseline.source.bodyLayoutInput);
    expect(streamed.source.blocks.body).toEqual(baseline.source.blocks.body);
    expect(streamed.source.section).toEqual(baseline.source.section);
    expect(documentFacts(streamed.source)).toEqual(documentFacts(baseline.source));
    expect(streamed.source.paintResources.descriptors)
      .toEqual(baseline.source.paintResources.descriptors);
    expect(streamed.source.blocks.body).toBe(ownedBody);
    const retainedNestedParagraph = (owned.body[0] as Extract<DocxDocumentModel['body'][number], { type: 'table' }>)
      .rows[0]!.cells[0]!.content[0]!;
    expect(retainedNestedParagraph).not.toBe(ownedNestedParagraph);
    const retainedBodyBlock = streamed.source.blocks.body[0]!;
    if (retainedBodyBlock.type !== 'table') throw new Error('Expected retained table block');
    expect(retainedBodyBlock.rows[0]!.cells[0]!.content[0]).toBe(retainedNestedParagraph);
  });

  it('owns its eager body input as an immutable data property', () => {
    const store = layoutSourceStore(documentWithUnavailableDrawing());
    const keys = Object.keys(store);
    const descriptor = Object.getOwnPropertyDescriptor(store, 'bodyLayoutInput');
    const input = store.bodyLayoutInput;

    expect(descriptor?.get).toBeUndefined();
    expect(descriptor?.value).toBe(input);
    expect(Object.isFrozen(input)).toBe(true);
    expect(store.bodyLayoutInput).toBe(input);
    expect(Object.isFrozen(store)).toBe(true);
    expect(Object.keys(store)).toEqual(keys);
    expect(Object.getOwnPropertyDescriptor(store, 'bodyLayoutInput')?.value)
      .toBe(descriptor?.value);
  });

  it('keeps chart and math payloads in manifests, not canonical paragraph runs', () => {
    const model = documentWithUnavailableDrawing();
    const chartSentinel = 'CHART_PAYLOAD_SENTINEL';
    const mathSentinel = 'MATH_PAYLOAD_SENTINEL';
    (model.body[0] as DocParagraph).runs = [{
      type: 'chart',
      widthPt: 40,
      heightPt: 20,
      anchor: false,
      chart: { title: { text: chartSentinel } },
    }, {
      type: 'math',
      display: false,
      fontSize: 12,
      nodes: [{ kind: 'run', text: 'x', payloadSentinel: mathSentinel }],
    }] as unknown as DocParagraph['runs'];

    const store = layoutSourceStore(model);
    const paragraph = store.blocks.resolve({
      story: 'body', storyInstance: 'body', path: [0],
    });
    const canonicalJson = JSON.stringify(paragraph);
    const manifestJson = JSON.stringify({
      math: store.mathOccurrences,
      paint: store.paintResources.descriptors,
    });

    expect(canonicalJson).not.toContain(chartSentinel);
    expect(canonicalJson).not.toContain(mathSentinel);
    expect(paragraph.type === 'paragraph' ? paragraph.runs : []).toMatchObject([
      { type: 'chart', resourceKey: expect.any(String) },
      { type: 'math', resourceKey: expect.any(String), fallbackText: 'x' },
    ]);
    expect(paragraph.type === 'paragraph' ? paragraph.runs[0] : null).not.toHaveProperty('chart');
    expect(paragraph.type === 'paragraph' ? paragraph.runs[1] : null).not.toHaveProperty('nodes');
    expect(manifestJson).toContain(chartSentinel);
    expect(manifestJson).toContain(mathSentinel);
  });

  it('removes parser-only paragraph, table, section, and VML wires from canonical JSON', () => {
    const model = documentWithUnavailableDrawing();
    const sentinel = 'PRIVATE_LAYOUT_WIRE_SENTINEL';
    const paragraph = model.body[0] as unknown as DocParagraph & Record<string, unknown>;
    paragraph.paragraphMarkFontFacts = { fontSize: 10, privateSentinel: sentinel };
    paragraph.numbering = {
      numId: 1, level: 0, format: 'decimal', text: '1.', indentLeft: 18, tab: 18,
      fontFacts: { fontSize: 10, privateSentinel: sentinel },
    } as never;
    paragraph.runs = [{
      type: 'shape', widthPt: 20, heightPt: 10, anchorXPt: 0, anchorYPt: 0,
      anchorXFromMargin: false, anchorYFromPara: false, zOrder: 0, subpaths: [],
      fill: null, stroke: null,
      textPath: { string: 'visible', privateSentinel: sentinel },
    }] as unknown as DocParagraph['runs'];
    const table = {
      type: 'table', colWidths: [100], __tableLayout: { privateSentinel: sentinel },
      rows: [{ __tableRowLayout: { privateSentinel: sentinel }, cells: [{
        __tableCellLayout: { privateSentinel: sentinel },
        content: [documentWithUnavailableDrawing().body[0]], colSpan: 1,
        vMerge: false, vAlign: 'top',
      }] }],
    };
    model.body.push(table as never);
    (model.section as unknown as Record<string, unknown>).__sectionPlacement = {
      privateSentinel: sentinel,
    };

    const store = layoutSourceStore(model);
    const canonicalJson = JSON.stringify({
      body: store.blocks.body,
      section: store.section,
    });

    expect(canonicalJson).not.toContain(sentinel);
    expect(canonicalJson).not.toMatch(/"__[^"]+"/);
    expect(canonicalJson).not.toContain('fontFacts');
    expect(canonicalJson).not.toContain('paragraphMarkFontFacts');
  });

  it('retains the same source through externally composed layout services', () => {
    const raw = documentWithUnavailableDrawing();
    const source = layoutSourceStore(raw);
    const services = createLayoutServices(source);
    const composed = Object.freeze({
      ...services,
      text: Object.freeze({ ...services.text }),
    });

    expect(layoutSourceStoreOf(services)).toBe(source);
    expect(layoutSourceStoreOf(composed)).toBe(source);
  });

  it('is recursively detached from later public-model mutation', () => {
    const adapted = layoutSourceModelAdapter(documentWithUnavailableDrawing());
    const retainedSectionWidth = adapted.source.section.pageWidth;
    const retainedBodyLength = adapted.source.blocks.body.length;

    adapted.document.section.pageWidth = 1;
    adapted.document.body.push(adapted.document.body[0]!);

    expect(adapted.source.section.pageWidth).toBe(retainedSectionWidth);
    expect(adapted.source.blocks.body).toHaveLength(retainedBodyLength);
    expect(Object.isFrozen(adapted.source.blocks.body)).toBe(true);
    expect(Object.isFrozen(adapted.source.blocks.body[0])).toBe(true);
    expect(Object.isFrozen(adapted.source.paintResources.descriptors)).toBe(true);
  });

  it('does not freeze or alias nested compatibility-model content', () => {
    const raw = documentWithUnavailableDrawing();
    const paragraph = raw.body[0] as DocParagraph;
    paragraph.runs = [{
      type: 'text', text: 'before', fontFamily: 'Arial', fontSize: 10,
      bold: false, italic: false,
    }] as unknown as DocParagraph['runs'];
    const adapted = layoutSourceModelAdapter(raw);
    const retained = adapted.source.blocks.resolve({
      story: 'body', storyInstance: 'body', path: [0],
    });
    if (retained.type !== 'paragraph') throw new Error('Expected paragraph fixture');

    const publicParagraph = adapted.document.body[0] as DocParagraph;
    expect(Object.isFrozen(publicParagraph)).toBe(false);
    expect(Object.isFrozen(publicParagraph.runs[0]!)).toBe(false);
    (publicParagraph.runs[0] as Extract<DocParagraph['runs'][number], { type: 'text' }>).text = 'after';
    publicParagraph.alignment = 'right';

    expect(retained.alignment).not.toBe('right');
    expect(retained.runs[0]).toMatchObject({ type: 'text', text: 'before' });
  });

  it('resolves exact nested and story block sources and rejects non-root story lookup', () => {
    const paragraph = (text: string) => ({
      type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0,
      indentFirst: 0, spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
      numbering: null, tabStops: [], runs: [{ type: 'text', text, fontSize: 10 }],
    }) as unknown as DocxDocumentModel['body'][number];
    const model = documentWithUnavailableDrawing();
    model.body = [{
      type: 'table', colWidths: [100], layout: 'fixed', alignment: 'left', indent: 0,
      bidiVisual: false, cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      cellMarginRight: 0, rows: [{ isHeader: false, cantSplit: false, gridBefore: 0,
        gridAfter: 0, rowHeight: null, rowHeightRule: 'auto', cells: [{ colSpan: 1,
          rowSpan: 1, vAlign: 'top', marginTop: null, marginBottom: null,
          marginLeft: null, marginRight: null, content: [paragraph('nested')] }] }],
    } as unknown as DocxDocumentModel['body'][number]];
    model.headers.default = { body: [paragraph('header')] };
    model.footnotes = [{ id: '7', content: [paragraph('note')] }];
    const blocks = layoutSourceStore(model).blocks;

    expect(blocks.resolve({ story: 'body', storyInstance: 'body', path: [0, 0, 0, 0] }).type)
      .toBe('paragraph');
    expect(blocks.resolve({ story: 'header', storyInstance: 'default', path: [0] }).type)
      .toBe('paragraph');
    expect(blocks.resolve({ story: 'footnote', storyInstance: '7', path: [0] }).type)
      .toBe('paragraph');
    expect(() => blocks.storyRoot({ story: 'header', storyInstance: 'default', path: [0] }))
      .toThrow(/root-only/);
  });

  it('seals and lays out a model-free source input directly', () => {
    const model = documentWithUnavailableDrawing();
    const adapted = layoutSourceStore(model);
    const paragraphRef = Object.freeze({
      story: 'body' as const, storyInstance: 'body', path: Object.freeze([0]),
    });
    const paragraph = adapted.blocks.resolve(paragraphRef);
    if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph fixture');
    const modelFree = sealLayoutSourceStore({
      bodyLayoutInput: adapted.bodyLayoutInput,
      blockRepository: Object.freeze({ body: adapted.blocks.body, stories: Object.freeze([]),
        footnotes: adapted.blocks.footnotes, endnotes: adapted.blocks.endnotes }),
      acquisitionFacts: Object.freeze({
        paragraphs: Object.freeze([Object.freeze({
          source: paragraphRef,
          publicAnchorBridges: Object.freeze(paragraph.runs.map(() => null)),
          numberingMarkerFallbackFontSizePt: null,
        })]),
        tables: Object.freeze([]),
      }),
      section: adapted.section,
      documentLayoutFacts: documentFacts(adapted),
      fonts: adapted.fonts,
      fontFamilyCharsets: adapted.fontFamilyCharsets,
      mathOccurrences: adapted.mathOccurrences,
      imageMetadata: adapted.imageMetadata,
      paintDescriptors: adapted.paintResources.descriptors,
      hasPaginationFields: adapted.hasPaginationFields,
      requiresDomVerticalGlyphLayout: adapted.requiresDomVerticalGlyphLayout,
      fatalParse: adapted.fatalParse,
    });
    const context = {
      font: '', letterSpacing: '0px', fontKerning: 'auto',
      measureText: () => ({ width: 0, actualBoundingBoxAscent: 0,
        actualBoundingBoxDescent: 0, fontBoundingBoxAscent: 0, fontBoundingBoxDescent: 0 }),
    } as unknown as CanvasRenderingContext2D;

    const direct = layoutDocument(modelFree, createLayoutServices(modelFree, { measureContext: context }));
    const full = layoutDocument(adapted, createLayoutServices(adapted, { measureContext: context }));
    expect(direct).toEqual(full);
    expect(modelFree.paintResources.descriptors).toEqual(adapted.paintResources.descriptors);
    expect('document' in modelFree).toBe(false);
  });

  it('rejects duplicate repository sources and unresolved acquisition facts', () => {
    expect(() => createLayoutBlockRepository({
      body: Object.freeze([]),
      stories: Object.freeze([
        { source: { story: 'header', storyInstance: 'default', path: [] }, body: Object.freeze([]) },
        { source: { story: 'header', storyInstance: 'default', path: [] }, body: Object.freeze([]) },
      ]),
      footnotes: Object.freeze([]),
      endnotes: Object.freeze([]),
    })).toThrow(/duplicate story source/i);
    expect(() => createLayoutBlockRepository({
      body: Object.freeze([]),
      stories: Object.freeze([{
        source: { story: 'footnote', storyInstance: 'misfiled', path: [] },
        body: Object.freeze([]),
      }]),
      footnotes: Object.freeze([]),
      endnotes: Object.freeze([]),
    })).toThrow(/unsupported repository story kind/i);

    const model = documentWithUnavailableDrawing();
    model.body = [];
    const source = layoutSourceStore(model);
    expect(() => sealLayoutSourceStore({
      bodyLayoutInput: source.bodyLayoutInput,
      blockRepository: Object.freeze({ body: source.blocks.body, stories: Object.freeze([]),
        footnotes: source.blocks.footnotes, endnotes: source.blocks.endnotes }),
      acquisitionFacts: Object.freeze({
        paragraphs: Object.freeze([Object.freeze({
          source: Object.freeze({ story: 'body' as const, storyInstance: 'body', path: Object.freeze([99]) }),
          publicAnchorBridges: Object.freeze([]),
          numberingMarkerFallbackFontSizePt: null,
        })]),
        tables: Object.freeze([]),
      }),
      section: source.section,
      documentLayoutFacts: documentFacts(source),
      fonts: source.fonts,
      fontFamilyCharsets: source.fontFamilyCharsets,
      mathOccurrences: source.mathOccurrences,
      imageMetadata: source.imageMetadata,
      paintDescriptors: source.paintResources.descriptors,
      hasPaginationFields: false,
      requiresDomVerticalGlyphLayout: false,
      fatalParse: null,
    })).toThrow(/unknown block source/i);
  });

  it('rejects acquisition fact fields outside the canonical exact shape', () => {
    const source = layoutSourceStore(documentWithUnavailableDrawing());
    const ref = Object.freeze({
      story: 'body' as const, storyInstance: 'body', path: Object.freeze([0]),
    });
    const paragraph = source.blocks.resolve(ref);
    if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph fixture');

    expect(() => sealLayoutSourceStore({
      bodyLayoutInput: source.bodyLayoutInput,
      blockRepository: Object.freeze({
        body: source.blocks.body, stories: Object.freeze([]),
        footnotes: source.blocks.footnotes, endnotes: source.blocks.endnotes,
      }),
      acquisitionFacts: Object.freeze({
        paragraphs: Object.freeze([Object.freeze({
          source: ref,
          publicAnchorBridges: Object.freeze(paragraph.runs.map(() => null)),
          numberingMarkerFallbackFontSizePt: null,
          obsoletePayloadFacade: Object.freeze({ runs: Object.freeze([]) }),
        })]),
        tables: Object.freeze([]),
      }),
      section: source.section,
      documentLayoutFacts: documentFacts(source),
      fonts: source.fonts,
      fontFamilyCharsets: source.fontFamilyCharsets,
      mathOccurrences: source.mathOccurrences,
      imageMetadata: source.imageMetadata,
      paintDescriptors: source.paintResources.descriptors,
      hasPaginationFields: false,
      requiresDomVerticalGlyphLayout: false,
      fatalParse: null,
    })).toThrow(/unexpected fields/i);
  });

  it('indexes numbering markers by identity and fallback size without lookup rescans', () => {
    const model = documentWithUnavailableDrawing();
    const numbering = { numId: 1, level: 0, format: 'decimal', text: '1.',
      indentLeft: 18, tab: 18, suff: 'tab' };
    model.body = Array.from({ length: 40 }, (_, index) => ({
      type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0,
      indentFirst: 0, spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
      numbering, tabStops: [], runs: [{ type: 'text', text: 'x',
        fontSize: index % 2 === 0 ? 10 : 20 }],
    })) as unknown as DocxDocumentModel['body'];
    const source = layoutSourceStore(model);
    const facts = source.blocks.body.map((_block, index) => {
      const ref = Object.freeze({ story: 'body' as const, storyInstance: 'body', path: Object.freeze([index]) });
      const paragraph = source.blocks.resolve(ref);
      if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph fixture');
      return Object.freeze({
        source: ref,
        publicAnchorBridges: Object.freeze(paragraph.runs.map(() => null)),
        numberingMarkerFallbackFontSizePt: index % 2 === 0 ? 10 : 20,
      });
    });
    const indexed = sealLayoutSourceStore({
      bodyLayoutInput: source.bodyLayoutInput,
      blockRepository: Object.freeze({ body: source.blocks.body, stories: Object.freeze([]),
        footnotes: source.blocks.footnotes, endnotes: source.blocks.endnotes }),
      acquisitionFacts: Object.freeze({ paragraphs: Object.freeze(facts), tables: Object.freeze([]) }),
      section: source.section, documentLayoutFacts: documentFacts(source),
      fonts: source.fonts, fontFamilyCharsets: source.fontFamilyCharsets,
      mathOccurrences: source.mathOccurrences, imageMetadata: source.imageMetadata,
      paintDescriptors: source.paintResources.descriptors, hasPaginationFields: false,
      requiresDomVerticalGlyphLayout: false, fatalParse: null,
    });
    for (let iteration = 0; iteration < 100; iteration += 1) {
      for (const index of [0, 1]) {
        const paragraph = indexed.blocks.resolve({
          story: 'body', storyInstance: 'body', path: [index],
        });
        if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph fixture');
        const fallback = index === 0 ? 10 : 20;
        expect(indexed.acquisition.acquisitionInputs.numberingMarkerShapeInput(
          paragraph.numbering!, fallback,
        ).fontSizePt).toBe(fallback);
      }
    }
  });
  it('registers reserved separator stories under closed canonical roots', () => {
    const settings = (store: LayoutSourceStore) => store.bodyLayoutInput.noteLayoutSettings;
    const first = layoutSourceStore(reservedSeparatorDocument());
    // The sealed layout input holds canonical role refs and no raw
    // parser-shaped separator wire.
    expect(settings(first)).not.toHaveProperty('reservedSeparators');
    const roles = settings(first)?.reservedSeparatorRoles?.footnote;
    expect(roles?.separator.root).toEqual({ story: 'footnote', storyInstance: 'reserved:separator', path: [] });
    expect(roles?.separator.rule).toMatchObject({
      mark: 'short', control: { storyInstance: 'reserved:separator', path: [0, 0] },
    });
    expect(roles?.continuationSeparator.rule?.mark).toBe('full');
    const paragraph = first.blocks.resolve(roles!.separator.paragraph!.source);
    if (paragraph.type !== 'paragraph') throw new Error('Expected a reserved separator paragraph');
    // Control and content mark stay distinct participants with their own properties.
    expect(paragraph.runs.map((run) => run.type === 'text'
      ? [run.noteSeparatorCharacter, run.fontSize] : null))
      .toEqual([['rule-control', 14], ['paragraph-mark', 9]]);
    // An empty notice stays addressable without a fabricated paragraph, and
    // reserved roots are not numbered note bodies.
    expect(first.blocks.storyRoot(roles!.continuationNotice.root)).toEqual([]);
    expect(roles?.continuationNotice.paragraph).toBeUndefined();
    expect(first.blocks.footnotes.map((note) => note.id)).toEqual(['1']);
    // A face used only by the control joins the document font ownership.
    expect(first.fonts.renderedFamilies).toContain('Control Face');

    const streamed = layoutSourceModelAdapterFromOwnedModel(
      reservedSeparatorDocument(), reservedSeparatorDocument(),
    ).source;
    expect(streamed.bodyLayoutInput).toEqual(first.bodyLayoutInput);
    expect(streamed.blocks.resolve(roles!.separator.paragraph!.source)).toEqual(paragraph);
  });

  it.each([false, true])('resolves reserved-story automatic margins with document fixed=%s before immutable acquisition', (fixed) => {
    const document = reservedSeparatorDocument() as DocxDocumentModel & {
      __noteLayoutSettings: { reservedSeparators: { footnote: Record<string, Record<string, unknown>> } };
    };
    document.settings = { doNotUseHtmlParagraphAutoSpacing: fixed };
    const stories = document.__noteLayoutSettings.reservedSeparators.footnote;
    const wrapper = stories.separator!.paragraph as {
      paragraph: DocParagraph;
      contentMark: unknown;
    };
    Object.assign(wrapper.paragraph, { spaceBefore: 5, spaceAfter: 7, beforeAutospacing: true, afterAutospacing: false });
    stories.continuationNotice = {
      paragraph: { ...wrapper, paragraph: { ...wrapper.paragraph, afterAutospacing: true } },
    };
    const original = structuredClone(document);
    const source = layoutSourceStore(document);
    const roles = source.bodyLayoutInput.noteLayoutSettings?.reservedSeparatorRoles?.footnote;
    if (!roles?.separator.paragraph || !roles.continuationNotice.paragraph) throw new Error('Expected reserved paragraphs');
    const rule = source.blocks.resolve(roles.separator.paragraph.source);
    const notice = source.blocks.resolve(roles.continuationNotice.paragraph.source);
    if (rule.type !== 'paragraph' || notice.type !== 'paragraph') throw new Error('Expected reserved paragraphs');
    // Automatic sides take the 14pt HTML amount; neither the 9pt paragraph
    // base nor the rule control changes it.
    expect([rule.spaceBefore, rule.spaceAfter]).toEqual(fixed ? [5, 7] : [14, 7]);
    expect([notice.spaceBefore, notice.spaceAfter]).toEqual(fixed ? [5, 7] : [14, 14]);
    expect(document).toEqual(original);
    expect(Object.isFrozen(rule)).toBe(true);
    const streamed = layoutSourceModelAdapterFromOwnedModel(structuredClone(document), structuredClone(document)).source;
    expect(streamed.blocks.resolve(roles.separator.paragraph.source)).toEqual(rule);
    // Settings and source mutations cannot change an acquired story or its owner.
    document.settings.doNotUseHtmlParagraphAutoSpacing = !fixed;
    wrapper.paragraph.spaceBefore = 90;
    expect(source.blocks.resolve(roles.separator.paragraph.source)).toBe(rule);
    expect([rule.spaceBefore, rule.spaceAfter]).toEqual(fixed ? [5, 7] : [14, 7]);
  });

  it.each(['rule-control', 'rule-mark', 'notice-mark'] as const)(
    'rejects a reserved %s glyph before metric-only normalization can discard it',
    (participant) => {
      const document = reservedSeparatorDocument() as DocxDocumentModel & {
        __noteLayoutSettings: { reservedSeparators: { footnote: Record<string, Record<string, unknown>> } };
      };
      const stories = document.__noteLayoutSettings.reservedSeparators.footnote;
      const paragraph = stories.separator!.paragraph as Record<string, unknown>;
      let character: { run: { text: string } };
      if (participant === 'rule-control') {
        character = (stories.separator!.rule as { control: typeof character }).control;
      } else if (participant === 'rule-mark') {
        character = paragraph.contentMark as typeof character;
      } else {
        // A paragraph-only notice follows the ordinary empty-paragraph path,
        // so its content mark must be checked even without a rule participant.
        const mark = structuredClone(paragraph.contentMark) as typeof character;
        stories.continuationNotice = { paragraph: { ...paragraph, contentMark: mark } };
        character = mark;
      }
      // A producer's symbol substitution can give an empty control text.
      character.run.text = '\u03a9';
      expect(() => layoutSourceStore(document)).toThrow('carries unsupported authored text');
    },
  );

  it('rejects reserved root collisions and inconsistent reserved separator shapes before layout', () => {
    expect(() => layoutSourceStore(reservedSeparatorDocument('reserved:separator')))
      .toThrow('collides with a reserved note separator story');
    const mutate = (edit: (stories: Record<string, Record<string, unknown>>) => void) => {
      const document = reservedSeparatorDocument() as DocxDocumentModel & {
        __noteLayoutSettings: { reservedSeparators: { footnote: Record<string, Record<string, unknown>> } };
      };
      edit(document.__noteLayoutSettings.reservedSeparators.footnote);
      return () => layoutSourceStore(document);
    };
    // Malformed records fail closed instead of sealing as empty stories,
    // which would suppress the scalar separator of their note kind.
    for (const [label, edit] of [
      ['number story', (stories) => { stories.separator = 42 as never; }],
      ['array story', (stories) => { stories.separator = [] as never; }],
      ['null rule', (stories) => { stories.separator!.rule = null; }],
      ['null paragraph', (stories) => { stories.continuationNotice!.paragraph = null; }],
      ['missing role', (stories) => { delete stories.continuationNotice; }],
      ['unknown story field', (stories) => { stories.continuationNotice!.guardCp = 7; }],
      ['non-record control', (stories) => {
        (stories.separator!.rule as Record<string, unknown>).control = 'x';
      }],
      ['non-boolean flag', (stories) => {
        (stories.separator!.paragraph as Record<string, unknown>).numbered = 0;
      }],
    ] as const satisfies readonly (readonly [string, (stories: Record<string, Record<string, unknown>>) => void])[]) {
      expect(mutate(edit), label).toThrow('Unsupported reserved note separator story');
    }
    const kind = (value: unknown) => {
      const document = reservedSeparatorDocument() as DocxDocumentModel & {
        __noteLayoutSettings: { reservedSeparators: Record<string, unknown> };
      };
      document.__noteLayoutSettings.reservedSeparators.endnote = value;
      return () => layoutSourceStore(document);
    };
    expect(kind(null)).toThrow('Unsupported reserved note separator story');
    expect(kind([])).toThrow('Unsupported reserved note separator story');
    // An absent kind is not a malformed one.
    expect(kind(undefined)).not.toThrow();
    // A rule always comes with the paragraph whose mark ends it.
    expect(mutate((stories) => { delete stories.separator!.paragraph; }))
      .toThrow('Unsupported reserved note separator story');
    // Presence-only frame/numbering flags mean the owned properties were not kept.
    expect(mutate((stories) => {
      (stories.separator!.paragraph as Record<string, unknown>).framed = true;
    })).toThrow('Unsupported reserved note separator story');
  });
});
