import { describe, expect, it, vi } from 'vitest';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';
import {
  attachDocumentLayoutVariants,
  selectDocumentLayoutPage,
} from './layout/document-layout-variants.js';
import {
  attachDocumentLayoutRuntime,
  documentLayoutRuntimeOf,
  layoutVariantStoreOf,
} from './layout/runtime-state.js';
import type { DeepReadonly, DocumentLayout, LayoutServices } from './layout/types.js';
import type {
  DocumentMeta,
  RenderWorkerRequest,
  RenderWorkerResponse,
} from './worker-protocol.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { DocxDocument } from './document.js';
import { retainRenderWorkerDocumentLayout } from './render-worker-layout.js';
import {
  projectRenderWorkerLayoutMeta,
  renderWorkerLayoutMeta,
} from './render-worker-metadata.js';
import { stableFingerprint } from './layout/fingerprint.js';
import { buildBookmarkPageMap } from './bookmark-nav.js';
import { textRunsForSelectedPage } from './text-run-projection.js';
import { DEFAULT_OOXML_RESOURCE_LIMITS } from '@silurus/ooxml-core/worker';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import {
  createLocalDocumentPullTransport,
  DocumentPullWorker,
  MaterializedDocumentCursorArchive,
} from './document-pull-worker.js';
import { syntheticDocxModel } from './testing/synthetic-document.js';

function services(): LayoutServices {
  return Object.freeze({
    text: { fingerprint: 'text:equal' },
    images: { fingerprint: 'images:equal' },
    math: { fingerprint: 'math:equal' },
  }) as LayoutServices;
}

function layout(currentDateMs: number): DocumentLayout {
  const variant = currentDateMs === 10 ? 'default' : 'dated';
  const count = variant === 'default' ? 1 : 2;
  return {
    pages: Array.from({ length: count }, (_, pageIndex) => ({
      pageIndex,
      geometry: {
        widthPt: variant === 'default' ? 612 : 792,
        heightPt: variant === 'default' ? 792 : 612,
      },
      variant,
    }) as never),
    diagnostics: [],
  };
}

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '', letterSpacing: '0px', fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 8,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
    }),
  } as unknown as CanvasRenderingContext2D;
}

function realModel(): DocxDocumentModel {
  return {
    section: {
      pageWidth: 612, pageHeight: 792,
      marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
    },
    body: [{
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
      defaultFontSize: 10,
      defaultFontFamily: 'Times New Roman',
      widowControl: false,
      runs: [{
        type: 'field',
        fieldType: 'date',
        instruction: ' DATE \\@ "yyyy" ',
        fallbackText: 'cached',
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 10,
        color: null,
        fontFamily: 'Times New Roman',
        background: null,
        vertAlign: null,
      }],
    }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  };
}

interface PrivateSectionPlacementWire {
  readonly sectionId: string;
  readonly sectionBidi: boolean;
}

type PrivateSectionProps = SectionProps & {
  readonly __sectionPlacement: PrivateSectionPlacementWire;
};

function finalNextColumnRtlModel(): DocxDocumentModel {
  const base = realModel();
  const columns = {
    count: 2,
    spacePt: 20,
    equalWidth: true,
    sep: false,
    cols: [],
  };
  const endingSection = {
    type: 'sectionBreak',
    kind: 'nextPage',
    geom: base.section,
    columns,
    textDirection: null,
    pageNumType: null,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    titlePage: false,
    __sectionPlacement: {
      sectionId: 'section:outgoing',
      sectionBidi: true,
    },
  } as unknown as BodyElement;
  const finalSection = {
    ...base.section,
    sectionStart: 'nextColumn',
    columns,
    __sectionPlacement: {
      sectionId: 'section:final',
      sectionBidi: true,
    },
  } as PrivateSectionProps;

  return {
    ...base,
    body: [base.body[0]!, endingSection, structuredClone(base.body[0]!)],
    section: finalSection,
  };
}

function metadataForDefaultLayout(
  model: DocxDocumentModel,
  layout: DeepReadonly<DocumentLayout>,
): DocumentMeta {
  return {
    pageCount: layout.pages.length,
    revisions: model.revisions ?? [],
    comments: model.comments ?? [],
    footnotes: model.footnotes ?? [],
    endnotes: model.endnotes ?? [],
    pageSizes: layout.pages.map((page) => ({
      widthPt: page.geometry.widthPt,
      heightPt: page.geometry.heightPt,
    })),
    bookmarkPages: [...buildBookmarkPageMap(layout)],
  };
}

describe('render worker canonical layout parity', () => {
  it('does not traverse page geometry when a publication has no review data', () => {
    const layout = {
      pages: [{
        pageIndex: 0,
        geometry: { widthPt: 612, heightPt: 792 },
        bookmarkStarts: [],
      }],
      diagnostics: [],
    } as unknown as DocumentLayout;
    const meta = projectRenderWorkerLayoutMeta(
      layout,
      {} as ReturnType<typeof layoutSourceStore>,
      { comments: [], revisions: [] },
      { provisional: true },
    );
    expect(meta.commentAnchorRanges).toEqual([]);
    expect(meta.revisionAnchorRanges).toEqual([]);
  });

  it('withholds future review anchors inside a split paragraph until its final fragment', () => {
    const model = syntheticDocxModel('plain', { paragraphs: 1 });
    const paragraph = model.body[0] as Extract<BodyElement, { type: 'paragraph' }>;
    const baseRun = paragraph.runs[0] as Extract<DocParagraph['runs'][number], { type: 'text' }>;
    paragraph.runs = [
      // Keep the review anchor beyond the first page without making this
      // metadata test depend on laying out an unnecessarily large document.
      { ...baseRun, text: 'opening '.repeat(1_000) },
      {
        ...baseRun,
        text: 'future review anchor',
        revision: { kind: 'deletion', id: '84' },
      },
    ];
    paragraph.commentMarks = [
      { id: '42', kind: 'rangeStart', runIndex: 1 },
      { id: '42', kind: 'rangeEnd', runIndex: 2 },
      { id: '42', kind: 'reference', runIndex: 2 },
    ];
    model.comments = [{ id: '42', text: 'future comment' }];
    model.revisions = [{ kind: 'deletion', id: '84', text: 'future revision' }];
    const source = layoutSourceStore(model);
    const layoutServices = createLayoutServices(source, { measureContext: measureContext() });
    const full = layoutDocument(model, layoutServices, { currentDateMs: 0 });
    expect(full.pages.length).toBeGreaterThan(1);
    const prefix = { ...full, pages: full.pages.slice(0, 1) } as DocumentLayout;
    const review = { comments: model.comments, revisions: model.revisions };

    const provisional = projectRenderWorkerLayoutMeta(
      prefix,
      source,
      review,
      { provisional: true },
    );
    expect(provisional.commentAnchorRanges).toEqual([]);
    expect(provisional.revisionAnchorRanges).toEqual([]);

    const authoritative = projectRenderWorkerLayoutMeta(full, source, review);
    expect(authoritative.commentAnchorRanges).toEqual([
      expect.objectContaining({ commentId: '42' }),
    ]);
    expect(authoritative.revisionAnchorRanges).toEqual([
      expect.objectContaining({ revisionIndex: 0, geometryFallback: expect.any(Object) }),
    ]);
  });

  it('projects page metadata from the runtime-selected tracked-changes variant', () => {
    const model = syntheticDocxModel('tracked', { paragraphs: 120 });
    const source = layoutSourceStore(model);
    const layoutServices = createLayoutServices(source, { measureContext: measureContext() });
    const retained = retainRenderWorkerDocumentLayout(source, layoutServices, 0);
    const review = { comments: model.comments ?? [], revisions: model.revisions ?? [] };

    const finalMeta = renderWorkerLayoutMeta(retained, review, 0, false);
    const markupMeta = renderWorkerLayoutMeta(retained, review, 0, true);

    expect(markupMeta.pageCount).not.toBe(finalMeta.pageCount);
    expect(markupMeta.pageSizes).toHaveLength(markupMeta.pageCount);
    expect(finalMeta.pageSizes).toHaveLength(finalMeta.pageCount);
  });

  it('exposes identical review metadata in main and worker modes', () => {
    const model = realModel();
    model.revisions = [{ kind: 'insertion', author: 'Reviewer', text: 'field' }];
    model.comments = [{ id: '7', author: 'Reviewer', text: 'Check field' }];
    const paragraph = model.body[0] as Extract<BodyElement, { type: 'paragraph' }>;
    paragraph.commentMarks = [
      { id: '7', kind: 'rangeStart', runIndex: 0 },
      { id: '7', kind: 'rangeEnd', runIndex: 1 },
      { id: '7', kind: 'reference', runIndex: 1 },
    ];
    const source = layoutSourceStore(model);
    const main = Object.create(DocxDocument.prototype) as DocxDocument;
    Object.assign(main, {
      _mode: 'main', _document: model, _source: source, _meta: null,
      _commentAnchorRanges: null,
      _revisionAnchorRanges: null,
    });
    attachDocumentLayoutRuntime(main, 0);
    const layoutServices = createLayoutServices(model, { measureContext: measureContext() });
    const variants = attachDocumentLayoutVariants({
      source,
      services: layoutServices,
      defaultCurrentDateMs: 0,
      buildLayout: (options) => layoutDocument(model, layoutServices, options),
    });
    documentLayoutRuntimeOf(main).services = layoutServices;
    const layout = variants.store.defaultLayout;
    const worker = Object.create(DocxDocument.prototype) as DocxDocument;
    Object.assign(worker, {
      _mode: 'worker', _document: null, _source: null,
      _meta: {
        ...metadataForDefaultLayout(model, layout),
        commentAnchorRanges: main.commentAnchorRanges(),
        revisionAnchorRanges: main.revisionAnchorRanges(),
      },
    });

    expect(worker.comments).toEqual(main.comments);
    expect(worker.revisions).toEqual(main.revisions);
    expect(worker.comments).toBe(worker.comments);
    expect(main.revisions).toBe(main.revisions);
    expect(Object.isFrozen(worker.comments)).toBe(true);
    expect(Object.isFrozen(worker.comments[0])).toBe(true);
    expect(Object.isFrozen(main.revisions)).toBe(true);
    expect(() => {
      (worker.comments[0] as { text: string }).text = 'caller mutation';
    }).toThrow(TypeError);
    expect(worker.comments[0]?.text).toBe('Check field');
    expect(worker.commentAnchorRanges()).toEqual(main.commentAnchorRanges());
    expect(worker.revisionAnchorRanges()).toEqual(main.revisionAnchorRanges());
  });

  it('switches the effective mode to main when the worker returns a vertical model', async () => {
    const model = realModel();
    model.section.textDirection = 'tbRl';
    const identity = { sessionId: 1, operationId: 1, generation: 1 };
    const fallbackArchive = new MaterializedDocumentCursorArchive(model);
    const fallbackWorker = new DocumentPullWorker(() => fallbackArchive);
    fallbackWorker.open(identity);
    const fallbackTransport = createLocalDocumentPullTransport(fallbackWorker);
    const requests: RenderWorkerRequest[] = [];
    const document = Object.create(DocxDocument.prototype) as DocxDocument;
    Object.assign(document, {
      _mode: 'worker',
      _document: null,
      _meta: null,
      _bridge: {
        request: async (factory: (id: number) => RenderWorkerRequest) => {
          requests.push(factory(7));
          return { type: 'mainThreadVerticalFallback', id: 7, ...identity };
        },
        transport: () => fallbackTransport,
      },
    });
    attachDocumentLayoutRuntime(document, 10);

    await (document as unknown as {
      _parse(buffer: ArrayBuffer, max?: number, google?: boolean): Promise<void>;
    })._parse(new ArrayBuffer(1), undefined, false);

    expect(requests[0]?.type).toBe('parse');
    expect(document.mode).toBe('main');
    expect((document as unknown as { _meta: unknown })._meta).toBeNull();
    expect(document.document.section.textDirection).toBe('tbRl');
  });

  it('dispatches variant page validation to the worker for render and collect', async () => {
    const requests: RenderWorkerRequest[] = [];
    const bitmap = {} as ImageBitmap;
    const document = Object.create(DocxDocument.prototype) as DocxDocument;
    Object.assign(document, {
      _mode: 'worker',
      _meta: {
        pageCount: 1, revisions: [], comments: [], footnotes: [], endnotes: [],
        pageSizes: [{ widthPt: 612, heightPt: 792 }], bookmarkPages: [],
      },
      _bridge: {
        request: async (factory: (id: number) => RenderWorkerRequest) => {
          const request = factory(9);
          requests.push(request);
          return request.type === 'renderPage'
            ? { type: 'pageRendered', id: 9, bitmap, runs: [] }
            : { type: 'runsCollected', id: 9, runs: [] };
        },
      },
    });
    // Real instances always carry the layout runtime (the constructor attaches
    // it); the active-view fill-in reads it on every render/collect call.
    attachDocumentLayoutRuntime(document, 0);

    await expect(document.renderPageToBitmap(1, { currentDate: 20 })).resolves.toBe(bitmap);
    await expect(document.collectPageRuns(1, { currentDate: 20 })).resolves.toEqual([]);
    expect(requests.map(({ type }) => type)).toEqual(['renderPage', 'collectRuns']);
    expect(requests.map((request) => 'pageIndex' in request ? request.pageIndex : null))
      .toEqual([1, 1]);
  });

  it('uses the real canonical builder with equal main and worker service snapshots', () => {
    const model = realModel();
    const mainServices = createLayoutServices(model, { measureContext: measureContext() });
    const workerServices = createLayoutServices(model, { measureContext: measureContext() });
    attachDocumentLayoutVariants({
      source: layoutSourceStore(model), services: mainServices, defaultCurrentDateMs: 10,
      buildLayout: (options) => layoutDocument(model, mainServices, options),
    });
    attachDocumentLayoutVariants({
      source: layoutSourceStore(model), services: workerServices, defaultCurrentDateMs: 10,
      buildLayout: (options) => layoutDocument(model, workerServices, options),
    });

    const main = selectDocumentLayoutPage(mainServices, {
      currentDate: undefined, defaultCurrentDateMs: 10,
    }, 0);
    const worker = selectDocumentLayoutPage(workerServices, {
      currentDate: undefined, defaultCurrentDateMs: 10,
    }, 0);

    expect(worker.key).toBe(main.key);
    expect(worker.options).toEqual(main.options);
    expect(worker.layout.pages.length).toBe(main.layout.pages.length);
    expect(worker.page.geometry).toEqual(main.page.geometry);
    expect(workerServices.text.fingerprint).toBe(mainServices.text.fingerprint);
    expect(workerServices.images.fingerprint).toBe(mainServices.images.fingerprint);
    expect(workerServices.math.fingerprint).toBe(mainServices.math.fingerprint);
  });

  it('projects equal main and worker run arrays from equal retained variants', () => {
    const model = realModel();
    const mainServices = createLayoutServices(model, { measureContext: measureContext() });
    const workerServices = createLayoutServices(model, { measureContext: measureContext() });
    attachDocumentLayoutVariants({
      source: layoutSourceStore(model), services: mainServices, defaultCurrentDateMs: 10,
      buildLayout: (options) => layoutDocument(model, mainServices, options),
    });
    retainRenderWorkerDocumentLayout(layoutSourceStore(model), workerServices, 10);

    const options = {
      currentDate: Date.UTC(2222, 0, 1),
      defaultCurrentDateMs: 10,
      width: 816,
    };
    expect(textRunsForSelectedPage(workerServices, 0, options)).toEqual(
      textRunsForSelectedPage(mainServices, 0, options),
    );
  });

  it('collects main-thread runs without constructing any Canvas', async () => {
    const model = realModel();
    const layoutServices = createLayoutServices(model, {
      measureContext: measureContext(),
    });
    attachDocumentLayoutVariants({
      source: layoutSourceStore(model),
      services: layoutServices,
      defaultCurrentDateMs: 10,
      buildLayout: (options) => layoutDocument(model, layoutServices, options),
    });
    const document = Object.create(DocxDocument.prototype) as DocxDocument;
    attachDocumentLayoutRuntime(document, 10);
    documentLayoutRuntimeOf(document).services = layoutServices;
    Object.assign(document, { _mode: 'main' });
    const offscreen = vi.fn(() => {
      throw new Error('collectPageRuns must not construct OffscreenCanvas');
    });
    const createElement = vi.fn(() => {
      throw new Error('collectPageRuns must not construct HTMLCanvasElement');
    });
    vi.stubGlobal('OffscreenCanvas', offscreen);
    vi.stubGlobal('document', { createElement });
    try {
      const runs = await document.collectPageRuns(0, {
        width: 816,
        currentDate: Date.UTC(2222, 0, 1),
      });
      expect(runs.map((run) => run.text).join('')).toBe('2222');
      expect(offscreen).not.toHaveBeenCalled();
      expect(createElement).not.toHaveBeenCalled();
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('retains complete default and dated layouts through the production worker wiring seam', () => {
    const model = realModel();
    const mainServices = createLayoutServices(model, { measureContext: measureContext() });
    const mainVariants = attachDocumentLayoutVariants({
      source: layoutSourceStore(model),
      services: mainServices,
      defaultCurrentDateMs: 10,
      buildLayout: (options) => layoutDocument(model, mainServices, options),
    });
    const workerServices = createLayoutServices(model, { measureContext: measureContext() });
    const workerState = retainRenderWorkerDocumentLayout(
      layoutSourceStore(model),
      workerServices,
      10,
    );

    expect(Object.keys(workerState).sort()).toEqual([
      'defaultCurrentDateMs',
      'layoutServices',
      'layoutVariants',
    ]);
    expect(workerState.layoutServices).toBe(workerServices);
    expect(layoutVariantStoreOf(workerServices)).toBe(workerState.layoutVariants);

    const fingerprints = (currentDate: number | undefined) => {
      const mainInput = { currentDate, defaultCurrentDateMs: 10 };
      const workerInput = {
        currentDate,
        defaultCurrentDateMs: workerState.defaultCurrentDateMs,
      };
      const main = selectDocumentLayoutPage(mainServices, mainInput, 0);
      const worker = selectDocumentLayoutPage(
        workerState.layoutServices,
        workerInput,
        0,
      );
      expect(worker.key).toBe(main.key);
      expect(worker.options).toEqual(main.options);
      expect(worker.layout.pages[0]?.layers.body.length).toBeGreaterThan(0);
      return {
        key: worker.key,
        main: stableFingerprint('document-layout', main.layout),
        worker: stableFingerprint('document-layout', worker.layout),
      };
    };

    const defaultVariant = fingerprints(undefined);
    const datedVariant = fingerprints(Date.UTC(2222, 0, 1));
    expect(selectDocumentLayoutPage(workerState.layoutServices, {
      currentDate: undefined,
      defaultCurrentDateMs: workerState.defaultCurrentDateMs,
    }, 0).layout).toBe(workerState.layoutVariants.defaultLayout);
    expect(stableFingerprint(
      'document-metadata',
      metadataForDefaultLayout(model, workerState.layoutVariants.defaultLayout),
    )).toBe(stableFingerprint(
      'document-metadata',
      metadataForDefaultLayout(model, mainVariants.store.defaultLayout),
    ));
    expect(defaultVariant.worker).toBe(defaultVariant.main);
    expect(datedVariant.worker).toBe(datedVariant.main);
    expect(datedVariant.key).not.toBe(defaultVariant.key);
  });

  it('retains parser-private final RTL nextColumn layout identically in main and worker', () => {
    const model = finalNextColumnRtlModel();
    const mainServices = createLayoutServices(model, { measureContext: measureContext() });
    attachDocumentLayoutVariants({
      source: layoutSourceStore(model),
      services: mainServices,
      defaultCurrentDateMs: 10,
      buildLayout: (options) => layoutDocument(model, mainServices, options),
    });
    const workerState = retainRenderWorkerDocumentLayout(
      layoutSourceStore(model),
      createLayoutServices(model, { measureContext: measureContext() }),
      10,
    );
    const main = selectDocumentLayoutPage(mainServices, {
      currentDate: undefined,
      defaultCurrentDateMs: 10,
    }, 0);
    const worker = selectDocumentLayoutPage(workerState.layoutServices, {
      currentDate: undefined,
      defaultCurrentDateMs: workerState.defaultCurrentDateMs,
    }, 0);

    expect(stableFingerprint('document-layout', worker.layout)).toBe(
      stableFingerprint('document-layout', main.layout),
    );
    expect(main.layout.pages).toHaveLength(1);
    expect(main.page.sectionRegions.map((region) => ({
      section: region.sectionOccurrenceId,
      direction: region.columnFlowDirection,
      columns: region.columnIndexes,
    }))).toEqual([
      { section: 'section:outgoing', direction: 'rtl', columns: [1] },
      { section: 'section:final', direction: 'rtl', columns: [0] },
    ]);
  });

  it('selects equal keys, fingerprints, page counts, sizes, and variants from equal normalized inputs', () => {
    const model = realModel();
    const mainServices = services();
    const workerServices = services();
    let mainBuilds = 0;
    let workerBuilds = 0;
    attachDocumentLayoutVariants({
      source: layoutSourceStore(model), services: mainServices, defaultCurrentDateMs: 10,
      buildLayout: (options) => { mainBuilds += 1; return layout(options.currentDateMs); },
    });
    attachDocumentLayoutVariants({
      source: layoutSourceStore(model), services: workerServices, defaultCurrentDateMs: 10,
      buildLayout: (options) => { workerBuilds += 1; return layout(options.currentDateMs); },
    });

    const main = selectDocumentLayoutPage(mainServices, {
      currentDate: new Date(20), defaultCurrentDateMs: 10,
    }, 1);
    const worker = selectDocumentLayoutPage(workerServices, {
      currentDate: 20, defaultCurrentDateMs: 10,
    }, 1);

    expect(worker.options).toEqual(main.options);
    expect(worker.key).toBe(main.key);
    expect(worker.layout.pages.length).toBe(main.layout.pages.length);
    expect(worker.page.geometry).toEqual(main.page.geometry);
    expect((worker.page as unknown as { variant: string }).variant).toBe(
      (main.page as unknown as { variant: string }).variant,
    );
    expect(main.layout).not.toBe(worker.layout);
    expect([mainBuilds, workerBuilds]).toEqual([1, 1]);

    expect(selectDocumentLayoutPage(mainServices, {
      currentDate: 10, defaultCurrentDateMs: 10,
    }, 0).key).not.toBe(main.key);
    expect([mainBuilds, workerBuilds]).toEqual([2, 1]);
  });

  it('pins worker request and response protocol shapes including structured failures', () => {
    const parse = {
      type: 'parse', id: 1, data: new ArrayBuffer(0), useGoogleFonts: false,
      resourcePolicy: DEFAULT_OOXML_RESOURCE_LIMITS,
      defaultCurrentDateMs: 10,
    } satisfies RenderWorkerRequest;
    const render = {
      type: 'renderPage', id: 2, pageIndex: 0, opts: { currentDate: 20 },
    } satisfies RenderWorkerRequest;
    const collect = {
      type: 'collectRuns', id: 3, pageIndex: 0, opts: { currentDate: 20 },
    } satisfies RenderWorkerRequest;
    const layoutMetaRequest = {
      type: 'selectLayoutView', id: 6, currentDateMs: 20, showTrackedChanges: true,
    } satisfies RenderWorkerRequest;
    const parsed = {
      type: 'parsedMeta', id: 1,
      meta: { pageCount: 1, revisions: [], comments: [], footnotes: [], endnotes: [], pageSizes: [], bookmarkPages: [] },
    } satisfies RenderWorkerResponse;
    const verticalFallback = {
      type: 'mainThreadVerticalFallback', id: 1,
      sessionId: 1, operationId: 1, generation: 1,
    } satisfies RenderWorkerResponse;
    const pageRendered = {
      type: 'pageRendered', id: 2, bitmap: {} as ImageBitmap, runs: [],
    } satisfies RenderWorkerResponse;
    const runsCollected = {
      type: 'runsCollected', id: 3, runs: [],
    } satisfies RenderWorkerResponse;
    const layoutMetaResponse = {
      type: 'layoutViewSelected', id: 6,
      meta: {
        pageCount: 1,
        pageSizes: [{ widthPt: 595, heightPt: 842 }],
        bookmarkPages: [],
        commentAnchorRanges: [],
        revisionAnchorRanges: [],
      },
    } satisfies RenderWorkerResponse;
    const progressiveParse = {
      type: 'parse', id: 5, data: new ArrayBuffer(0), useGoogleFonts: false,
      resourcePolicy: DEFAULT_OOXML_RESOURCE_LIMITS,
      defaultCurrentDateMs: 10,
      currentDateMs: 20,
      showTrackedChanges: true,
      progressiveLayout: true,
    } satisfies RenderWorkerRequest;
    // Uncorrelated pushes: `forId`, never `id`. The bridge resolves a pending
    // request on the first response `correlate` matches, so an `id` here would
    // settle the parse before the authoritative metadata exists.
    const layoutPartial = {
      type: 'layoutPartial', forId: 5,
      partial: {
        pageCount: 2,
        pageSizes: [{ widthPt: 595, heightPt: 842 }, { widthPt: 595, heightPt: 842 }],
        bookmarkPages: [],
        commentAnchorRanges: [],
        revisionAnchorRanges: [],
        exact: false,
        review: { revisions: [], comments: [], footnotes: [], endnotes: [] },
      },
    } satisfies RenderWorkerResponse;
    const layoutProgress = {
      type: 'layoutProgress', forId: 5, committedPages: 12,
    } satisfies RenderWorkerResponse;
    const error = {
      type: 'error', id: 4, message: 'unsupported transition',
      errorName: 'UnsupportedPageFlowTransitionError',
      code: 'NEXT_COLUMN_DESTINATION_UNAVAILABLE',
      reason: 'physical-column',
      outgoingColumnIndex: 0,
      outgoingColumnCount: 3,
      incomingColumnCount: 1,
    } satisfies RenderWorkerResponse;

    expect(Object.keys(parse).sort()).toEqual([
      'data', 'defaultCurrentDateMs', 'id', 'resourcePolicy', 'type', 'useGoogleFonts',
    ]);
    expect(Object.keys(render).sort()).toEqual(['id', 'opts', 'pageIndex', 'type']);
    expect(Object.keys(collect).sort()).toEqual(['id', 'opts', 'pageIndex', 'type']);
    expect(Object.keys(layoutMetaRequest).sort()).toEqual([
      'currentDateMs', 'id', 'showTrackedChanges', 'type',
    ]);
    expect(Object.keys(parsed).sort()).toEqual(['id', 'meta', 'type']);
    expect(Object.keys(verticalFallback).sort()).toEqual([
      'generation', 'id', 'operationId', 'sessionId', 'type',
    ]);
    expect(Object.keys(pageRendered).sort()).toEqual(['bitmap', 'id', 'runs', 'type']);
    expect(Object.keys(runsCollected).sort()).toEqual(['id', 'runs', 'type']);
    expect(Object.keys(layoutMetaResponse).sort()).toEqual(['id', 'meta', 'type']);
    expect(Object.keys(progressiveParse).sort()).toEqual([
      'currentDateMs', 'data', 'defaultCurrentDateMs', 'id', 'progressiveLayout',
      'resourcePolicy', 'showTrackedChanges', 'type', 'useGoogleFonts',
    ]);
    expect(Object.keys(layoutPartial).sort()).toEqual(['forId', 'partial', 'type']);
    expect(Object.keys(layoutPartial.partial).sort()).toEqual([
      'bookmarkPages', 'commentAnchorRanges', 'exact', 'pageCount', 'pageSizes',
      'review', 'revisionAnchorRanges',
    ]);
    expect(Object.keys(layoutProgress).sort()).toEqual(['committedPages', 'forId', 'type']);
    // The push arms must not be correlatable, or they would settle the parse.
    expect('id' in layoutPartial).toBe(false);
    expect('id' in layoutProgress).toBe(false);
    expect(Object.keys(error).sort()).toEqual([
      'code',
      'errorName',
      'id',
      'incomingColumnCount',
      'message',
      'outgoingColumnCount',
      'outgoingColumnIndex',
      'reason',
      'type',
    ]);
  });
});
