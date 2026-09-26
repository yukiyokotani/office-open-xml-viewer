import { beforeEach, describe, expect, it, vi } from 'vitest';
import type {
  ModelSourceModuleDescriptor,
  OpenedModelSourceModule,
  WasmParserHost,
} from '@silurus/ooxml-core';
import type { DocxDocumentModel } from './types.js';
import {
  createLocalDocumentPullTransport,
  DocumentPullWorker,
  MaterializedDocumentCursorArchive,
} from './document-pull-worker.js';
import { materializeDocumentPullOwnedModelsSession } from './document-pull-client.js';
import {
  WorkerDocumentSourceOwner,
  type DocxModelSourceArchive,
  type OoxmlWorkerDocumentArchive,
} from './internal/worker-document-source.js';

type Opened = OpenedModelSourceModule<DocxModelSourceArchive>;

const workerMocks = vi.hoisted(() => ({
  opens: [] as Array<(...args: unknown[]) => Promise<unknown>>,
  parserInit: vi.fn(),
  metrics: vi.fn(async (): Promise<{ faces: FontFace[]; routes: Record<string, never> }> => (
    { faces: [], routes: {} }
  )),
  unloadLocal: vi.fn(),
  paginate: vi.fn(async () => undefined),
  verticalFallback: vi.fn(() => false),
}));
vi.mock('@silurus/ooxml-core', async importOriginal => ({
  ...await importOriginal<typeof import('@silurus/ooxml-core')>(),
  openModelSourceModule: (...args: unknown[]) => workerMocks.opens.at(-1)!(...args),
  loadOfficeFontFallbacks: workerMocks.metrics,
  unloadOfficeFontFallbacks: workerMocks.unloadLocal,
}));
vi.mock('./wasm/docx_parser.js', () => ({
  default: workerMocks.parserInit, reinit: vi.fn(), DocxArchive: class {},
}));
vi.mock('./google-fonts.js', () => ({
  DOCX_GOOGLE_FONTS: {}, docxFontPreloadNames: () => [], docxOfficeFontFallbackRequests: () => [],
}));
vi.mock('./embedded-fonts.js', () => ({
  loadEmbeddedFonts: async () => ({ faces: [], metrics: {}, routes: [] }),
}));
vi.mock('./renderer.js', () => ({ prepareMathRuns: vi.fn(), renderLayoutSourceToCanvas: vi.fn() }));
vi.mock('./vertical-render-capability.js', () => ({
  documentRequiresDomVerticalGlyphLayout: workerMocks.verticalFallback,
}));
vi.mock('./layout-source-model-adapter.js', () => ({
  layoutSourceModelAdapterFromOwnedModel: (document: DocxDocumentModel) => ({
    document, source: { fatalParse: null, mathOccurrences: [] },
  }),
}));
vi.mock('./layout-runtime.js', () => ({ createLayoutServices: () => ({}) }));
vi.mock('./render-worker-layout.js', () => ({
  retainRenderWorkerDocumentLayout: () => ({
    defaultCurrentDateMs: 0, layoutServices: {},
    layoutVariants: { layoutFor: () => ({ pages: [] }) },
  }),
}));
vi.mock('./render-worker-progressive.js', () => ({
  paginateRenderWorkerDocumentProgressively: workerMocks.paginate,
}));
vi.mock('./render-worker-metadata.js', () => ({
  projectRenderWorkerLayoutMeta: () => ({ pageCount: 0, pageSizes: [], bookmarkPages: [] }),
  renderWorkerLayoutMeta: () => ({ pageCount: 0, pageSizes: [], bookmarkPages: [] }),
}));

const descriptor: ModelSourceModuleDescriptor = {
  protocol: 'ooxml-model-source-module/v1',
  target: 'docx',
  moduleUrl: 'https://example.test/source.mjs',
  config: {},
};

type Dispatch = (event: MessageEvent) => Promise<void>;

async function loadWorker(): Promise<{ dispatch: Dispatch; posts: ReturnType<typeof vi.fn> }> {
  vi.resetModules();
  const posts = vi.fn();
  Object.assign(globalThis, { self: { postMessage: posts, onmessage: null } });
  await import('./render-worker.js');
  const dispatch = (globalThis.self as unknown as { onmessage: Dispatch }).onmessage;
  return { dispatch, posts };
}

function posted(posts: ReturnType<typeof vi.fn>, type: string, id: number): Record<string, unknown>[] {
  return posts.mock.calls
    .map(([message]) => message as Record<string, unknown>)
    .filter((message) => message.type === type && (message.id ?? message.forId) === id);
}

beforeEach(() => {
  workerMocks.opens.length = 0;
  workerMocks.parserInit.mockClear();
  workerMocks.metrics.mockReset();
  workerMocks.metrics.mockResolvedValue({ faces: [], routes: {} });
  workerMocks.unloadLocal.mockClear();
  workerMocks.paginate.mockReset();
  workerMocks.paginate.mockResolvedValue(undefined);
  workerMocks.verticalFallback.mockReset();
  workerMocks.verticalFallback.mockReturnValue(false);
});

describe('DOCX render worker with a model source', () => {
  it('drains a model-source archive through the real pull consumer and keeps images readable', async () => {
    const source = modelSource({
      headers: { default: { body: [{ type: 'paragraph', runs: [] }] }, first: null, even: null },
    });
    const host = {
      archive: null, run: vi.fn(), disposeArchive: vi.fn(), ensureReady: vi.fn(), setWasmInput: vi.fn(),
    } as unknown as WasmParserHost<OoxmlWorkerDocumentArchive>;
    const owner = new WorkerDocumentSourceOwner(host, async () => source.opened);
    await owner.openModelSource(new Uint8Array([1, 2, 3]), descriptor);
    const pull = new DocumentPullWorker(() => owner.cursor(), operation => owner.execute(operation));
    const identity = { sessionId: 1, operationId: 1, generation: 1 };
    pull.open(identity);
    const pulled = await materializeDocumentPullOwnedModelsSession(
      createLocalDocumentPullTransport(pull), identity,
    );
    await pull.reset();

    expect(pulled.document.body.map(element => element.type)).toEqual(['pageBreak']);
    expect(pulled.document.headers.default?.body).toHaveLength(1);
    expect(owner.execute(archive => archive.extract_image('media/1'))).toEqual(new Uint8Array([1, 2, 3]));
    expect(host.ensureReady).not.toHaveBeenCalled();
    expect(host.run).not.toHaveBeenCalled();
    expect(source.close).not.toHaveBeenCalled();
  });

  it('parses without the OOXML runtime, applies the source view default and degrades missing capabilities', async () => {
    const { dispatch, posts } = await loadWorker();
    const source = modelSource({}, { showTrackedChanges: true });
    const open = vi.fn(async () => source.opened);
    workerMocks.opens.push(open);
    const transfer = [new ArrayBuffer(2)];

    // No `init` message: a model-source parse must not need the OOXML WASM input.
    await dispatch({ data: { ...parseRequest(11), sourceTransfer: transfer } } as MessageEvent);
    expect(open).toHaveBeenCalledExactlyOnceWith(
      descriptor, new Uint8Array(1), expect.any(Function), undefined, transfer,
    );
    expect(posted(posts, 'parsedMeta', 11)).toEqual([
      expect.objectContaining({ usage: undefined, showTrackedChanges: true }),
    ]);
    expect(workerMocks.parserInit).not.toHaveBeenCalled();

    await dispatch({ data: { type: 'extractImage', id: 12, path: 'media/1' } } as MessageEvent);
    expect(posts).toHaveBeenCalledWith(
      expect.objectContaining({ type: 'imageExtracted', id: 12 }), expect.any(Array),
    );
    await dispatch({ data: { type: 'resourceUsage', id: 13 } } as MessageEvent);
    expect(posted(posts, 'resourceUsage', 13)).toEqual([expect.objectContaining({ usage: undefined })]);
    await dispatch({ data: { type: 'toMarkdown', id: 14 } } as MessageEvent);
    expect(posted(posts, 'error', 14)).toEqual([expect.objectContaining({
      message: 'Markdown conversion is unsupported for this source',
    })]);
    expect(source.close).not.toHaveBeenCalled();
  });

  it('lets an explicit caller view beat the source view default in both published views', async () => {
    const { dispatch, posts } = await loadWorker();
    workerMocks.opens.push(async () => modelSource({}, { showTrackedChanges: true }).opened);
    workerMocks.paginate.mockImplementationOnce(async (...args: unknown[]) => {
      const sink = args[2] as { publish(publication: unknown): void };
      sink.publish({ pageCount: 1 });
    });
    await dispatch({
      data: { ...parseRequest(15), showTrackedChanges: false, progressiveLayout: true },
    } as MessageEvent);
    expect(posted(posts, 'layoutPartial', 15)).toEqual([
      expect.objectContaining({ showTrackedChanges: false }),
    ]);
    expect(posted(posts, 'parsedMeta', 15)).toEqual([
      expect.objectContaining({ showTrackedChanges: false }),
    ]);
  });

  it('forwards validated source view defaults with the main-thread vertical fallback', async () => {
    const { dispatch, posts } = await loadWorker();
    workerMocks.verticalFallback.mockReturnValue(true);
    workerMocks.opens.push(async () => modelSource({}, { showTrackedChanges: true }).opened);
    await dispatch({ data: parseRequest(16) } as MessageEvent);
    expect(posted(posts, 'mainThreadVerticalFallback', 16)).toEqual([
      expect.objectContaining({ viewDefaults: { showTrackedChanges: true } }),
    ]);
  });

  it('does not let a deferred old parse publish or close the newer source', async () => {
    const { dispatch, posts } = await loadWorker();
    let release!: () => void;
    const staleFace = {} as FontFace;
    const oldMetrics = new Promise<{ faces: FontFace[]; routes: Record<string, never> }>(resolve => {
      release = () => resolve({ faces: [staleFace], routes: {} });
    });
    let metricCall = 0;
    workerMocks.metrics.mockImplementation(() => metricCall++ === 0
      ? oldMetrics
      : Promise.resolve({ faces: [], routes: {} }));
    const old = modelSource(); const current = modelSource();
    workerMocks.opens.push(vi.fn().mockResolvedValueOnce(old.opened).mockResolvedValueOnce(current.opened));
    const stale = dispatch({ data: parseRequest(20) } as MessageEvent);
    await vi.waitFor(() => expect(metricCall).toBe(1));
    await dispatch({ data: parseRequest(21) } as MessageEvent);
    release();
    await stale;

    expect(posted(posts, 'parsedMeta', 21)).toHaveLength(1);
    expect(posted(posts, 'parsedMeta', 20)).toHaveLength(0);
    expect(posted(posts, 'error', 20)).toHaveLength(1);
    expect(old.close).toHaveBeenCalledOnce();
    expect(current.close).not.toHaveBeenCalled();
    expect(workerMocks.unloadLocal).toHaveBeenCalledWith([staleFace]);
  });

  it('releases request fonts and closes the source when pagination fails', async () => {
    const { dispatch, posts } = await loadWorker();
    const face = {} as FontFace;
    workerMocks.metrics.mockResolvedValue({ faces: [face], routes: {} });
    workerMocks.paginate.mockRejectedValueOnce(new Error('pagination failed'));
    const source = modelSource();
    workerMocks.opens.push(async () => source.opened);
    await dispatch({ data: { ...parseRequest(30), progressiveLayout: true } } as MessageEvent);

    expect(posted(posts, 'error', 30)).toEqual([
      expect.objectContaining({ message: 'pagination failed' }),
    ]);
    expect(workerMocks.unloadLocal).toHaveBeenCalledWith([face]);
    expect(source.close).toHaveBeenCalledOnce();
  });
});

function modelSource(
  overrides: Partial<Record<keyof DocxDocumentModel, unknown>> = {},
  viewDefaults: Readonly<Record<string, boolean>> = {},
): { close: ReturnType<typeof vi.fn>; opened: Opened } {
  const model = {
    section: {}, body: [{ type: 'pageBreak' }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    ...overrides,
  } as unknown as DocxDocumentModel;
  const archive = Object.assign(new MaterializedDocumentCursorArchive(model), {
    assert_healthy: vi.fn(),
    extract_image: vi.fn(() => new Uint8Array([1, 2, 3])),
  }) as unknown as DocxModelSourceArchive;
  const close = vi.fn();
  return { close, opened: { archive, viewDefaults, close } };
}

function parseRequest(id: number) {
  return {
    type: 'parse' as const, id, data: new ArrayBuffer(1), source: descriptor,
    resourcePolicy: {}, useGoogleFonts: false, defaultCurrentDateMs: 0,
  };
}
