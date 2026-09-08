import { describe, expect, it, vi } from 'vitest';
import type { WasmParserHost } from '@silurus/ooxml-core';
import type { OwnedLegacyDocSource } from '@silurus/ooxml-legacy-converter/internal/direct-doc-engine';
import type { DocxDocumentModel } from './types.js';
import {
  createLocalDocumentPullTransport,
  DocumentPullWorker,
  MaterializedDocumentCursorArchive,
} from './document-pull-worker.js';
import { materializeDocumentPullOwnedModelsSession } from './document-pull-client.js';
import {
  WorkerDocumentSourceOwner,
  type OoxmlWorkerDocumentArchive,
} from './internal/worker-document-source.js';

const workerMocks = vi.hoisted(() => ({
  opens: [] as Array<(...args: unknown[]) => Promise<OwnedLegacyDocSource>>,
  metrics: vi.fn(async (): Promise<{ faces: FontFace[]; metrics: undefined }> => (
    { faces: [], metrics: undefined }
  )),
  unloadLocal: vi.fn(),
  paginate: vi.fn(async () => undefined),
}));
vi.mock('@silurus/ooxml-core', async importOriginal => ({
  ...await importOriginal<typeof import('@silurus/ooxml-core')>(),
  unloadLocalFontMetrics: workerMocks.unloadLocal,
}));
vi.mock('./wasm/docx_parser.js', () => ({ default: vi.fn(), reinit: vi.fn(), DocxArchive: class {} }));
vi.mock('@silurus/ooxml-legacy-converter/internal/direct-doc-engine', () => ({
  openLegacyDocSource: (...args: unknown[]) => workerMocks.opens.at(-1)!(...args),
}));
vi.mock('./local-font-metrics.js', () => ({ loadDocxLocalFontMetrics: workerMocks.metrics }));
vi.mock('./google-fonts.js', () => ({ DOCX_GOOGLE_FONTS: {}, docxFontPreloadNames: () => [] }));
vi.mock('./embedded-fonts.js', () => ({ loadEmbeddedFonts: async () => [] }));
vi.mock('./renderer.js', () => ({ prepareMathRuns: vi.fn(), renderLayoutSourceToCanvas: vi.fn() }));
vi.mock('./vertical-render-capability.js', () => ({ documentRequiresDomVerticalGlyphLayout: () => false }));
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

const descriptor = {
  protocol: 'ooxml-legacy-doc-source/v1' as const,
  builtin: 'doc' as const,
  wasmUrl: 'https://example.test/direct-doc.wasm',
};

describe('render-worker direct DOC source route', () => {
  it('drains the native archive through the real pull consumer and retains images', async () => {
    const model = {
      section: {}, body: [{ type: 'pageBreak' }],
      headers: { default: { body: [{ type: 'paragraph', runs: [] }] }, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as unknown as DocxDocumentModel;
    const cursor = new MaterializedDocumentCursorArchive(model);
    const native = Object.assign(cursor, {
      free: vi.fn(), assert_healthy: vi.fn(),
      extract_image: vi.fn(() => new Uint8Array([7, 8, 9])),
    });
    const closeArchive = vi.fn();
    const open = vi.fn(async (): Promise<OwnedLegacyDocSource> => ({
      archive: native, sourceByteLength: 3, closeArchive,
    }));
    const host = {
      archive: null, run: vi.fn(), disposeArchive: vi.fn(), ensureReady: vi.fn(), setWasmInput: vi.fn(),
    } as unknown as WasmParserHost<OoxmlWorkerDocumentArchive>;
    const owner = new WorkerDocumentSourceOwner(host, open);
    await owner.openNative(new Uint8Array([1, 2, 3]), descriptor);
    const pull = new DocumentPullWorker(() => owner.cursor(), operation => owner.execute(operation));
    const identity = { sessionId: 1, operationId: 1, generation: 1 };
    pull.open(identity);
    const pulled = await materializeDocumentPullOwnedModelsSession(
      createLocalDocumentPullTransport(pull), identity,
    );
    await pull.reset();

    expect(pulled.document.body.map(element => element.type)).toEqual(['pageBreak']);
    expect(pulled.document.headers.default?.body).toHaveLength(1);
    expect(owner.execute(archive => archive.extract_image('legacy-doc/image/1')))
      .toEqual(new Uint8Array([7, 8, 9]));
    expect(host.ensureReady).not.toHaveBeenCalled();
    expect(host.run).not.toHaveBeenCalled();
    expect(closeArchive).not.toHaveBeenCalled();
    owner.closeNative(); owner.closeNative();
    expect(closeArchive).toHaveBeenCalledOnce();
  });

  it('dispatches the actual worker native route and keeps images after terminal ACK', async () => {
    vi.resetModules();
    const posts = vi.fn();
    Object.assign(globalThis, { self: { postMessage: posts, onmessage: null } });
    const direct = nativeSource();
    workerMocks.opens.push(vi.fn(async () => direct.owned));
    await import('./render-worker.js');
    const dispatch = (globalThis.self as unknown as { onmessage(e: MessageEvent): Promise<void> }).onmessage;
    await dispatch({ data: parseRequest(11) } as MessageEvent);
    expect(posts).toHaveBeenCalledWith(expect.objectContaining({ type: 'parsedMeta', id: 11, usage: undefined }), undefined);
    await dispatch({ data: { type: 'extractImage', id: 12, path: 'legacy-doc/image/1' } } as MessageEvent);
    expect(posts).toHaveBeenCalledWith(
      expect.objectContaining({ type: 'imageExtracted', id: 12 }), expect.any(Array),
    );
    await dispatch({ data: { type: 'resourceUsage', id: 13 } } as MessageEvent);
    await dispatch({ data: { type: 'toMarkdown', id: 14 } } as MessageEvent);
    expect(posts).toHaveBeenCalledWith(expect.objectContaining({
      type: 'error', id: 13, message: expect.stringContaining('unsupported for direct legacy DOC'),
    }), undefined);
    expect(posts).toHaveBeenCalledWith(expect.objectContaining({
      type: 'error', id: 14, message: expect.stringContaining('unsupported for direct legacy DOC'),
    }), undefined);
    expect(direct.close).not.toHaveBeenCalled();
  });

  it('does not let a deferred old parse publish or close the newer native owner', async () => {
    vi.resetModules();
    workerMocks.metrics.mockClear();
    const posts = vi.fn();
    Object.assign(globalThis, { self: { postMessage: posts, onmessage: null } });
    let release!: () => void;
    const staleFace = {} as FontFace;
    const oldMetrics = new Promise<{ faces: FontFace[]; metrics: undefined }>(resolve => {
      release = () => resolve({ faces: [staleFace], metrics: undefined });
    });
    let metricCall = 0;
    workerMocks.metrics.mockImplementation(() => metricCall++ === 0
      ? oldMetrics
      : Promise.resolve({ faces: [], metrics: undefined }));
    const old = nativeSource(); const current = nativeSource();
    workerMocks.opens.push(vi.fn().mockResolvedValueOnce(old.owned).mockResolvedValueOnce(current.owned));
    await import('./render-worker.js');
    const dispatch = (globalThis.self as unknown as { onmessage(e: MessageEvent): Promise<void> }).onmessage;
    const stale = dispatch({ data: parseRequest(20) } as MessageEvent);
    await vi.waitFor(() => expect(metricCall).toBe(1));
    await dispatch({ data: parseRequest(21) } as MessageEvent);
    release();
    await stale;
    expect(posts).toHaveBeenCalledWith(expect.objectContaining({ type: 'parsedMeta', id: 21, usage: undefined }), undefined);
    expect(posts.mock.calls.filter(([message]) => (
      message as { type?: string; id?: number }
    ).type === 'parsedMeta' && (message as { id?: number }).id === 20)).toHaveLength(0);
    expect(posts).toHaveBeenCalledWith(expect.objectContaining({ type: 'error', id: 20 }), undefined);
    expect(old.close).toHaveBeenCalledOnce();
    expect(current.close).not.toHaveBeenCalled();
    expect(workerMocks.unloadLocal).toHaveBeenCalledWith([staleFace]);
  });

  it('releases transferred request fonts when later pagination fails', async () => {
    vi.resetModules();
    workerMocks.metrics.mockReset();
    workerMocks.unloadLocal.mockClear();
    workerMocks.paginate.mockReset();
    const face = {} as FontFace;
    workerMocks.metrics.mockResolvedValue({ faces: [face], metrics: undefined });
    workerMocks.paginate.mockRejectedValueOnce(new Error('pagination failed'));
    const posts = vi.fn();
    Object.assign(globalThis, { self: { postMessage: posts, onmessage: null } });
    const direct = nativeSource();
    workerMocks.opens.push(vi.fn(async () => direct.owned));
    await import('./render-worker.js');
    const dispatch = (globalThis.self as unknown as { onmessage(e: MessageEvent): Promise<void> }).onmessage;
    await dispatch({ data: { ...parseRequest(30), progressiveLayout: true } } as MessageEvent);

    expect(posts).toHaveBeenCalledWith(expect.objectContaining({
      type: 'error', id: 30, message: 'pagination failed',
    }), undefined);
    expect(workerMocks.unloadLocal).toHaveBeenCalledWith([face]);
    expect(direct.close).toHaveBeenCalledOnce();
  });
});

function nativeSource() {
  const model = {
    section: {}, body: [{ type: 'pageBreak' }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
  const cursor = new MaterializedDocumentCursorArchive(model);
  const archive = Object.assign(cursor, {
    free: vi.fn(), assert_healthy: vi.fn(),
    extract_image: vi.fn(() => new Uint8Array([1, 2, 3])),
  });
  const close = vi.fn();
  return { close, owned: { archive, sourceByteLength: 1, closeArchive: close } as OwnedLegacyDocSource };
}

function parseRequest(id: number) {
  return {
    type: 'parse' as const, id, data: new ArrayBuffer(1), source: descriptor,
    resourcePolicy: {}, useGoogleFonts: false, defaultCurrentDateMs: 0,
  };
}
