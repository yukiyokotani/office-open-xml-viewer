import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { ModelSource, ModelSourceModuleDescriptor } from '@silurus/ooxml-core';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';
import { DocxDocument } from './document.js';
import { activeDocxLayoutViewOf } from './document-layout-view.js';
import {
  DocumentPullWorker,
  isDocumentPullCommand,
  MaterializedDocumentCursorArchive,
} from './document-pull-worker.js';
import { installStubCanvas, syntheticDocxModel } from './testing/synthetic-document.js';
import type { DocumentMeta } from './worker-protocol.js';

vi.mock('@silurus/ooxml-core', async (load) => ({
  ...await load<typeof import('@silurus/ooxml-core')>(),
  loadOfficeFontFallbacks: async () => ({ faces: [], routes: {} }),
}));

// DocxDocument.load() with LoadOptions.modelSources, driven through the real
// WorkerBridge against a scripted worker that speaks the parse / pull protocol.
// View precedence is one rule for every source: explicit caller option > the
// source's view default > the renderer default (final view).

type Message = Record<string, unknown>;
type Script = (worker: ProtocolWorker, message: Message) => void | Promise<void>;

class ProtocolWorker {
  static instances: ProtocolWorker[] = [];
  static script: Script = () => undefined;
  terminated = false;
  readonly messages: Message[] = [];
  readonly transfers: (Transferable[] | undefined)[] = [];
  private readonly listeners = new Set<(event: MessageEvent) => void>();
  constructor() { ProtocolWorker.instances.push(this); }
  postMessage(message: unknown, transfer?: Transferable[]): void {
    this.messages.push(message as Message);
    this.transfers.push(transfer);
    const script = ProtocolWorker.script;
    queueMicrotask(() => { void script(this, message as Message); });
  }
  addEventListener(type: string, listener: (event: MessageEvent) => void): void {
    if (type === 'message') this.listeners.add(listener);
  }
  removeEventListener(type: string, listener: (event: MessageEvent) => void): void {
    if (type === 'message') this.listeners.delete(listener);
  }
  reply(data: unknown): void {
    for (const listener of [...this.listeners]) listener({ data } as MessageEvent);
  }
  terminate(): void { this.terminated = true; }
  parseRequests(): Message[] { return this.messages.filter((message) => message.type === 'parse'); }
}

const globals = globalThis as Record<string, unknown>;
const originals = { Worker: globals.Worker, location: globals.location };
const descriptor: ModelSourceModuleDescriptor = {
  protocol: 'ooxml-model-source-module/v1',
  target: 'docx',
  moduleUrl: 'https://example.test/source.mjs',
  config: { variant: 'test' },
};
// A CFB container: the OOXML path rejects it before any worker exists, so a
// load that gets past container resolution proves the source claimed it.
const cfbBytes = () => buildCfbFixture(['Root Entry', 'WordDocument', '1Table']);

function fakeSource(options: {
  claim?: (bytes: Uint8Array) => boolean;
  transfer?: Transferable[];
  target?: string;
} = {}) {
  const release = vi.fn();
  const claim = vi.fn(options.claim ?? (() => true));
  const beginLoad = vi.fn(() => ({
    module: descriptor,
    ...(options.transfer ? { transfer: options.transfer } : {}),
    release,
  }));
  const source = { target: options.target ?? 'docx', claim, beginLoad } as unknown as ModelSource<'docx'>;
  return { source, claim, beginLoad, release };
}

function meta(pageCount = 1): DocumentMeta {
  return {
    pageCount,
    revisions: [], comments: [], footnotes: [], endnotes: [],
    pageSizes: Array.from({ length: pageCount }, () => ({ widthPt: 595, heightPt: 842 })),
    bookmarkPages: [], commentAnchorRanges: [], revisionAnchorRanges: [],
  } as unknown as DocumentMeta;
}

/** Render-worker script: echo the effective view the way render-worker.ts does. */
function renderWorkerScript(sourceDefault: boolean | undefined, options: { partial?: boolean } = {}): Script {
  return (worker, message) => {
    if (message.type === 'parse') {
      const effective = (message.showTrackedChanges as boolean | undefined) ?? sourceDefault ?? false;
      if (options.partial) {
        worker.reply({
          type: 'layoutPartial', forId: message.id, showTrackedChanges: effective,
          partial: { ...meta(), exact: false },
        });
      }
      // After a publication the view is already fixed; omit the echo so a
      // partial-driven test observes only the publication's view.
      worker.reply({
        type: 'parsedMeta', id: message.id, meta: meta(),
        ...(options.partial ? {} : { showTrackedChanges: effective }),
      });
    } else if (message.type === 'resourceUsage') {
      worker.reply({ type: 'resourceUsage', id: message.id, usage: undefined });
    }
  };
}

/** Parse-worker script: open a pull session over a materialized model. */
function parseWorkerScript(viewDefaults: Record<string, boolean> | undefined): Script {
  let pull: DocumentPullWorker | undefined;
  return async (worker, message) => {
    if (isDocumentPullCommand(message)) {
      await pull!.dispatch(message, (response) => worker.reply(response));
    } else if (message.type === 'parse') {
      const archive = new MaterializedDocumentCursorArchive(syntheticDocxModel('tracked', { paragraphs: 6 }));
      pull = new DocumentPullWorker(() => archive);
      const identity = { sessionId: 1, operationId: 1, generation: 1 };
      pull.open(identity);
      worker.reply({
        type: 'documentSessionOpened', id: message.id, ...identity,
        ...(viewDefaults ? { viewDefaults } : {}),
      });
    } else if (message.type === 'resourceUsage') {
      worker.reply({ type: 'resourceUsage', id: message.id, usage: undefined });
    }
  };
}

beforeAll(() => {
  // Also satisfies worker mode's OffscreenCanvas feature check.
  installStubCanvas();
});

afterEach(() => {
  globals.Worker = originals.Worker;
  globals.location = originals.location;
  ProtocolWorker.instances = [];
  ProtocolWorker.script = () => undefined;
});

function install(script: Script): void {
  globals.Worker = ProtocolWorker;
  globals.location = { href: 'http://localhost/' };
  ProtocolWorker.script = script;
}

describe('DocxDocument.load with model sources', () => {
  it.each([
    { mode: 'worker' as const, progressiveLayout: false },
    { mode: 'worker' as const, progressiveLayout: true },
  ])('forwards a claimed load to the worker parse ($mode, progressive=$progressiveLayout)', async (options) => {
    install(renderWorkerScript(undefined));
    const transfer = new ArrayBuffer(4);
    const { source, beginLoad, release } = fakeSource({ transfer: [transfer] });
    const document = await DocxDocument.load(cfbBytes(), { ...options, modelSources: [source] });

    const [parse, ...others] = ProtocolWorker.instances[0]!.parseRequests();
    expect(others).toEqual([]);
    expect(parse).toMatchObject({ source: descriptor, sourceTransfer: [transfer] });
    const parseIndex = ProtocolWorker.instances[0]!.messages.indexOf(parse!);
    expect(ProtocolWorker.instances[0]!.transfers[parseIndex]).toContain(transfer);
    expect(beginLoad).toHaveBeenCalledOnce();
    expect(release).toHaveBeenCalledOnce();
    // A source without resource accounting still reports metrics, just without a usage snapshot.
    await expect(document.getResourceMetrics()).resolves.toBeDefined();
    document.destroy();
  });

  it('streams a claimed load in main mode and releases it once', async () => {
    install(parseWorkerScript(undefined));
    const { source, release } = fakeSource();
    const document = await DocxDocument.load(cfbBytes(), { modelSources: [source] });
    expect(ProtocolWorker.instances[0]!.parseRequests()).toEqual([
      expect.objectContaining({ source: descriptor }),
    ]);
    expect(ProtocolWorker.instances[0]!.parseRequests()[0]).not.toHaveProperty('sourceTransfer');
    expect(document.pageCount).toBeGreaterThan(0);
    expect(activeDocxLayoutViewOf(document).showTrackedChanges).toBe(false);
    expect(release).toHaveBeenCalledOnce();
    document.destroy();
  });

  it('keeps unclaimed input on the OOXML path', async () => {
    install(renderWorkerScript(true));
    const { source, claim, beginLoad } = fakeSource({ claim: () => false });
    await expect(DocxDocument.load(cfbBytes(), { modelSources: [source] }))
      .rejects.toMatchObject({ code: 'legacy-binary-format' });
    expect(claim).toHaveBeenCalledOnce();
    expect(beginLoad).not.toHaveBeenCalled();
    expect(ProtocolWorker.instances).toEqual([]);

    const document = await DocxDocument.load(new ArrayBuffer(4), { mode: 'worker', modelSources: [source] });
    const [parse] = ProtocolWorker.instances[0]!.parseRequests();
    expect(parse).not.toHaveProperty('source');
    expect(parse).not.toHaveProperty('showTrackedChanges');
    document.destroy();
  });

  it('fails closed on a throwing claim and on a source for another target', async () => {
    install(renderWorkerScript(undefined));
    const failure = new Error('claim failed');
    const throwing = fakeSource({ claim: () => { throw failure; } });
    await expect(DocxDocument.load(new ArrayBuffer(4), { modelSources: [throwing.source] }))
      .rejects.toBe(failure);
    expect(throwing.beginLoad).not.toHaveBeenCalled();

    const mismatched = fakeSource({ target: 'xlsx' });
    await expect(DocxDocument.load(new ArrayBuffer(4), { modelSources: [mismatched.source] }))
      .rejects.toThrow(TypeError);
    expect(mismatched.claim).not.toHaveBeenCalled();
    expect(ProtocolWorker.instances).toEqual([]);
  });

  it('releases once and does not fall back to OOXML when the source parse fails', async () => {
    install((worker, message) => {
      if (message.type === 'parse') {
        worker.reply({ type: 'error', id: message.id, name: 'Error', message: 'source parse failed' });
      }
    });
    const { source, release } = fakeSource();
    await expect(DocxDocument.load(cfbBytes(), { mode: 'worker', modelSources: [source] }))
      .rejects.toThrow('source parse failed');
    expect(release).toHaveBeenCalledOnce();
    expect(ProtocolWorker.instances[0]!.parseRequests()).toHaveLength(1);
    expect(ProtocolWorker.instances[0]!.terminated).toBe(true);
  });

  it('sends the caller view tri-state and adopts the worker effective view', async () => {
    const cases = [
      { caller: undefined, sent: undefined, active: true },
      { caller: false, sent: false, active: false },
      { caller: true, sent: true, active: true },
    ];
    for (const { caller, sent, active } of cases) {
      install(renderWorkerScript(true));
      const { source } = fakeSource();
      const document = await DocxDocument.load(cfbBytes(), {
        mode: 'worker',
        modelSources: [source],
        ...(caller === undefined ? {} : { showTrackedChanges: caller }),
      });
      const [parse] = ProtocolWorker.instances.at(-1)!.parseRequests();
      if (sent === undefined) expect(parse).not.toHaveProperty('showTrackedChanges');
      else expect(parse).toMatchObject({ showTrackedChanges: sent });
      expect(activeDocxLayoutViewOf(document).showTrackedChanges).toBe(active);
      document.destroy();
    }
  });

  it('adopts the worker effective view from the first progressive publication', async () => {
    install(renderWorkerScript(true, { partial: true }));
    const { source } = fakeSource();
    const document = await DocxDocument.load(cfbBytes(), {
      mode: 'worker', progressiveLayout: true, modelSources: [source],
    });
    expect(activeDocxLayoutViewOf(document).showTrackedChanges).toBe(true);
    document.destroy();
  });

  it('applies the source view default in main mode only when the caller did not choose', async () => {
    install(parseWorkerScript({ showTrackedChanges: true }));
    const followed = await DocxDocument.load(cfbBytes(), { modelSources: [fakeSource().source] });
    expect(activeDocxLayoutViewOf(followed).showTrackedChanges).toBe(true);
    followed.destroy();

    install(parseWorkerScript({ showTrackedChanges: true }));
    const explicit = await DocxDocument.load(cfbBytes(), {
      showTrackedChanges: false, modelSources: [fakeSource().source],
    });
    expect(activeDocxLayoutViewOf(explicit).showTrackedChanges).toBe(false);
    explicit.destroy();
  });
});
