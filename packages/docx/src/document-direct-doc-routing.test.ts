import { afterEach, describe, expect, it, vi } from 'vitest';
import type { WorkerLike } from '@silurus/ooxml-core';
import * as core from '@silurus/ooxml-core';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';
import { DocxDocument } from './document.js';
import * as localFontMetrics from './local-font-metrics.js';
import * as renderer from './renderer.js';

class SilentWorker implements WorkerLike {
  static instances: SilentWorker[] = [];
  terminated = false;
  readonly messages: unknown[] = [];
  constructor() { SilentWorker.instances.push(this); }
  postMessage(message: unknown): void { this.messages.push(message); }
  addEventListener(): void {}
  removeEventListener(): void {}
  terminate(): void { this.terminated = true; }
}

const globals = globalThis as Record<string, unknown>;
const originals = { Worker: globals.Worker, OffscreenCanvas: globals.OffscreenCanvas, location: globals.location };
const source = { protocol: 'ooxml-legacy-doc-source/v1', builtin: 'doc',
  wasmUrl: 'https://example.test/direct-doc.wasm' } as const;
const docBuffer = () => buildCfbFixture(['Root Entry', 'WordDocument', '1Table']);

afterEach(() => {
  vi.restoreAllMocks();
  globals.Worker = originals.Worker;
  globals.OffscreenCanvas = originals.OffscreenCanvas;
  globals.location = originals.location;
  SilentWorker.instances = [];
});

function install() {
  globals.Worker = SilentWorker;
  globals.OffscreenCanvas = class {};
  globals.location = { href: 'http://localhost/' };
  const parse = vi.spyOn(
    DocxDocument.prototype as unknown as { _parse(...args: unknown[]): Promise<void> },
    '_parse',
  ).mockResolvedValue(undefined);
  return parse;
}

describe('DocxDocument direct DOC routing', () => {
  it.each([
    { mode: 'main' as const, progressiveLayout: false },
    { mode: 'worker' as const, progressiveLayout: false },
    { mode: 'worker' as const, progressiveLayout: true },
  ])('forwards the admitted descriptor in $mode mode (progressive=$progressiveLayout)', async (options) => {
    const parse = install();
    const usage = vi.spyOn(
      DocxDocument.prototype as unknown as { _resourceUsage(timeoutMs: number): Promise<unknown> },
      '_resourceUsage',
    );
    const document = await DocxDocument.load(docBuffer(), {
      ...options, legacyConversion: { doc: { source } },
    });
    expect(parse).toHaveBeenCalledTimes(1);
    expect(parse.mock.calls[0]?.[7]).toEqual(source);
    expect(usage).not.toHaveBeenCalled();
    await expect(document.getResourceMetrics()).rejects.toThrow(
      'resource usage is unsupported for direct legacy DOC sources',
    );
    document.destroy();
  });

  it('keeps ordinary input and unrelated format opt-ins off the native DOC route', async () => {
    const parse = install();
    const ordinary = await DocxDocument.load(new ArrayBuffer(0));
    expect(parse.mock.calls[0]?.[7]).toBeUndefined();
    ordinary.destroy();
    await expect(DocxDocument.load(docBuffer(), {
      legacyConversion: { xls: { converter: { convert: vi.fn() } } },
    })).rejects.toMatchObject({ code: 'legacy-binary-format' });
    expect(parse).toHaveBeenCalledTimes(1);
  });

  it('does not fall back when native parsing fails', async () => {
    const failure = new Error('native parse failed');
    const parse = install().mockRejectedValueOnce(failure);
    await expect(DocxDocument.load(docBuffer(), {
      legacyConversion: { doc: { source } },
    })).rejects.toBe(failure);
    expect(parse).toHaveBeenCalledTimes(1);
    expect(SilentWorker.instances[0]?.terminated).toBe(true);
  });

  it('destroys the owned worker when the retained native signal aborts during parse', async () => {
    let settle!: () => void;
    const parsePending = new Promise<void>((resolve) => { settle = resolve; });
    const parse = install().mockReturnValueOnce(parsePending);
    const controller = new AbortController();
    const loading = DocxDocument.load(docBuffer(), {
      legacyConversion: { doc: { source, signal: controller.signal } },
    });
    await vi.waitFor(() => expect(parse).toHaveBeenCalledOnce());
    controller.abort();
    await expect(loading).rejects.toMatchObject({ name: 'AbortError' });
    expect(SilentWorker.instances[0]?.terminated).toBe(true);
    settle();
  });

  it('rejects a pre-aborted source without starting parse', async () => {
    const parse = install();
    const controller = new AbortController();
    controller.abort();
    await expect(DocxDocument.load(docBuffer(), {
      legacyConversion: { doc: { source, signal: controller.signal } },
    })).rejects.toThrow();
    expect(parse).not.toHaveBeenCalled();
  });

  it('forwards the descriptor through the actual worker parse request', async () => {
    globals.Worker = SilentWorker;
    globals.OffscreenCanvas = class {};
    globals.location = { href: 'http://localhost/' };
    const controller = new AbortController();
    const loading = DocxDocument.load(docBuffer(), {
      mode: 'worker',
      legacyConversion: { doc: { source, signal: controller.signal } },
    });
    await vi.waitFor(() => expect(SilentWorker.instances[0]?.messages.some(
      (message) => (message as { type?: string }).type === 'parse',
    )).toBe(true));
    expect(SilentWorker.instances[0]?.messages.find(
      (message) => (message as { type?: string }).type === 'parse',
    )).toMatchObject({ type: 'parse', source });
    controller.abort();
    await expect(loading).rejects.toMatchObject({ name: 'AbortError' });
  });

  it('rejects abort after parse while post-parse setup is pending', async () => {
    let finishMetrics!: () => void;
    const metricsPending = new Promise<void>((resolve) => { finishMetrics = resolve; });
    const face = {} as FontFace;
    const unload = vi.spyOn(core, 'unloadLocalFontMetrics');
    vi.spyOn(localFontMetrics, 'loadDocxLocalFontMetrics').mockImplementation(async () => {
      await metricsPending;
      return { faces: [face], metrics: new Map() } as never;
    });
    const parse = install().mockImplementationOnce(async function (this: unknown) {
      (this as { _document: unknown })._document = { body: [] };
    });
    const controller = new AbortController();
    const loading = DocxDocument.load(docBuffer(), {
      legacyConversion: { doc: { source, signal: controller.signal } },
    });
    await vi.waitFor(() => expect(localFontMetrics.loadDocxLocalFontMetrics).toHaveBeenCalledOnce());
    controller.abort();
    await expect(loading).rejects.toMatchObject({ name: 'AbortError' });
    finishMetrics();
    await vi.waitFor(() => expect(unload).toHaveBeenCalledWith([face]));
    expect(parse).toHaveBeenCalledOnce();
    expect(SilentWorker.instances[0]?.terminated).toBe(true);
  });

  it('rejects abort while deferred math preparation remains pending', async () => {
    let finishMath!: () => void;
    const mathPending = new Promise<void>((resolve) => { finishMath = resolve; });
    vi.spyOn(localFontMetrics, 'loadDocxLocalFontMetrics').mockResolvedValue({
      faces: [], metrics: new Map(),
    } as never);
    vi.spyOn(renderer, 'documentHasMath').mockReturnValue(true);
    const prepare = vi.spyOn(renderer, 'prepareMathRuns').mockImplementation(async () => {
      await mathPending;
      return { records: [], drawables: new Map() };
    });
    const parse = install().mockImplementationOnce(async function (this: unknown) {
      (this as { _document: unknown })._document = { body: [] };
    });
    const controller = new AbortController();
    const loading = DocxDocument.load(docBuffer(), {
      math: { loadMathJax: vi.fn(), mathMLToSvg: vi.fn() },
      legacyConversion: { doc: { source, signal: controller.signal } },
    });
    await vi.waitFor(() => expect(prepare).toHaveBeenCalledOnce());
    controller.abort();
    await expect(loading).rejects.toMatchObject({ name: 'AbortError' });
    finishMath();
    await Promise.resolve();
    expect(parse).toHaveBeenCalledOnce();
    expect(SilentWorker.instances[0]?.terminated).toBe(true);
  });
});
