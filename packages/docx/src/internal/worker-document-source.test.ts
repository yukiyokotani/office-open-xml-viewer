import { describe, expect, it, vi } from 'vitest';
import type {
  ModelSourceModuleDescriptor,
  OpenedModelSourceModule,
  WasmParserHost,
} from '@silurus/ooxml-core';
import {
  WorkerDocumentSourceOwner,
  type DocxModelSourceArchive,
  type OoxmlWorkerDocumentArchive,
} from './worker-document-source.js';

const descriptor: ModelSourceModuleDescriptor = {
  protocol: 'ooxml-model-source-module/v1',
  target: 'docx',
  moduleUrl: 'https://example.test/source.mjs',
  config: {},
};

type Opened = OpenedModelSourceModule<DocxModelSourceArchive>;

describe('WorkerDocumentSourceOwner', () => {
  it('opens a model source without touching the DOCX runtime and keeps images after the cursor closes', async () => {
    const source = archive();
    const close = vi.fn();
    const transfer = [new ArrayBuffer(1)];
    const open = vi.fn(async (): Promise<Opened> => ({ archive: source, viewDefaults: {}, close }));
    const host = hostFor(null);
    const owner = new WorkerDocumentSourceOwner(host.value, open);

    await expect(owner.openModelSource(new Uint8Array([1, 2, 3]), descriptor, transfer))
      .resolves.toEqual({});
    expect(open).toHaveBeenCalledExactlyOnceWith(descriptor, new Uint8Array([1, 2, 3]), transfer);
    expect(owner.kind).toBe('model-source');
    owner.execute((current) => current.close_document_session());
    expect(owner.execute((current) => current.extract_image('media/1'))).toEqual(new Uint8Array([9]));
    expect(host.run).not.toHaveBeenCalled();
    expect(close).not.toHaveBeenCalled();
    owner.closeModelSource();
    owner.closeModelSource();
    expect(close).toHaveBeenCalledTimes(1);
    expect(owner.kind).toBe('ooxml');
  });

  it('admits only a boolean showTrackedChanges view default and closes a source reporting others', async () => {
    const accepted = new WorkerDocumentSourceOwner(hostFor(null).value, async () => ({
      archive: archive(), viewDefaults: { showTrackedChanges: true }, close: vi.fn(),
    }));
    await expect(accepted.openModelSource(new Uint8Array([1]), descriptor))
      .resolves.toEqual({ showTrackedChanges: true });
    expect(accepted.viewDefaults).toEqual({ showTrackedChanges: true });

    const close = vi.fn();
    const rejected = new WorkerDocumentSourceOwner(hostFor(null).value, async () => ({
      archive: archive(), viewDefaults: { showTrackedChanges: true, showComments: false }, close,
    }));
    await expect(rejected.openModelSource(new Uint8Array([1]), descriptor))
      .rejects.toThrow('unsupported DOCX model source view default: showComments');
    expect(close).toHaveBeenCalledTimes(1);
    expect(rejected.cursor()).toBeNull();
    expect(rejected.viewDefaults).toEqual({});
  });

  it('runs an existing OOXML archive under the parser host without opening a source', () => {
    const ooxml = archive() as OoxmlWorkerDocumentArchive;
    ooxml.resource_usage = vi.fn(() => new Uint8Array([4]));
    ooxml.to_markdown = vi.fn(() => 'markdown');
    const host = hostFor(ooxml);
    const open = vi.fn();
    const owner = new WorkerDocumentSourceOwner(host.value, open);

    expect(owner.execute((current) => current.extract_image('word/media/a.png')))
      .toEqual(new Uint8Array([9]));
    expect(owner.resourceUsage()).toEqual(new Uint8Array([4]));
    expect(owner.toMarkdown()).toBe('markdown');
    expect(host.run).toHaveBeenCalledTimes(3);
    expect(open).not.toHaveBeenCalled();
  });

  it('reports missing optional capabilities instead of reaching for the OOXML archive', async () => {
    const owner = new WorkerDocumentSourceOwner(hostFor(null).value, async () => ({
      archive: archive(), viewDefaults: {}, close: vi.fn(),
    }));
    await owner.openModelSource(new Uint8Array(), descriptor);
    expect(owner.resourceUsage()).toBeUndefined();
    expect(() => owner.toMarkdown()).toThrow('Markdown conversion is unsupported for this source');

    const capable = { ...archive(), resource_usage: () => new Uint8Array([7]), to_markdown: () => '# md' };
    const full = new WorkerDocumentSourceOwner(hostFor(null).value, async () => ({
      archive: capable, viewDefaults: {}, close: vi.fn(),
    }));
    await full.openModelSource(new Uint8Array(), descriptor);
    expect(full.resourceUsage()).toEqual(new Uint8Array([7]));
    expect(full.toMarkdown()).toBe('# md');
  });

  it('closes the source when one of its calls traps, but not on an ordinary error', async () => {
    const close = vi.fn();
    const owner = new WorkerDocumentSourceOwner(hostFor(null).value, async () => ({
      archive: archive(), viewDefaults: {}, close,
    }));
    await owner.openModelSource(new Uint8Array(), descriptor);
    expect(() => owner.execute(() => { throw new Error('bad path'); })).toThrow('bad path');
    expect(close).not.toHaveBeenCalled();
    const trap = Object.assign(new Error('unreachable'), { name: 'RuntimeError' });
    expect(() => owner.execute(() => { throw trap; })).toThrow(trap);
    expect(close).toHaveBeenCalledTimes(1);
    expect(owner.cursor()).toBeNull();
  });

  it('does not fall back to OOXML when opening fails and allows a later open', async () => {
    const failure = new Error('open failed');
    const close = vi.fn();
    const open = vi.fn()
      .mockRejectedValueOnce(failure)
      .mockResolvedValueOnce({ archive: archive(), viewDefaults: {}, close });
    const host = hostFor(null);
    const owner = new WorkerDocumentSourceOwner(host.value, open);
    await expect(owner.openModelSource(new Uint8Array(), descriptor)).rejects.toBe(failure);
    expect(owner.cursor()).toBeNull();
    expect(host.run).not.toHaveBeenCalled();
    await expect(owner.openModelSource(new Uint8Array(), descriptor)).resolves.toEqual({});
    expect(owner.kind).toBe('model-source');
  });

  it('rejects replacing a loaded source until it is closed', async () => {
    const close = vi.fn();
    const open = vi.fn(async (): Promise<Opened> => ({ archive: archive(), viewDefaults: {}, close }));
    const owner = new WorkerDocumentSourceOwner(hostFor(null).value, open);
    await owner.openModelSource(new Uint8Array(), descriptor);
    await expect(owner.openModelSource(new Uint8Array(), descriptor)).rejects.toThrow('already loaded');
    expect(open).toHaveBeenCalledTimes(1);
    owner.closeModelSource();
    await owner.openModelSource(new Uint8Array(), descriptor);
    expect(open).toHaveBeenCalledTimes(2);
    owner.closeModelSource();
    expect(close).toHaveBeenCalledTimes(2);

    const loaded = new WorkerDocumentSourceOwner(hostFor(archive() as OoxmlWorkerDocumentArchive).value, open);
    await expect(loaded.openModelSource(new Uint8Array(), descriptor)).rejects.toThrow('already loaded');
    expect(open).toHaveBeenCalledTimes(2);
  });

  it('rejects overlapping opens and closes a result invalidated while pending', async () => {
    let resolve!: (source: Opened) => void;
    const open = vi.fn(() => new Promise<Opened>((accept) => { resolve = accept; }));
    const owner = new WorkerDocumentSourceOwner(hostFor(null).value, open);
    const first = owner.openModelSource(new Uint8Array([1]), descriptor);
    await expect(owner.openModelSource(new Uint8Array([2]), descriptor)).rejects.toThrow('opening');
    owner.closeModelSource();
    const close = vi.fn();
    resolve({ archive: archive(), viewDefaults: {}, close });
    await expect(first).rejects.toThrow('superseded');
    expect(close).toHaveBeenCalledTimes(1);
    expect(owner.cursor()).toBeNull();
  });

  it('does not let a late old open clear or overwrite a newer pending open', async () => {
    const resolvers: ((source: Opened) => void)[] = [];
    const open = vi.fn(() => new Promise<Opened>((resolve) => { resolvers.push(resolve); }));
    const owner = new WorkerDocumentSourceOwner(hostFor(null).value, open);
    const oldOpen = owner.openModelSource(new Uint8Array([1]), descriptor);
    owner.closeModelSource();
    const newOpen = owner.openModelSource(new Uint8Array([2]), descriptor);
    const oldClose = vi.fn();
    resolvers[0]!({ archive: archive(), viewDefaults: {}, close: oldClose });
    await expect(oldOpen).rejects.toThrow('superseded');
    await expect(owner.openModelSource(new Uint8Array([3]), descriptor)).rejects.toThrow('opening');
    const current = archive();
    const newClose = vi.fn();
    resolvers[1]!({ archive: current, viewDefaults: { showTrackedChanges: false }, close: newClose });
    await expect(newOpen).resolves.toEqual({ showTrackedChanges: false });
    expect(owner.cursor()).toBe(current);
    expect(oldClose).toHaveBeenCalledTimes(1);
    owner.closeModelSource();
    expect(newClose).toHaveBeenCalledTimes(1);
  });
});

function hostFor(current: OoxmlWorkerDocumentArchive | null) {
  const run = vi.fn(<T>(operation: () => T): T => operation());
  return {
    run,
    value: { archive: current, run } as unknown as WasmParserHost<OoxmlWorkerDocumentArchive>,
  };
}

function archive(): DocxModelSourceArchive & { free(): void } {
  return {
    free: vi.fn(), assert_healthy: vi.fn(),
    open_document_cursor: vi.fn(), pull_document_chunk: vi.fn(() => new Uint8Array()),
    document_chunk_done: vi.fn(() => true), acknowledge_document_chunk: vi.fn(),
    cancel_document_cursor: vi.fn(), close_document_session: vi.fn(),
    extract_image: vi.fn(() => new Uint8Array([9])),
  } as unknown as DocxModelSourceArchive & { free(): void };
}
