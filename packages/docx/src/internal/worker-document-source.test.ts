import { describe, expect, it, vi } from 'vitest';
import type { WasmParserHost } from '@silurus/ooxml-core';
import type { LegacyDocDirectSourceDescriptor } from '@silurus/ooxml-core/internal/legacy-doc-source';
import type { OwnedLegacyDocSource } from '@silurus/ooxml-legacy-converter/internal/direct-doc-engine';
import {
  WorkerDocumentSourceOwner,
  type OoxmlWorkerDocumentArchive,
  type WorkerDocumentArchive,
} from './worker-document-source.js';

const descriptor: LegacyDocDirectSourceDescriptor = {
  protocol: 'ooxml-legacy-doc-source/v1', builtin: 'doc',
  wasmUrl: 'https://example.test/direct-doc.wasm',
};

describe('WorkerDocumentSourceOwner', () => {
  it('opens native lazily without touching the DOCX runtime and retains images after cursor close', async () => {
    const native = archive();
    const closeArchive = vi.fn();
    const open = vi.fn(async (): Promise<OwnedLegacyDocSource> => ({
      archive: native, sourceByteLength: 3, closeArchive,
    }));
    const host = hostFor(null);
    const owner = new WorkerDocumentSourceOwner(host.value, open);
    const signal = new AbortController().signal;

    expect(await owner.openNative(new Uint8Array([1, 2, 3]), descriptor, signal)).toBe(native);
    expect(open).toHaveBeenCalledExactlyOnceWith(new Uint8Array([1, 2, 3]), descriptor, signal);
    expect(host.run).not.toHaveBeenCalled();
    expect(owner.kind).toBe('legacy-doc');
    owner.execute((current) => current.close_document_session());
    expect(owner.execute((current) => current.extract_image('legacy-doc/image/1')))
      .toEqual(new Uint8Array([9]));
    expect(closeArchive).not.toHaveBeenCalled();
    owner.closeNative(); owner.closeNative();
    expect(closeArchive).toHaveBeenCalledTimes(1);
  });

  it('uses an existing OOXML archive without loading legacy glue', () => {
    const ooxml = archive() as OoxmlWorkerDocumentArchive;
    ooxml.resource_usage = vi.fn(() => new Uint8Array([4]));
    ooxml.to_markdown = vi.fn(() => 'markdown');
    const host = hostFor(ooxml);
    const open = vi.fn();
    const owner = new WorkerDocumentSourceOwner(host.value, open);

    expect(owner.execute((current) => current.extract_image('word/media/a.png')))
      .toEqual(new Uint8Array([9]));
    expect(host.run).toHaveBeenCalledTimes(1);
    expect(owner.ooxml('resource usage')).toBe(ooxml);
    expect(open).not.toHaveBeenCalled();
  });

  it('does not fall back to OOXML when native opening fails', async () => {
    const failure = new Error('native failed');
    const host = hostFor(null);
    const owner = new WorkerDocumentSourceOwner(host.value, vi.fn(async () => { throw failure; }));
    await expect(owner.openNative(new Uint8Array(), descriptor)).rejects.toBe(failure);
    expect(owner.cursor()).toBeNull();
    expect(host.run).not.toHaveBeenCalled();
  });

  it('rejects OOXML-only operations and replacement while native is owned', async () => {
    const native = archive();
    const closeArchive = vi.fn();
    const open = vi.fn(async () => ({ archive: native, sourceByteLength: 0, closeArchive }));
    const owner = new WorkerDocumentSourceOwner(hostFor(null).value, open);
    await owner.openNative(new Uint8Array(), descriptor);
    expect(() => owner.ooxml('resource usage')).toThrow(
      'resource usage is unsupported for direct legacy DOC sources',
    );
    await expect(owner.openNative(new Uint8Array(), descriptor)).rejects.toThrow('already loaded');
    expect(open).toHaveBeenCalledTimes(1);
    owner.closeNative();
    await owner.openNative(new Uint8Array(), descriptor);
    expect(open).toHaveBeenCalledTimes(2);
    owner.closeNative();
    expect(closeArchive).toHaveBeenCalledTimes(2);
  });

  it('rejects overlapping opens and disposes a result invalidated while pending', async () => {
    let resolve!: (source: OwnedLegacyDocSource) => void;
    const pending = new Promise<OwnedLegacyDocSource>((accept) => { resolve = accept; });
    const open = vi.fn(() => pending);
    const owner = new WorkerDocumentSourceOwner(hostFor(null).value, open);
    const first = owner.openNative(new Uint8Array([1]), descriptor);
    await expect(owner.openNative(new Uint8Array([2]), descriptor)).rejects.toThrow('opening');
    owner.closeNative();
    const closeArchive = vi.fn();
    resolve({ archive: archive(), sourceByteLength: 1, closeArchive });
    await expect(first).rejects.toThrow('superseded');
    expect(closeArchive).toHaveBeenCalledTimes(1);
    expect(owner.cursor()).toBeNull();
  });

  it('does not let a late old open clear or overwrite a newer pending generation', async () => {
    const resolvers: ((source: OwnedLegacyDocSource) => void)[] = [];
    const open = vi.fn(() => new Promise<OwnedLegacyDocSource>((resolve) => { resolvers.push(resolve); }));
    const owner = new WorkerDocumentSourceOwner(hostFor(null).value, open);
    const oldOpen = owner.openNative(new Uint8Array([1]), descriptor);
    owner.closeNative();
    const newOpen = owner.openNative(new Uint8Array([2]), descriptor);
    const oldClose = vi.fn();
    resolvers[0]!({ archive: archive(), sourceByteLength: 1, closeArchive: oldClose });
    await expect(oldOpen).rejects.toThrow('superseded');
    await expect(owner.openNative(new Uint8Array([3]), descriptor)).rejects.toThrow('opening');
    const current = archive();
    const newClose = vi.fn();
    resolvers[1]!({ archive: current, sourceByteLength: 1, closeArchive: newClose });
    await expect(newOpen).resolves.toBe(current);
    expect(oldClose).toHaveBeenCalledTimes(1);
    owner.closeNative();
    expect(newClose).toHaveBeenCalledTimes(1);
  });

  it('clears failed pending state and post-validates cancellation', async () => {
    const controller = new AbortController();
    const closeArchive = vi.fn();
    const native = archive();
    const open = vi.fn()
      .mockRejectedValueOnce(new Error('failed'))
      .mockImplementationOnce(async () => {
        controller.abort();
        return { archive: native, sourceByteLength: 0, closeArchive };
      })
      .mockResolvedValueOnce({ archive: native, sourceByteLength: 0, closeArchive });
    const owner = new WorkerDocumentSourceOwner(hostFor(null).value, open);
    await expect(owner.openNative(new Uint8Array(), descriptor)).rejects.toThrow('failed');
    await expect(owner.openNative(new Uint8Array(), descriptor, controller.signal))
      .rejects.toMatchObject({ name: 'AbortError' });
    expect(closeArchive).toHaveBeenCalledTimes(1);
    await expect(owner.openNative(new Uint8Array(), descriptor)).resolves.toBe(native);
    owner.closeNative();
    expect(closeArchive).toHaveBeenCalledTimes(2);
  });
});

function hostFor(current: OoxmlWorkerDocumentArchive | null) {
  const run = vi.fn(<T>(operation: () => T): T => operation());
  return {
    run,
    value: { archive: current, run } as unknown as WasmParserHost<OoxmlWorkerDocumentArchive>,
  };
}

function archive(): WorkerDocumentArchive & { free(): void } {
  return {
    free: vi.fn(), assert_healthy: vi.fn(),
    open_document_cursor: vi.fn(), pull_document_chunk: vi.fn(() => new Uint8Array()),
    document_chunk_done: vi.fn(() => true), acknowledge_document_chunk: vi.fn(),
    cancel_document_cursor: vi.fn(), close_document_session: vi.fn(),
    extract_image: vi.fn(() => new Uint8Array([9])),
  };
}
