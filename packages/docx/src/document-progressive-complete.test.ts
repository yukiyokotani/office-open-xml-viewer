import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { OoxmlResourceUsageSnapshot, WorkerLike } from '@silurus/ooxml-core';
import { DocxDocument } from './document.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { installStubCanvas, syntheticDocxModel } from './testing/synthetic-document.js';

// ─────────────────────────────────────────────────────────────────────────────
// The `onLayoutComplete` terminal-callback contract for main-mode progressive
// loads: exactly one success notification per load, whether or not the
// document was long enough to publish partials first.
//
// `layoutDocumentProgressively` only publishes a preview when the body has
// more than MIN_PROGRESSIVE_ENTRIES entries, so a short document's drain
// completes with `publishedLayout === null`. The success notification used to
// be gated on that local, so consumers of a fast document never learned the
// layout was complete — the terminal callback contract silently depended on
// document speed.
//
// These tests drive the REAL `DocxDocument.load` progressive block: `_parse`
// is stubbed to install a synthetic model/source (the WASM parse is not what
// is under test), and the final usage probe is stubbed because the inert
// inline worker cannot answer it.
// ─────────────────────────────────────────────────────────────────────────────

class SilentWorker implements WorkerLike {
  postMessage(): void {}
  addEventListener(): void {}
  removeEventListener(): void {}
  terminate(): void {}
}

const globals = globalThis as Record<string, unknown>;
const originals = {
  Worker: globals.Worker,
  location: globals.location,
};

const USAGE: OoxmlResourceUsageSnapshot = {
  archiveEntryCount: 1,
  declaredInflatedBytes: 0,
  largestInflatedEntryBytes: 0,
  distinctInflatedBytes: 0,
  operationInflatedBytes: 0,
};

beforeAll(() => {
  installStubCanvas();
});

afterEach(() => {
  vi.restoreAllMocks();
  globals.Worker = originals.Worker;
  globals.location = originals.location;
});

/** Stub the WASM parse with a synthetic document of `paragraphs` body entries. */
function installMainModeParse(paragraphs: number): void {
  globals.Worker = SilentWorker;
  globals.location = { href: 'http://localhost/' };
  vi.spyOn(
    DocxDocument.prototype as unknown as {
      _parse(
        buffer: ArrayBuffer,
        resourcePolicy: unknown,
        useGoogleFonts?: boolean,
        timeoutMs?: number,
        onUsage?: unknown,
        renderers?: unknown,
        progressive?: unknown,
      ): Promise<void>;
    },
    '_parse',
  ).mockImplementation(async function (this: DocxDocument) {
    const doc = this as unknown as {
      _document: unknown;
      _source: unknown;
      _meta: unknown;
    };
    const model = syntheticDocxModel('plain', { paragraphs });
    doc._document = model;
    doc._source = layoutSourceStore(model);
    doc._meta = null;
  });
  vi.spyOn(
    DocxDocument.prototype as unknown as {
      _resourceUsage(timeoutMs: number): Promise<OoxmlResourceUsageSnapshot>;
    },
    '_resourceUsage',
  ).mockResolvedValue(USAGE);
}

describe('main-mode progressive load: onLayoutComplete contract', () => {
  it('notifies completion exactly once for a fast document that publishes nothing early', async () => {
    // Three body entries stay under the preview threshold, so the drain
    // completes before any partial could be shown: no onLayoutPartial, and
    // load() resolves on the finished layout itself. The completion callback
    // must still fire — load() resolving is not a substitute for it.
    installMainModeParse(3);
    const completions: unknown[] = [];
    const partials: unknown[] = [];

    const doc = await DocxDocument.load(new ArrayBuffer(0), {
      progressiveLayout: true,
      onLayoutComplete: (error) => completions.push(error),
      onLayoutPartial: (partial) => partials.push(partial),
    });
    await doc.waitUntilLayoutComplete();

    expect(partials).toHaveLength(0);
    expect(completions).toEqual([undefined]);
    expect(doc.layoutComplete).toBe(true);
    doc.destroy();
  });

  it('notifies completion exactly once after partials for a long document', async () => {
    // Long enough to cross several page-count checkpoints: load() resolves on
    // the opening publication and the drain keeps publishing. The completion
    // callback fires once, after the last partial, never twice.
    installMainModeParse(600);
    const completions: unknown[] = [];
    const partials: number[] = [];

    const doc = await DocxDocument.load(new ArrayBuffer(0), {
      progressiveLayout: true,
      onLayoutComplete: (error) => completions.push(error),
      onLayoutPartial: (partial) => partials.push(partial.availableUnits),
    });
    await doc.waitUntilLayoutComplete();

    expect(partials.length).toBeGreaterThan(0);
    expect(completions).toEqual([undefined]);
    expect(doc.layoutComplete).toBe(true);
    doc.destroy();
  }, 300_000);
});
