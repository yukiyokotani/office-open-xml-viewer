import { describe, expect, it, vi } from 'vitest';
import type {
  PullSessionIdentity,
} from '@silurus/ooxml-core/worker';
import {
  materializeDocumentPullLayoutSession,
  materializeDocumentPullSession,
  materializeDocumentPullOwnedModelsSession,
} from './document-pull-client.js';
import { layoutSourceModelAdapterFromOwnedModel } from './layout-source-model-adapter.js';
import {
  createLocalDocumentPullTransport,
  DocumentPullWorker,
  MaterializedDocumentCursorArchive,
  type DocxDocumentCursorArchive,
} from './document-pull-worker.js';
import type { DocxDocumentModel } from './types.js';

const encoder = new TextEncoder();

class FakeArchive implements DocxDocumentCursorArchive {
  private sequence = 0;
  private delivered = false;
  private currentDone = false;
  opened = false;
  canceled = false;
  readonly units: Uint8Array[];

  constructor(units?: unknown[]) {
    this.units = (units ?? [
      { kind: 'body', body: [{ type: 'pageBreak' }] },
      { kind: 'body', body: [{ type: 'columnBreak' }] },
      { kind: 'complete', document: { body: [] } },
    ]).map((unit) => encoder.encode(JSON.stringify(unit)));
  }

  open_document_cursor(): void { this.opened = true; }

  pull_document_chunk(sequence: number, _operation: number, _generation: number, credit: number): Uint8Array {
    if (!this.opened) throw new Error('not open');
    if (sequence !== this.sequence) throw new Error('bad sequence');
    if (this.delivered) throw new Error('document unit must be acknowledged before another pull');
    const bytes = this.units[sequence]!;
    if (bytes.byteLength > credit) {
      throw new Error(`document unit requires ${bytes.byteLength} bytes but credit is ${credit}`);
    }
    this.delivered = true;
    this.currentDone = sequence + 1 === this.units.length;
    return bytes.slice();
  }

  document_chunk_done(): boolean { return this.currentDone; }

  acknowledge_document_chunk(sequence: number): void {
    if (!this.delivered || sequence !== this.sequence) throw new Error('bad acknowledgement');
    this.delivered = false;
    this.sequence += 1;
  }

  cancel_document_cursor(): void { this.canceled = true; }
  close_document_session(): void { this.canceled = true; }
}

const identity: PullSessionIdentity<number> = {
  sessionId: 1,
  operationId: 1,
  generation: 1,
};

describe('DOCX document pull integration', () => {
  it('materializes acknowledged body units and a body-free terminal envelope', async () => {
    const archive = new FakeArchive();
    const worker = new DocumentPullWorker(() => archive);
    worker.open(identity);

    const document = await materializeDocumentPullSession(
      createLocalDocumentPullTransport(worker),
      identity,
    );

    expect(document.body.map((element) => element.type)).toEqual(['pageBreak', 'columnBreak']);
    expect(archive.canceled).toBe(false);
  });

  it('builds the public model and immutable layout ownership graph per bounded unit', async () => {
    const archive = new FakeArchive([
      { kind: 'body', body: [{ type: 'pageBreak' }] },
      { kind: 'body', body: [{ type: 'columnBreak' }] },
      {
        kind: 'complete',
        document: {
          body: [],
          section: {},
          headers: { default: null, first: null, even: null },
          footers: { default: null, first: null, even: null },
        },
      },
    ]);
    const worker = new DocumentPullWorker(() => archive);
    worker.open(identity);
    const nativeStructuredClone = globalThis.structuredClone;
    const clonedDocumentBodyLengths: number[] = [];
    const clone = vi.spyOn(globalThis, 'structuredClone').mockImplementation((value) => {
      if (!!value && typeof value === 'object'
        && !Array.isArray(value)
        && Array.isArray((value as { body?: unknown }).body)
        && 'section' in value) {
        clonedDocumentBodyLengths.push((value as unknown as { body: unknown[] }).body.length);
      }
      return nativeStructuredClone(value);
    });

    try {
      const models = await materializeDocumentPullOwnedModelsSession(
        createLocalDocumentPullTransport(worker),
        identity,
      );
      const ownedBody = models.ownedLayoutDocument.body;
      const adapted = layoutSourceModelAdapterFromOwnedModel(
        models.document,
        models.ownedLayoutDocument,
      );

      expect(adapted.document.body.map((element) => element.type))
        .toEqual(['pageBreak', 'columnBreak']);
      expect(adapted.source.blocks.body.map((element) => element.type))
        .toEqual(['pageBreak', 'columnBreak']);
      expect(adapted.source.bodyLayoutInput.sequence)
        .toEqual([
          expect.objectContaining({ kind: 'authored-break', break: 'page' }),
          expect.objectContaining({ kind: 'authored-break', break: 'column' }),
        ]);

      const retainedFirst = adapted.source.blocks.body[0]!;
      (adapted.document.body[0] as { type: string }).type = 'columnBreak';
      expect(retainedFirst.type).toBe('pageBreak');
      expect(Object.isFrozen(retainedFirst)).toBe(true);
      expect(adapted.source.blocks.body).toBe(ownedBody);

      // The only body-array clones are the individually transferred units.
      // In particular, the completed two-element compatibility body is never
      // passed to structuredClone at the terminal adapter boundary.
      expect(clonedDocumentBodyLengths).toEqual([0]);
    } finally {
      clone.mockRestore();
    }
  });

  it('builds the Node layout source without constructing a compatibility graph', async () => {
    const archive = new FakeArchive([
      { kind: 'body', body: [{ type: 'pageBreak' }] },
      { kind: 'body', body: [{ type: 'columnBreak' }] },
      {
        kind: 'complete',
        document: {
          body: [],
          section: {},
          headers: { default: null, first: null, even: null },
          footers: { default: null, first: null, even: null },
        },
      },
    ]);
    const worker = new DocumentPullWorker(() => archive);
    worker.open(identity);

    const source = await materializeDocumentPullLayoutSession(
      createLocalDocumentPullTransport(worker),
      identity,
    );

    expect(source.blocks.body.map((element) => element.type))
      .toEqual(['pageBreak', 'columnBreak']);
    expect(source.bodyLayoutInput.sequence).toEqual([
      expect.objectContaining({ kind: 'authored-break', break: 'page' }),
      expect.objectContaining({ kind: 'authored-break', break: 'column' }),
    ]);
    expect(Object.isFrozen(source.blocks.body[0])).toBe(true);
  });

  it('cancels producer ownership when consumer validation rejects a unit', async () => {
    const archive = new FakeArchive([{ kind: 'unexpected' }]);
    const worker = new DocumentPullWorker(() => archive);
    worker.open(identity);

    await expect(
      materializeDocumentPullOwnedModelsSession(
        createLocalDocumentPullTransport(worker),
        identity,
      ),
    ).rejects.toThrow('unknown shape');
    expect(archive.canceled).toBe(true);
  });

  it('streams units when the package has no document cursor checkpoint', async () => {
    // A corrupt package is streamed as a placeholder document before any
    // cursor ledger exists. The absent checkpoint must not fail the chunk
    // (same policy as the PPTX/XLSX cursors); real usage errors still escape.
    class UnledgeredArchive extends FakeArchive {
      document_cursor_resource_usage(): Uint8Array {
        throw new Error('document cursor usage is unavailable');
      }
    }
    const unledgered = new UnledgeredArchive();
    const worker = new DocumentPullWorker(() => unledgered);
    worker.open(identity);
    const document = await materializeDocumentPullSession(
      createLocalDocumentPullTransport(worker),
      identity,
    );
    expect(document.body.map((element) => element.type)).toEqual(['pageBreak', 'columnBreak']);

    class BrokenUsageArchive extends FakeArchive {
      document_cursor_resource_usage(): Uint8Array {
        throw new Error('resource ledger corrupted');
      }
    }
    const brokenArchive = new BrokenUsageArchive();
    const broken = new DocumentPullWorker(() => brokenArchive);
    broken.open(identity);
    await expect(
      materializeDocumentPullSession(createLocalDocumentPullTransport(broken), identity),
    ).rejects.toThrow('resource ledger corrupted');
  });

  it('streams an already-materialized fallback model without a whole-model envelope', async () => {
    const source = {
      section: {},
      body: [{ type: 'pageBreak' }, { type: 'columnBreak' }],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as unknown as DocxDocumentModel;
    const archive = new MaterializedDocumentCursorArchive(source);
    const worker = new DocumentPullWorker(() => archive);
    worker.open(identity);

    // Ownership moves into the bounded cursor immediately; the compatibility
    // model returned to the consumer is reconstructed from transferred units.
    expect(source.body).toEqual([]);
    const { document } = await materializeDocumentPullOwnedModelsSession(
      createLocalDocumentPullTransport(worker),
      identity,
    );

    expect(document).not.toBe(source);
    expect(document.body.map((element) => element.type)).toEqual(['pageBreak', 'columnBreak']);
  });
});
