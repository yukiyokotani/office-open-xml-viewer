import { afterEach, describe, expect, it, vi } from 'vitest';
import type { PullSessionIdentity } from '@silurus/ooxml-core/worker';
import type { DocxDocumentCursorArchive } from './document-pull-worker.js';

/**
 * #2248 — the DOCX error page is code only a document that failed to parse
 * needs. The pull materializers hand such a document out only once that code
 * has loaded, so the synchronous layout that follows finds it; the wait ends
 * with its owner, and a document that parsed neither loads nor awaits it.
 */
const encoder = new TextEncoder();
const identity: PullSessionIdentity<number> = { sessionId: 1, operationId: 1, generation: 1 };
const MESSAGE = 'word/document.xml: unexpected end of stream';

function terminal(parseError?: string) {
  return {
    kind: 'complete',
    document: {
      body: [],
      section: {
        pageWidth: 612, pageHeight: 792,
        marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
        headerDistance: 36, footerDistance: 36,
        titlePage: false, evenAndOddHeaders: false,
      },
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      ...(parseError === undefined ? {} : { parseError }),
    },
  };
}

class OneUnitArchive implements DocxDocumentCursorArchive {
  private delivered = false;
  canceled = false;
  private readonly unit: Uint8Array;
  constructor(unit: unknown) { this.unit = encoder.encode(JSON.stringify(unit)); }
  open_document_cursor(): void {}
  pull_document_chunk(): Uint8Array {
    if (this.delivered) throw new Error('document unit must be acknowledged before another pull');
    this.delivered = true;
    return this.unit.slice();
  }
  document_chunk_done(): boolean { return true; }
  acknowledge_document_chunk(): void { this.delivered = false; }
  cancel_document_cursor(): void { this.canceled = true; }
  close_document_session(): void { this.canceled = true; }
}

/** Fresh modules whose error-page chunk stays in flight until `release()`. */
async function coldModules() {
  vi.resetModules();
  let release!: () => void;
  const stalled = new Promise<void>((resolve) => { release = resolve; });
  vi.doMock('./layout/error-page.js', async () => {
    await stalled;
    return vi.importActual('./layout/error-page.js');
  });
  const client = await import('./document-pull-client.js');
  const pull = await import('./document-pull-worker.js');
  const variants = await import('./layout/document-layout-variants.js');
  const { ensureDocumentLayoutVariants } = await import('./layout/document.js');
  const { createLayoutServices } = await import('./layout-runtime.js');
  const transportFor = (parseError?: string) => {
    const archive = new OneUnitArchive(terminal(parseError));
    const worker = new pull.DocumentPullWorker(() => archive);
    worker.open(identity);
    return { archive, transport: pull.createLocalDocumentPullTransport(worker) };
  };
  return { client, variants, ensureDocumentLayoutVariants, createLayoutServices, transportFor, release };
}

afterEach(() => {
  vi.doUnmock('./layout/error-page.js');
});

describe('DOCX pull materializers and the on-demand error page (#2248)', () => {
  it('hand out a document that parsed without loading or awaiting anything', async () => {
    const { client, variants, transportFor } = await coldModules();
    const layout = await client.materializeDocumentPullLayoutSession(transportFor().transport, identity);
    const adapter = await client.materializeDocumentPullAdapterSession(transportFor().transport, identity);
    expect(layout.fatalParse).toBeNull();
    expect(adapter.source.fatalParse).toBeNull();
    // The chunk is still stalled and was never requested by these two.
    expect(variants.errorPageLayoutReady(layout)).toBeUndefined();
  });

  it.each(['layout', 'adapter'] as const)('%s session: a failed document waits for its error page', async (kind) => {
    const { client, variants, ensureDocumentLayoutVariants, createLayoutServices, transportFor, release } = await coldModules();
    const { transport } = transportFor(MESSAGE);
    let settled = false;
    const pending = (kind === 'layout'
      ? client.materializeDocumentPullLayoutSession(transport, identity)
      : client.materializeDocumentPullAdapterSession(transport, identity).then((adapter) => adapter.source)
    ).then((source) => { settled = true; return source; });
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(settled).toBe(false);
    release();
    const source = await pending;
    expect(source.fatalParse?.message).toBe(MESSAGE);
    // The synchronous layout that follows finds the page's code.
    const services = createLayoutServices(source);
    ensureDocumentLayoutVariants(services, 0, () => source);
    expect(variants.selectDocumentLayoutPage(services, { defaultCurrentDateMs: 0 }, 0).layout.pages).toHaveLength(1);
  });

  it.each(['layout', 'adapter'] as const)('%s session: the wait ends when its owner does', async (kind) => {
    const { client, transportFor } = await coldModules();
    const { transport } = transportFor(MESSAGE);
    const owner = new AbortController();
    const options = { lifetime: () => owner.signal };
    const pending = kind === 'layout'
      ? client.materializeDocumentPullLayoutSession(transport, identity, options)
      : client.materializeDocumentPullAdapterSession(transport, identity, options);
    const outcome = pending.then(() => 'resolved', (error: Error) => error.name);
    await new Promise((resolve) => setTimeout(resolve, 20));
    owner.abort();
    // The chunk is still in flight: only the owner's wait has ended.
    await expect(outcome).resolves.toBe('AbortError');
  });

  it('rejects at once for an owner that already ended, and keeps the shared load', async () => {
    const { client, variants, transportFor, release } = await coldModules();
    const owner = new AbortController();
    owner.abort();
    await expect(client.materializeDocumentPullAdapterSession(
      transportFor(MESSAGE).transport, identity, { signal: undefined, lifetime: () => owner.signal },
    )).rejects.toMatchObject({ name: 'AbortError' });
    release();
    const other = await client.materializeDocumentPullAdapterSession(transportFor(MESSAGE).transport, identity);
    expect(variants.errorPageLayoutReady(other.source, owner.signal)).toBeUndefined();
  });
});
