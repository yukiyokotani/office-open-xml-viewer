import {
  HARD_MAX_PPTX_SLIDE_JSON_BYTES,
  PULL_SESSION_PROTOCOL,
  type PullSessionCommand,
} from '@silurus/ooxml-core/worker';
import { afterEach, describe, expect, it, vi } from 'vitest';

/**
 * The main thread sends its host locale with every worker parse request
 * (presentation.ts `hostLocale: hostLocale()`); the real worker handler adopts
 * it before the font-demand collector sees a chart, so a grouped chart label
 * demands the locale's group separator (U+202F for fr-FR) even though this
 * worker's own navigator is en-US. Dropping the field from the request makes
 * the demand fall back to the worker's own locale.
 */

const resourcePolicy = {
  maxArchiveEntryBytes: null,
  maxTotalInflatedBytes: null,
  maxArchiveEntries: null,
} as const;

const identity = { sessionId: 1, operationId: 1, generation: 1 } as const;

function encode(value: unknown): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(value));
}

/** One-slide deck holding a chart. */
class FakePptxArchive {
  presentation_bootstrap(): Uint8Array {
    return encode({
      slideCount: 1,
      slideWidth: 12_192_000,
      slideHeight: 6_858_000,
      defaultTextColor: null,
      majorFont: null,
      minorFont: null,
      hlinkColor: null,
      folHlinkColor: null,
      embeddedFonts: [],
      slides: [{ index: 0, partName: 'ppt/slides/slide1.xml' }],
    });
  }

  pull_slide(): Uint8Array {
    return encode({
      index: 0,
      slideNumber: 1,
      partName: 'ppt/slides/slide1.xml',
      background: null,
      elements: [{ type: 'chart', chart: { chartType: 'bar', categories: [], series: [] } }],
    });
  }

  slide_cursor_resource_usage(): Uint8Array {
    return encode({
      archiveEntryCount: 1,
      declaredInflatedBytes: 2,
      distinctInflatedBytes: 3,
      operationInflatedBytes: 4,
    });
  }

  acknowledge_slide(): void {}
  cancel_slide(): void {}
  close_presentation_session(): void {}
  assert_healthy(): void {}
  free(): void {}
}

vi.mock('./wasm/pptx_parser.js', () => ({
  default: () => Promise.resolve(),
  reinit: () => Promise.resolve(),
  PptxArchive: FakePptxArchive,
}));

interface FakeSelf {
  onmessage: ((event: MessageEvent) => void) | null;
  posted: { kind: string }[];
  postMessage: (message: unknown) => void;
}

/** Import the worker fresh (its top-level `self.onmessage = …` runs on import),
 *  after `self` and the WASM mock are installed — the pattern used by the
 *  worker init-hang tests, so no real Worker or WASM is needed. */
async function startWorker(): Promise<FakeSelf> {
  const posted: { kind: string }[] = [];
  const fake: FakeSelf = {
    onmessage: null,
    posted,
    postMessage: (message) => { posted.push(message as { kind: string }); },
  };
  vi.stubGlobal('self', fake);
  vi.resetModules();
  await import('./worker.js');
  return fake;
}

function send(worker: FakeSelf, data: unknown): void {
  worker.onmessage?.({ data } as MessageEvent);
}

function pullCommand(
  requestId: number,
  body: { kind: 'pull'; sequence: number; byteCredit: number } | { kind: 'ack'; sequence: number },
): PullSessionCommand<number> {
  return { protocol: PULL_SESSION_PROTOCOL, requestId, ...identity, ...body };
}

async function reply<T extends { kind: string }>(worker: FakeSelf, kind: string): Promise<T> {
  return vi.waitFor(() => {
    const message = worker.posted.find((posted) => posted.kind === kind);
    expect(message, `worker never responded with ${kind}: ${JSON.stringify(worker.posted)}`)
      .toBeDefined();
    return message as T;
  }, { interval: 1 });
}

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const FONT_DEMAND_SIDECAR = new URL('./font-preload-demand.ts', import.meta.url).href;
const NARROW_NO_BREAK_SPACE = 0x202f;

async function fontDemandFor(request: Record<string, unknown>): Promise<unknown> {
  const worker = await startWorker();
  send(worker, { kind: 'init', wasmUrl: 'x' });
  send(worker, {
    kind: 'parse', id: 1, buffer: new ArrayBuffer(4), resourcePolicy, cjkFallback: 'jp',
    collectFontDemand: FONT_DEMAND_SIDECAR, ...request,
  });
  await reply(worker, 'presentationOpened');
  send(worker, { kind: 'openSlideSession', id: 2, slideIndex: 0, ...identity });
  await reply(worker, 'slideSessionOpened');
  send(worker, pullCommand(1, { kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES }));
  await reply(worker, 'chunk');
  send(worker, pullCommand(2, { kind: 'ack', sequence: 0 }));
  await reply(worker, 'accepted');
  send(worker, { kind: 'finishPresentationPreflight', id: 3 });
  const ready = await reply<{ kind: string; preflight: { fontPreloadDemand?: number[] } }>(
    worker, 'presentationPreflightReady',
  );
  return ready.preflight.fontPreloadDemand;
}

describe('PPTX worker host locale', () => {
  it('adopts the parse request locale before chart font demand', async () => {
    expect(await fontDemandFor({ hostLocale: 'fr-FR' })).toContain(NARROW_NO_BREAK_SPACE);
  });

  it('keeps the worker own locale when the request carries none', async () => {
    expect(await fontDemandFor({})).not.toContain(NARROW_NO_BREAK_SPACE);
  });
});

describe('PPTX loader parse request', () => {
  it.each(['worker', 'main'] as const)('sends the main thread host locale with the %s-mode parse request', async (mode) => {
    const { PptxPresentation } = await import('./presentation.js');
    const posted: Record<string, unknown>[] = [];
    vi.stubGlobal('navigator', { language: 'fr-FR' });
    vi.stubGlobal('location', { href: 'http://localhost/' });
    vi.stubGlobal('OffscreenCanvas', class {});
    vi.stubGlobal('Worker', class {
      postMessage(message: Record<string, unknown>): void { posted.push(message); }
      addEventListener(): void {}
      removeEventListener(): void {}
      terminate(): void {}
    });
    const loading = PptxPresentation.load(new ArrayBuffer(8), { mode }).catch(() => undefined);
    await vi.waitFor(() => expect(posted.some(message => message.kind === 'parse')).toBe(true));
    expect(posted.find(message => message.kind === 'parse')).toMatchObject({ hostLocale: 'fr-FR' });
    void loading;
  });
});
