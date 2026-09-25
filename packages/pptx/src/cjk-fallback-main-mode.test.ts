import {
  HARD_MAX_PPTX_SLIDE_JSON_BYTES,
  PULL_SESSION_PROTOCOL,
  type PullSessionCommand,
} from '@silurus/ooxml-core/worker';
import { afterEach, describe, expect, it, vi } from 'vitest';

/**
 * A main-mode load used to drop the resolved regional Han fallback on the way
 * to the preflight that produces its font-preload set: an ambiguous-Han deck
 * preloaded the `jp` default while the renderer painted with the resolved
 * region's stack, so shared Han took Japanese glyph variants out of the only
 * family that had been fetched. `worker-protocol.ts` records why the region
 * has to cross this protocol at all.
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

/** Ambiguous deck: shared Han text under theme fonts that name no CJK region. */
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
      elements: [{
        type: 'shape',
        textBody: { paragraphs: [{ runs: [{ type: 'text', text: '漢字' }] }] },
      }],
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

describe('PPTX main-mode cjkFallback', () => {
  it('preloads the region the parse request carried, not the jp default', async () => {
    const worker = await startWorker();
    send(worker, { kind: 'init', wasmUrl: 'x' });
    send(worker, { kind: 'parse', id: 1, buffer: new ArrayBuffer(4), resourcePolicy, cjkFallback: 'sc' });
    await reply(worker, 'presentationOpened');

    send(worker, { kind: 'openSlideSession', id: 2, slideIndex: 0, ...identity });
    await reply(worker, 'slideSessionOpened');
    send(worker, pullCommand(1, {
      kind: 'pull', sequence: 0, byteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
    }));
    await reply(worker, 'chunk');
    send(worker, pullCommand(2, { kind: 'ack', sequence: 0 }));
    await reply(worker, 'accepted');

    send(worker, { kind: 'finishPresentationPreflight', id: 3 });
    const ready = await reply<{ kind: string; preflight: { fontPreloadNames: string[] } }>(
      worker,
      'presentationPreflightReady',
    );
    expect(ready.preflight.fontPreloadNames).toEqual(['Noto Sans SC', 'Noto Serif SC']);
  });
});
