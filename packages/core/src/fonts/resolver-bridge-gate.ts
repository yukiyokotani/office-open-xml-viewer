/**
 * Main-thread end of the `fontResolver` worker bridge for one worker-mode
 * document (`resolver-bridge.ts`): the only bridge code in a format entry's
 * static graph. Library resource policy, not OOXML behaviour.
 *
 * The host that answers the worker (`resolver-bridge-host.ts`, with the realm
 * source cache, the document's failure record and the face downloads) is
 * loaded through `import()`, and only for a document whose `fonts` include a
 * `fontResolver`: no other document ever sends the worker's resolver calls.
 * The two realms agree on when the host exists through a handshake, so the
 * host sees exactly the message sequence it would see if it were static:
 *
 * - Attaching posts `ooxmlFontBridgeState: 'pending'` before the document's
 *   parse request. From then on the worker client holds every bridge message
 *   it would post, in order (`WorkerFontResolverClient`).
 * - Once the host has loaded and been constructed, `'ready'`: the worker
 *   posts what it held, in order, then posts directly. The host therefore
 *   receives every request, acquisition and cancellation in the worker's
 *   order. Messages that reach this gate before the host (the worker posted
 *   them before it saw `'pending'`) are handed to it in arrival order too.
 * - Deadlines: a bridged call's deadline starts when the host receives it
 *   (`resolver-bridge-host.ts`), that is after `'ready'`; the `remainingMs`
 *   the worker sends is the time its session had left when it made the call.
 *   The wait for the host is bounded separately, by the ready ceiling below.
 * - The host chunk fails to load, or (DOCX, PPTX) has not loaded within
 *   {@link FONT_LOAD_CEILING_MS} of attaching: the bridge is unavailable for
 *   this document. One warning here; `'unavailable'` tells the worker, which
 *   fails every held and later call: each resolver family fails and falls
 *   back to the next source or the host font, as any run-time font-source
 *   failure does (one family's failure never rejects `load()`). A host that
 *   loads after that is not constructed, so the outcome stays one for the
 *   whole document. Messages that reach this gate meanwhile are answered as
 *   failed, so nothing waits on them.
 * - A disposed gate (the document was destroyed or replaced, also before
 *   `'ready'`) posts nothing more and never constructs a host; its worker is
 *   terminated, or reset by the next document's parse, which abandons the
 *   held calls (the next attachment's `'pending'` drops the held messages).
 *
 * XLSX passes its own loader and no ready ceiling: its viewer-realm sessions
 * share the failure record with the worker, so the host comes with the
 * workbook's font-source module (one load for both realms, see `workbook.ts`),
 * and `load()` already waits for that module, so a stalled chunk cannot leave
 * the worker waiting where the load would not.
 */
import { FONT_LOAD_CEILING_MS } from './font-registry.js';
import { isWorkerFontResolveMessage } from './resolver-bridge-messages.js';
import type { FontResolverBridgeHost } from './resolver-bridge-host.js';
import type {
  MainFontAcquireReply,
  MainFontBridgeState,
  MainFontResolveReply,
  WorkerFontResolveMessage,
} from './resolver-bridge.js';
import type { FontSourceDescriptor } from './source.js';
import type { FontResolverFailures } from './resolver-failures.js';

type Post = (message: unknown, transfer?: Transferable[]) => void;

/** The reason every call of a document whose host could not load fails with. */
export const FONT_RESOLVER_BRIDGE_UNAVAILABLE = 'font resolver bridge unavailable';

/** What the gate needs of the host module. */
export interface FontResolverBridgeHostModule {
  readonly FontResolverBridgeHost: typeof FontResolverBridgeHost;
}

export interface FontResolverBridgeOptions {
  readonly post: Post;
  /** The document's validated `fonts` option. */
  readonly sources: readonly FontSourceDescriptor[] | undefined;
  /** The document's resolver failure record shared with its viewer-realm
   * sessions (XLSX); see `resolver-failures.ts`. */
  readonly resolverFailures?: FontResolverFailures;
  /** Loads the host; defaults to the lazy import of `resolver-bridge-host.ts`. */
  readonly loadHost?: () => Promise<FontResolverBridgeHostModule>;
  /** How long the host may take to load before the bridge is unavailable;
   * `Infinity` waits for the load to settle (see the module comment). */
  readonly readyCeilingMs?: number;
}

const isNonNegativeSafeInteger = (value: unknown): value is number =>
  typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;

export class FontResolverBridge {
  private readonly postToWorker: Post;
  private host: FontResolverBridgeHost | null = null;
  /** `none`: the document has no resolver; no host is ever loaded. */
  private state: 'none' | 'pending' | 'ready' | 'unavailable';
  /** Messages received before the host (see the module comment). */
  private early: unknown[] = [];
  private timer: ReturnType<typeof setTimeout> | undefined;
  private disposed = false;

  constructor(options: FontResolverBridgeOptions) {
    this.postToWorker = options.post;
    const { sources } = options;
    if (!Array.isArray(sources) || !sources.some((source) => source.kind === 'font-resolver')) {
      this.state = 'none';
      return;
    }
    this.state = 'pending';
    this.signal({ kind: 'ooxmlFontBridgeState', state: 'pending' });
    const ceiling = options.readyCeilingMs ?? FONT_LOAD_CEILING_MS;
    if (Number.isFinite(ceiling)) {
      this.timer = setTimeout(() => this.unavailable(`it did not load within ${ceiling} ms`), ceiling);
    }
    let loading: Promise<FontResolverBridgeHostModule>;
    try {
      loading = (options.loadHost ?? (() => import('./resolver-bridge-host.js')))();
    } catch (error) {
      loading = Promise.reject(error);
    }
    loading.then(({ FontResolverBridgeHost: Host }) => {
      if (this.disposed || this.state !== 'pending') return;
      clearTimeout(this.timer);
      const host = new Host({
        post: options.post,
        sources,
        ...(options.resolverFailures ? { resolverFailures: options.resolverFailures } : {}),
      });
      this.host = host;
      this.state = 'ready';
      this.signal({ kind: 'ooxmlFontBridgeState', state: 'ready' });
      for (const message of this.early.splice(0)) host.accept(message);
    }).catch((error: unknown) => {
      this.unavailable(error instanceof Error ? error.message : String(error));
    });
  }

  /** Route `message`; returns true when it belongs to the bridge. */
  accept(message: unknown): boolean {
    if (!isWorkerFontResolveMessage(message)) return false;
    if (this.disposed) return true;
    switch (this.state) {
      case 'ready': return (this.host as FontResolverBridgeHost).accept(message);
      case 'pending': this.early.push(message); return true;
      case 'none': this.refuse(message, 'font resolver function is not available in this realm', false); return true;
      case 'unavailable': this.refuse(message, FONT_RESOLVER_BRIDGE_UNAVAILABLE, true); return true;
    }
  }

  /** The host could not load (in time): see the module comment. */
  private unavailable(detail: string): void {
    if (this.disposed || this.state !== 'pending') return;
    clearTimeout(this.timer);
    this.state = 'unavailable';
    console.warn(`[ooxml] font resolver bridge failed to load: ${detail}; fontResolver families fall back`);
    this.signal({ kind: 'ooxmlFontBridgeState', state: 'unavailable', reason: FONT_RESOLVER_BRIDGE_UNAVAILABLE });
    for (const message of this.early.splice(0)) this.refuse(message, FONT_RESOLVER_BRIDGE_UNAVAILABLE, true);
  }

  /** Answer a call no host will run, so the worker never waits for it. */
  private refuse(message: unknown, reason: string, acquisitions: boolean): void {
    const bridged = message as WorkerFontResolveMessage;
    if (!isNonNegativeSafeInteger(bridged.generation) || !isNonNegativeSafeInteger(bridged.requestId)) return;
    const { generation, requestId } = bridged;
    if (bridged.kind === 'ooxmlResolveFonts') {
      this.signal({
        kind: 'ooxmlFontsResolved', generation, requestId, outcome: { status: 'failed', reason },
      } satisfies MainFontResolveReply);
    } else if (bridged.kind === 'ooxmlAcquireFonts' && acquisitions) {
      this.signal({
        kind: 'ooxmlFontsAcquired', generation, requestId, outcome: { status: 'failed', reason },
      } satisfies MainFontAcquireReply);
    }
  }

  private signal(message: MainFontBridgeState | MainFontResolveReply | MainFontAcquireReply): void {
    try { this.postToWorker(message); } catch { /* worker terminated */ }
  }

  /** The document is destroyed or replaced. */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    clearTimeout(this.timer);
    this.early = [];
    this.host?.dispose();
    this.host = null;
  }
}
