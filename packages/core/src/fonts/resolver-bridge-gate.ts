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
 * - Deadlines: no clock of the document's resolver calls runs before
 *   `'ready'`. A font-source session with a resolver starts each source walk,
 *   and with it the font ceiling whose remainder each bridged call carries
 *   (`remainingMs`), only once the host is ready (`ResolverBridgeState`), and
 *   the host starts a call's deadline when it receives it. So once ready,
 *   every deadline behaves as with a static host. The wait for the host is
 *   bounded separately, by the ready ceiling: {@link FONT_LOAD_CEILING_MS}
 *   from attaching.
 * - The host chunk fails to load, or has not loaded within the ready
 *   ceiling: the bridge is unavailable for this document. One warning here;
 *   `'unavailable'` tells the worker, which fails every held and later call.
 *   Every family of the document's resolver sources then fails with
 *   {@link FONT_RESOLVER_BRIDGE_UNAVAILABLE} and paints with the host font,
 *   warned, as a run-time font-source failure does (#2203: it never rejects
 *   `load()`), whatever the realm cache holds, in the worker and (XLSX) the
 *   viewer realm alike, which follows the same state through
 *   {@link FontResolverBridge.sessionState}. A host that loads after that is
 *   not constructed, so the outcome stays one for the whole document.
 *   Messages that reach this gate meanwhile are answered as failed, so
 *   nothing waits on them.
 * - A disposed gate (the document was destroyed or replaced, also before
 *   `'ready'`) posts nothing more and never constructs a host; its worker is
 *   terminated, or reused by the next document. Every attachment posts its
 *   state (`'ready'` at once for a document without a resolver, whose gate
 *   answers stray resolve requests as failed), so nothing waits on a host
 *   that is gone: the next attachment's state supersedes the held messages,
 *   and the next parse's reset abandons their calls and lets the abandoned
 *   document's sessions waiting for the host go on, as it rejects the calls
 *   they made.
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
import type { ResolverBridgeState } from './source-session.js';

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
  /** Loads the host; defaults to the lazy import of `resolver-bridge-host.ts`
   * (replaced only by tests of a slow or failing chunk). */
  readonly loadHost?: () => Promise<FontResolverBridgeHostModule>;
}

const isNonNegativeSafeInteger = (value: unknown): value is number =>
  typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;

export class FontResolverBridge {
  private postToWorker: Post;
  private host: FontResolverBridgeHost | null = null;
  /** What the host is constructed with, until it is (or never will be): a
   * stalled import must not keep the document's sources alive. */
  private construction: Pick<FontResolverBridgeOptions, 'post' | 'sources' | 'resolverFailures'> | null = null;
  /** `none`: the document has no resolver; no host is ever loaded. */
  private state: 'none' | 'pending' | 'ready' | 'unavailable';
  /** Messages received before the host (see the module comment). */
  private early: unknown[] = [];
  private timer: ReturnType<typeof setTimeout> | undefined;
  /** The ready ceiling by wall clock as well as by timer: a busy thread may
   * run the import's continuation before an overdue timer. */
  private readyDeadline = 0;
  private disposed = false;
  private waiters: (() => void)[] = [];

  /** The host as this document's viewer-realm sessions see it (XLSX): the
   * same state the worker's sessions follow (`ResolverBridgeState`). */
  readonly sessionState: ResolverBridgeState = {
    ready: () => (this.state === 'pending' && !this.disposed
      ? new Promise<void>((resolve) => { this.waiters.push(resolve); }) : undefined),
    unavailable: () => (this.state === 'unavailable' ? FONT_RESOLVER_BRIDGE_UNAVAILABLE : undefined),
  };

  constructor(options: FontResolverBridgeOptions) {
    this.postToWorker = options.post;
    const { sources } = options;
    if (!Array.isArray(sources) || !sources.some((source) => source.kind === 'font-resolver')) {
      this.state = 'none';
      // Every attachment states its host, so a worker reused from a document
      // whose host never became ready holds nothing for this one.
      this.signal({ kind: 'ooxmlFontBridgeState', state: 'ready' });
      return;
    }
    this.state = 'pending';
    this.construction = {
      post: options.post, sources,
      ...(options.resolverFailures ? { resolverFailures: options.resolverFailures } : {}),
    };
    this.signal({ kind: 'ooxmlFontBridgeState', state: 'pending' });
    const late = `it did not load within ${FONT_LOAD_CEILING_MS} ms`;
    this.readyDeadline = Date.now() + FONT_LOAD_CEILING_MS;
    this.timer = setTimeout(() => this.unavailable(late), FONT_LOAD_CEILING_MS);
    let loading: Promise<FontResolverBridgeHostModule>;
    try {
      loading = (options.loadHost ?? (() => import('./resolver-bridge-host.js')))();
    } catch (error) {
      loading = Promise.reject(error);
    }
    loading.then(({ FontResolverBridgeHost: Host }) => {
      const construction = this.construction;
      if (this.disposed || this.state !== 'pending' || !construction) return;
      if (Date.now() >= this.readyDeadline) { this.unavailable(late); return; }
      clearTimeout(this.timer);
      this.construction = null;
      const host = new Host(construction);
      this.host = host;
      this.state = 'ready';
      this.signal({ kind: 'ooxmlFontBridgeState', state: 'ready' });
      for (const message of this.early.splice(0)) host.accept(message);
      for (const resolve of this.waiters.splice(0)) resolve();
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
    this.construction = null;
    this.state = 'unavailable';
    console.warn(`[ooxml] font resolver bridge failed to load: ${detail}; fontResolver families fall back`);
    this.signal({ kind: 'ooxmlFontBridgeState', state: 'unavailable', reason: FONT_RESOLVER_BRIDGE_UNAVAILABLE });
    for (const message of this.early.splice(0)) this.refuse(message, FONT_RESOLVER_BRIDGE_UNAVAILABLE, true);
    for (const resolve of this.waiters.splice(0)) resolve();
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
    this.construction = null;
    this.postToWorker = () => undefined;
    this.early = [];
    // Waiting sessions belong to the destroyed document: their release ends them.
    for (const resolve of this.waiters.splice(0)) resolve();
    this.host?.dispose();
    this.host = null;
  }
}
