import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import { OoxmlError } from '../errors/ooxml-error.js';
import { HARD_MAX_FONT_SOURCE_DECLARATIONS, HARD_MAX_FONT_SOURCE_FACE_BYTES } from '../worker/resource-policy.generated.js';
import { WorkerFontResolverClient } from './resolver-bridge.js';
import { liveSourceCacheCount } from './source-cache.js';
import { FontSourceSession, retainedFontSourceBytes, type FontSessionDemand } from './source-session.js';
import {
  assertFontSourceDescriptors,
  type FontDemand,
  type FontFaceDeclaration,
  type FontResolveContext,
  type FontResolveFn,
  type FontResolverOptions,
  type FontSource,
} from './source.js';

// The worker bridge of `fontResolver` (design §3.4, §4.3–§4.5). Each realm is
// its own instance of the module graph, with its own caches, registries and
// byte totals, as a window and its render workers are: the statically
// imported modules above are worker realm 1, `worker2` a second worker (a
// second worker-mode document), `main` the window. Messages cross a real
// structuredClone with the sender's transfer list (a transferred ArrayBuffer
// is detached on the sending side) and are delivered asynchronously.

interface WorkerRealm {
  readonly FontSourceSession: typeof FontSourceSession;
  readonly WorkerFontResolverClient: typeof WorkerFontResolverClient;
  readonly assertFontSourceDescriptors: typeof assertFontSourceDescriptors;
  readonly retainedFontSourceBytes: () => number;
  readonly liveSourceCacheCount: () => number;
}
interface MainRealm extends WorkerRealm {
  readonly fontResolver: typeof import('./source.js').fontResolver;
  readonly FontResolverBridgeHost: typeof import('./resolver-bridge-host.js').FontResolverBridgeHost;
  readonly FontResolverResponder: typeof import('./resolver-bridge-responder.js').FontResolverResponder;
  readonly copyBridgedFontDemand: typeof import('./resolver-bridge-responder.js').copyBridgedFontDemand;
  readonly sourceCacheEntryCount: () => number;
  readonly responderModule: typeof import('./resolver-bridge-responder.js');
  readonly acquireFaceUrls: typeof import('./face-download.js').acquireFaceUrls;
  readonly releaseAcquiredFaces: typeof import('./face-download.js').releaseAcquiredFaces;
  readonly acquiredFaceCount: typeof import('./face-download.js').acquiredFaceCount;
  readonly inRealmResolverFaceAcquirer: typeof import('./face-download.js').inRealmResolverFaceAcquirer;
  readonly acquireResolverFacesInRealm: typeof import('./face-download.js').acquireResolverFacesInRealm;
  readonly FONT_DEADLINE_REASON: symbol;
}
const worker1: WorkerRealm = {
  FontSourceSession, WorkerFontResolverClient, assertFontSourceDescriptors, retainedFontSourceBytes, liveSourceCacheCount,
};
let worker2: WorkerRealm;
let main: MainRealm;

async function loadRealm() {
  vi.resetModules();
  const [source, session, bridge, host, responder, cache, acquisition] = await Promise.all([
    import('./source.js'), import('./source-session.js'), import('./resolver-bridge.js'),
    import('./resolver-bridge-host.js'),
    import('./resolver-bridge-responder.js'),
    import('./source-cache.js'),
    import('./face-download.js'),
  ]);
  return {
    FontSourceSession: session.FontSourceSession,
    WorkerFontResolverClient: bridge.WorkerFontResolverClient,
    assertFontSourceDescriptors: source.assertFontSourceDescriptors,
    retainedFontSourceBytes: session.retainedFontSourceBytes,
    liveSourceCacheCount: cache.liveSourceCacheCount,
    fontResolver: source.fontResolver,
    FontResolverBridgeHost: host.FontResolverBridgeHost,
    FontResolverResponder: responder.FontResolverResponder,
    copyBridgedFontDemand: responder.copyBridgedFontDemand,
    sourceCacheEntryCount: cache.sourceCacheEntryCount,
    responderModule: responder,
    acquireFaceUrls: acquisition.acquireFaceUrls,
    releaseAcquiredFaces: acquisition.releaseAcquiredFaces,
    acquiredFaceCount: acquisition.acquiredFaceCount,
    inRealmResolverFaceAcquirer: acquisition.inRealmResolverFaceAcquirer,
    acquireResolverFacesInRealm: acquisition.acquireResolverFacesInRealm,
    FONT_DEADLINE_REASON: acquisition.FONT_DEADLINE_REASON,
  };
}
beforeAll(async () => {
  worker2 = await loadRealm();
  // Main last: the host's lazy responder import resolves in this graph.
  main = await loadRealm();
});
const fontResolver = (resolve: FontResolveFn, options?: FontResolverOptions) => main.fontResolver(resolve, options);

const G = globalThis as Record<string, unknown>;
const ORIGINAL_FONT_FACE = G.FontFace;

class FakeFontFace {
  status: FontFaceLoadStatus = 'unloaded';
  readonly unicodeRange: string;
  readonly bytes: number;
  constructor(readonly family: string, source: string | ArrayBuffer, descriptors: FontFaceDescriptors = {}) {
    this.unicodeRange = descriptors.unicodeRange ?? 'U+0-10FFFF';
    this.bytes = typeof source === 'string' ? 0 : source.byteLength;
  }
  load(): Promise<FakeFontFace> {
    this.status = 'loaded';
    return Promise.resolve(this);
  }
}

function fakeSet() {
  const faces: FakeFontFace[] = [];
  return {
    faces,
    add: (face: FakeFontFace) => { faces.push(face); },
    delete: (face: FakeFontFace) => { const at = faces.indexOf(face); if (at >= 0) faces.splice(at, 1); return at >= 0; },
  };
}
const asSet = (set: ReturnType<typeof fakeSet>) => set as unknown as FontFaceSet;

const demand = (families: string[], scalars: FontSessionDemand['scalars'] = [0x41]): FontSessionDemand =>
  ({ families, fallbackFamilies: [], scalars, substituteScripts: Object.fromEntries(families.map((family) => [family.toLowerCase(), []])) });

type Post = (message: unknown, transfer?: Transferable[]) => void;

const sessions: { release(): void }[] = [];
const hosts: { dispose(): void }[] = [];

/** One worker-mode document: its main-side host and its worker realm,
 * joined by structured clone. `toMain` / `toWorker` record the traffic. */
function bridge(
  sources: readonly FontSource[],
  options: {
    realm?: WorkerRealm;
    failures?: Map<string, string>;
    loadResponder?: () => Promise<typeof import('./resolver-bridge-responder.js')>;
  } = {},
) {
  const realm = options.realm ?? worker1;
  const toMain: unknown[] = [];
  const toWorker: unknown[] = [];
  let hold = false;
  const held: (() => void)[] = [];
  const deliver = (fn: () => void, toWorkerRealm: boolean) => {
    if (hold && toWorkerRealm) held.push(fn);
    else void Promise.resolve().then(fn);
  };
  let client: WorkerFontResolverClient;
  const mainPost: Post = (message, transfer) => {
    const copy = structuredClone(message, { transfer: transfer as Transferable[] });
    toWorker.push(copy);
    deliver(() => client.accept(copy), true);
  };
  const host = new main.FontResolverBridgeHost({
    post: mainPost, sources, ...(options.failures ? { resolverFailures: options.failures } : {}),
    ...(options.loadResponder ? { loadResponder: options.loadResponder } : {}),
  });
  hosts.push(host);
  const workerPost: Post = (message, transfer) => {
    const copy = structuredClone(message, { transfer: transfer as Transferable[] });
    toMain.push(copy);
    deliver(() => host.accept(copy), false);
  };
  client = new realm.WorkerFontResolverClient(workerPost);
  // The worker receives the descriptors as a structured clone and validates them.
  const workerSources = realm.assertFontSourceDescriptors(structuredClone(sources));
  return {
    client, host, toMain, toWorker,
    /** Hold main → worker replies until `resume()` (a busy worker). */
    pause() { hold = true; },
    resume() { hold = false; for (const fn of held.splice(0)) void Promise.resolve().then(fn); },
    session(set = fakeSet()) {
      const session = new realm.FontSourceSession({ sources: workerSources, fontSet: asSet(set), ...client.sessionOptions });
      sessions.push(session);
      return { session, set };
    },
  };
}

/** A main-realm (viewer or main-mode) session. */
function mainSession(sources: readonly FontSource[], failures?: Map<string, string>, set = fakeSet()) {
  const session = new main.FontSourceSession({ sources, fontSet: asSet(set), ...(failures ? { resolverFailures: failures } : {}) });
  sessions.push(session);
  return { session, set };
}

const FACE = (bytes = 64, family = 'Private Face'): FontFaceDeclaration => ({ family, data: new Uint8Array(bytes).fill(7).buffer });

let warn: ReturnType<typeof vi.spyOn>;
beforeEach(() => {
  G.FontFace = FakeFontFace;
  warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
});
afterEach(async () => {
  vi.useRealTimers();
  for (const session of sessions.splice(0)) session.release();
  for (const host of hosts.splice(0)) host.dispose();
  // A host disposes its (lazily loaded) responder asynchronously.
  await new Promise((resolve) => setTimeout(resolve, 0));
  // Every realm returns everything it took.
  for (const realm of [worker1, worker2, main]) {
    expect(realm.liveSourceCacheCount()).toBe(0);
    expect(realm.retainedFontSourceBytes()).toBe(0);
  }
  G.FontFace = ORIGINAL_FONT_FACE;
  vi.restoreAllMocks();
});

describe('fontResolver bridge: main and worker sessions resolve alike', () => {
  it('registers and charges the same faces as a direct session; main keeps the original', async () => {
    const answered: ArrayBuffer[] = [];
    const seen: FontDemand[] = [];
    const resolve: FontResolveFn = (request) => {
      seen.push(request);
      const face = FACE(64);
      answered.push(face.data as ArrayBuffer);
      return request.families.map(() => face);
    };
    const source = fontResolver(resolve);
    const { session: direct } = mainSession([source]);
    const directLoaded = await direct.ensure(demand(['Private Face']));
    const directBytes = main.retainedFontSourceBytes();
    direct.release();
    expect(main.retainedFontSourceBytes()).toBe(0);

    const link = bridge([source]);
    const { session, set } = link.session();
    const loaded = await session.ensure(demand(['Private Face']));
    expect(loaded.faces.map(({ family }) => family)).toEqual(directLoaded.faces.map(({ family }) => family));
    expect(set.faces.map((face) => face.bytes)).toEqual([64]);
    // Each realm charges what it keeps, once: the worker as a direct session
    // does (its cached copy and its registration), main only the cached
    // original it shares with later documents (it registers nothing).
    expect(retainedFontSourceBytes()).toBe(directBytes);
    expect(main.retainedFontSourceBytes()).toBe(64);
    expect(answered[1]?.byteLength).toBe(64);
    // The application saw the same demand in both modes: family names only.
    expect(seen[1]).toEqual(seen[0]);
    expect(seen[0]).toEqual({ families: ['Private Face'], fallbackFamilies: [] });
    session.release();
    link.host.dispose();
    await vi.waitFor(() => expect(main.retainedFontSourceBytes()).toBe(0));
  });

  it('sends one copy per distinct buffer', async () => {
    const face = FACE(32);
    const source = fontResolver(() => [face, { ...face, weight: '700' }]);
    const link = bridge([source]);
    const { set, session } = link.session();
    await session.ensure(demand(['Private Face']));
    expect((face.data as ArrayBuffer).byteLength).toBe(32);
    expect(set.faces.map((entry) => entry.bytes)).toEqual([32, 32]);
    const reply = link.toWorker[0] as { outcome: { families: { declarations: FontFaceDeclaration[] }[] } };
    const [first, second] = reply.outcome.families[0]?.declarations ?? [];
    expect(first?.data).toBe(second?.data);
  });

  for (const order of ['worker first', 'viewer first', 'concurrently'] as const) {
    it(`runs the resolver once per family when the viewer realm also paints (${order})`, async () => {
      let calls = 0;
      const source = fontResolver((request) => { calls++; return request.families.map((family) => FACE(16, family)); });
      const failures = new Map<string, string>();
      const link = bridge([source], { failures });
      const viewer = mainSession([source], failures);
      const worker = link.session();
      const ensureViewer = () => viewer.session.ensure(demand(['Private Face']));
      const ensureWorker = () => worker.session.ensure(demand(['Private Face']));
      if (order === 'worker first') { await ensureWorker(); await ensureViewer(); }
      else if (order === 'viewer first') { await ensureViewer(); await ensureWorker(); }
      else await Promise.all([ensureWorker(), ensureViewer()]);
      expect(calls).toBe(1);
      expect(viewer.set.faces.map((face) => face.bytes)).toEqual([16]);
      expect(worker.set.faces.map((face) => face.bytes)).toEqual([16]);
      await Promise.all([worker.session.ensure(demand(['Second Face'])), viewer.session.ensure(demand(['Second Face']))]);
      expect(calls).toBe(2);
      // Main keeps the entries while either the viewer or the bridge holds them.
      viewer.session.release();
      expect(main.liveSourceCacheCount()).toBe(1);
      link.host.dispose();
      await vi.waitFor(() => expect(main.liveSourceCacheCount()).toBe(0));
    });
  }

  it('joins a main-mode document\'s resolution and never detaches its cached buffer', async () => {
    let calls = 0;
    const face = FACE(16);
    const source = fontResolver(() => { calls++; return [face]; });
    await mainSession([source]).session.ensure(demand(['Private Face']));
    const link = bridge([source]);
    const { set, session } = link.session();
    await session.ensure(demand(['Private Face']));
    expect(calls).toBe(1);
    expect(set.faces.map((entry) => entry.bytes)).toEqual([16]);
    expect((face.data as ArrayBuffer).byteLength).toBe(16);
  });
});

describe('fontResolver bridge: review findings', () => {
  // Finding 1: two worker-mode documents on one family and one fresh buffer.
  for (const timing of ['concurrently', 'sequentially while both are live'] as const) {
    it(`shares one resolution between two worker documents (${timing}) without detaching it`, async () => {
      let calls = 0;
      const face = FACE(24);
      const source = fontResolver(() => { calls++; return [face]; });
      const first = bridge([source]);
      const second = bridge([source], { realm: worker2 });
      const a = first.session();
      const b = second.session();
      if (timing === 'concurrently') {
        await Promise.all([a.session.ensure(demand(['Private Face'])), b.session.ensure(demand(['Private Face']))]);
      } else {
        await a.session.ensure(demand(['Private Face']));
        await b.session.ensure(demand(['Private Face']));
      }
      expect(calls).toBe(1);
      expect(a.set.faces.map((entry) => entry.bytes)).toEqual([24]);
      expect(b.set.faces.map((entry) => entry.bytes)).toEqual([24]);
      expect((face.data as ArrayBuffer).byteLength).toBe(24);
      expect(warn).not.toHaveBeenCalled();
    });
  }

  // Finding 2: two configurations of one function.
  it('names the exact source, so a substitute configuration of the same function is honoured', async () => {
    const resolve: FontResolveFn = () => [FACE(8, 'Served')];
    const fonts = [fontResolver(resolve), fontResolver(resolve, { substitutes: [{ for: 'Missing', use: 'Served' }] })];
    const { session: direct } = mainSession(fonts);
    const mainLoaded = await direct.ensure(demand(['Missing']));
    direct.release();
    const link = bridge(fonts);
    const workerLoaded = await link.session().session.ensure(demand(['Missing']));
    expect(mainLoaded.substitutions.get('missing')).toEqual({ family: 'Served' });
    expect(workerLoaded.substitutions.get('missing')).toEqual(mainLoaded.substitutions.get('missing'));
    expect(workerLoaded.faces.map(({ family, sourceIndex }) => [family, sourceIndex]))
      .toEqual(mainLoaded.faces.map(({ family, sourceIndex }) => [family, sourceIndex]));
    const request = link.toMain[1] as { sourceIndex: number };
    expect(request.sourceIndex).toBe(1);
  });

  // Finding 3: a resolver that fails once, then succeeds.
  for (const order of ['viewer first', 'worker first'] as const) {
    it(`gives a workbook's viewer and worker one outcome when the resolver fails once (${order})`, async () => {
      let calls = 0;
      const source = fontResolver(() => {
        calls++;
        if (calls === 1) throw new Error('flaky backend');
        return [FACE(8)];
      });
      const failures = new Map<string, string>();
      const link = bridge([source], { failures });
      const viewer = mainSession([source], failures);
      const worker = link.session();
      const ensureViewer = () => viewer.session.ensure(demand(['Private Face']));
      const ensureWorker = () => worker.session.ensure(demand(['Private Face']));
      const [firstLoaded, secondLoaded] = order === 'viewer first'
        ? [await ensureViewer(), await ensureWorker()]
        : [await ensureWorker(), await ensureViewer()];
      expect(firstLoaded.faces).toEqual([]);
      expect(secondLoaded.faces).toEqual([]);
      expect(calls).toBe(1);
      // Another document is not bound by this one's failure: it retries.
      const other = mainSession([source]);
      expect((await other.session.ensure(demand(['Private Face']))).faces).toHaveLength(1);
      expect(calls).toBe(2);
    });
  }

  // Round 2, finding 1: a timeout is the document's outcome too.
  for (const order of ['worker first', 'viewer first', 'overlapping, answer after the worker deadline'] as const) {
    it(`gives a workbook's viewer and worker one outcome when the resolver times out (${order})`, async () => {
      vi.useFakeTimers();
      let calls = 0;
      const source = fontResolver(() => {
        calls++;
        // The first call answers after 17 s (or never); a retry would succeed at once.
        if (calls > 1) return [FACE(8)];
        return order === 'overlapping, answer after the worker deadline'
          ? new Promise((resolve) => setTimeout(() => resolve([FACE(8)]), 17_000))
          : new Promise(() => undefined);
      });
      const failures = new Map<string, string>();
      const link = bridge([source], { failures });
      const viewer = mainSession([source], failures);
      const worker = link.session();
      const ensureViewer = () => viewer.session.ensure(demand(['Private Face']));
      const ensureWorker = () => worker.session.ensure(demand(['Private Face']));
      let viewerLoaded: Awaited<ReturnType<typeof ensureViewer>>;
      let workerLoaded: Awaited<ReturnType<typeof ensureWorker>>;
      if (order === 'worker first') {
        const pending = ensureWorker();
        await vi.advanceTimersByTimeAsync(15_000);
        workerLoaded = await pending;
        viewerLoaded = await ensureViewer();
      } else if (order === 'viewer first') {
        const pending = ensureViewer();
        await vi.advanceTimersByTimeAsync(15_000);
        viewerLoaded = await pending;
        workerLoaded = await ensureWorker();
      } else {
        // The worker asks at 0 s, the viewer joins at 5 s; main times the
        // worker's call out at 15 s, the answer arrives at 17 s, within the
        // viewer's own ceiling (20 s).
        const workerPending = ensureWorker();
        await vi.advanceTimersByTimeAsync(5_000);
        const viewerPending = ensureViewer();
        await vi.advanceTimersByTimeAsync(13_000);
        [workerLoaded, viewerLoaded] = await Promise.all([workerPending, viewerPending]);
      }
      expect(calls).toBe(1);
      expect(workerLoaded.faces).toEqual([]);
      expect(viewerLoaded.faces).toEqual([]);
      expect(viewer.set.faces.filter((face) => face.status === 'loaded')).toEqual([]);
      expect(String(warn.mock.calls.at(-1)?.[0])).toContain('Private Face (timed out)');
    });
  }

  // Round 3, finding 1: the deadline is cumulative over the source walk.
  it('keeps main mode\'s per-ensure deadline across resolvers in worker mode', async () => {
    vi.useFakeTimers();
    const later = <T>(ms: number, value: T) => new Promise<T>((resolve) => setTimeout(() => resolve(value), ms));
    const fonts = [
      fontResolver(() => later(10_000, [])),
      fontResolver(() => later(10_000, [FACE(8)])),
    ];
    const run = async (session: { ensure(request: FontSessionDemand): Promise<{ faces: readonly unknown[] }> }) => {
      const pending = session.ensure(demand(['Private Face']));
      await vi.advanceTimersByTimeAsync(20_000);
      return pending;
    };
    const direct = mainSession(fonts).session;
    const mainLoaded = await run(direct);
    // A separate document: nothing of the main-mode run stays cached.
    direct.release();
    const workerLoaded = await run(bridge(fonts).session().session);
    // Main mode: the first resolver used 10 s, the second had 5 s left.
    expect(mainLoaded.faces).toEqual([]);
    expect(workerLoaded.faces).toEqual(mainLoaded.faces);
    expect(warn.mock.calls.filter(([message]: unknown[]) => String(message).includes('Private Face (timed out)'))).toHaveLength(2);
  });

  // Round 3, finding 2: a bridge that cannot load is the document's outcome.
  it('records a failed responder bootstrap so the viewer realm reports the same', async () => {
    let calls = 0;
    const source = fontResolver(() => { calls++; return [FACE(8)]; });
    const failures = new Map<string, string>();
    const link = bridge([source], {
      failures,
      // The lazily loaded chunk is missing (a 404 behind a stale deploy).
      loadResponder: () => Promise.reject(new Error('Failed to fetch dynamically imported module')),
    });
    const workerLoaded = await link.session().session.ensure(demand(['Private Face']));
    expect(workerLoaded.faces).toEqual([]);
    expect(String(warn.mock.calls.find(([message]: unknown[]) => String(message).includes('bridge failed to load'))?.[0]))
      .toContain('Failed to fetch');
    const viewer = mainSession([source], failures);
    expect((await viewer.session.ensure(demand(['Private Face']))).faces).toEqual([]);
    expect(calls).toBe(0);
    expect(String(warn.mock.calls.at(-1)?.[0])).toContain('Private Face (font resolver bridge unavailable)');
  });

  // Round 4: the responder's bootstrap counts against the call's deadline.
  for (const repro of ['10 s bootstrap, 10 s resolver', '35 s bootstrap, immediate resolver'] as const) {
    it(`starts a bridged call's deadline on arrival (${repro})`, async () => {
      vi.useFakeTimers();
      const slowBootstrap = repro.startsWith('10') ? 10_000 : 35_000;
      let calls = 0;
      const source = fontResolver(() => {
        calls++;
        return repro.startsWith('10')
          ? new Promise((resolve) => setTimeout(() => resolve([FACE(8)]), 10_000))
          : [FACE(8)];
      });
      const failures = new Map<string, string>();
      const link = bridge([source], {
        failures,
        loadResponder: () => new Promise((resolve) => setTimeout(() => resolve(main.responderModule), slowBootstrap)),
      });
      let settled = false;
      const pending = link.session().session.ensure(demand(['Private Face'])).then((loaded) => { settled = true; return loaded; });
      await vi.advanceTimersByTimeAsync(15_000);
      // Answered at the 15 s ceiling, as main mode times out: never by the
      // worker's backstop, and never with the face.
      expect(settled).toBe(true);
      expect((await pending).faces).toEqual([]);
      expect(String(warn.mock.calls.at(-1)?.[0])).toContain('Private Face (timed out)');
      await vi.advanceTimersByTimeAsync(30_000);
      // An expired call is not run once the responder loads (35 s case); the
      // 10 s case ran it with the 5 s left.
      expect(calls).toBe(repro.startsWith('10') ? 1 : 0);
      // The viewer realm of the same workbook reports the same outcome.
      const viewer = mainSession([source], failures);
      expect((await viewer.session.ensure(demand(['Private Face']))).faces).toEqual([]);
      expect(calls).toBe(repro.startsWith('10') ? 1 : 0);
    });
  }

  // Round 5: one outcome per family per document, success included,
  // whatever the bridge's bootstrap state.
  const delayedResponder = (ms: number) => () =>
    new Promise<typeof main.responderModule>((resolve) => setTimeout(() => resolve(main.responderModule), ms));
  const missingResponder = () => Promise.reject(new Error('Failed to fetch dynamically imported module'));
  for (const [label, loadResponder] of [['delayed 17 s', delayedResponder(17_000)], ['missing', missingResponder]] as const) {
    it(`gives the worker the viewer realm's loaded face while the responder is ${label}`, async () => {
      vi.useFakeTimers();
      let calls = 0;
      const source = fontResolver(() => { calls++; return [FACE(12)]; });
      const failures = new Map<string, string>();
      const viewer = mainSession([source], failures);
      expect((await viewer.session.ensure(demand(['Private Face']))).faces).toHaveLength(1);
      const link = bridge([source], { failures, loadResponder });
      const worker = link.session();
      const pending = worker.session.ensure(demand(['Private Face']));
      await vi.advanceTimersByTimeAsync(0);
      // Answered from the document's settled outcome, before any deadline.
      expect((await pending).faces).toHaveLength(1);
      expect(worker.set.faces.map((face) => face.bytes)).toEqual([12]);
      expect(calls).toBe(1);
      await vi.advanceTimersByTimeAsync(20_000);
    });

    it(`fails the worker without a call after the viewer realm failed, while the responder is ${label}`, async () => {
      vi.useFakeTimers();
      let calls = 0;
      const source = fontResolver(() => { calls++; if (calls === 1) throw new Error('flaky backend'); return [FACE(12)]; });
      const failures = new Map<string, string>();
      const viewer = mainSession([source], failures);
      expect((await viewer.session.ensure(demand(['Private Face']))).faces).toEqual([]);
      const link = bridge([source], { failures, loadResponder });
      const pending = link.session().session.ensure(demand(['Private Face']));
      await vi.advanceTimersByTimeAsync(0);
      expect((await pending).faces).toEqual([]);
      expect(String(warn.mock.calls.at(-1)?.[0])).toContain('Private Face (flaky backend)');
      expect(calls).toBe(1);
      await vi.advanceTimersByTimeAsync(20_000);
    });
  }

  it('adopts the viewer realm\'s outcome reached while the request waits for the responder', async () => {
    vi.useFakeTimers();
    let calls = 0;
    const source = fontResolver(() => { calls++; return [FACE(12)]; });
    const failures = new Map<string, string>();
    const link = bridge([source], { failures, loadResponder: delayedResponder(17_000) });
    const worker = link.session();
    const pending = worker.session.ensure(demand(['Private Face']));
    await vi.advanceTimersByTimeAsync(1_000);
    // The viewer realm resolves and loads the face while the request waits.
    expect((await mainSession([source], failures).session.ensure(demand(['Private Face']))).faces).toHaveLength(1);
    // At the request's deadline it adopts that success instead of failing.
    await vi.advanceTimersByTimeAsync(14_000);
    expect((await pending).faces).toHaveLength(1);
    expect(calls).toBe(1);
    await vi.advanceTimersByTimeAsync(5_000);
  });

  // Round 6: an answer taken from the cache is held by the receiving document.
  for (const producer of ['a main-mode document', 'another worker-mode document'] as const) {
    it(`keeps a family answered from ${producer}'s resolution for the receiving document`, async () => {
      let calls = 0;
      const source = fontResolver(() => {
        calls++;
        if (calls > 1) throw new Error('token expired');
        return [FACE(12)];
      });
      // Document X resolves the family first.
      let releaseX: () => void;
      if (producer === 'a main-mode document') {
        const x = mainSession([source]);
        expect((await x.session.ensure(demand(['Private Face']))).faces).toHaveLength(1);
        releaseX = () => x.session.release();
      } else {
        const x = bridge([source]);
        const xWorker = x.session();
        expect((await xWorker.session.ensure(demand(['Private Face']))).faces).toHaveLength(1);
        releaseX = () => { xWorker.session.release(); x.host.dispose(); };
      }
      // Workbook Y (worker mode) receives it from the cache, its responder
      // chunk unavailable, so only the host's settled answer can serve it.
      const failures = new Map<string, string>();
      // Its own worker (each worker-mode document has one).
      const y = bridge([source], { failures, loadResponder: missingResponder, realm: worker2 });
      const yWorker = y.session();
      expect((await yWorker.session.ensure(demand(['Private Face']))).faces).toHaveLength(1);
      // X goes away; Y's viewer realm must still see Y's outcome.
      releaseX();
      await new Promise((resolve) => setTimeout(resolve, 0));
      const yViewer = mainSession([source], failures);
      expect((await yViewer.session.ensure(demand(['Private Face']))).faces).toHaveLength(1);
      expect(calls).toBe(1);
      // Everything returns to 0/0 once Y is gone too (checked after each test).
      yViewer.session.release();
      yWorker.session.release();
      y.host.dispose();
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(main.liveSourceCacheCount()).toBe(0);
      expect(main.retainedFontSourceBytes()).toBe(0);
    });
  }

  // Round 7: a batch ends per family; decided families keep their outcome.
  const recordedFamilies = (failures: Map<string, string>) =>
    [...failures.keys()].map((key) => key.split('\u0000')[1]).sort();
  it('keeps a family the worker already resolved when the viewer\'s batch times out', async () => {
    vi.useFakeTimers();
    let calls = 0;
    const source = fontResolver((request) => {
      calls++;
      if (request.families.includes('Slow')) return new Promise(() => undefined);
      return request.families.map((family) => FACE(8, family));
    });
    const failures = new Map<string, string>();
    const worker = bridge([source], { failures }).session();
    expect((await worker.session.ensure(demand(['Fast']))).faces).toHaveLength(1);
    const viewer = mainSession([source], failures);
    const pending = viewer.session.ensure(demand(['Fast', 'Slow']));
    await vi.advanceTimersByTimeAsync(15_000);
    const loaded = await pending;
    expect(loaded.faces.map(({ family }) => family)).toEqual(['Fast']);
    // Only the undecided family is recorded.
    expect(recordedFamilies(failures)).toEqual(['slow']);
    expect(calls).toBe(2);
  });

  for (const failing of ['viewer', 'worker'] as const) {
    it(`concludes a ${failing} batch that breaches a ceiling per family`, async () => {
      let calls = 0;
      const source = fontResolver((request) => {
        calls++;
        if (request.families.includes('Big')) return Array.from({ length: HARD_MAX_FONT_SOURCE_DECLARATIONS + 1 }, () => FACE(1, 'Big'));
        return request.families.map((family) => FACE(8, family));
      });
      const failures = new Map<string, string>();
      const link = bridge([source], { failures });
      const viewer = mainSession([source], failures);
      const worker = link.session();
      // The other realm decides Fast first.
      const decider = failing === 'viewer' ? worker : viewer;
      expect((await decider.session.ensure(demand(['Fast']))).faces).toHaveLength(1);
      const breaching = failing === 'viewer' ? viewer : worker;
      const error = await breaching.session.ensure(demand(['Fast', 'Big'])).catch((caught: unknown) => caught);
      expect((error as OoxmlError).violation).toMatchObject({ quota: 'declarations' });
      expect(recordedFamilies(failures)).toEqual(['big']);
      // The decided realm keeps Fast; a later requester of the document gets
      // Big's recorded failure with no further call.
      const later = mainSession([source], failures);
      const laterLoaded = await later.session.ensure(demand(['Fast', 'Big']));
      expect(laterLoaded.faces.map(({ family }) => family)).toEqual(['Fast']);
      expect(calls).toBe(2);
    });
  }

  // Round 8: main is the single authority; lateness is judged by wall clock.
  for (const [label, blockMs, loads] of [['blocks main for 31 s', 31_000, false], ['answers at 14 s', 14_000, true]] as const) {
    for (const first of ['worker', 'viewer'] as const) {
      it(`gives both realms one outcome when the resolver ${label} (${first} first)`, async () => {
        vi.useFakeTimers();
        let calls = 0;
        const source = fontResolver(() => {
          calls++;
          // A synchronous block: the clock moves, no timer can fire.
          vi.setSystemTime(Date.now() + blockMs);
          return [FACE(8)];
        });
        const failures = new Map<string, string>();
        const worker = bridge([source], { failures }).session();
        const viewer = mainSession([source], failures);
        const ensureWorker = async () => {
          const pending = worker.session.ensure(demand(['Private Face']));
          await vi.advanceTimersByTimeAsync(0);
          return pending;
        };
        const ensureViewer = () => viewer.session.ensure(demand(['Private Face']));
        const [a, b] = first === 'worker'
          ? [await ensureWorker(), await ensureViewer()]
          : [await ensureViewer(), await ensureWorker()];
        expect(a.faces).toHaveLength(loads ? 1 : 0);
        expect(b.faces).toHaveLength(loads ? 1 : 0);
        expect(calls).toBe(1);
        expect([...failures.values()]).toEqual(loads ? [] : ['timed out']);
        // Main's timers that were overdue fire now; nothing changes.
        await vi.advanceTimersByTimeAsync(60_000);
        expect(worker.set.faces.filter((face) => face.status === 'loaded')).toHaveLength(loads ? 1 : 0);
      });
    }
  }

  it('waits for main without a deadline of its own, and still cancels on reset', async () => {
    vi.useFakeTimers();
    // A main realm that holds the request (it decides; the worker never does).
    const posted: unknown[] = [];
    const client = new WorkerFontResolverClient((message) => { posted.push(message); });
    const workerSources = assertFontSourceDescriptors(structuredClone([fontResolver(() => [FACE(8)])]));
    const session = new FontSourceSession({ sources: workerSources, fontSet: asSet(fakeSet()), ...client.sessionOptions });
    sessions.push(session);
    let settled = false;
    const pending = session.ensure(demand(['Private Face'])).then((loaded) => { settled = true; return loaded; });
    await vi.advanceTimersByTimeAsync(120_000);
    expect(settled).toBe(false);
    expect(client.pendingCount).toBe(1);
    client.reset();
    expect((await pending).faces).toEqual([]);
    expect(client.pendingCount).toBe(0);
    expect(posted.map((message) => (message as { kind: string }).kind)).toEqual(['ooxmlResolveFonts', 'ooxmlCancelResolveFonts']);
  });

  it('cancels a waiting call when the worker document is destroyed', async () => {
    vi.useFakeTimers();
    let signal: AbortSignal | undefined;
    const source = fontResolver((_demand, context) => { signal = context.signal; return new Promise(() => undefined); });
    const link = bridge([source]);
    const { session } = link.session();
    const pending = session.ensure(demand(['Private Face']));
    await vi.advanceTimersByTimeAsync(1_000);
    session.release();
    await vi.advanceTimersByTimeAsync(0);
    expect(signal?.aborted).toBe(true);
    expect((await pending).faces).toEqual([]);
    expect(link.client.pendingCount).toBe(0);
  });

  // Finding 4: a worker reset abandons its calls on main too.
  it('cancels every pending call on reset, so a late answer is never posted or detached', async () => {
    let signal: AbortSignal | undefined;
    let finish!: (faces: readonly FontFaceDeclaration[]) => void;
    const source = fontResolver((_demand, context: FontResolveContext) => {
      signal = context.signal;
      return new Promise((resolve) => { finish = resolve; });
    });
    const link = bridge([source]);
    const { session, set } = link.session();
    const pending = session.ensure(demand(['Private Face']));
    await vi.waitFor(() => expect(signal).toBeDefined());
    // The worker starts another document.
    link.client.reset();
    await vi.waitFor(() => expect(signal?.aborted).toBe(true));
    expect(link.toMain.map((message) => (message as { kind: string }).kind))
      .toEqual(['ooxmlResolveFonts', 'ooxmlCancelResolveFonts']);
    const late = FACE(8);
    finish([late]);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(link.toWorker).toEqual([]);
    expect((late.data as ArrayBuffer).byteLength).toBe(8);
    expect((await pending).faces).toEqual([]);
    expect(set.faces).toHaveLength(0);
    expect(link.client.pendingCount).toBe(0);
  });

  it('abandons calls of an older generation when a newer request arrives', async () => {
    const signals: AbortSignal[] = [];
    const source = fontResolver((_demand, context) => { signals.push(context.signal); return new Promise(() => undefined); });
    const posts: unknown[] = [];
    const responder = new main.FontResolverResponder({ post: (message) => posts.push(message), sources: [source] });
    const request = (generation: number, requestId: number, family: string) => ({
      kind: 'ooxmlResolveFonts', generation, requestId, sourceIndex: 0, resolverId: source.id, demand: { families: [family], fallbackFamilies: [] },
    } as never);
    responder.accept(request(0, 1, 'A'));
    await vi.waitFor(() => expect(signals).toHaveLength(1));
    responder.accept(request(1, 1, 'B'));
    await vi.waitFor(() => expect(signals).toHaveLength(2));
    expect(signals[0]?.aborted).toBe(true);
    expect(signals[1]?.aborted).toBe(false);
    // A request of the abandoned generation is not started.
    responder.accept(request(0, 2, 'C'));
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(signals).toHaveLength(2);
    expect(posts).toEqual([]);
    responder.dispose();
  });
});

describe('fontResolver bridge: abort and malformed replies', () => {
  it('aborts the application call when the worker abandons it and posts no late answer', async () => {
    let signal: AbortSignal | undefined;
    let finish!: (faces: readonly FontFaceDeclaration[]) => void;
    const source = fontResolver((_demand, context: FontResolveContext) => {
      signal = context.signal;
      return new Promise((resolve) => { finish = resolve; });
    });
    const link = bridge([source]);
    const { session, set } = link.session();
    const pending = session.ensure(demand(['Private Face']));
    await vi.waitFor(() => expect(signal).toBeDefined());
    session.release();
    await vi.waitFor(() => expect(signal?.aborted).toBe(true));
    finish([FACE(8)]);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(link.toWorker).toEqual([]);
    expect((await pending).faces).toEqual([]);
    expect(set.faces).toHaveLength(0);
  });

  it('aborts every in-flight call when the document is destroyed', async () => {
    let signal: AbortSignal | undefined;
    const source = fontResolver((_demand, context) => {
      signal = context.signal;
      return new Promise(() => undefined);
    });
    const link = bridge([source]);
    const { session } = link.session();
    void session.ensure(demand(['Private Face']));
    await vi.waitFor(() => expect(signal).toBeDefined());
    link.host.dispose();
    await vi.waitFor(() => expect(signal?.aborted).toBe(true));
    expect(link.toWorker).toEqual([]);
  });

  it('fails closed on a malformed reply: host fallback, nothing registered', async () => {
    const source = fontResolver(() => [FACE(8)]);
    const workerSources = assertFontSourceDescriptors(structuredClone([source]));
    for (const outcome of [
      undefined,
      { status: 'families', families: 'nope' },
      { status: 'families', families: [{ family: 'private face', kind: 'claimed', declarations: new Array(HARD_MAX_FONT_SOURCE_DECLARATIONS + 1).fill(null) }] },
      { status: 'families', families: [{ family: 'private face', kind: 'claimed', declarations: [{ family: 'Private Face', data: new Uint8Array(4) }] }] },
      { status: 'families', families: [{ family: 1, kind: 'absent' }] },
      { status: 'families', families: [{ family: 'private face', kind: 'failed' }] },
      { status: 'limit', violation: { resource: 'font-source', quota: 'face', metric: 'bytes', limit: 2, observed: 1 } },
      { status: 'unknown' },
    ]) {
      const client: WorkerFontResolverClient = new WorkerFontResolverClient((message) => {
        const { kind, generation, requestId } = message as { kind: string; generation: number; requestId: number };
        if (kind !== 'ooxmlResolveFonts') return;
        void Promise.resolve().then(() => client.accept({ kind: 'ooxmlFontsResolved', generation, requestId, outcome }));
      });
      const set = fakeSet();
      const session = new FontSourceSession({ sources: workerSources, fontSet: asSet(set), runResolver: client.run });
      expect((await session.ensure(demand(['Private Face']))).faces).toEqual([]);
      expect(set.faces).toHaveLength(0);
      session.release();
    }
    expect(warn).toHaveBeenCalled();
  });

  it('rejects the load for a ceiling breach found by the main-side admission', async () => {
    const source = fontResolver(() => Array.from({ length: HARD_MAX_FONT_SOURCE_DECLARATIONS + 1 }, () => FACE(1)));
    const { session } = bridge([source]).session();
    const error = await session.ensure(demand(['Private Face'])).catch((caught: unknown) => caught);
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).violation).toMatchObject({ quota: 'declarations', observed: HARD_MAX_FONT_SOURCE_DECLARATIONS + 1 });
  });

  it('keeps an application error a family failure, with the same reason as main mode', async () => {
    const source = fontResolver(() => { throw new Error('token expired'); });
    const { session } = bridge([source]).session();
    expect((await session.ensure(demand(['Private Face']))).faces).toEqual([]);
    expect(String(warn.mock.calls.at(-1)?.[0])).toContain('Private Face (token expired)');
  });

  it('runs only resolvers the document was configured with', async () => {
    const configured = fontResolver(() => [FACE(4)]);
    const other = fontResolver(() => [FACE(4, 'Other')]);
    const link = bridge([configured]);
    const session = new FontSourceSession({
      sources: assertFontSourceDescriptors(structuredClone([other])),
      fontSet: asSet(fakeSet()),
      runResolver: link.client.run,
    });
    sessions.push(session);
    expect((await session.ensure(demand(['Other']))).faces).toEqual([]);
    expect(String(warn.mock.calls.at(-1)?.[0])).toContain('not available in this realm');
  });

  it('times the application call out after the font ceiling', async () => {
    vi.useFakeTimers();
    let signal: AbortSignal | undefined;
    const source = fontResolver((_demand, context) => {
      signal = context.signal;
      return new Promise(() => undefined);
    });
    const posts: unknown[] = [];
    const responder = new main.FontResolverResponder({ post: (message) => posts.push(message), sources: [source] });
    responder.accept({ kind: 'ooxmlResolveFonts', generation: 0, requestId: 1, sourceIndex: 0, resolverId: source.id, demand: { families: ['A'], fallbackFamilies: [] } } as never);
    await vi.advanceTimersByTimeAsync(15_000);
    expect(signal?.aborted).toBe(true);
    expect(posts).toEqual([{
      kind: 'ooxmlFontsResolved', generation: 0, requestId: 1,
      outcome: { status: 'families', families: [{ family: 'a', kind: 'failed', reason: 'timed out' }] },
    }]);
    expect(responder.pendingCount).toBe(0);
    responder.dispose();
  });
});

describe('fontResolver bridge: the page fetches resolver URLs', () => {
  const ORIGINAL_FETCH = G.fetch;
  type Route = { bytes?: number; status?: number; hang?: boolean };
  function installFetch(routes: Record<string, Route>) {
    const calls: { url: string; init?: RequestInit }[] = [];
    G.fetch = vi.fn(async (url: string, init?: RequestInit) => {
      calls.push({ url, init });
      const route = routes[url];
      if (route?.hang) {
        await new Promise((_resolve, reject) => init?.signal?.addEventListener('abort',
          () => reject(Object.assign(new Error('aborted'), { name: 'AbortError' }))));
      }
      if (!route) return { ok: false, status: 404, arrayBuffer: async () => new ArrayBuffer(0) };
      return {
        ok: (route.status ?? 200) < 400, status: route.status ?? 200,
        arrayBuffer: async () => new Uint8Array(route.bytes ?? 0).fill(9).buffer,
      };
    });
    return calls;
  }
  afterEach(async () => {
    vi.useRealTimers();
    G.fetch = ORIGINAL_FETCH;
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(main.acquiredFaceCount()).toBe(0);
  });

  const URL_FACES = (family = 'Private Face'): FontFaceDeclaration[] => [
    { family, url: 'https://fonts.app.example/private-400.woff2', weight: '400' },
    { family, url: 'https://fonts.app.example/private-700.woff2', weight: '700' },
  ];
  const kinds = (messages: unknown[]) => messages.map((message) => (message as { kind: string }).kind);

  it('fetches the claimed URLs in the page with page credentials and sends the worker bytes', async () => {
    const calls = installFetch({
      'https://fonts.app.example/private-400.woff2': { bytes: 40 },
      'https://fonts.app.example/private-700.woff2': { bytes: 70 },
    });
    const source = fontResolver(() => [...URL_FACES(), { family: 'Private Face', url: 'data:font/woff2;base64,AAAA', style: 'italic' }]);

    // Main mode is unchanged: the browser loads the URLs (FontFace sources).
    const { session: direct, set: directSet } = mainSession([source]);
    await direct.ensure(demand(['Private Face']));
    expect(directSet.faces.map((face) => face.bytes)).toEqual([0, 0, 0]);
    expect(calls).toHaveLength(0);
    direct.release();

    const link = bridge([source]);
    const { session, set } = link.session();
    const loaded = await session.ensure(demand(['Private Face']));
    expect(loaded.faces.map(({ weight, style }) => `${weight}/${style}`)).toEqual(['400/normal', '700/normal', 'normal/italic']);
    // http(s) faces arrive as bytes; the inline data: URL stays as it is.
    expect(set.faces.map((face) => face.bytes)).toEqual([40, 70, 0]);
    expect(kinds(link.toMain)).toEqual(['ooxmlResolveFonts', 'ooxmlAcquireFonts']);
    expect(kinds(link.toWorker)).toEqual(['ooxmlFontsResolved', 'ooxmlFontsAcquired']);
    // Requested by the page, once per URL, with the CSS font-fetch mode.
    expect(calls.map(({ url, init }) => [url, init?.credentials, init?.mode])).toEqual([
      ['https://fonts.app.example/private-400.woff2', 'same-origin', 'cors'],
      ['https://fonts.app.example/private-700.woff2', 'same-origin', 'cors'],
    ]);
    // The worker charges what it keeps: its cached answer (the data: URL),
    // the transferred bodies and the registrations; main the acquired
    // bodies it keeps for the document.
    expect(retainedFontSourceBytes()).toBe(3 + 110 + 113);
    expect(main.retainedFontSourceBytes()).toBe(3 + 110);
    session.release();
    link.host.dispose();
    await vi.waitFor(() => expect(main.retainedFontSourceBytes()).toBe(0));
  });

  it('fetches only URLs the document resolved for the named family', async () => {
    const calls = installFetch({ 'https://evil.example/x.woff2': { bytes: 4 } });
    const source = fontResolver(() => URL_FACES().slice(0, 1));
    const link = bridge([source]);
    const { session } = link.session();
    installFetch({ 'https://fonts.app.example/private-400.woff2': { bytes: 40 } });
    await session.ensure(demand(['Private Face']));
    const posts: unknown[] = [];
    const host = new main.FontResolverBridgeHost({ post: (message) => posts.push(message), sources: [source] });
    hosts.push(host);
    host.accept({
      kind: 'ooxmlAcquireFonts', generation: 0, requestId: 9, sourceIndex: 0, resolverId: source.id, remainingMs: 1000,
      families: [{ family: 'Private Face', urls: ['https://evil.example/x.woff2'] }],
    });
    await vi.waitFor(() => expect(posts).toHaveLength(1));
    expect(calls).toHaveLength(0);
    expect((posts[0] as { outcome: unknown }).outcome).toEqual({
      status: 'faces', faces: [{ url: 'https://evil.example/x.woff2', reason: 'not a font this document resolved' }],
    });
    session.release();
    link.host.dispose();
  });

  it('shares one request and one outcome with the XLSX viewer realm of the document', async () => {
    const calls = installFetch({ 'https://fonts.app.example/private-400.woff2': { bytes: 40 } });
    const resolve = vi.fn(() => URL_FACES());
    const source = fontResolver(resolve);
    const failures = new Map<string, string>();
    const viewerSet = fakeSet();
    const viewer = new main.FontSourceSession({
      sources: [source], fontSet: asSet(viewerSet), resolverFailures: failures,
      acquireResolverFaces: main.inRealmResolverFaceAcquirer(failures),
    });
    sessions.push(viewer);
    const viewerLoaded = await viewer.ensure(demand(['Private Face']));
    const link = bridge([source], { failures });
    const { session, set } = link.session();
    const loaded = await session.ensure(demand(['Private Face']));
    expect(resolve).toHaveBeenCalledTimes(1);
    // One fetch per URL for both realms, including the failed one.
    expect(calls.map(({ url }) => url)).toEqual([
      'https://fonts.app.example/private-400.woff2', 'https://fonts.app.example/private-700.woff2',
    ]);
    expect(viewerLoaded.faces.map(({ weight }) => weight)).toEqual(['400']);
    expect(loaded.faces.map(({ weight }) => weight)).toEqual(['400']);
    expect(viewerSet.faces.map((face) => face.bytes)).toEqual([40]);
    expect(set.faces.map((face) => face.bytes)).toEqual([40]);
    session.release();
    viewer.release();
    link.host.dispose();
  });

  it('gives overlapping requesters of one document one outcome per URL', async () => {
    let release: () => void = () => undefined;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const calls: string[] = [];
    G.fetch = vi.fn(async (url: string) => {
      calls.push(url);
      await gate;
      return { ok: true, status: 200, arrayBuffer: async () => new Uint8Array(16).buffer };
    });
    const failures = new Map<string, string>();
    const source = fontResolver(() => URL_FACES().slice(0, 1));
    const url = 'https://fonts.app.example/private-400.woff2';
    const sourceId = (await import('./source-cache.js')).fontSourceId(source);
    const viewer = { holder: {} };
    const worker = { holder: {} };
    // The viewer realm's request reaches its deadline first; the worker's
    // request for the same URL is still waiting on the shared download.
    const viewerAbort = new AbortController();
    const viewerPending = main.acquireResolverFacesInRealm(failures, sourceId, [url],
      { signal: viewerAbort.signal, deadline: Date.now() - 1, holder: viewer.holder });
    const workerPending = main.acquireResolverFacesInRealm(failures, sourceId, [url],
      { signal: new AbortController().signal, deadline: Date.now() + 60_000, holder: worker.holder });
    viewerAbort.abort(main.FONT_DEADLINE_REASON);
    expect(await viewerPending).toEqual(new Map([[url, { reason: 'timed out' }]]));
    release();
    // The worker adopts the document's outcome instead of painting bytes the
    // viewer never measured with.
    expect(await workerPending).toEqual(new Map([[url, { reason: 'timed out' }]]));
    const later = await main.acquireResolverFacesInRealm(failures, sourceId, [url],
      { signal: new AbortController().signal, deadline: Date.now() + 60_000, holder: worker.holder });
    expect(later).toEqual(new Map([[url, { reason: 'timed out' }]]));
    expect(calls).toEqual([url]);
    main.releaseAcquiredFaces(viewer.holder);
    main.releaseAcquiredFaces(worker.holder);
  });

  it('lets the worker adopt the page\'s acquisition verdict after its own ceiling', async () => {
    vi.useFakeTimers();
    let release: () => void = () => undefined;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    G.fetch = vi.fn(async () => { await gate; return { ok: true, status: 200, arrayBuffer: async () => new Uint8Array(16).buffer }; });
    const failures = new Map<string, string>();
    const source = fontResolver(() => URL_FACES().slice(0, 1));
    const link = bridge([source], { failures });
    const { session, set } = link.session();
    const pending = session.ensure(demand(['Private Face']));
    await vi.waitFor(() => expect(kinds(link.toMain)).toContain('ooxmlAcquireFonts'));
    // The page acquires in time, but its reply reaches a busy worker only
    // after the worker's own face-loading ceiling.
    link.pause();
    release();
    await vi.waitFor(() => expect(kinds(link.toWorker)).toContain('ooxmlFontsAcquired'));
    await vi.advanceTimersByTimeAsync(20_000);
    link.resume();
    await vi.advanceTimersByTimeAsync(0);
    const loaded = await pending;
    expect(loaded.faces).toHaveLength(1);
    expect(set.faces.map((face) => face.bytes)).toEqual([16]);
    // An XLSX viewer realm of the same document paints the same face.
    const viewer = new main.FontSourceSession({
      sources: [source], fontSet: asSet(fakeSet()), resolverFailures: failures,
      acquireResolverFaces: main.inRealmResolverFaceAcquirer(failures),
    });
    sessions.push(viewer);
    expect((await viewer.ensure(demand(['Private Face']))).faces).toHaveLength(1);
    session.release();
    viewer.release();
    link.host.dispose();
  });

  it('never records a cancelled request as the document\'s outcome', async () => {
    let release: () => void = () => undefined;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    G.fetch = vi.fn(async () => { await gate; return { ok: true, status: 200, arrayBuffer: async () => new Uint8Array(16).buffer }; });
    const failures = new Map<string, string>();
    const source = fontResolver(() => URL_FACES().slice(0, 1));
    const url = 'https://fonts.app.example/private-400.woff2';
    const sourceId = (await import('./source-cache.js')).fontSourceId(source);
    const closing = new AbortController();
    const live = {};
    // Cancelled (its session released) before its deadline; it resumes only
    // after that deadline has passed.
    const cancelled = main.acquireResolverFacesInRealm(failures, sourceId, [url],
      { signal: closing.signal, deadline: Date.now() + 5, holder: {} });
    const other = main.acquireResolverFacesInRealm(failures, sourceId, [url],
      { signal: new AbortController().signal, deadline: Date.now() + 60_000, holder: live });
    closing.abort();
    await new Promise((resolve) => setTimeout(resolve, 20));
    await cancelled;
    release();
    expect('data' in ((await other).get(url) ?? {})).toBe(true);
    expect([...failures.keys()].some((key) => key.includes('\u0000acquire\u0000'))).toBe(false);
    main.releaseAcquiredFaces(live);
  });

  it('takes no hold for a released requester that adopts the document\'s bytes', async () => {
    // Two XLSX viewer-realm sessions of one document share a download and
    // the outcome record. The first is released somewhere between the
    // download settling and its continuation, while the second records the
    // success the first then finds: whichever microtask the release lands
    // in, nothing stays held once both sessions are released.
    for (let delay = 0; delay < 16; delay++) {
      const failures = new Map<string, string>();
      const url = `https://fonts.app.example/adopt-${delay}.woff2`;
      const source = fontResolver(() => [{ family: 'Private Face', url }]);
      let first: FontSourceSession | undefined;
      G.fetch = vi.fn(async () => ({
        ok: true, status: 200,
        arrayBuffer: async () => {
          let step: Promise<void> = Promise.resolve();
          for (let index = 0; index < delay; index++) step = step.then(() => undefined);
          void step.then(() => first?.release());
          return new Uint8Array(16).buffer;
        },
      }));
      const [a, b] = [fakeSet(), fakeSet()].map((set) => new main.FontSourceSession({
        sources: [source], fontSet: asSet(set), resolverFailures: failures,
        acquireResolverFaces: main.inRealmResolverFaceAcquirer(failures),
      }));
      first = a;
      sessions.push(a, b);
      // The second session's continuation runs ahead of the first's.
      const second = b.ensure(demand(['Private Face']));
      await Promise.all([a.ensure(demand(['Private Face'])).catch(() => undefined), second]);
      b.release();
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(main.acquiredFaceCount(), `delay ${delay}`).toBe(0);
      expect(main.retainedFontSourceBytes(), `delay ${delay}`).toBe(0);
    }
  });

  it('adopts bytes another requester of the document received', async () => {
    installFetch({ 'https://fonts.app.example/private-400.woff2': { bytes: 16 } });
    const failures = new Map<string, string>();
    const source = fontResolver(() => URL_FACES().slice(0, 1));
    const url = 'https://fonts.app.example/private-400.woff2';
    const sourceId = (await import('./source-cache.js')).fontSourceId(source);
    const first = {};
    const second = {};
    const got = await main.acquireResolverFacesInRealm(failures, sourceId, [url],
      { signal: new AbortController().signal, deadline: Date.now() + 60_000, holder: first });
    // A requester whose own deadline already passed still paints the
    // document's bytes, without another request.
    const adopted = await main.acquireResolverFacesInRealm(failures, sourceId, [url],
      { signal: new AbortController().signal, deadline: Date.now() - 1, holder: second });
    expect((adopted.get(url) as { data: ArrayBuffer }).data).toBe((got.get(url) as { data: ArrayBuffer }).data);
    main.releaseAcquiredFaces(first);
    expect(main.acquiredFaceCount()).toBe(1);
    main.releaseAcquiredFaces(second);
  });

  it('reuses the bytes of a face a later delta claims again, charging the copy once', async () => {
    const calls = installFetch({ 'https://fonts.app.example/sub.woff2': { bytes: 32 } });
    const source = fontResolver((request) => request.families.map((family) => ({ family, url: 'https://fonts.app.example/sub.woff2' })));
    const link = bridge([source]);
    const { session } = link.session();
    await session.ensure(demand(['First']));
    const afterFirst = retainedFontSourceBytes();
    await session.ensure(demand(['Second']));
    expect(kinds(link.toMain).filter((kind) => kind === 'ooxmlAcquireFonts')).toHaveLength(1);
    expect(calls).toHaveLength(1);
    // The second family registers its own face (another CSS family); the
    // transferred body is kept and charged once.
    expect(retainedFontSourceBytes()).toBe(afterFirst + 32);
    session.release();
    link.host.dispose();
  });

  it('keeps claim precedence when a claimed face cannot be fetched', async () => {
    installFetch({ 'https://fonts.app.example/sub.woff2': { bytes: 8 } });
    const source = fontResolver(() => [
      { family: 'Private Face', url: 'https://fonts.app.example/gone.woff2' },
      { family: 'Sub Face', url: 'https://fonts.app.example/sub.woff2' },
    ], { substitutes: [{ for: 'Private Face', use: 'Sub Face' }] });
    const link = bridge([source]);
    const { session, set } = link.session();
    const loaded = await session.ensure(demand(['Private Face']));
    // As in main mode: the authored face claims the family, fails, and the
    // family keeps its host fallback; the substitute is never promoted.
    expect(loaded.faces).toEqual([]);
    expect(loaded.substitutions.size).toBe(0);
    expect(set.faces).toHaveLength(0);
    expect(String(warn.mock.calls.at(-1)?.[0])).toContain('Private Face (HTTP 404)');
    session.release();
    link.host.dispose();
  });

  it('leaves out faces whose download fails and warns their family', async () => {
    installFetch({ 'https://fonts.app.example/private-400.woff2': { bytes: 40 } });
    const source = fontResolver(() => URL_FACES());
    const link = bridge([source]);
    const { session, set } = link.session();
    const loaded = await session.ensure(demand(['Private Face']));
    expect(loaded.faces.map(({ weight }) => weight)).toEqual(['400']);
    expect(set.faces.map((face) => face.bytes)).toEqual([40]);
    expect(String(warn.mock.calls.at(-1)?.[0])).toContain('Private Face (HTTP 404)');
    session.release();
    link.host.dispose();
  });

  it('acquires within the face-loading ceiling', async () => {
    vi.useFakeTimers();
    installFetch({ 'https://fonts.app.example/private-400.woff2': { hang: true } });
    const source = fontResolver(() => URL_FACES().slice(0, 1));
    const link = bridge([source]);
    const { session } = link.session();
    const pending = session.ensure(demand(['Private Face']));
    await vi.advanceTimersByTimeAsync(15_000);
    expect((await pending).faces).toEqual([]);
    expect(String(warn.mock.calls.at(-1)?.[0])).toContain('Private Face (timed out)');
    session.release();
    link.host.dispose();
  });

  it('posts nothing for an acquisition the worker cancels, and aborts its requests', async () => {
    const calls = installFetch({ 'https://fonts.app.example/private-400.woff2': { hang: true } });
    const source = fontResolver(() => URL_FACES().slice(0, 1));
    const link = bridge([source]);
    const { session } = link.session();
    const pending = session.ensure(demand(['Private Face']));
    await vi.waitFor(() => expect(calls).toHaveLength(1));
    link.client.reset();
    await vi.waitFor(() => expect((calls[0].init?.signal as AbortSignal).aborted).toBe(true));
    await pending.catch(() => undefined);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(kinds(link.toWorker)).toEqual(['ooxmlFontsResolved']);
    session.release();
    link.host.dispose();
  });

  it('rejects the worker load for a body beyond the face ceiling and aborts the rest of the batch', async () => {
    const calls = installFetch({
      'https://fonts.app.example/private-400.woff2': { bytes: HARD_MAX_FONT_SOURCE_FACE_BYTES + 1 },
      'https://fonts.app.example/private-700.woff2': { hang: true },
    });
    const source = fontResolver(() => URL_FACES());
    const link = bridge([source]);
    const { session } = link.session();
    await expect(session.ensure(demand(['Private Face'])))
      .rejects.toMatchObject({ code: 'font-source-limit', violation: { quota: 'face' } });
    const sibling = calls.find(({ url }) => url.endsWith('private-700.woff2'));
    expect((sibling?.init?.signal as AbortSignal).aborted).toBe(true);
    expect(main.retainedFontSourceBytes()).toBe(0);
    link.host.dispose();
  });
});

describe('face acquisition', () => {
  const ORIGINAL_FETCH = G.fetch;
  afterEach(() => { G.fetch = ORIGINAL_FETCH; vi.useRealTimers(); });

  it('runs at most six downloads at a time', async () => {
    const started: string[] = [];
    G.fetch = vi.fn((url: string, init?: RequestInit) => {
      started.push(url);
      return new Promise((_resolve, reject) => init?.signal?.addEventListener('abort',
        () => reject(Object.assign(new Error('aborted'), { name: 'AbortError' }))));
    });
    const controller = new AbortController();
    const urls = Array.from({ length: 10 }, (_, index) => `https://fonts.app.example/${index}.woff2`);
    const holder = {};
    const pending = main.acquireFaceUrls(urls, { credentials: 'omit', signal: controller.signal, deadline: Date.now() + 60_000, holder });
    await vi.waitFor(() => expect(started).toHaveLength(6));
    await new Promise((resolve) => setTimeout(resolve, 10));
    expect(started).toHaveLength(6);
    controller.abort();
    const outcomes = await pending;
    expect([...outcomes.values()].every((outcome) => 'reason' in outcome && outcome.reason === 'timed out')).toBe(true);
    main.releaseAcquiredFaces(holder);
    expect(main.acquiredFaceCount()).toBe(0);
  });

  it('keeps a holder\'s bytes when one of its overlapping joins is cancelled', async () => {
    let release: () => void = () => undefined;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    G.fetch = vi.fn(async () => { await gate; return { ok: true, status: 200, arrayBuffer: async () => new Uint8Array(16).buffer }; });
    const holder = {};
    const url = 'https://fonts.app.example/twice.woff2';
    const cancelled = new AbortController();
    const one = main.acquireFaceUrls([url], { credentials: 'omit', signal: cancelled.signal, deadline: Date.now() + 60_000, holder });
    const two = main.acquireFaceUrls([url], { credentials: 'omit', signal: new AbortController().signal, deadline: Date.now() + 60_000, holder });
    cancelled.abort();
    await one;
    release();
    expect('data' in ((await two).get(url) ?? {})).toBe(true);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(main.acquiredFaceCount()).toBe(1);
    expect(main.retainedFontSourceBytes()).toBe(16);
    main.releaseAcquiredFaces(holder);
    expect(main.retainedFontSourceBytes()).toBe(0);
  });

  it('takes no hold for a requester released while its completed join resumes', async () => {
    // The release lands somewhere between the download settling and the
    // join's continuation: whichever microtask it is, nothing stays held.
    for (let delay = 0; delay < 12; delay++) {
      const controller = new AbortController();
      const holder = {};
      G.fetch = vi.fn(async () => ({
        ok: true, status: 200,
        arrayBuffer: async () => {
          let step: Promise<void> = Promise.resolve();
          for (let index = 0; index < delay; index++) step = step.then(() => undefined);
          // As FontSourceSession.release(): abort, then return every hold.
          void step.then(() => { controller.abort(); main.releaseAcquiredFaces(holder); });
          return new Uint8Array(16).buffer;
        },
      }));
      await main.acquireFaceUrls([`https://fonts.app.example/race-${delay}.woff2`],
        { credentials: 'omit', signal: controller.signal, deadline: Date.now() + 60_000, holder });
      // The release above was the holder's last: nothing may hold it again.
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(main.acquiredFaceCount(), `delay ${delay}`).toBe(0);
      expect(main.retainedFontSourceBytes(), `delay ${delay}`).toBe(0);
    }
  });

  it('fails a body that completes after the deadline, whatever the timers did', async () => {
    vi.useFakeTimers();
    const start = Date.now();
    G.fetch = vi.fn(async () => {
      vi.setSystemTime(start + 20_000);
      return { ok: true, status: 200, arrayBuffer: async () => new Uint8Array(4).buffer };
    });
    const holder = {};
    const outcomes = await main.acquireFaceUrls(['https://fonts.app.example/late.woff2'], {
      credentials: 'same-origin', signal: new AbortController().signal, deadline: start + 10_000, holder,
    });
    expect(outcomes.get('https://fonts.app.example/late.woff2')).toEqual({ reason: 'timed out' });
    main.releaseAcquiredFaces(holder);
    expect(main.retainedFontSourceBytes()).toBe(0);
  });
});

describe('fontResolver bridge: demand copy', () => {
  it('accepts family names only and rejects malformed demand', () => {
    expect(main.copyBridgedFontDemand({ families: ['A'], fallbackFamilies: ['B'] }))
      .toEqual({ families: ['A'], fallbackFamilies: ['B'] });
    expect(main.copyBridgedFontDemand({ families: [1], fallbackFamilies: [] })).toBeUndefined();
    expect(main.copyBridgedFontDemand({ families: ['A'] })).toBeUndefined();
    // Nothing else crosses the bridge: scalars, tuples and scopes stay in the worker.
    expect(main.copyBridgedFontDemand({ families: ['A'], fallbackFamilies: [], scalars: 'all' })).toBeUndefined();
    expect(main.copyBridgedFontDemand({ families: ['A'], fallbackFamilies: [], substituteScripts: {} })).toBeUndefined();
  });
});
