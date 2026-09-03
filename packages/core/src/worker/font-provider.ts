import {
  FontProviderSession,
  providerFontFamily,
  registerResolvedFonts,
  type FontFamilyRoutes,
  type ResolvedFonts,
} from '../fonts/provider.js';
import { releaseFaces } from '../fonts/font-registry.js';

export const FONT_PROVIDER_PROTOCOL = 'ooxml-font-v1' as const;

export interface WorkerFontRequest {
  readonly protocol: typeof FONT_PROVIDER_PROTOCOL;
  readonly kind: 'resolve';
  readonly fontRequestId: number;
  readonly generation: number;
  readonly families: readonly string[];
}

export interface WorkerFontResponse {
  readonly protocol: typeof FONT_PROVIDER_PROTOCOL;
  readonly kind: 'resolved';
  readonly fontRequestId: number;
  readonly generation: number;
  readonly strict?: boolean;
  readonly resolved?: ResolvedFonts;
  readonly error?: string;
}

export function isWorkerFontRequest(value: unknown): value is WorkerFontRequest {
  const candidate = value as Partial<WorkerFontRequest> | null;
  return candidate?.protocol === FONT_PROVIDER_PROTOCOL
    && candidate.kind === 'resolve'
    && Number.isSafeInteger(candidate.fontRequestId)
    && Number.isSafeInteger(candidate.generation)
    && Array.isArray(candidate.families)
    && candidate.families.every((family) => typeof family === 'string');
}

export function isWorkerFontResponse(value: unknown): value is WorkerFontResponse {
  const candidate = value as Partial<WorkerFontResponse> | null;
  return candidate?.protocol === FONT_PROVIDER_PROTOCOL
    && candidate.kind === 'resolved'
    && Number.isSafeInteger(candidate.fontRequestId)
    && Number.isSafeInteger(candidate.generation);
}

type Post = (message: unknown, transfer?: Transferable[]) => void;

/** Main-thread endpoint for font requests emitted while a render worker parses. */
export class FontProviderHost {
  constructor(
    private readonly session: FontProviderSession,
    private readonly post: Post,
    private readonly target: FontFaceSet | null = null,
  ) {}

  async accept(value: unknown): Promise<boolean> {
    if (!isWorkerFontRequest(value)) return false;
    try {
      const resolved = await this.session.ensure(value.families, this.target);
      const copies: ResolvedFonts = {
        routes: { ...resolved.routes },
        faces: resolved.faces.map((face) => ({ ...face, data: face.data.slice(0) })),
      };
      this.post({
        protocol: FONT_PROVIDER_PROTOCOL,
        kind: 'resolved',
        fontRequestId: value.fontRequestId,
        generation: value.generation,
        strict: this.session.strict,
        resolved: copies,
      } satisfies WorkerFontResponse, copies.faces.map((face) => face.data));
    } catch (error) {
      try {
        this.post({
          protocol: FONT_PROVIDER_PROTOCOL,
          kind: 'resolved',
          fontRequestId: value.fontRequestId,
          generation: value.generation,
          error: error instanceof Error ? error.message : String(error),
        } satisfies WorkerFontResponse);
      } catch {
        // The worker may have terminated while provider work was in flight.
      }
    }
    return true;
  }
}

interface Pending {
  readonly generation: number;
  readonly resolve: (routes: FontFamilyRoutes) => void;
  readonly reject: (error: Error) => void;
}

/** Render-worker endpoint that requests faces from the host and owns worker registrations. */
export class FontProviderClient {
  private nextId = 1;
  private readonly pending = new Map<number, Pending>();
  private faces: FontFace[] = [];
  private readonly aliases = new Set<string>();
  private readonly families = new Set<string>();
  private routes: FontFamilyRoutes = {};
  private tail: Promise<void> = Promise.resolve();
  private epoch = 0;

  constructor(private readonly post: Post) {}

  resolve(families: readonly string[], generation: number): Promise<FontFamilyRoutes> {
    const epoch = this.epoch;
    const operation = this.tail.then(() => {
      if (epoch !== this.epoch) throw new Error('font provider request canceled');
      const missing = families.filter((family) => {
        const key = family.trim().toLocaleLowerCase('en-US');
        if (!key || this.families.has(key)) return false;
        this.families.add(key);
        return true;
      });
      if (missing.length === 0) return { ...this.routes };
      const fontRequestId = this.nextId++;
      return new Promise<FontFamilyRoutes>((resolve, reject) => {
        this.pending.set(fontRequestId, { generation, resolve, reject });
        this.post({
          protocol: FONT_PROVIDER_PROTOCOL,
          kind: 'resolve',
          fontRequestId,
          generation,
          families: missing,
        } satisfies WorkerFontRequest);
      });
    });
    this.tail = operation.then(() => undefined, () => undefined);
    return operation;
  }

  async accept(value: unknown): Promise<boolean> {
    if (!isWorkerFontResponse(value)) return false;
    const pending = this.pending.get(value.fontRequestId);
    if (!pending || pending.generation !== value.generation) return true;
    this.pending.delete(value.fontRequestId);
    if (value.error) {
      pending.reject(new Error(value.error));
      return true;
    }
    const resolved = value.resolved ?? { routes: {}, faces: [] };
    const freshAliases = new Set(Object.values(resolved.routes)
      .map((route) => typeof route === 'string' ? route : route.family)
      .filter((alias) => !this.aliases.has(alias)));
    const fresh = {
      routes: Object.fromEntries(
        Object.entries(resolved.routes).filter(([, route]) => (
          freshAliases.has(typeof route === 'string' ? route : route.family)
        )),
      ),
      faces: resolved.faces.filter((face) => freshAliases.has(face.alias)),
    } satisfies ResolvedFonts;
    const loaded = await registerResolvedFonts(fresh);
    if (
      value.strict &&
      (
        loaded.faces.length !== fresh.faces.length ||
        Object.keys(loaded.routes).length !== Object.keys(fresh.routes).length
      )
    ) {
      releaseFaces(loaded.faces);
      pending.reject(new Error('font provider face failed to load'));
      return true;
    }
    this.faces.push(...loaded.faces);
    for (const route of Object.values(loaded.routes)) {
      this.aliases.add(typeof route === 'string' ? route : route.family);
    }
    this.routes = {
      ...this.routes,
      ...Object.fromEntries(
        Object.entries(resolved.routes).filter(([family]) => {
          const alias = providerFontFamily(resolved.routes, family);
          return alias ? this.aliases.has(alias) : false;
        }),
      ),
    };
    pending.resolve({ ...this.routes });
    return true;
  }

  reset(): void {
    this.epoch += 1;
    releaseFaces(this.faces);
    this.faces = [];
    this.aliases.clear();
    this.families.clear();
    this.routes = {};
    for (const pending of this.pending.values()) {
      pending.reject(new Error('font provider request canceled'));
    }
    this.pending.clear();
  }
}
