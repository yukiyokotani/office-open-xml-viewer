import { HARD_MAX_EMBEDDED_FONT_BYTES } from '../worker/resource-policy.generated.js';
import { retainFace, releaseFaces } from './font-registry.js';
import { activeFontSet, withFontCeiling } from './preload.js';

export type FontFailure = 'fallback' | 'error';

export type FontAssetSource =
  | Readonly<{ url: string | URL }>
  | Readonly<{ data: ArrayBuffer }>;

export interface FontAsset {
  /** Authored OOXML family this face satisfies. */
  readonly family: string;
  readonly source: FontAssetSource;
  readonly descriptors?: FontFaceDescriptors;
}

export interface FontResolveOptions {
  readonly signal: AbortSignal;
}

/** Main-thread extension point for application-owned font sources. */
export abstract class FontProvider {
  abstract resolve(
    families: readonly string[],
    options: Readonly<FontResolveOptions>,
  ): Promise<readonly FontAsset[]>;
}

export type FontFamilyRoutes = Readonly<Record<string, string>>;

export interface ResolvedFontFace {
  readonly family: string;
  readonly alias: string;
  readonly data: ArrayBuffer;
  readonly descriptors: FontFaceDescriptors;
}

export interface ResolvedFonts {
  readonly routes: FontFamilyRoutes;
  readonly faces: readonly ResolvedFontFace[];
}

const HARD_MAX_PROVIDER_FAMILIES = 256;
const HARD_MAX_PROVIDER_FACES = 1024;
const HARD_MAX_PROVIDER_BYTES = 256 * 1024 * 1024;
let nextSessionId = 1;

function normalizedFamily(value: string): string {
  return value.trim().toLocaleLowerCase('en-US');
}

function safeFamily(value: string): string {
  return [...value].map((char) => (char.codePointAt(0) ?? 0).toString(16)).join('_');
}

function copyDescriptors(value: FontFaceDescriptors | undefined): FontFaceDescriptors {
  return value ? { ...value } : {};
}

async function sourceBytes(source: FontAssetSource, signal: AbortSignal): Promise<ArrayBuffer> {
  if ('data' in source) {
    if (source.data.byteLength === 0 || source.data.byteLength > HARD_MAX_EMBEDDED_FONT_BYTES) {
      throw new Error(`face is ${source.data.byteLength} bytes`);
    }
    return source.data.slice(0);
  }
  if (typeof fetch === 'undefined') throw new Error('fetch is unavailable');
  const base = typeof document !== 'undefined' ? document.baseURI : globalThis.location?.href;
  const url = new URL(source.url, base).href;
  const response = await fetch(url, { signal });
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const declared = Number(response.headers.get('content-length'));
  if (Number.isFinite(declared) && declared > HARD_MAX_EMBEDDED_FONT_BYTES) {
    throw new Error(`face is ${declared} bytes`);
  }
  if (!response.body) {
    const data = await response.arrayBuffer();
    if (data.byteLength === 0 || data.byteLength > HARD_MAX_EMBEDDED_FONT_BYTES) {
      throw new Error(`face is ${data.byteLength} bytes`);
    }
    return data;
  }
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > HARD_MAX_EMBEDDED_FONT_BYTES) {
      await reader.cancel();
      throw new Error(`face exceeds ${HARD_MAX_EMBEDDED_FONT_BYTES} bytes`);
    }
    chunks.push(value);
  }
  if (total === 0) throw new Error('face is 0 bytes');
  const data = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    data.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return data.buffer;
}

/** Look up the private fallback alias for an authored family. */
export function providerFontFamily(
  routes: FontFamilyRoutes | undefined,
  family: string | null | undefined,
): string | undefined {
  return family ? routes?.[normalizedFamily(family)] : undefined;
}

interface StoredFace extends ResolvedFontFace {
  readonly key: string;
}

interface SetRegistration {
  readonly keys: Set<string>;
  readonly faces: FontFace[];
}

/** One document's provider resolution, byte ownership, and FontFace registrations. */
export class FontProviderSession {
  private readonly id = nextSessionId++;
  private readonly abort = new AbortController();
  private readonly families = new Map<string, { name: string; alias: string; faces: StoredFace[] }>();
  private readonly registrations = new Map<FontFaceSet, SetRegistration>();
  private tail: Promise<void> = Promise.resolve();
  private destroyed = false;

  constructor(
    private readonly provider: FontProvider,
    private readonly failure: FontFailure = 'fallback',
  ) {
    if (failure !== 'fallback' && failure !== 'error') {
      throw new TypeError(`invalid fontFailure: ${String(failure)}`);
    }
  }

  get strict(): boolean {
    return this.failure === 'error';
  }

  async ensure(
    values: Iterable<string | null | undefined>,
    target: FontFaceSet | null = activeFontSet(),
  ): Promise<ResolvedFonts> {
    const requested = new Map<string, string>();
    for (const value of values) {
      const name = value?.trim();
      if (!name) continue;
      requested.set(normalizedFamily(name), name);
    }
    const run = async (): Promise<void> => {
      if (this.destroyed) throw new Error('font provider session destroyed');
      const missing = [...requested].filter(([key]) => !this.families.has(key));
      if (missing.length === 0) return;
      if (this.families.size + missing.length > HARD_MAX_PROVIDER_FAMILIES) {
        throw new Error(`font provider requested more than ${HARD_MAX_PROVIDER_FAMILIES} families`);
      }
      const result = await withFontCeiling(this.provider.resolve(
        missing.map(([, name]) => name),
        { signal: this.abort.signal },
      ));
      if (!Array.isArray(result)) {
        this.abort.abort();
        throw new Error('font provider timed out');
      }
      const storedFaceCount = [...this.families.values()]
        .reduce((sum, family) => sum + family.faces.length, 0);
      let storedBytes = [...this.families.values()]
        .flatMap((family) => family.faces)
        .reduce((sum, face) => sum + face.data.byteLength, 0);
      if (storedFaceCount + result.length > HARD_MAX_PROVIDER_FACES) {
        throw new Error(`font provider returned more than ${HARD_MAX_PROVIDER_FACES} faces`);
      }
      const byFamily = new Map<string, FontAsset[]>();
      for (const asset of result) {
        const key = normalizedFamily(asset.family);
        if (!requested.has(key)) continue;
        byFamily.set(key, [...(byFamily.get(key) ?? []), asset]);
      }
      for (const [key, name] of missing) {
        const alias = `__ooxml_provider_${this.id}_${safeFamily(key)}`;
        const stored: StoredFace[] = [];
        for (const [index, asset] of (byFamily.get(key) ?? []).entries()) {
          try {
            if (this.destroyed) throw new Error('font provider session destroyed');
            const data = await withFontCeiling(sourceBytes(asset.source, this.abort.signal));
            if (!(data instanceof ArrayBuffer)) {
              this.abort.abort();
              throw new Error('font provider source timed out');
            }
            if (this.destroyed) throw new Error('font provider session destroyed');
            if (storedBytes + data.byteLength > HARD_MAX_PROVIDER_BYTES) {
              throw new Error(`font provider exceeds ${HARD_MAX_PROVIDER_BYTES} bytes`);
            }
            storedBytes += data.byteLength;
            stored.push({
              key: `${key}:${index}`,
              family: name,
              alias,
              data,
              descriptors: copyDescriptors(asset.descriptors),
            });
          } catch (error) {
            if (this.failure === 'error') throw error;
          }
        }
        if (stored.length === 0) {
          if (this.failure === 'error') throw new Error(`font provider did not resolve ${name}`);
          console.warn(`[ooxml] font provider did not resolve ${name}; using local fallback`);
        }
        this.families.set(key, { name, alias, faces: stored });
      }
    };
    const pending = this.tail.then(run);
    this.tail = pending.catch(() => undefined);
    try {
      await pending;
      if (target) await this.register(target, requested.keys());
    } catch (error) {
      if (this.failure === 'error') throw error;
      console.warn(`[ooxml] font provider failed; using local fallback: ${error instanceof Error ? error.message : String(error)}`);
    }
    return this.snapshot(requested.keys());
  }

  private async register(target: FontFaceSet, keys: Iterable<string>): Promise<void> {
    if (typeof FontFace === 'undefined') {
      if (this.failure === 'error') throw new Error('FontFace is unavailable');
      return;
    }
    const registration = this.registrations.get(target) ?? { keys: new Set<string>(), faces: [] };
    this.registrations.set(target, registration);
    const added: FontFace[] = [];
    for (const key of keys) {
      for (const stored of this.families.get(key)?.faces ?? []) {
        if (registration.keys.has(stored.key)) continue;
        const signature = `provider:${this.id}:${stored.key}`;
        const { face } = retainFace(signature, target, () => {
          const created = new FontFace(stored.alias, stored.data.slice(0), stored.descriptors);
          target.add(created);
          return created;
        });
        registration.keys.add(stored.key);
        registration.faces.push(face);
        added.push(face);
      }
    }
    const loaded = await withFontCeiling(Promise.allSettled(added.map((face) => face.load())));
    if (!Array.isArray(loaded) || loaded.some((result) => result.status === 'rejected')) {
      if (this.failure === 'error') throw new Error('font provider face failed to load');
      console.warn('[ooxml] a font provider face failed to load; using local fallback');
    }
  }

  private snapshot(keys: Iterable<string>): ResolvedFonts {
    const routes: Record<string, string> = {};
    const faces: ResolvedFontFace[] = [];
    for (const key of keys) {
      const family = this.families.get(key);
      if (!family || family.faces.length === 0) continue;
      routes[key] = family.alias;
      for (const face of family.faces) {
        faces.push({ ...face, data: face.data.slice(0) });
      }
    }
    return { routes, faces };
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.abort.abort();
    for (const registration of this.registrations.values()) releaseFaces(registration.faces);
    this.registrations.clear();
    this.families.clear();
  }
}

/** Register host-resolved faces in a worker realm and return the loaded routes. */
export async function registerResolvedFonts(
  resolved: ResolvedFonts,
  target: FontFaceSet | null = activeFontSet(),
): Promise<Readonly<{ routes: FontFamilyRoutes; faces: FontFace[] }>> {
  if (!target || typeof FontFace === 'undefined') return { routes: {}, faces: [] };
  const faces: FontFace[] = [];
  for (const [index, resolvedFace] of resolved.faces.entries()) {
    const signature = `provider-wire:${resolvedFace.alias}:${index}`;
    const { face } = retainFace(signature, target, () => {
      const created = new FontFace(
        resolvedFace.alias,
        resolvedFace.data.slice(0),
        resolvedFace.descriptors,
      );
      target.add(created);
      return created;
    });
    faces.push(face);
  }
  const loaded = await withFontCeiling(Promise.allSettled(faces.map((face) => face.load())));
  if (!Array.isArray(loaded)) {
    releaseFaces(faces);
    return { routes: {}, faces: [] };
  }
  const good = faces.filter((_, index) => loaded[index]?.status === 'fulfilled');
  for (const [index, face] of faces.entries()) {
    if (loaded[index]?.status === 'rejected') releaseFaces([face]);
  }
  const goodAliases = new Set(good.map((face) => face.family.replace(/^['"]|['"]$/g, '')));
  return {
    routes: Object.fromEntries(
      Object.entries(resolved.routes).filter(([, alias]) => goodAliases.has(alias)),
    ),
    faces: good,
  };
}
