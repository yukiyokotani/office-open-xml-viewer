/**
 * Refcounted FontFace registry shared by the embedded-font loader
 * ({@link ./embedded.ts}) and the library's probe registrations (route
 * sentinels, local-metric probes). Font sources register per document under
 * private families instead (`source-session.ts`).
 *
 * Each DOM/worker realm owns a FontFaceSet. Opening the same font repeatedly in
 * one set would otherwise add duplicate faces and leak them; same-origin child
 * windows additionally require their own registrations. Its users share
 * this per-set dedup + refcount so a face is added once to each target set and
 * removed only when that set's last holder releases it.
 *
 * The registry is deliberately format-agnostic: callers own how they compute a
 * face's signature (embedded fonts hash the de-obfuscated bytes) and how they
 * build the `FontFace`. The registry owns only the shared concern — *this
 * exact face is referenced by N holders; delete it from its set at N = 0* — so
 * no user duplicates the refcount / last-release / double-release-safety logic.
 */

/** One shared, refcounted FontFace registration. */
interface FontRegistration {
  face: FontFace;
  /** The FontFaceSet the face was added to (`document.fonts` / `self.fonts`).
   *  Held so release can `delete()` from the SAME set the retain added to, and
   *  so a stale signature colliding across two different sets never mixes. */
  set: FontFaceSet;
  refs: number;
}

/** Signature → registrations keyed by FontFaceSet. A same-origin popup owns a
 * distinct set even though it shares the opener's JavaScript module instance,
 * so the set identity is part of the registration key. */
const registry = new Map<string, Map<FontFaceSet, FontRegistration>>();

/** Result of {@link retainFace}: the shared `FontFace` this caller now holds a
 *  reference to, and whether THIS call created it. Loaders may use `isNew` to
 *  avoid redundant work where their failure semantics permit it; loaders that
 *  must observe an in-flight shared load can call idempotent `FontFace.load()`
 *  for both new and reused registrations. */
export interface RetainResult {
  face: FontFace;
  /** `true` when this retain created + added the face (first holder); `false`
   *  when it reused an existing shared registration (refs bumped). */
  isNew: boolean;
}

/**
 * Retain a shared FontFace for `sig` in `set`, bumping its refcount.
 *
 * - First holder of `sig` (in this set): `create()` builds the `FontFace`, it is
 *   added to the set, and `{ face, isNew: true }` is returned.
 * - A later holder of the same `sig`: the existing shared face is reused, its
 *   refcount bumped, and `{ face, isNew: false }` returned. The face may still
 *   be loading if two holders registered concurrently.
 *
 * A signature whose registration lives in a DIFFERENT set (e.g. a stale
 * cross-context collision) is treated as absent: a fresh registration replaces
 * it, so a face is never handed back from a set the caller is not adding to.
 *
 * `create()` must both construct the `FontFace` AND add it to `set` (the two are
 * inseparable — the registry cannot know a loader's `set.add` semantics), then
 * return the face. It runs ONLY on the first-holder path.
 */
export function retainFace(sig: string, set: FontFaceSet, create: () => FontFace): RetainResult {
  const registrations = registry.get(sig);
  const existing = registrations?.get(set);
  if (existing) {
    existing.refs++;
    return { face: existing.face, isNew: false };
  }
  const face = create();
  const bySet = registrations ?? new Map<FontFaceSet, FontRegistration>();
  bySet.set(set, { face, set, refs: 1 });
  registry.set(sig, bySet);
  return { face, isNew: true };
}

/**
 * Release a set of shared `FontFace` objects (as returned by the loaders'
 * retain paths). Each face's refcount is decremented; the face is removed from
 * its FontFaceSet only when the last holder releases it, so a font shared by two
 * open documents survives until both are destroyed. Safe to call with faces the
 * registry does not know (no-op).
 *
 * **Idempotent / double-release safe (refs are never over-decremented).** Two
 * independent guards protect a font another document is still using from being
 * evicted by a stray double-release:
 *
 * - *Within one call*: the same `FontFace` appearing twice in `faces` (a caller
 *   passing a list with duplicates) is decremented AT MOST ONCE — a per-call
 *   `seen` set skips repeats. Without this, `release([F, F])` would drop refs by
 *   2 and could delete `F` while a second holder still references it.
 * - *Across calls*: once a face's refcount reaches 0 its registry entry is
 *   removed, so a later call that passes the same (now-unregistered) face finds
 *   no entry and is a no-op — it can never push another registration's refs
 *   negative.
 */
export function releaseFaces(faces: Iterable<FontFace>): void {
  // Guard against the same face appearing more than once in THIS call so a
  // duplicate cannot decrement a shared registration's refcount twice.
  const seen = new Set<FontFace>();
  for (const face of faces) {
    if (seen.has(face)) continue;
    seen.add(face);
    // Find the registration for this face (identity match). The registry is
    // small (one entry per distinct face), so a linear scan is fine. A face that
    // was already fully released has no entry → this loop finds nothing and the
    // release is a no-op (cross-call idempotency).
    for (const [sig, registrations] of registry) {
      let matched = false;
      for (const [set, reg] of registrations) {
        if (reg.face !== face) continue;
        matched = true;
        reg.refs--;
        if (reg.refs <= 0) {
          try {
            reg.set.delete(face);
          } catch {
            /* a set without delete() (older shim / mock): drop the entry anyway */
          }
          registrations.delete(set);
          if (registrations.size === 0) registry.delete(sig);
        }
        break;
      }
      if (matched) break;
    }
  }
}

// ---- realm font-set primitives -----------------------------------------------
// Shared by every font loader (font sources, embedded fonts, installed-face
// probes and line-metric preflights): library resource policy, not OOXML
// behaviour. Each JS realm owns one FontFaceSet.

/**
 * Hard ceiling so a wedged network (a stylesheet fetch that never settles, or a
 * `FontFace.load()` that never resolves) cannot hang the caller forever. This is
 * a SAFETY NET, not the normal exit: on a reachable network every awaited
 * promise settles well within it, so first paint is deterministic. It is
 * intentionally generous — racing font loads against a short timeout could
 * resolve while faces were still downloading, which is exactly the cold-cache
 * flicker the await-before-paint contract prevents.
 */
export const FONT_LOAD_CEILING_MS = 15000;

/** Race a font-load promise against {@link FONT_LOAD_CEILING_MS} so a wedged
 *  network or a `FontFace.load()` that never settles cannot hang the caller
 *  forever. Shared by the embedded-font and local-face loaders (same
 *  first-paint determinism contract); font sources abort at the same ceiling
 *  instead of racing it (`source-session.ts`). */
export function withFontCeiling<T>(p: Promise<T>): Promise<T | void> {
  return Promise.race([
    p,
    new Promise<void>((resolve) => setTimeout(resolve, FONT_LOAD_CEILING_MS)),
  ]);
}

/** The FontFaceSet of the current context: `document.fonts` on the main
 *  thread, `self.fonts` in a worker, null elsewhere (Node without a shim).
 *  Every loader resolves its default target through this one rule, so faces of
 *  one document land in the same set. */
export function activeFontSet(): FontFaceSet | null {
  if (typeof document !== 'undefined' && document && document.fonts) return document.fonts;
  if (typeof self !== 'undefined' && self && 'fonts' in self) {
    return (self as unknown as { fonts: FontFaceSet }).fonts;
  }
  return null;
}

/** Demanded Unicode scalars of one load (sorted or first-seen), or `'all'`:
 *  the conservative answer that loads every `unicode-range` subset. */
export type FontPreloadDemand = 'all' | readonly number[];
