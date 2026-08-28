/**
 * Refcounted FontFace registry shared by the embedded-font loader
 * ({@link ./embedded.ts}) and the Google-Fonts preloader ({@link ./preload.ts}).
 *
 * Each DOM/worker realm owns a FontFaceSet. Opening the same font repeatedly in
 * one set would otherwise add duplicate faces and leak them; same-origin child
 * windows additionally require their own registrations. Both loaders share
 * this per-set dedup + refcount so a face is added once to each target set and
 * removed only when that set's last holder releases it.
 *
 * The registry is deliberately format-agnostic: callers own how they compute a
 * face's signature (embedded fonts hash the de-obfuscated bytes; Google Fonts
 * key on the CSS url + family + descriptors) and how they build the `FontFace`.
 * The registry owns only the shared concern — *this exact face is referenced by
 * N holders; delete it from its set at N = 0* — so neither loader duplicates the
 * refcount / last-release / double-release-safety logic.
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

/** Test hook — clears the shared refcount registry (does NOT touch any
 *  FontFaceSet; tests install a fresh fake set per case). */
export function _resetFontRegistryForTests(): void {
  registry.clear();
}

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
