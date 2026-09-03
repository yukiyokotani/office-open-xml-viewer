/**
 * Compatibility Google Fonts preload utility.
 *
 * The viewers use {@link GoogleFontsProvider} through the shared provider
 * session. These functions remain for callers of the older low-level API.
 * Callers pass the set of font-family
 * names they want available, plus a static map from a lower-cased key to a
 * Google Fonts CSS URL (and optionally an alternate FontFaceSet family name
 * for Office substitutes such as Calibri → Carlito). Names without a map
 * entry are skipped (the renderer falls back to the system font).
 *
 * Rather than inject a `<link rel="stylesheet">` and read `document.fonts`
 * (which is impossible inside a Web Worker), this fetches the Google Fonts CSS
 * directly, parses its `@font-face` rules, and registers `FontFace` objects
 * into whichever FontFaceSet exists in the current JS context —
 * `document.fonts` on the main thread, `self.fonts` in a worker. This keeps the
 * loader FontFaceSet-agnostic so both the main-thread and worker rendering
 * modes share one code path.
 *
 * Font load is forced via `face.load()` rather than `FontFaceSet.load()`
 * because canvas-only rendering does not put glyphs into the DOM, so the
 * unicode-range gating in modern Google Fonts CSS would otherwise leave the
 * `FontFace` entries in the `unloaded` state — the first paint would then
 * use a system fallback and shift once a later interaction re-rasterized
 * the canvas after the font landed.
 */
import { retainFace, releaseFaces } from './font-registry.js';
import { activeFontSet, withFontCeiling } from './preload.js';

export interface FontPreloadEntry {
  /** Google Fonts CSS URL — `display=swap` recommended. */
  url: string;
  /**
   * Family name to drive {@link FontFaceSet} loading when the substitute
   * differs from the requested face (e.g. Calibri → Carlito). Defaults to
   * the requested name when omitted.
   */
  loadFamily?: string;
}

/** In-flight (or completed) stylesheet FETCH promise per CSS url, keyed by url.
 *  Resolves with the PARSED `@font-face` rules of that stylesheet (empty on a
 *  failed fetch). This dedups only the NETWORK fetch: a concurrent caller with
 *  the same url JOINs the first fetch instead of re-downloading, and the join
 *  cannot resolve before the fetch settles (first-paint determinism). The actual
 *  `FontFace` objects are dedup + refcounted separately per call in the shared
 *  {@link ./font-registry.ts} (so holders in one FontFaceSet share one face;
 *  another document set receives its own face). The stored
 *  promise ALWAYS resolves: a failed fetch records the failed families + deletes
 *  the entry (so a later call retries) inside the producing call, so a cache-hit
 *  awaiter never throws. */
const cssFetches = new Map<string, Promise<ParsedFontFace[]>>();

/** Test hook — clears the per-context CSS fetch cache. */
export function _resetCssCacheForTests(): void {
  cssFetches.clear();
}

export interface ParsedFontFace {
  family: string;
  src: string;
  descriptors: FontFaceDescriptors;
}

/** Fetch and parse one Google Fonts stylesheet with cross-document request
 * deduplication. Failed requests are evicted so a later document can retry. */
export async function loadGoogleFontRules(url: string): Promise<ParsedFontFace[]> {
  const inFlight = cssFetches.get(url);
  if (inFlight) return inFlight;
  const request = (async (): Promise<ParsedFontFace[]> => {
    try {
      const response = await fetch(url);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      return parseFontFaceRules(await response.text());
    } catch {
      cssFetches.delete(url);
      return [];
    }
  })();
  cssFetches.set(url, request);
  return request;
}

/** Extract @font-face rules from a Google Fonts stylesheet. Deliberately
 *  minimal: Google's CSS is machine-generated (one declaration per line, no
 *  nesting), so a brace-block regex is sufficient and avoids a CSS parser. */
export function parseFontFaceRules(css: string): ParsedFontFace[] {
  const faces: ParsedFontFace[] = [];
  const blockRe = /@font-face\s*\{([^}]*)\}/g;
  let m: RegExpExecArray | null;
  while ((m = blockRe.exec(css))) {
    const body = m[1];
    const prop = (name: string): string | undefined =>
      body.match(new RegExp(`(?:^|;|\\n)\\s*${name}\\s*:\\s*([^;]+)`, 'i'))?.[1].trim();
    const familyRaw = prop('font-family');
    const src = prop('src');
    if (!familyRaw || !src) continue;
    const descriptors: FontFaceDescriptors = {};
    const style = prop('font-style');
    if (style) descriptors.style = style;
    const weight = prop('font-weight');
    if (weight) descriptors.weight = weight;
    const stretch = prop('font-stretch');
    if (stretch) descriptors.stretch = stretch;
    const unicodeRange = prop('unicode-range');
    if (unicodeRange) descriptors.unicodeRange = unicodeRange;
    faces.push({ family: familyRaw.replace(/^['"]|['"]$/g, ''), src, descriptors });
  }
  return faces;
}

/** Stable signature for one Google-Fonts `@font-face`, keyed to the CSS url it
 *  came from + its identity (family + all CSS descriptors). `gfonts:` namespaces
 *  it away from embedded-font signatures in the shared registry. Two documents
 *  requesting the same web font compute the SAME signature and, within one
 *  FontFaceSet, share one refcounted FontFace; distinct subsets
 *  key distinctly. */
function googleFaceSignature(url: string, f: ParsedFontFace): string {
  const d = f.descriptors;
  return [
    'gfonts',
    url,
    f.family.toLowerCase(),
    d.style ?? '',
    d.weight ?? '',
    d.stretch ?? '',
    d.unicodeRange ?? '',
    // `src` last: two rules identical in every descriptor but pointing at
    // different files (theoretically) still key apart.
    f.src,
  ].join('|');
}

/** @deprecated Prefer `useGoogleFonts` or `GoogleFontsProvider`. */
export async function preloadGoogleFonts(
  fontNames: Iterable<string | null | undefined>,
  map: Record<string, FontPreloadEntry>,
  targetFontSet: FontFaceSet | null = activeFontSet(),
): Promise<FontFace[]> {
  const fonts = targetFontSet;
  if (!fonts || typeof FontFace === 'undefined' || typeof fetch === 'undefined') return [];

  const seen = new Set<string>();
  const targetFamilies = new Set<string>();
  const cssUrls = new Set<string>();
  // url → lower-cased target families requested from that stylesheet. Built in
  // the main loop so the failure path can attribute a failed fetch to the
  // families this call actually asked for, without re-scanning `map` under the
  // implicit assumption that a map key equals the lower-cased requested name.
  const urlTargets = new Map<string, Set<string>>();
  // Families that could not be made available (stylesheet fetch failed, or every
  // matching FontFace.load() rejected). Surfaced once at the end via a single
  // console.warn so a failed web font no longer falls back to a system face
  // completely silently. The renderer still degrades gracefully — this is
  // diagnostics only.
  const failedFamilies = new Set<string>();

  for (const name of fontNames) {
    if (!name) continue;
    const key = name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    const entry = map[key];
    if (!entry) continue;
    cssUrls.add(entry.url);
    const family = (entry.loadFamily ?? name).toLowerCase();
    targetFamilies.add(family);
    let targets = urlTargets.get(entry.url);
    if (!targets) {
      targets = new Set<string>();
      urlTargets.set(entry.url, targets);
    }
    targets.add(family);
  }
  if (targetFamilies.size === 0) return [];

  // 1) Fetch each stylesheet once per JS context (network dedup via cssFetches)
  //    and PARSE its `@font-face` rules. The rules are shared data; the FontFace
  //    objects are created + refcounted per call in step 2 so holders targeting
  //    the same set share one face and it leaves only after all release it.
  const parsedGroups = await withFontCeiling(
    Promise.all(
      [...cssUrls].map(async (url): Promise<{ url: string; rules: ParsedFontFace[] }> => {
        // Cache hit: AWAIT the in-flight fetch so concurrent callers join the
        // first download rather than racing past it. The stored promise never
        // rejects (see cssFetches), so a joined-then-failed fetch yields [].
        const rules = await loadGoogleFontRules(url);
        if (rules.length === 0) {
          for (const family of urlTargets.get(url) ?? []) failedFamilies.add(family);
        }
        return { url, rules };
      }),
    ),
  );

  // 2) Retain a shared, refcounted FontFace per parsed rule. The FIRST holder of
  //    a rule (across all open documents) creates the FontFace, adds it to the
  //    set, and must force-`load()` it; a later holder reuses the shared face
  //    (already loaded) and only bumps its refcount. `held` is what THIS call
  //    references — returned so the caller can release it in `destroy()`, the
  //    fix for the SPA leak where every opened document left its Google FontFace
  //    objects in `document.fonts` forever.
  const held: FontFace[] = [];
  const toLoad: FontFace[] = [];
  for (const group of Array.isArray(parsedGroups) ? parsedGroups : []) {
    for (const rule of group.rules) {
      const sig = googleFaceSignature(group.url, rule);
      const { face, isNew } = retainFace(sig, fonts, () => {
        const created = new FontFace(rule.family, rule.src, rule.descriptors);
        fonts.add(created);
        return created;
      });
      held.push(face);
      if (isNew) toLoad.push(face);
    }
  }

  // 3) Force-load only the FontFaces this call newly created and AWAIT them —
  //    no timeout race that could resolve mid-download. `face.load()` is required
  //    because unicode-range gating would otherwise leave the faces `unloaded`
  //    and the first canvas paint would use a system fallback, shifting once a
  //    later interaction re-rasterized. We load the FontFace objects WE created
  //    rather than re-selecting them from the set by family name: `FontFace.family`
  //    serializes a multi-word name back WITH quotes (e.g. `"Nunito Sans"`), so a
  //    `family`-string filter silently matches nothing and the fonts never load —
  //    the bug this avoids. (Reused faces were already loaded by their first holder.)
  if (toLoad.length > 0) {
    await withFontCeiling(
      Promise.allSettled(toLoad.map((f) => f.load())).then((results) => {
        results.forEach((res, i) => {
          if (res.status === 'rejected') {
            failedFamilies.add(toLoad[i].family.replace(/['"]/g, '').toLowerCase());
          }
        });
        return fonts.ready;
      }),
    );
  }

  if (failedFamilies.size > 0) {
    console.warn(
      `[ooxml] failed to preload web font(s): ${[...failedFamilies].join(', ')}; ` +
        `falling back to system fonts (text may shift or differ).`,
    );
  }

  return held;
}

/**
 * Release the Google-Fonts `FontFace` objects a document/presentation/workbook
 * preloaded (the array returned by {@link preloadGoogleFonts}). Refcounted +
 * dedup-safe via the shared {@link ./font-registry.ts}: each face is removed
 * from its FontFaceSet only when the LAST holder releases it, so a web font used
 * by two holders in the same document survives until both are destroyed. Double-release safe
 * (a face passed twice, or a re-release after full release, is a no-op) and safe
 * in a context without a FontFaceSet. Twin of `unregisterEmbeddedFonts`; called
 * from each viewer's `destroy()` to fix the SPA leak where every opened document
 * left its Google FontFace objects in `document.fonts` forever.
 */
/** @deprecated Prefer session-owned cleanup through `GoogleFontsProvider`. */
export function unloadGoogleFonts(faces: Iterable<FontFace>): void {
  releaseFaces(faces);
}
