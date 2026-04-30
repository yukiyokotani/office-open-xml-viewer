/**
 * Google Fonts preload utility shared by docx / pptx / xlsx viewers.
 *
 * The contract is intentionally narrow: callers pass the set of font-family
 * names they want available, plus a static map from a lower-cased key to a
 * Google Fonts CSS URL (and optionally an alternate FontFaceSet family name
 * for Office substitutes such as Calibri → Carlito). Names without a map
 * entry are skipped (the renderer falls back to the system font).
 *
 * Font load is forced via `face.load()` rather than `document.fonts.load()`
 * because canvas-only rendering does not put glyphs into the DOM, so the
 * unicode-range gating in modern Google Fonts CSS would otherwise leave the
 * `FontFace` entries in the `unloaded` state — the first paint would then
 * use a system fallback and shift once a later interaction re-rasterized
 * the canvas after the font landed.
 */
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

const PRELOAD_TIMEOUT_MS = 3000;
const REGISTRATION_POLL_MS = 20;
const REGISTRATION_POLL_ATTEMPTS = 25;

export async function preloadGoogleFonts(
  fontNames: Iterable<string | null | undefined>,
  map: Record<string, FontPreloadEntry>,
): Promise<void> {
  if (typeof document === 'undefined') return;

  const seen = new Set<string>();
  const targetFamilies = new Set<string>();

  for (const name of fontNames) {
    if (!name) continue;
    const key = name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    const entry = map[key];
    if (!entry) continue;

    if (!document.querySelector(`link[href="${entry.url}"]`)) {
      try {
        const link = document.createElement('link');
        link.rel = 'stylesheet';
        link.href = entry.url;
        document.head.appendChild(link);
      } catch {
        // Network or DOM error — silently skip; renderer falls back to system.
      }
    }

    targetFamilies.add((entry.loadFamily ?? name).toLowerCase());
  }

  if (targetFamilies.size === 0) return;

  // After a fresh `<link>` append, FontFaceSet briefly reports zero matching
  // FontFace entries while the stylesheet parses. Poll for up to ~500 ms.
  let registered: FontFace[] = [];
  for (let i = 0; i < REGISTRATION_POLL_ATTEMPTS; i++) {
    registered = [...document.fonts].filter((f) =>
      targetFamilies.has(f.family.toLowerCase()),
    );
    if (registered.length > 0) break;
    await new Promise<void>((r) => setTimeout(r, REGISTRATION_POLL_MS));
  }

  await Promise.race([
    Promise.allSettled(registered.map((f) => f.load())).then(() =>
      document.fonts.ready,
    ),
    new Promise<void>((resolve) => setTimeout(resolve, PRELOAD_TIMEOUT_MS)),
  ]);
}
