/**
 * Embedded-font registration shared by docx / pptx / xlsx viewers.
 *
 * OOXML documents may ship the fonts they use inside the package so the text
 * renders with the authored typeface even when it is absent from the host
 * system. This module turns those embedded font parts into `FontFace` objects
 * registered into the active FontFaceSet (`document.fonts` on the main thread,
 * `self.fonts` in a worker), so a subsequent `ctx.font = '… FamilyName'` selects
 * the real font instead of a substitute.
 *
 * Two obfuscation conventions exist across the formats, handled here:
 *
 * - **WordprocessingML (docx)** stores fonts as *obfuscated* parts
 *   (`.odttf`, content type `application/vnd.openxmlformats-officedocument.obfuscatedFont`).
 *   The first 32 bytes are XOR-masked with the `w:fontKey` GUID per ECMA-376
 *   §17.8.1 (Font Embedding). {@link deobfuscateOdttf} reverses it.
 * - **PresentationML (pptx)** stores fonts as `.fntdata` parts, content type
 *   `application/x-font-ttf` (raw sfnt) or `application/x-fontdata` (EOT). Per
 *   ECMA-376 §15 (the Font part), the §17.8.1 obfuscation is *only* permitted
 *   for WordprocessingML, so pptx font bytes are consumed as-is (no XOR).
 *
 * The registration mechanism (build `FontFace` from bytes, add to the set,
 * force `load()`, await `fonts.ready`) mirrors {@link preloadGoogleFonts}: the
 * two loaders share {@link activeFontSet} and the same first-paint determinism
 * contract (fonts must be ready before the caller measures/paints text).
 */
import { activeFontSet, withFontCeiling } from './preload.js';
import { retainFace, releaseFaces, _resetFontRegistryForTests } from './font-registry.js';
import { HARD_MAX_EMBEDDED_FONT_BYTES } from '../worker/resource-policy.generated.js';

/** Shared admission check for work that must inspect an embedded font before
 * registration. Keeping the generated ceiling private prevents format loaders
 * from copying the policy value or exposing it as an unrelated public constant. */
export function embeddedFontBytesAreWithinLimit(
  bytes: Readonly<{ byteLength: number }>,
  maxBytes = HARD_MAX_EMBEDDED_FONT_BYTES,
): boolean {
  return bytes.byteLength > 0 && bytes.byteLength <= maxBytes;
}

/** Test hook — clears the shared FontFace refcount registry (does NOT touch any
 *  FontFaceSet; tests install a fresh fake set per case). Re-exported here under
 *  the embedded name so existing embedded-font tests keep their reset call; the
 *  registry itself is now shared with the Google-Fonts preloader. */
export const _resetEmbeddedRegistryForTests = _resetFontRegistryForTests;

/**
 * ECMA-376 §17.8.1 — de-obfuscate a WordprocessingML embedded font (`.odttf`).
 *
 * The algorithm (verbatim from the spec): reverse the order of the bytes in the
 * `fontKey` GUID (big-endian ordering), then XOR that 16-byte key against the
 * first 32 bytes of the font binary — once against bytes 0–15, once against
 * 16–31. The transform is its own inverse, so the same routine de-obfuscates.
 *
 * `fontKey` is the GUID string from `<w:embedRegular w:fontKey="{…}"/>` &c.,
 * e.g. `"{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}"`. Braces and hyphens are
 * optional; any 32 hex digits are accepted.
 *
 * Returns a fresh `Uint8Array` (the input is not mutated). Throws if the GUID
 * does not yield exactly 16 bytes — a malformed key must not silently produce a
 * garbage font.
 */
export function deobfuscateOdttf(bytes: Uint8Array, fontKey: string): Uint8Array {
  const key = fontKeyBytes(fontKey);
  const out = bytes.slice();
  const n = Math.min(32, out.length);
  for (let i = 0; i < n; i++) {
    // Bytes 0–15 XOR key[0..15]; bytes 16–31 XOR the SAME key again (i % 16).
    out[i] ^= key[i % 16];
  }
  return out;
}

/** Parse a `w:fontKey` GUID into its 16 bytes, reversed (big-endian) as §17.8.1
 *  requires. The reversal is over the full 16-byte string-order representation
 *  (the spec example `001B70DC-AA60-4AD5-90EC-18A0948E1EAE` →
 *  `AE1E8E94-A018-EC90-D54A-60AADC701B00`), NOT the .NET mixed-endian GUID
 *  layout. Throws on any input that is not exactly 32 hex digits. */
function fontKeyBytes(fontKey: string): Uint8Array {
  const hex = fontKey.replace(/[{}\-\s]/g, '');
  if (hex.length !== 32 || /[^0-9a-fA-F]/.test(hex)) {
    throw new Error(`invalid fontKey GUID: ${fontKey}`);
  }
  const raw = new Uint8Array(16);
  for (let i = 0; i < 16; i++) {
    raw[i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  }
  return raw.reverse();
}

/** One embedded font face to register: the family name it is drawn with, the
 *  raw part bytes, whether the part is `.odttf`-obfuscated (docx) and, if so,
 *  the `w:fontKey` GUID, plus the CSS bold/italic descriptors that map the
 *  style slot (regular / bold / italic / boldItalic). */
export interface EmbeddedFontFace {
  /** The FontFaceSet family name — the document's font name (e.g. "Ubuntu").
   *  The renderer sets `ctx.font` with this exact name so the browser selects
   *  the embedded face. */
  family: string;
  /** Raw bytes of the embedded font part (still obfuscated when `odttf`). */
  bytes: Uint8Array;
  /** `true` for docx `.odttf` parts (needs §17.8.1 de-obfuscation with
   *  `fontKey`); `false` for pptx `.fntdata` (raw sfnt / EOT). */
  odttf: boolean;
  /** The `w:fontKey` GUID; required when `odttf` is true, ignored otherwise. */
  fontKey?: string;
  /** CSS `font-weight` for this style slot ('normal' | 'bold'). */
  weight: 'normal' | 'bold';
  /** CSS `font-style` for this style slot ('normal' | 'italic'). */
  style: 'normal' | 'italic';
}

/**
 * Embedded fonts dedup + refcount through the shared {@link ./font-registry.ts}:
 * `document.fonts` (the FontFaceSet) is a process-global singleton, so opening
 * the same embedded font in two documents — or the same document twice in an SPA
 * — would otherwise add a second, byte-identical `FontFace` every time and leak
 * them all (nothing ever removed them). We dedup by a stable content signature
 * ({@link contentSignature}) and refcount: the first registration adds the
 * `FontFace`; later ones reuse it and bump refs; {@link unregisterEmbeddedFonts}
 * decrements and only `document.fonts.delete()`s the face when the last holder
 * releases it. The refcount machinery is shared with the Google-Fonts preloader.
 */

/** Cheap 32-bit FNV-1a over the (de-obfuscated) font bytes — the content key for
 *  dedup. Combined with family/weight/style so two distinct faces that happen to
 *  collide on a hash are still only merged when they truly are the same slot. */
function contentSignature(family: string, weight: string, style: string, bytes: Uint8Array): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < bytes.length; i++) {
    h ^= bytes[i];
    h = Math.imul(h, 0x01000193);
  }
  return `${family}|${weight}|${style}|${bytes.length}|${(h >>> 0).toString(16)}`;
}

/**
 * Register a set of embedded font faces into the active FontFaceSet and await
 * their load, so the caller can measure/paint text with the real typefaces.
 *
 * De-obfuscation (docx `.odttf`) is applied per face. A face whose bytes are
 * corrupt, whose `fontKey` is malformed, or that exceeds {@link maxBytes} is
 * skipped with a `console.warn` — one bad embedded font must never abort the
 * whole document. Sizes are capped as a zip-bomb / memory safety net.
 *
 * Registration is **deduped + refcounted**: registering a byte-identical face
 * (same family/weight/style/bytes) that is already in the set reuses the
 * existing `FontFace` instead of adding a duplicate. The returned array is the
 * set of shared `FontFace` objects this call holds a reference to; pass it to
 * {@link unregisterEmbeddedFonts} (e.g. from a document's `destroy()`) to release
 * them — the face leaves the FontFaceSet only when the last holder releases it.
 *
 * No-ops (resolves to `[]`) when there is no FontFaceSet or `FontFace` in the
 * current context (e.g. Node without a shim), matching {@link preloadGoogleFonts}.
 *
 * @param faces      the embedded faces to register
 * @param maxBytes   per-face size ceiling; faces larger than this are skipped
 *                   (default 30 MB — comfortably above a full CJK font, well
 *                   below a memory hazard)
 * @returns the shared `FontFace` objects registered/reused for this call
 */
export async function registerEmbeddedFonts(
  faces: Iterable<EmbeddedFontFace>,
  maxBytes = HARD_MAX_EMBEDDED_FONT_BYTES,
): Promise<FontFace[]> {
  const set = activeFontSet();
  if (!set || typeof FontFace === 'undefined') return [];

  const held: FontFace[] = []; // unique faces this call references
  const retainedSignatures = new Set<string>();
  const failed: string[] = [];
  for (const face of faces) {
    try {
      if (!embeddedFontBytesAreWithinLimit(face.bytes, maxBytes)) {
        failed.push(face.family);
        continue;
      }
      const data = face.odttf
        ? deobfuscateOdttf(face.bytes, face.fontKey ?? '')
        : face.bytes;
      // `embedded:` namespaces the signature so an embedded face can never share
      // a registry entry with a Google-Fonts face (a different keyspace).
      const sig = `embedded:${contentSignature(face.family, face.weight, face.style, data)}`;
      // One package can legally point multiple slots at the same part. Retain a
      // content-identical face only once per caller so release remains exactly
      // balanced with the registry refcount.
      if (retainedSignatures.has(sig)) continue;
      retainedSignatures.add(sig);

      // Dedup + refcount via the shared registry: a byte-identical face already
      // in THIS set → reuse + bump refs; otherwise build + add it once.
      const { face: ff } = retainFace(sig, set, () => {
        // Copy into a standalone ArrayBuffer: FontFace(source) reads the buffer,
        // and `data` may be a subarray view into a larger WASM/transfer buffer.
        const buf = data.buffer.slice(
          data.byteOffset,
          data.byteOffset + data.byteLength,
        ) as ArrayBuffer;
        const created = new FontFace(face.family, buf, {
          weight: face.weight,
          style: face.style,
        });
        set.add(created);
        return created;
      });
      held.push(ff);
    } catch {
      // Malformed fontKey / unreadable part: skip this face, keep the document.
      failed.push(face.family);
    }
  }

  let loaded = held;
  if (held.length > 0) {
    // FontFace.load() is idempotent. Calling it for reused faces also closes the
    // concurrent-holder race: every caller observes the shared face's real load
    // result instead of assuming that the first holder already succeeded.
    const results = await withFontCeiling(
      Promise.allSettled(held.map((font) => Promise.resolve().then(() => font.load()))),
    );
    if (!Array.isArray(results)) {
      // A wedged face reached the hard time ceiling. Do not advertise an
      // indeterminate registration as usable or suppress a substitute font.
      failed.push(...held.map((font) => font.family));
      releaseFaces(held);
      loaded = [];
    } else {
      loaded = [];
      results.forEach((result, index) => {
        const font = held[index];
        if (result.status === 'fulfilled') {
          loaded.push(font);
        } else {
          failed.push(font.family);
          releaseFaces([font]);
        }
      });
      // Loading the exact faces above is the readiness proof. This extra wait
      // lets the FontFaceSet settle without making a wedged `ready` promise turn
      // already-loaded faces back into failures.
      await withFontCeiling(set.ready);
    }
  }

  if (failed.length > 0) {
    console.warn(
      `[ooxml] failed to register embedded font(s): ${[...new Set(failed)].join(', ')}; ` +
        `falling back to substitute fonts (text may shift or differ).`,
    );
  }

  return loaded;
}

/**
 * Release the embedded `FontFace` objects a document registered (the array
 * returned by {@link registerEmbeddedFonts}). Each face's refcount is
 * decremented; the face is removed from its FontFaceSet only when the last
 * holder releases it, so a font shared by two open documents survives until both
 * are destroyed. Safe to call with faces from a context without a FontFaceSet
 * (no-op). Prevents the SPA leak where every opened document left its FontFace
 * objects in `document.fonts` forever.
 *
 * **Idempotent / double-release safe (refs are never over-decremented).** Two
 * independent guards protect a font another document is still using from being
 * evicted by a stray double-release:
 *
 * - *Within one call*: the same `FontFace` appearing twice in `faces` (a caller
 *   passing a list with duplicates) is decremented AT MOST ONCE — a per-call
 *   `seen` set skips repeats. Without this, `unregister([F, F])` would drop refs
 *   by 2 and could delete `F` while a second holder still references it.
 * - *Across calls*: once a face's refcount reaches 0 its registry entry is
 *   removed, so a later call that passes the same (now-unregistered) face finds
 *   no entry and is a no-op — it can never push another registration's refs
 *   negative.
 *
 * These make releasing the exact same face list twice harmless. Callers should
 * still drop their reference after releasing (e.g. `DocxDocument.destroy()`
 * clears its held array), but a mistaken double `destroy()` no longer corrupts
 * the shared refcount.
 */
export function unregisterEmbeddedFonts(faces: Iterable<FontFace>): void {
  releaseFaces(faces);
}
