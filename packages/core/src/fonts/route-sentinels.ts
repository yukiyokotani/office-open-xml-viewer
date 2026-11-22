import { retainFace, releaseFaces, withFontCeiling } from './font-registry.js';
import type { CanvasFontBoxProbeContext } from './canvas-font-box.js';
import { quoteCssFontFamily } from '../internal/css-font-family.js';

// ── Route sentinels ─────────────────────────────────────────────────────────
//
// A CSS family list names faces by family, and a family that is not installed
// is skipped silently. Neither which entry Canvas uses as the first available
// face nor which face draws a given cluster is observable by name. Two
// generated registrations make both observable. Each maps every Unicode scalar
// (cmap format 13: many-to-one ranges, the format of a last-resort font) to one
// glyph, and they differ only in their advance and font box:
//
//   sentinel A: advance 0.50 em, box 0.70 / 0.20 em, ink 8 em up / 3 em down
//   sentinel B: advance 0.75 em, box 0.60 / 0.30 em, ink 9 em up / 4 em down
//
// A sentinel glyph therefore makes the cluster's ink box at least 11 em tall
// wherever it is placed, unless the cluster's own ink is that tall already
// (then ownership is not claimed).
//
// Appending a sentinel to a family-list prefix ends font fallback there: every
// cluster the prefix cannot draw is drawn by the sentinel, never by a later
// family or a system fallback face.
//
// - First available family: `family, A` and `family, B` report the same Canvas
//   font box exactly when the family resolves to a face (the box is that
//   face's); otherwise they report the two sentinel boxes. The first available
//   family of a list is its first family whose boxes agree.
// - Ownership: the families before the first available one resolve to no
//   face, so a cluster is drawn by that family in both `family, A` and
//   `family, B`, or (partly) by a sentinel in each. Equal advance and ink extents therefore prove that no
//   sentinel glyph, i.e. no face after that family, drew any part of it; any
//   sentinel glyph changes the advance or the ink box. This holds only for a
//   cluster the engine routes through the family list at all, which a
//   sentinel-only control establishes first (nativeRouteClusterOwned).
//   Families before the first available one are system families that matched
//   nothing: callers decline a route when one of them is a registered
//   FontFace (a unicode-range resource can draw some clusters without being
//   the first available face).
//
// This proves which faces of the list draw a cluster without a coverage table,
// a family-name guess or size-adjust (WebKit keeps ink extents under
// size-adjust, so the size-adjust oracle of local-line-metrics declines there).
// Both registrations validate themselves before use; an engine that cannot
// load or measure them declines.

export const SENTINELS = [
  { family: '__ooxml_route_sentinel_a_6d2e91__', advance: 512, ascent: 716, descent: 204, inkUp: 8192, inkDown: 3072 },
  { family: '__ooxml_route_sentinel_b_6d2e91__', advance: 768, ascent: 614, descent: 307, inkUp: 9216, inkDown: 4096 },
] as const;
/** Equal to the probe size, so every sentinel metric is a whole pixel. */
export const SENTINEL_UPM = 1024;
export const SENTINEL_FAMILIES = SENTINELS.map((sentinel) => sentinel.family);
/** One size for every probe, which is also the half-size strut: engines cache
 * font lookups per size, so the listed families are resolved at two sizes in
 * all. */
export const PROBE_EM = 1024;

/** Minimal TrueType face: .notdef plus one box glyph that every scalar maps to. */
function sentinelFontBytes(
  advance: number, ascent: number, descent: number, inkUp: number, inkDown: number, name: string,
): Uint8Array {
  const u16 = (value: number) => [(value >> 8) & 255, value & 255];
  const i16 = (value: number) => u16(value & 0xffff);
  const u32 = (value: number) => [(value >>> 24) & 255, (value >>> 16) & 255, (value >>> 8) & 255, value & 255];
  const zeros = (count: number) => new Array<number>(count).fill(0);
  const inkTop = inkUp;
  const glyph = [...i16(1), ...i16(0), ...i16(-inkDown), ...i16(advance), ...i16(inkTop), ...u16(3), ...u16(0), 1, 1, 1, 1];
  const points = [[0, -inkDown], [advance, -inkDown], [advance, inkTop], [0, inkTop]];
  const xs: number[] = [];
  const ys: number[] = [];
  let previousX = 0;
  let previousY = 0;
  for (const [x, y] of points) {
    xs.push(...i16(x! - previousX));
    ys.push(...i16(y! - previousY));
    previousX = x!;
    previousY = y!;
  }
  glyph.push(...xs, ...ys);
  while (glyph.length % 4) glyph.push(0);
  const nameRecords: number[] = [];
  const nameStorage: number[] = [];
  for (const [id, value] of [[1, name], [2, 'Regular'], [4, name], [6, name]] as const) {
    const bytes = [...value].flatMap((character) => u16(character.charCodeAt(0)));
    nameRecords.push(...u16(3), ...u16(1), ...u16(0x409), ...u16(id), ...u16(bytes.length), ...u16(nameStorage.length));
    nameStorage.push(...bytes);
  }
  // Surrogate code points are not scalars; every other value maps to glyph 1.
  const groups = [[0x0000, 0xd7ff], [0xe000, 0x10ffff]] as const;
  const cmap13 = [...u16(13), ...u16(0), ...u32(16 + 12 * groups.length), ...u32(0), ...u32(groups.length)];
  for (const [start, end] of groups) cmap13.push(...u32(start), ...u32(end), ...u32(1));
  const tables: Record<string, number[]> = {
    'OS/2': [
      ...u16(4), ...i16(advance), ...u16(400), ...u16(5), ...u16(0), ...zeros(16),
      ...i16(50), ...i16(250), ...i16(0), ...zeros(10), ...zeros(16), 0x4f, 0x4f, 0x58, 0x4d,
      ...u16(0x00c0), ...u16(0x20), ...u16(0xffff),
      ...i16(ascent), ...i16(-descent), ...i16(0), ...u16(ascent), ...u16(descent),
      ...u32(1), ...u32(0), ...i16(inkTop), ...i16(inkTop), ...u16(0), ...u16(0x20), ...u16(1),
    ],
    cmap: [...u16(0), ...u16(1), ...u16(3), ...u16(10), ...u32(12), ...cmap13],
    glyf: glyph,
    head: [
      ...u32(0x00010000), ...u32(0x00010000), ...u32(0), ...u32(0x5f0f3cf5), ...u16(0x000b),
      ...u16(SENTINEL_UPM), ...zeros(16), ...i16(0), ...i16(-inkDown), ...i16(advance), ...i16(inkTop),
      ...u16(0), ...u16(8), ...i16(2), ...i16(0), ...i16(0),
    ],
    hhea: [
      ...u32(0x00010000), ...i16(ascent), ...i16(-descent), ...i16(0), ...u16(advance),
      ...i16(0), ...i16(0), ...i16(advance), ...i16(1), ...i16(0), ...i16(0), ...zeros(8), ...i16(0), ...u16(2),
    ],
    hmtx: [...u16(advance), ...i16(0), ...u16(advance), ...i16(0)],
    loca: [...u16(0), ...u16(0), ...u16(glyph.length / 2)],
    maxp: [...u32(0x00010000), ...u16(2), ...u16(4), ...u16(1), ...zeros(4), ...u16(2), ...zeros(16)],
    name: [...u16(0), ...u16(4), ...u16(6 + nameRecords.length), ...nameRecords, ...nameStorage],
    post: [...u32(0x00030000), ...zeros(4), ...i16(-100), ...i16(50), ...zeros(20)],
  };
  const tags = Object.keys(tables).sort();
  let selector = 0;
  while ((1 << (selector + 1)) <= tags.length) selector += 1;
  const searchRange = (1 << selector) * 16;
  const checksum = (bytes: ArrayLike<number>) => {
    let sum = 0;
    for (let index = 0; index < bytes.length; index += 4) {
      sum = (sum + ((((bytes[index] ?? 0) << 24) >>> 0) + ((bytes[index + 1] ?? 0) << 16)
        + ((bytes[index + 2] ?? 0) << 8) + (bytes[index + 3] ?? 0))) >>> 0;
    }
    return sum;
  };
  const directory: number[] = [];
  const body: number[] = [];
  let offset = 12 + 16 * tags.length;
  let headOffset = 0;
  for (const tag of tags) {
    const data = tables[tag]!;
    if (tag === 'head') headOffset = offset;
    directory.push(...[...tag].map((character) => character.charCodeAt(0)),
      ...u32(checksum(data)), ...u32(offset), ...u32(data.length));
    body.push(...data);
    while (body.length % 4) body.push(0);
    offset = 12 + 16 * tags.length + body.length;
  }
  const font = new Uint8Array([
    ...u32(0x00010000), ...u16(tags.length), ...u16(searchRange), ...u16(selector),
    ...u16(tags.length * 16 - searchRange), ...directory, ...body,
  ]);
  font.set(u32((0xb1b0afba - checksum(font)) >>> 0), headOffset + 8);
  return font;
}

export function sentinelsWork(context: CanvasFontBoxProbeContext): boolean {
  try {
    return SENTINELS.every((sentinel) => {
      context.font = `${PROBE_EM}px ${quoteCssFontFamily(sentinel.family)}`;
      const measured = context.measureText('A');
      const scale = PROBE_EM / SENTINEL_UPM;
      return Math.abs(measured.width - sentinel.advance * scale) < 0.01
        && Math.abs(measured.actualBoundingBoxAscent - sentinel.inkUp * scale) < 1
        && Math.abs(measured.fontBoundingBoxAscent - sentinel.ascent * scale) < 0.01
        && Math.abs(measured.fontBoundingBoxDescent - sentinel.descent * scale) < 0.01;
    });
  } catch {
    return false;
  }
}

/** Register both sentinels in `fontSet` (refcounted). Returns the retained
 * faces for the caller's release list, or null (nothing retained) when this
 * realm cannot register or load them. Registration is not proof that a realm
 * applies them: every consumer (measurement, verification, ownership)
 * validates the sentinel geometry itself before relying on it, and declines
 * otherwise. */
export async function loadNativeRouteSentinels(fontSet: FontFaceSet | null): Promise<FontFace[] | null> {
  if (!fontSet || typeof FontFace === 'undefined') return null;
  const faces: FontFace[] = [];
  try {
    for (const sentinel of SENTINELS) {
      faces.push(retainFace(`native-route-sentinel:${sentinel.family}`, fontSet, () => {
        const face = new FontFace(sentinel.family, sentinelFontBytes(
          sentinel.advance, sentinel.ascent, sentinel.descent, sentinel.inkUp, sentinel.inkDown,
          sentinel.family.replace(/_/g, ''),
        ) as BufferSource);
        fontSet.add(face);
        return face;
      }).face);
    }
    const loaded = await withFontCeiling(Promise.all(faces.map((face) => face.load())));
    if (loaded) return faces;
  } catch { /* Unregistrable sentinels decline every native route. */ }
  releaseFaces(faces);
  return null;
}

// ── Family lists ──────────────────────────────────────────────────────────────

export { parseFamilyList, type FamilyEntry } from '../internal/css-font-family.js';
