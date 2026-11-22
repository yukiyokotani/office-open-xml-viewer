import { playEmf } from '../image/emf.js';
import { isEmf, isWmf, playWmf } from '../image/wmf.js';

/** Every text the WMF / EMF players measure or paint for `bytes`, played on
 * recording contexts under both host outcomes the WMF symbol-charset
 * substitution can see (the authored face present or absent). Font-demand
 * coverage tests compare it with core `metafileDemandScalars`. */
export function metafilePaintedTexts(bytes: Uint8Array): string[] {
  if (!isWmf(bytes) && !isEmf(bytes)) return [];
  const texts: string[] = [];
  for (const hostHasFace of [false, true]) {
    let font = '10px serif';
    const metrics = (text: string) => {
      // The player compares a face against an absent control family: equal
      // geometry means the host lacks the face.
      const control = /^[^"]*"__ooxml_wmf_missing_face_control_[0-9a-f]+__"/u.test(font);
      const width = [...text].length * (hostHasFace && !control ? 6 : 5);
      return {
        width,
        fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
        actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
        actualBoundingBoxLeft: 0, actualBoundingBoxRight: width,
      };
    };
    const noop = () => {};
    const state: Record<string | symbol, unknown> = {
      get font() { return font; },
      set font(value: string) { font = value; },
      measureText: (text: string) => { texts.push(text); return metrics(text); },
      fillText: (text: string) => { texts.push(text); },
      strokeText: (text: string) => { texts.push(text); },
      getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
      createLinearGradient: () => ({ addColorStop: noop }),
      createRadialGradient: () => ({ addColorStop: noop }),
      createPattern: () => null,
      getImageData: (_x: number, _y: number, w: number, h: number) => ({ data: new Uint8ClampedArray(Math.max(0, w * h * 4)) }),
      createImageData: (w: number, h: number) => ({ data: new Uint8ClampedArray(Math.max(0, w * h * 4)), width: w, height: h }),
    };
    // Every other context member is a no-op.
    const ctx = new Proxy({} as CanvasRenderingContext2D, {
      get(_target, key) { return key in state ? state[key] : noop; },
      set(_target, key, value) { state[key] = value; return true; },
    });
    try {
      if (isWmf(bytes)) playWmf(bytes, ctx, 800, 600);
      else playEmf(bytes, ctx, 800, 600);
    } catch {
      // A resource ceiling fails the whole picture: nothing it drew paints.
    }
  }
  return texts;
}

const encoder = new TextEncoder();

/** One WMF record: function code and 16-bit parameter words or raw bytes. */
export type WmfFixtureRecord = readonly [fn: number, params: Uint8Array];

/** A standard (non-placeable) WMF from `records` (EOF appended). */
export function wmfFixture(records: readonly WmfFixtureRecord[]): Uint8Array {
  const body = records.map(([fn, params]) => {
    const padded = params.length & 1 ? [...params, 0] : [...params];
    const words = 3 + padded.length / 2;
    return [words & 0xff, (words >> 8) & 0xff, (words >> 16) & 0xff, words >>> 24, fn & 0xff, fn >> 8, ...padded];
  }).flat();
  const eof = [3, 0, 0, 0, 0, 0];
  const size = (18 + body.length + eof.length) / 2;
  const header = [1, 0, 9, 0, 0, 3, size & 0xff, (size >> 8) & 0xff, (size >> 16) & 0xff, size >>> 24, 8, 0, 0, 0, 0, 0, 0, 0];
  return new Uint8Array([...header, ...body, ...eof]);
}

const u16 = (...values: number[]) => values.flatMap((value) => [value & 0xff, (value >> 8) & 0xff]);

/** WMF fixture records (MS-WMF 2.3): window mapping, a font, and text. */
export const wmfRecords = {
  setWindowExt: (x: number, y: number): WmfFixtureRecord => [0x020c, new Uint8Array(u16(y, x))],
  setMapMode: (mode: number): WmfFixtureRecord => [0x0103, new Uint8Array(u16(mode))],
  /** META_CREATEFONTINDIRECT: Height, Weight, Italic, CharSet, PitchAndFamily, FaceName. */
  createFont: (face: Uint8Array | string, options: { height?: number; weight?: number; charset?: number; italic?: boolean; pitchAndFamily?: number } = {}): WmfFixtureRecord => {
    const name = typeof face === 'string' ? encoder.encode(face) : face;
    return [0x02fb, new Uint8Array([
      ...u16((options.height ?? -20) & 0xffff, 0, 0, 0, options.weight ?? 400),
      options.italic ? 1 : 0, 0, 0, options.charset ?? 0, 0, 0, 0, options.pitchAndFamily ?? 0,
      ...name, 0,
    ])];
  },
  selectObject: (index: number): WmfFixtureRecord => [0x012d, new Uint8Array(u16(index))],
  deleteObject: (index: number): WmfFixtureRecord => [0x01f0, new Uint8Array(u16(index))],
  saveDc: (): WmfFixtureRecord => [0x001e, new Uint8Array()],
  restoreDc: (level: number): WmfFixtureRecord => [0x0127, new Uint8Array(u16(level & 0xffff))],
  createPen: (): WmfFixtureRecord => [0x02fa, new Uint8Array(u16(0, 1, 0, 0, 0))],
  /** META_TEXTOUT of `string` bytes at (10, 20). */
  textOut: (string: Uint8Array): WmfFixtureRecord => [0x0521, new Uint8Array([
    ...u16(string.length), ...string, ...(string.length & 1 ? [0] : []), ...u16(20, 10),
  ])],
  /** META_EXTTEXTOUT of `string` bytes at (10, 20) with `options`. */
  extTextOut: (string: Uint8Array, options = 0): WmfFixtureRecord => [0x0a32, new Uint8Array([
    ...u16(20, 10, string.length, options),
    ...(options & 0x6 ? u16(0, 0, 100, 100) : []),
    ...string, ...(string.length & 1 ? [0] : []),
  ])],
};

/** An EMF with one EXTCREATEFONTINDIRECTW font and EXTTEXTOUTW strings. */
export function emfTextFixture(face: string, strings: readonly string[], options: { charset?: number; textOptions?: number } = {}): Uint8Array {
  const records: number[][] = [];
  const u32 = (...values: number[]) => values.flatMap((value) => [value & 0xff, (value >> 8) & 0xff, (value >> 16) & 0xff, (value >>> 24) & 0xff]);
  const record = (type: number, data: number[]) => {
    const padded = data.length % 4 ? [...data, ...new Array(4 - (data.length % 4)).fill(0)] : data;
    records.push([...u32(type, 8 + padded.length), ...padded]);
  };
  // EMR_EXTCREATEFONTINDIRECTW: ihFont, then LOGFONTW (92 bytes).
  const faceUnits = [...face].flatMap((ch) => u16(ch.charCodeAt(0)));
  const logfont = [...u32(-20 >>> 0, 0, 0, 0, 400), 0, 0, 0, options.charset ?? 0, 0, 0, 0, 0,
    ...faceUnits, ...new Array(64 - faceUnits.length).fill(0)];
  record(82, [...u32(1), ...logfont]);
  record(37, u32(1)); // EMR_SELECTOBJECT
  for (const string of strings) {
    const textOptions = options.textOptions ?? 0;
    // rclBounds, iGraphicsMode, exScale, eyScale, then EMRTEXT: Reference,
    // Chars, offString, Options, Rectangle, offDx; the string follows.
    const header = 76;
    const units = [...string].flatMap((ch) => {
      const cp = ch.codePointAt(0)!;
      return cp > 0xffff ? u16(0xd800 + ((cp - 0x10000) >> 10), 0xdc00 + ((cp - 0x10000) & 0x3ff)) : u16(cp);
    });
    record(84, [...u32(0, 0, 100, 100, 1, 0, 0, 10, 20, units.length / 2, header, textOptions, 0, 0, 100, 100, 0), ...units]);
  }
  const body = records.flat();
  // EMR_HEADER (88 bytes): bounds, frame, signature, version, size, records…
  const header = [
    ...u32(1, 88, 0, 0, 100, 100, 0, 0, 2646, 2646, 0x464d4520, 0x10000, 88 + body.length + 20, records.length + 2, 0, 0),
    ...u32(0, 0, 1024, 768, 270, 203),
  ];
  const eof = u32(14, 20, 0, 0, 20);
  return new Uint8Array([...header, ...body, ...eof]);
}
