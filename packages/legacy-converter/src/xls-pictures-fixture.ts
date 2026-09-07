/** Authored passive BIFF8 fixture; no private or Office-generated data. */
import { concat, little16, little32 } from './test-fixtures.js';
import { buildCfbWithStreams } from '@silurus/ooxml-core/testing';

export const picturePng = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLttAAAAABJRU5ErkJggg==', 'base64'));
const wmfRecord = (fn: number, words: number[] = []) => concat(little32(3 + words.length), little16(fn), ...words.map(little16));
const wmfBody = concat(wmfRecord(0x020c, [100, 100]), wmfRecord(0x0325, [2, 10, 10, 90, 90]), wmfRecord(0));
export const pictureWmf = concat(little16(1), little16(9), little16(0x100), little32((18 + wmfBody.length) / 2), little16(0), little32(8), little16(7), wmfBody);
const biff = (kind: number, data: Uint8Array = new Uint8Array()) => concat(little16(kind), little16(data.length), data);
const art = (kind: number, options: number, data: Uint8Array) => concat(little16(options), little16(kind), little32(data.length), data);
const bof = (kind: number) => biff(0x809, concat(little16(0x600), little16(kind)));

export function buildXlsPicturesFixture(options: {
  hiddenRoot?: boolean; nested?: boolean; malformedImage?: boolean;
  unknownFont?: boolean; behavior?: 0 | 2 | 3; dx?: number;
  rotation?: number; flipH?: boolean; flipV?: boolean;
  normalFont?: { family: string; sizePoints: number; bold: boolean; italic: boolean };
  imageFormat?: 'png' | 'wmf';
  wmfTrailer?: Uint8Array; wmfMalformedBeforeEof?: boolean;
} = {}): Uint8Array {
  const font = (name: string, normal = false) => {
    const format = normal ? options.normalFont : undefined;
    const data = new Uint8Array(16);
    data.set(little16((format?.sizePoints ?? 11) * 20));
    data.set(little16(format?.italic ? 2 : 0), 2);
    data.set(little16(format?.bold ? 700 : 400), 6); data[14] = name.length;
    // MS-XLS 2.4.122 requires fontName.fHighByte=1 even for ASCII names.
    // ShortXLUnicodeString counts UTF-16 units, not UTF-8 bytes.
    data[15] = 1;
    const chars = Array.from({ length: name.length }, (_, i) => name.charCodeAt(i));
    return biff(0x31, concat(data, concat(...chars.map(little16))));
  };
  const xf = new Uint8Array(20); xf[0] = 1; xf[4] = 4;
  if (options.unknownFont) xf[4] = 0;
  const blip = options.imageFormat === 'wmf'
    ? art(0xf01b, 0x2160, concat(new Uint8Array(16), (() => { let image: Uint8Array<ArrayBufferLike> = options.malformedImage ? new Uint8Array([1]) : new Uint8Array(pictureWmf); if (options.wmfMalformedBeforeEof) { image = new Uint8Array(image); image[image.length - 2] = 1; } if (options.wmfTrailer) { image = concat(image, options.wmfTrailer); new DataView(image.buffer, image.byteOffset).setUint32(6, image.length / 2, true); } const h = new Uint8Array(34); const v = new DataView(h.buffer); v.setUint32(0, image.length, true); v.setUint32(28, image.length, true); h[32] = h[33] = 0xfe; return concat(h, image); })()))
    : art(0xf01e, 0x6e00, concat(new Uint8Array(17), options.malformedImage ? new Uint8Array([1]) : picturePng));
  const store = biff(0xeb, art(0xf000, 15, art(0xf001, 31, blip)));
  const fsp = (id: number, flags: number) => art(0xf00a, 2, concat(little32(id), little32(flags)));
  const root = art(0xf004, 15, concat(fsp(1, 5), ...(options.hiddenRoot
    ? [art(0xf00b, 0x13, concat(little16(0x3bf), little32(0x20002)))] : [])));
  const corner = [options.behavior ?? 2, 0, options.dx ?? 512, 0, 128, 2, 256, 3, 64];
  const pictureProps = options.rotation == null
    ? concat(little16(0x4104), little32(1))
    // Keep this authored fixture in ascending property-ID order.
    : concat(little16(4), little32(options.rotation >>> 0), little16(0x4104), little32(1));
  const shapeFlags = 0xa00 | (options.flipH ? 64 : 0) | (options.flipV ? 128 : 0);
  const shape = art(0xf004, 15, concat(fsp(2, shapeFlags),
    art(0xf00b, options.rotation == null ? 0x13 : 0x23, pictureProps),
    art(0xf010, 0, concat(...corner.map(little16))), art(0xf011, 0, new Uint8Array()),
  ));
  const shapes = options.nested
    ? art(0xf003, 15, concat(art(0xf004, 15, fsp(3, 1)), shape)) : shape;
  const drawing = art(0xf002, 15, art(0xf003, 15, concat(root, shapes)));
  const obj = new Uint8Array(38);
  obj.set(concat(little16(0x15), little16(18), little16(8), little16(7), little16(0x11)));
  obj.set(concat(little16(7), little16(2), little16(9), little16(8), little16(2), little16(1)), 22);
  const bound = concat(little32(0), new Uint8Array([0, 0, 1, 0, 83]));
  const globals = () => concat(bof(5), biff(0x85, bound), font('Unused'), font(options.normalFont?.family ?? 'Arial', true), biff(0xe0, xf), store, biff(10));
  bound.set(little32(globals().length));
  const number = new Uint8Array(14); new DataView(number.buffer).setFloat64(6, 42.5, true);
  const workbook = concat(globals(), bof(16), biff(0x225, new Uint8Array([0, 0, 44, 1])),
    biff(0x55, little16(8)), biff(0x99, little16(2560)), biff(0x203, number), biff(0xec, drawing), biff(0x5d, obj), biff(10));
  return new Uint8Array(buildCfbWithStreams([{ name: 'Workbook', data: workbook }]));
}
