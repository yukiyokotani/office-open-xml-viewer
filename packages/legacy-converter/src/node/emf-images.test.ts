// Passive EMF/WMF pictures (MS-ODRAW OfficeArtBlipEMF/WMF, MS-EMF, MS-WMF)
// through the direct PPT reader. The blip decoding matrix (raw and
// zlib, one or two UIDs, declared sizes, post-EOF payloads, malformed
// framing) is Rust-covered in officeart/metafile.rs; these cases check the
// wiring from the binary file to the model, the retained bytes and the core
// metafile players.
import { deflateSync } from 'node:zlib';
import { expect, it, vi } from 'vitest';
import { playEmf } from '../../../core/src/image/emf.js';
import { playWmf } from '../../../core/src/image/wmf.js';
import { buildPptFixture, concat, little16, little32 } from '../test-fixtures.js';
import { testPptSource } from '../test-sources.js';
import { openPptxPresentation } from './node-facade.js';

const record = (kind: number, options: number, bytes: Uint8Array) => concat(little16(options), little16(kind), little32(bytes.length), bytes);

// MS-EMF: header, explicit stock black pen, rectangle, EOF. No external resources.
const emf = new Uint8Array(144);
const emfView = new DataView(emf.buffer);
for (const [offset, value] of [[0, 1], [4, 88], [16, 100], [20, 100], [32, 2540], [36, 2540],
  [40, 0x464d4520], [44, 0x10000], [48, 144], [52, 4], [56, 1], [72, 96], [76, 96], [80, 25], [84, 25],
  [88, 37], [92, 12], [96, 0x80000007],
  [100, 43], [104, 24], [108, 10], [112, 20], [116, 70], [120, 80], [124, 14], [128, 20], [140, 20]]) {
  emfView.setUint32(offset, value, true);
}
const wmfRecord = (fn: number, words: number[] = []) => concat(little32(3 + words.length), little16(fn), ...words.map(little16));
const wmfBody = concat(
  wmfRecord(0x020c, [100, 100]),
  wmfRecord(0x02fa, [0, 1, 0, 0, 0]),
  wmfRecord(0x012d, [0]),
  wmfRecord(0x0325, [2, 10, 10, 90, 90]),
  wmfRecord(0),
);
const wmf = concat(little16(1), little16(9), little16(0x300), little32((18 + wmfBody.length) / 2), little16(1), little32(8), little16(0), wmfBody);

type Kind = 'emf' | 'wmf';
function blip(kind: Kind, source: Uint8Array, compressed: boolean, two: boolean, declared = source.length): Uint8Array {
  const image = compressed ? new Uint8Array(deflateSync(source)) : source;
  const header = new Uint8Array(34);
  const view = new DataView(header.buffer);
  view.setUint32(0, declared, true);
  view.setUint32(28, image.length, true);
  header[32] = compressed ? 0 : 0xfe;
  header[33] = 0xfe;
  const [type, instance] = kind === 'emf' ? [0xf01a, 0x3d4] : [0xf01b, 0x216];
  return record(type, (instance + (two ? 1 : 0)) << 4, concat(new Uint8Array(two ? 32 : 16), header, image));
}
function ppt(bytes: Uint8Array, delayed: boolean, kind: Kind): Uint8Array {
  const shape = record(0xf004, 15, concat(
    record(0xf00a, (75 << 4) | 2, concat(little32(42), little32(0xa00))),
    record(0xf00b, 0x13, concat(little16(0x4104), little32(1))),
    record(0xf010, 0, concat(...[0, 0, 576, 576].map(little32))),
  ));
  const bse = new Uint8Array(36);
  const view = new DataView(bse.buffer);
  bse[0] = bse[1] = kind === 'emf' ? 2 : 3;
  view.setUint32(20, bytes.length, true);
  view.setUint32(24, 1, true);
  const entry = record(0xf007, (kind === 'emf' ? 2 : 3) << 4 | 2, concat(bse, delayed ? new Uint8Array() : bytes));
  return buildPptFixture(record(1036, 15, record(0xf002, 15, shape)), new Uint8Array(), undefined,
    { entries: [entry], ...(delayed ? { pictures: bytes } : {}) });
}

interface Picture { imagePath: string; mimeType: string }
async function pptPictures(bytes: Uint8Array): Promise<{ pictures: Picture[]; extract(path: string): Uint8Array }> {
  // Slide images are admitted as their slide is pulled; read them meanwhile.
  const session = await openPptxPresentation(bytes, { modelSources: [testPptSource()] });
  const pictures: Picture[] = [];
  const images = new Map<string, Uint8Array>();
  try {
    for await (const slide of session.slides()) {
      for (const element of slide.elements as unknown as ({ type: string } & Picture)[]) {
        if (element.type !== 'picture') continue;
        pictures.push(element);
        images.set(element.imagePath, new Uint8Array(await (await session.getImage(element.imagePath, element.mimeType)).arrayBuffer()));
      }
    }
  } finally { await session.close(); }
  return { pictures, extract: path => images.get(path) as Uint8Array };
}

it.each([
  { kind: 'emf', compressed: false, two: false, delayed: false },
  { kind: 'wmf', compressed: true, two: true, delayed: true },
] as const)('retains a passive $kind through ppt (compressed=$compressed twoUIDs=$two delayed=$delayed)', async ({ kind, compressed, two, delayed }) => {
  const source = kind === 'emf' ? emf : wmf;
  const bytes = blip(kind, source, compressed, two);
  const { pictures, extract } = await pptPictures(ppt(bytes, delayed, kind));
  expect(pictures).toHaveLength(1);
  expect(pictures[0].mimeType).toBe(kind === 'emf' ? 'image/emf' : 'image/wmf');
  const extracted = extract(pictures[0].imagePath);
  expect(extracted).toEqual(source);
  const stroke = vi.fn();
  const context = new Proxy({ stroke }, { get(target, name) { return Reflect.get(target, name) ?? vi.fn(); } });
  expect((kind === 'emf' ? playEmf : playWmf)(extracted, context as unknown as CanvasRenderingContext2D, 100, 100)).toBe(true);
  expect(stroke).toHaveBeenCalled(); // The core metafile player receives visible geometry.
});

it('direct rejects a ppt picture whose WMF declares a payload after its EOF record', async () => {
  // The converter omitted the picture with a warning; the direct reader does
  // not drop authored drawing content (the Rust metafile validator returns
  // no image for any post-EOF payload).
  const tailed = concat(wmf, new Uint8Array(8).fill(0xa5));
  new DataView(tailed.buffer).setUint32(6, tailed.length / 2, true);
  const bytes = blip('wmf', tailed, true, false);
  await expect(pptPictures(ppt(bytes, false, 'wmf'))).rejects.toMatchObject({ message: 'UNSUPPORTED:PowerPoint picture BLIP is not a supported image' });
});

it('rejects malformed WMF data before EOF through ppt', async () => {
  const malformed = new Uint8Array(wmf); malformed[malformed.length - 2] = 1;
  const bytes = blip('wmf', malformed, false, false);
  await expect(pptPictures(ppt(bytes, false, 'wmf'))).rejects.toMatchObject({ message: 'UNSUPPORTED:missing WMF end record' });
});

it('rejects an EMF expansion-size bomb through ppt', async () => {
  const bytes = blip('emf', emf, true, false, 0xffffffff);
  await expect(pptPictures(ppt(bytes, true, 'emf'))).rejects.toMatchObject({ message: 'UNSUPPORTED:OfficeArt metafile byte budget exceeded' });
});
