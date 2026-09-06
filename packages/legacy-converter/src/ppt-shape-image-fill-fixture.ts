/** Authored passive PPT shape-picture-fill fixture; no private data. */
import { buildPptFixture, concat, little16, little32 } from './test-fixtures.js';

export const shapeFillPng = Uint8Array.from(Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAADklEQVR4nGP4z8AAQg0AD3oDfnfpf5cAAAAASUVORK5CYII=',
  'base64',
));

const record = (kind: number, bytes: Uint8Array, options = 0) =>
  concat(little16(options), little16(kind), little32(bytes.length), bytes);

export function buildPptShapeImageFillFixture(): Uint8Array {
  const blip = record(0xf01e, concat(new Uint8Array(17), shapeFillPng), 0x6e00);
  // Keep the authored OfficeArtFOPT property array in ascending property-ID order.
  const properties = record(0xf00b, concat(
    little16(0x180), little32(3),
    little16(0x182), little32(32768),
    little16(0x4186), little32(1),
    little16(0x1bf), little32(0x00200020),
    little16(0x1c0), little32(0xff),
  ), (5 << 4) | 3);
  const shape = record(0xf004, concat(
    record(0xf00a, concat(little32(42), little32(0x200)), (3 << 4) | 2),
    record(0xf010, concat(...[0, 0, 1152, 576].map(little32))),
    properties,
    record(0xf00d, record(4008, new TextEncoder().encode('Picture fill text')), 15),
  ), 15);
  return buildPptFixture(record(1036, record(0xf002, shape, 15), 15), new Uint8Array(), undefined,
    { entries: [blip] });
}
