import { afterEach, describe, expect, it, vi } from 'vitest';
import { EmfPlusPlayer, scanEmfPlus, type EmfPlusTarget } from './emf-plus.js';

// A 256-byte player ceiling makes the decoded-byte budget boundaries testable
// with bitmaps of a few pixels.
vi.mock('./pixel-budget.js', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./pixel-budget.js')>()),
  HARD_MAX_DECODED_IMAGE_BYTES: 256,
}));

afterEach(() => vi.unstubAllGlobals());

const u32 = (v: number) => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
const f32 = (v: number) => [...new Uint8Array(new Float32Array([v]).buffer)];
/** One EMF+ record: Type, Flags, Size, DataSize, data (4-byte aligned). */
const plusRecord = (type: number, flags: number, data: number[] = []) => {
  const body = [...data];
  while (body.length % 4) body.push(0);
  return [type & 255, type >>> 8, flags & 255, flags >>> 8, ...u32(12 + body.length), ...u32(data.length), ...body];
};
/** An EMF record: iType, nSize, data. */
const emfRecord = (type: number, data: number[] = []) => [...u32(type), ...u32(8 + data.length), ...data];
/** EMR_COMMENT carrying EMF+ records. */
const comment = (...records: number[][]) => {
  const payload = records.flat();
  return emfRecord(70, [...u32(4 + payload.length), ...u32(0x2b464d45), ...payload]);
};
const file = (...records: number[][]) => new Uint8Array([...emfRecord(1, new Array(80).fill(0)), ...records.flat(), ...emfRecord(14)]);
const header = plusRecord(0x4001, 1, [...u32(0xdbc01002), ...u32(1), ...u32(96), ...u32(96)]);
/** EmfPlusImageAttributes (2.2.1.5), object 0. */
const attributes = plusRecord(0x4008, 0x0800, [...u32(0xdbc01002), ...u32(0), ...u32(0), ...u32(0xffffffff), ...u32(0), ...u32(0)]);
/** A 4×4 32bpp ARGB bitmap object: 64 decoded bytes. */
const bitmap4 = (id: number) => plusRecord(0x4008, 0x0500 | id, [
  ...u32(0xdbc01002), ...u32(1), ...u32(4), ...u32(4), ...u32(16), ...u32(0x0026200a), ...u32(0), ...new Array(64).fill(255),
]);
const drawImage = (image: number) =>
  plusRecord(0x401a, image, [...u32(0), ...u32(2), ...[0, 0, 4, 4].flatMap(f32), ...[0, 0, 4, 4].flatMap(f32)]);

describe('EMF+ decoded-byte budget', () => {
  it('rejects continued-object fragments that overrun TotalObjectSize before allocating, in scan and playback', () => {
    // Declares a 24-byte ImageAttributes object but carries 324 bytes.
    const oversized = plusRecord(0x4008, 0x8800, [...u32(24), ...u32(0xdbc01002), ...new Array(320).fill(0)]);
    const bytes = file(comment(header, bitmap4(1), oversized, drawImage(1)));
    const allocated: number[] = [];
    const Original = Uint8Array;
    vi.stubGlobal('Uint8Array', new Proxy(Original, {
      construct(target, args) {
        if (typeof args[0] === 'number') allocated.push(args[0]);
        return Reflect.construct(target, args);
      },
    }));
    const scan = scanEmfPlus(bytes);
    expect(scan.play).toBe(false);
    expect(scan.failures).toContain('EMF+ continued object (fragments exceed TotalObjectSize)');

    const target: EmfPlusTarget = {
      ctx: {} as CanvasRenderingContext2D,
      W: 4, H: 4, left: 0, top: 0, boundsW: 4, boundsH: 4, drew: false, unsupported: new Set(),
    };
    const at = 88; // the comment after the 88-byte EMR_HEADER
    const dv = new DataView(bytes.buffer);
    new EmfPlusPlayer(target).playComment(dv, at, at + dv.getUint32(at + 4, true));
    expect(target.unsupported).toContain('EMF+ continued object (fragments exceed TotalObjectSize)');
    vi.unstubAllGlobals();
    expect(allocated.every((n) => n <= 256)).toBe(true);
  });

  it('accepts exactly the per-player table plus blit budget and rejects one more retained bitmap', () => {
    const withBitmaps = (count: number) =>
      file(comment(header, attributes, ...Array.from({ length: count }, (_, i) => bitmap4(i + 1)), drawImage(1)));
    expect(scanEmfPlus(withBitmaps(2)).play).toBe(true); // 2×64 retained + 2×64 blit = 256
    expect(scanEmfPlus(withBitmaps(3)).play).toBe(false);
  });
});
