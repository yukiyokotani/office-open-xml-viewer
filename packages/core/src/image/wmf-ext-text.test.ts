import { describe, expect, it, vi } from 'vitest';
import { playWmf } from './wmf.js';

const u16 = (value: number) => [value & 0xff, (value >>> 8) & 0xff];
const u32 = (value: number) => [
  value & 0xff,
  (value >>> 8) & 0xff,
  (value >>> 16) & 0xff,
  (value >>> 24) & 0xff,
];

function record(fn: number, params: number[]): Uint8Array {
  return new Uint8Array([...u32(3 + params.length / 2), ...u16(fn), ...params]);
}

function header(): Uint8Array {
  return new Uint8Array([...u16(1), ...u16(9), ...u16(0x300), ...u32(0), ...u16(8), ...u32(0), ...u16(0)]);
}

function join(...parts: Uint8Array[]): Uint8Array {
  const output = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

function font(height = -20, charset = 0): Uint8Array {
  const face = [...Buffer.from('Times New Roman\0', 'latin1')];
  return record(0x02fb, [
    ...u16(height), ...u16(0), ...u16(0), ...u16(0), ...u16(400),
    1, 0, 0, charset, 0, 0, 0, 0,
    ...face, ...new Array(32 - face.length).fill(0),
  ]);
}

function extText(text: number[], options = 0, dx: number[] = []): Uint8Array {
  return record(0x0a32, [
    ...u16(24), ...u16(12), ...u16(text.length), ...u16(options),
    ...text, ...(text.length & 1 ? [0] : []), ...dx,
  ]);
}

function base(tail: Uint8Array, fontRecord = font()): Uint8Array {
  return join(
    header(),
    record(0x020b, [...u16(0), ...u16(0)]),
    record(0x020c, [...u16(100), ...u16(100)]),
    record(0x0102, u16(1)),
    fontRecord,
    record(0x012d, u16(0)),
    record(0x012e, u16(0x18)),
    tail,
    record(0, []),
  );
}

function context() {
  const fillText = vi.fn();
  const ctx = new Proxy(
    { fillText, textAlign: 'left', textBaseline: 'top', font: '', fillStyle: '' },
    {
      get: (target, property) => Reflect.get(target, property) ?? vi.fn(),
      set: (target, property, value) => (Reflect.set(target, property, value), true),
    },
  ) as unknown as CanvasRenderingContext2D;
  return { ctx, fillText };
}

describe('WMF EXTTEXTOUT bounded ANSI subset', () => {
  it('draws Windows-1252 bytes including NUL with mapped coordinates and font fields', () => {
    const raw = [0x41, 0x80, 0, 0x42];
    const mock = context();

    expect(playWmf(base(extText(raw)), mock.ctx, 200, 100)).toBe(true);
    expect(mock.fillText).toHaveBeenCalledWith('A€\0B', 24, 24);
    expect(mock.ctx.textBaseline).toBe('alphabetic');
    expect(mock.ctx.font).toBe('italic 400 20px "Times New Roman"');
  });

  it('accepts an 80-byte record string in linear space', () => {
    const mock = context();
    expect(playWmf(base(extText(new Array(80).fill(0x41))), mock.ctx, 100, 100)).toBe(true);
    expect(mock.fillText.mock.calls[0][0]).toHaveLength(80);
  });

  it('decodes the complete Windows-1252 C1 index and isomorphic edge bytes', () => {
    const bytes = [0x7f, ...Array.from({ length: 32 }, (_, index) => index + 0x80), 0xa0, 0xff];
    const expectedCodePoints = [
      0x007f,
      0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021,
      0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008d, 0x017d, 0x008f,
      0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
      0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
      0x00a0, 0x00ff,
    ];
    const mock = context();

    playWmf(base(extText(bytes)), mock.ctx, 100, 100);
    const rendered = mock.fillText.mock.calls[0][0] as string;
    expect(Array.from(rendered, (character) => character.codePointAt(0))).toEqual(expectedCodePoints);
  });

  it.each([
    ['options', extText([65], 2), font()],
    ['dx', extText([65], 0, u16(1)), font()],
    ['symbol', extText([65]), font(-20, 2)],
    ['positive-height', extText([65]), font(20, 0)],
  ] as const)('validates but skips unsupported %s', (_name, text, fontRecord) => {
    const mock = context();
    expect(playWmf(base(text, fontRecord), mock.ctx, 100, 100)).toBe(false);
    expect(mock.fillText).not.toHaveBeenCalled();
  });

  it('does not read a short CREATEFONTINDIRECT record from the next record', () => {
    for (const length of [0, 10, 17, 18, 49]) {
      const mock = context();
      const shortFont = record(0x02fb, new Array(length + (length & 1)).fill(0));
      playWmf(join(header(), record(0x020c, [...u16(100), ...u16(100)]), record(0x0102, u16(1)), shortFont, record(0x012d, u16(0)), record(0x012e, u16(0x18)), extText([65]), record(0, [])), mock.ctx, 100, 100);
      expect(mock.fillText).not.toHaveBeenCalled();
    }
  });

  it('rejects every in-bounds short EXTTEXTOUT parameter record', () => {
    const params = [...u16(24), ...u16(12), ...u16(3), ...u16(0), 65, 66, 67, 0];
    for (let length = 0; length < params.length; length += 2) {
      const mock = context();
      playWmf(base(record(0x0a32, params.slice(0, length))), mock.ctx, 100, 100);
      expect(mock.fillText).not.toHaveBeenCalled();
    }
  });

  it('requires transparent mode and positive mapping, and latches unsafe state', () => {
    const files = [
      join(header(), record(0x0102, u16(1)), record(0x0103, u16(1)), font(), record(0x012d, u16(0)), record(0x012e, u16(0x18)), extText([65]), record(0, [])),
      join(header(), record(0x020c, [...u16(100), ...u16(100)]), font(), record(0x012d, u16(0)), record(0x012e, u16(0x18)), extText([65]), record(0, [])),
      join(header(), record(0x020c, [...u16(-100), ...u16(100)]), record(0x0102, u16(1)), font(), record(0x012d, u16(0)), record(0x012e, u16(0x18)), extText([65]), record(0, [])),
    ];
    for (const file of files) {
      const mock = context();
      playWmf(file, mock.ctx, 100, 100);
      expect(mock.fillText).not.toHaveBeenCalled();
    }
  });

  it('allows SETBKCOLOR and reserved BKMODE/TEXTALIGN words', () => {
    const mock = context();
    const file = join(header(), record(0x020c, [...u16(100), ...u16(100)]), record(0x0201, u32(0xffffff)), record(0x0102, [...u16(1), ...u16(0xbeef)]), font(), record(0x012d, u16(0)), record(0x012e, [...u16(0x18), ...u16(0xbeef)]), extText([65]), record(0, []));
    playWmf(file, mock.ctx, 100, 100);
    expect(mock.fillText).toHaveBeenCalledOnce();
  });

  it.each([[[1, 2], true], [[0x57, 0x4d, 0x46, 0x43], false]] as const)(
    'handles private comment %j with draw=%s',
    (data, draws) => {
      const comment = record(0x0626, [...u16(15), ...u16(data.length), ...data, ...(data.length & 1 ? [0] : [])]);
      const mock = context();
      playWmf(base(join(comment, extText([65]))), mock.ctx, 100, 100);
      expect(mock.fillText).toHaveBeenCalledTimes(draws ? 1 : 0);
    },
  );

  it.each([0x1c, 0x19, 0x118])('rejects invalid or stateful text alignment %#x', (align) => {
    const mock = context();
    const file = join(header(), record(0x020c, [...u16(100), ...u16(100)]), record(0x0102, u16(1)), font(), record(0x012d, u16(0)), record(0x012e, u16(align)), extText([65]), record(0, []));
    playWmf(file, mock.ctx, 100, 100);
    expect(mock.fillText).not.toHaveBeenCalled();
  });

  it('latches failed SELECTOBJECT and malformed DELETEOBJECT operations', () => {
    for (const operation of [record(0x012d, u16(99)), record(0x01f0, [])]) {
      const mock = context();
      playWmf(base(join(operation, extText([65]))), mock.ctx, 100, 100);
      expect(mock.fillText).not.toHaveBeenCalled();
    }
  });

  it('does not change the established TEXTOUT NUL-terminated path', () => {
    const textOut = record(0x0521, [...u16(3), 65, 0, 66, 0, ...u16(10), ...u16(20)]);
    const mock = context();
    playWmf(join(header(), record(0x020c, [...u16(100), ...u16(100)]), font(), record(0x012d, u16(0)), textOut, record(0, [])), mock.ctx, 100, 100);
    expect(mock.fillText).toHaveBeenCalledWith('A', 20, 10);
  });
});
