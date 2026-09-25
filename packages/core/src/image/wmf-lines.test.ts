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

function file(...records: Uint8Array[]): Uint8Array {
  const header = new Uint8Array([
    ...u16(1), ...u16(9), ...u16(0x300), ...u32(0), ...u16(8), ...u32(0), ...u16(0),
  ]);
  const parts = [header, ...records, record(0, [])];
  const output = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

function pen(): Uint8Array {
  return record(0x02fa, [...u16(0), ...u16(2), ...u16(0), ...u32(0x000000ff)]);
}

function context() {
  const beginPath = vi.fn();
  const moveTo = vi.fn();
  const lineTo = vi.fn();
  const stroke = vi.fn();
  const setLineDash = vi.fn();
  const save = vi.fn();
  const restore = vi.fn();
  const ctx = new Proxy(
    { beginPath, moveTo, lineTo, stroke, setLineDash, save, restore, strokeStyle: '', lineWidth: 0, lineCap: 'butt', lineJoin: 'miter' },
    {
      get: (target, property) => Reflect.get(target, property) ?? vi.fn(),
      set: (target, property, value) => (Reflect.set(target, property, value), true),
    },
  ) as unknown as CanvasRenderingContext2D;
  return { ctx, beginPath, moveTo, lineTo, stroke, setLineDash, save, restore };
}

const windowRecords = [
  record(0x020b, [...u16(20), ...u16(10)]),
  record(0x020c, [...u16(100), ...u16(200)]),
];

describe('WMF MOVETO and LINETO', () => {
  it('maps a MOVETO/LINETO chain, strokes with the selected pen, and updates current position', () => {
    const mock = context();
    const bytes = file(
      ...windowRecords,
      pen(),
      record(0x012d, u16(0)),
      record(0x0214, [...u16(30), ...u16(50)]),
      record(0x0213, [...u16(70), ...u16(90)]),
      record(0x0213, [...u16(40), ...u16(30)]),
    );

    expect(playWmf(bytes, mock.ctx, 400, 200)).toBe(true);
    expect(mock.beginPath).toHaveBeenCalledTimes(2);
    expect(mock.moveTo).toHaveBeenNthCalledWith(1, 80, 20);
    expect(mock.lineTo).toHaveBeenNthCalledWith(1, 160, 100);
    expect(mock.moveTo).toHaveBeenNthCalledWith(2, 160, 100);
    expect(mock.lineTo).toHaveBeenNthCalledWith(2, 40, 40);
    expect(mock.stroke).toHaveBeenCalledTimes(2);
    expect(mock.ctx.strokeStyle).toBe('#ff0000');
    expect(mock.ctx.lineWidth).toBe(4);
    expect(mock.ctx.lineCap).toBe('round');
    expect(mock.ctx.lineJoin).toBe('round');
    expect(mock.setLineDash).toHaveBeenCalledWith([]);
    expect(mock.save).toHaveBeenCalledTimes(2);
    expect(mock.restore).toHaveBeenCalledTimes(2);
  });

  it('updates current position even with a null pen and draws a later segment', () => {
    const mock = context();
    const nullPen = record(0x02fa, [...u16(5), ...u16(1), ...u16(0), ...u32(0)]);
    const bytes = file(
      ...windowRecords,
      nullPen,
      pen(),
      record(0x012d, u16(0)),
      record(0x0214, [...u16(30), ...u16(50)]),
      record(0x0213, [...u16(70), ...u16(90)]),
      record(0x012d, u16(1)),
      record(0x0213, [...u16(40), ...u16(30)]),
    );

    expect(playWmf(bytes, mock.ctx, 400, 200)).toBe(true);
    expect(mock.moveTo).toHaveBeenCalledOnce();
    expect(mock.moveTo).toHaveBeenCalledWith(160, 100);
    expect(mock.lineTo).toHaveBeenCalledWith(40, 40);
  });

  it('does not apply the optional boundary-frame suppression heuristic to LINETO', () => {
    const mock = context();
    const bytes = file(
      ...windowRecords,
      pen(),
      record(0x012d, u16(0)),
      record(0x0214, [...u16(20), ...u16(10)]),
      record(0x0213, [...u16(120), ...u16(10)]),
    );

    expect(playWmf(bytes, mock.ctx, 400, 200, true)).toBe(true);
    expect(mock.moveTo).toHaveBeenCalledWith(0, 0);
    expect(mock.lineTo).toHaveBeenCalledWith(0, 200);
    expect(mock.stroke).toHaveBeenCalledOnce();
  });

  it('maps a positive subpixel pen width without the legacy visibility clamp', () => {
    const mock = context();
    const oneUnitPen = record(0x02fa, [...u16(0), ...u16(1), ...u16(0), ...u32(0)]);
    const bytes = file(
      ...windowRecords,
      oneUnitPen,
      record(0x012d, u16(0)),
      record(0x0214, [...u16(30), ...u16(50)]),
      record(0x0213, [...u16(70), ...u16(90)]),
    );

    expect(playWmf(bytes, mock.ctx, 50, 200)).toBe(true);
    expect(mock.ctx.lineWidth).toBe(0.25);
  });

  it('does not guess a default or stale current position', () => {
    for (const records of [
      [...windowRecords, pen(), record(0x012d, u16(0)), record(0x0213, [...u16(70), ...u16(90)])],
      [...windowRecords, pen(), record(0x012d, u16(0)), record(0x0214, [...u16(30), ...u16(50)]), record(0x001e, []), record(0x0213, [...u16(70), ...u16(90)])],
      [...windowRecords, pen(), record(0x012d, u16(0)), record(0x001e, []), record(0x0214, [...u16(30), ...u16(50)]), record(0x0213, [...u16(70), ...u16(90)])],
    ]) {
      const mock = context();
      expect(playWmf(file(...records), mock.ctx, 400, 200)).toBe(false);
      expect(mock.stroke).not.toHaveBeenCalled();
    }
  });

  it('rejects malformed coordinate records without reading adjacent records', () => {
    for (const malformed of [record(0x0214, []), record(0x0213, [...u16(1), ...u16(2), ...u16(3)])]) {
      const mock = context();
      const bytes = file(
        ...windowRecords,
        pen(),
        record(0x012d, u16(0)),
        record(0x0214, [...u16(30), ...u16(50)]),
        malformed,
        record(0x0213, [...u16(70), ...u16(90)]),
      );
      expect(playWmf(bytes, mock.ctx, 400, 200)).toBe(false);
      expect(mock.stroke).not.toHaveBeenCalled();
    }
  });

  it('does not draw with a malformed or unsupported selected pen', () => {
    const shortPen = record(0x02fa, [...u16(0), ...u16(2), ...u16(0)]);
    const dashedPen = record(0x02fa, [...u16(1), ...u16(2), ...u16(0), ...u32(0)]);
    const negativeWidthPen = record(0x02fa, [...u16(0), ...u16(-1), ...u16(0), ...u32(0)]);
    for (const penRecord of [shortPen, dashedPen, negativeWidthPen]) {
      const mock = context();
      const bytes = file(
        ...windowRecords,
        penRecord,
        record(0x012d, u16(0)),
        record(0x0214, [...u16(30), ...u16(50)]),
        record(0x0213, [...u16(70), ...u16(90)]),
      );
      expect(playWmf(bytes, mock.ctx, 400, 200)).toBe(false);
      expect(mock.stroke).not.toHaveBeenCalled();
    }
  });

  it('latches a malformed SELECTOBJECT instead of retaining the previous pen', () => {
    const mock = context();
    const bytes = file(
      ...windowRecords,
      pen(),
      record(0x012d, u16(0)),
      record(0x012d, []),
      record(0x0214, [...u16(30), ...u16(50)]),
      record(0x0213, [...u16(70), ...u16(90)]),
    );
    expect(playWmf(bytes, mock.ctx, 400, 200)).toBe(false);
    expect(mock.stroke).not.toHaveBeenCalled();
  });

  it('invalidates the current position when text requests TA_UPDATECP', () => {
    const mock = context();
    const textOut = record(0x0521, [...u16(1), 65, 0, ...u16(40), ...u16(30)]);
    const bytes = file(
      ...windowRecords,
      pen(),
      record(0x012d, u16(0)),
      record(0x0214, [...u16(30), ...u16(50)]),
      record(0x012e, u16(1)),
      textOut,
      record(0x0213, [...u16(70), ...u16(90)]),
    );
    playWmf(bytes, mock.ctx, 400, 200);
    expect(mock.stroke).not.toHaveBeenCalled();
  });

  it('latches malformed text alignment and object creation state', () => {
    const malformedAlignment = record(0x012e, []);
    const shortFont = record(0x02fb, new Array(16).fill(0));
    for (const malformed of [malformedAlignment, shortFont]) {
      const mock = context();
      const bytes = file(
        ...windowRecords,
        pen(),
        record(0x012d, u16(0)),
        malformed,
        record(0x0214, [...u16(30), ...u16(50)]),
        record(0x0213, [...u16(70), ...u16(90)]),
      );
      playWmf(bytes, mock.ctx, 400, 200);
      expect(mock.stroke).not.toHaveBeenCalled();
    }
  });
});
