import { describe, expect, it, vi } from 'vitest';
import {
  admitMaximumDigitWidth,
  configureHostLayout,
  decodeHostLayoutRequest,
  respondToHostLayoutRequest,
  XLSX_HOST_LAYOUT_REQUEST,
  XLSX_HOST_LAYOUT_RESULT,
} from './host-layout.js';

const encode = (value: unknown) => new TextEncoder().encode(JSON.stringify(value));
const font = { family: 'Calibri', sizePt: 11, bold: false, italic: false };

describe('decodeHostLayoutRequest', () => {
  it('admits null and an exact font tuple', () => {
    expect(decodeHostLayoutRequest(encode(null))).toBeNull();
    expect(decodeHostLayoutRequest(encode(font))).toEqual(font);
  });

  it('enforces the 4 KiB request budget at its boundary', () => {
    // The whole JSON document is exactly 4096 / 4097 bytes; the family itself
    // is short enough to pass the name bound so only the budget decides.
    const padded = (total: number) => {
      const base = new TextEncoder().encode(JSON.stringify(font)).byteLength;
      const bytes = new Uint8Array(total);
      bytes.set(new TextEncoder().encode(JSON.stringify(font)));
      bytes.fill(0x20, base); // trailing JSON whitespace
      return bytes;
    };
    expect(decodeHostLayoutRequest(padded(4096))).toEqual(font);
    expect(() => decodeHostLayoutRequest(padded(4097))).toThrow(RangeError);
  });

  it('fails closed on malformed or widened requests', () => {
    expect(() => decodeHostLayoutRequest(new Uint8Array([0xff]))).toThrow(TypeError);
    expect(() => decodeHostLayoutRequest(encode({ ...font, extra: 1 }))).toThrow(TypeError);
    expect(() => decodeHostLayoutRequest(encode({ ...font, sizePt: 0 }))).toThrow(TypeError);
    expect(() => decodeHostLayoutRequest(encode({ ...font, family: '' }))).toThrow(TypeError);
  });
});

describe('host layout decision', () => {
  it('admits only integer widths from 1 to 4096 pixels', () => {
    expect(admitMaximumDigitWidth(7)).toBe(7);
    expect(admitMaximumDigitWidth(4096)).toBe(4096);
    for (const value of [0, 4097, 7.5, Number.NaN, Number.POSITIVE_INFINITY, '7']) {
      expect(admitMaximumDigitWidth(value)).toBeUndefined();
    }
  });

  it('configures exactly once, with undefined for a null request or an unusable width', async () => {
    const configure = vi.fn();
    await expect(configureHostLayout(
      { host_layout_request: () => encode(null), configure_host_layout: configure },
      () => 9,
    )).resolves.toBeUndefined();
    await expect(configureHostLayout(
      { host_layout_request: () => encode(font), configure_host_layout: configure },
      () => 4097,
    )).resolves.toBeUndefined();
    await expect(configureHostLayout(
      { host_layout_request: () => encode(font), configure_host_layout: configure },
      () => 8,
    )).resolves.toBe(8);
    expect(configure.mock.calls).toEqual([[undefined], [undefined], [8]]);
    await expect(configureHostLayout({}, () => 8)).resolves.toBeUndefined();
    await expect(configureHostLayout({ host_layout_request: () => encode(font) }, () => 8))
      .rejects.toThrow(TypeError);
  });

  it('answers a page request with the measured width and ignores other messages', () => {
    const post = vi.fn();
    expect(respondToHostLayoutRequest(post, { type: 'other' }, () => 8)).toBe(false);
    expect(respondToHostLayoutRequest(post, { type: XLSX_HOST_LAYOUT_REQUEST, requestId: 3, font }, () => 8))
      .toBe(true);
    expect(respondToHostLayoutRequest(post, { type: XLSX_HOST_LAYOUT_REQUEST, requestId: 4, font: {} }, () => 8))
      .toBe(true);
    expect(post.mock.calls).toEqual([
      [{ type: XLSX_HOST_LAYOUT_RESULT, requestId: 3, maximumDigitWidth: 8 }],
      [{ type: XLSX_HOST_LAYOUT_RESULT, requestId: 4 }],
    ]);
  });
});
