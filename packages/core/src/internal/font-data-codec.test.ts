import { expect, it } from 'vitest';
import { decodeFontRanges, decodeFontBitmap } from './font-data-codec.js';

it('decodes gaps and lengths without confusing an empty repertoire with unknown data', () => {
  expect(decodeFontRanges('')).toEqual([]);
  // (gap0,length0), then(gap127,length128): scalars0 and128..256.
  expect(decodeFontRanges(btoa(String.fromCharCode(0, 0, 127, 128, 1)))).toEqual([0, 0, 128, 256]);
  expect(decodeFontRanges('!')).toBeUndefined();
  expect(decodeFontRanges(btoa(String.fromCharCode(128)))).toBeUndefined();
  expect(decodeFontRanges(btoa(String.fromCharCode(255, 255, 255, 255, 255, 0)))).toBeUndefined();
});

it('rejects incomplete or amplified bitmap packets and preserves trailing zeros', () => {
  expect(decodeFontBitmap(btoa(String.fromCharCode(255, 7)), 4)).toEqual(new Uint8Array([7, 7, 0, 0]));
  expect(decodeFontBitmap(btoa(String.fromCharCode(2, 7)), 4)).toBeUndefined();
  expect(decodeFontBitmap(btoa(String.fromCharCode(255, 7)), 1)).toBeUndefined();
  expect(decodeFontBitmap('', Number.MAX_SAFE_INTEGER)).toBeUndefined();
});
