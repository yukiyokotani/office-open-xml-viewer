/** Authored BIFF8 rich-text workbook; no private or Office-generated data. */
import { buildXlsFixture, concat, little16 } from '../test-fixtures.js';

/**
 * B2 holds the shared string "base RED normal" with MS-XLS 2.5.293 FormatRuns
 * switching to font 5 (24pt red, BIFF skips font index 4) at "RED " and back
 * to the bold-free font 0 at "normal"; the trailing run uses the 0xFFFF end
 * sentinel font, which must not be resolved.
 */
export function buildXlsRichFixture(): Uint8Array {
  const record = (kind: number, data: Uint8Array) => concat(little16(kind), little16(data.length), data);
  const fonts: Uint8Array[] = [];
  for (let i = 0; i < 5; i++) {
    const font = new Uint8Array(21);
    const view = new DataView(font.buffer);
    view.setUint16(0, i === 4 ? 480 : 220, true);
    view.setUint16(4, i === 4 ? 10 : 0x7fff, true);
    view.setUint16(6, i === 1 ? 700 : 400, true);
    font[14] = 5; font.set(new TextEncoder().encode('Arial'), 16);
    fonts.push(record(0x31, font));
  }
  const xf = new Uint8Array(20); xf[0] = 1;
  const text = new TextEncoder().encode('base RED normal');
  return buildXlsFixture({
    sharedString: concat(little16(text.length), new Uint8Array([8]), little16(3), text,
      little16(5), little16(5), little16(9), little16(0), little16(15), little16(65535)),
    styleRecords: concat(...fonts, record(0xe0, xf)),
  });
}
