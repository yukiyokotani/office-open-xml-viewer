import { describe, expect, it } from 'vitest';
import { metafileDemandScalars, metafilePaintTexts } from '../internal/metafile-font-demand.js';
import { emfTextFixture, metafilePaintedTexts, wmfFixture, wmfRecords as r } from '../testing/metafile-paint.js';

const bytes = (...values: number[]) => new Uint8Array(values);
const ascii = (text: string) => new TextEncoder().encode(text);

/** Painted scalars the demand misses, as U+XXXX. */
function missing(metafile: Uint8Array): string[] {
  const demand = metafileDemandScalars(metafile);
  if (demand === 'all') throw new Error('unexpected all demand');
  const demanded = new Set(demand);
  const painted = new Set(metafilePaintedTexts(metafile).flatMap((text) => [...text].map((ch) => ch.codePointAt(0)!)));
  return [...painted].filter((cp) => !demanded.has(cp)).map((cp) => `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`);
}

const scalarsOf = (metafile: Uint8Array) => {
  const demand = metafileDemandScalars(metafile);
  return demand === 'all' ? demand : new Set(demand);
};

describe('metafile text demand covers what the players measure and paint', () => {
  it('WMF text in single-byte, double-byte and symbol character sets (MS-WMF 2.3.3.5)', () => {
    const wmf = wmfFixture([
      r.setWindowExt(1000, 1000),
      r.createFont('Arial'), // 0
      r.createFont('MS Mincho', { charset: 128 }), // 1: Shift-JIS
      r.createFont('Symbol', { charset: 2 }), // 2: published repertoire
      r.createFont('Wingdings', { charset: 2 }), // 3: no repertoire: host decides
      r.createFont('Arial', { charset: 161, weight: 600 }), // 4: Greek
      r.selectObject(0), r.textOut(bytes(0x43, 0x80, 0xe9, 0x20)),
      r.selectObject(1), r.extTextOut(bytes(0x93, 0xfa, 0x96, 0x7b, 0x41)),
      r.selectObject(2), r.textOut(bytes(0x61, 0x62, 0xb6)),
      r.selectObject(3), r.extTextOut(bytes(0x4a, 0xe0), 0x0004),
      r.selectObject(4), r.textOut(bytes(0xe1, 0xe2)),
    ]);
    expect(metafilePaintedTexts(wmf).join('')).toContain('日本');
    expect(missing(wmf)).toEqual([]);
    const demanded = scalarsOf(wmf) as Set<number>;
    for (const ch of 'C€é 日本Aαβ∂αβ') expect(demanded.has(ch.codePointAt(0)!), ch).toBe(true);
    // Wingdings keeps its font codes when the host has it.
    expect(demanded.has(0xf04a)).toBe(true);
  });

  it('follows the selected font through the object table and saved DC levels', () => {
    const wmf = wmfFixture([
      r.setWindowExt(1000, 1000),
      r.createPen(), // 0
      r.createFont('Arial'), // 1
      r.createFont('MS Gothic', { charset: 128 }), // 2
      r.selectObject(1),
      r.saveDc(),
      r.selectObject(2), r.textOut(bytes(0x82, 0xa0)), // あ
      r.restoreDc(-1),
      r.textOut(bytes(0x82, 0xa0)), // ANSI ‚ and a non-breaking space
      r.deleteObject(1), // the selected font: no font, nothing paints
      r.textOut(bytes(0x88, 0x9f)),
      r.deleteObject(0),
      r.createFont('MS Gothic', { charset: 128 }), // takes slot 0
      r.selectObject(0), r.textOut(bytes(0x88, 0x9f)), // 亜
    ]);
    expect(metafilePaintedTexts(wmf)).toEqual(expect.arrayContaining(['あ', '\u201a\u00a0', '亜']));
    expect(missing(wmf)).toEqual([]);
    const demanded = scalarsOf(wmf) as Set<number>;
    for (const ch of 'あ‚ 亜') expect(demanded.has(ch.codePointAt(0)!), ch).toBe(true);
    // The bytes painted with no font selected are not decoded at all.
    expect(demanded.has(0x02c6)).toBe(false);
  });

  it('a glyph-index string, a vertical face and an unselected text paint nothing', () => {
    const wmf = wmfFixture([
      r.setWindowExt(1000, 1000),
      r.textOut(ascii('ÿ')),
      r.createFont('@MS Mincho', { charset: 128 }), r.selectObject(0), r.textOut(bytes(0x93, 0xfa)),
      r.createFont('Arial'), r.selectObject(1), r.extTextOut(bytes(0x10, 0x11), 0x0010),
    ]);
    expect(metafilePaintedTexts(wmf).filter((text) => /[^\x20-\x7e]/u.test(text))).toEqual([]);
    expect(metafilePaintTexts(wmf)).toEqual([]);
  });

  it('EMF EXTTEXTOUTW strings, with the symbol repertoire conversion (MS-EMF 2.3.5.8)', () => {
    const plain = emfTextFixture('Arial', ['Größe 漢字', 'x→y']);
    expect(metafilePaintedTexts(plain)).toEqual(expect.arrayContaining(['Größe 漢字', 'x→y']));
    expect(missing(plain)).toEqual([]);
    const symbol = emfTextFixture('Symbol', [''], { charset: 2 });
    expect(metafilePaintedTexts(symbol)).toContain('α∂');
    expect(missing(symbol)).toEqual([]);
  });

  it('bytes no player plays demand nothing beyond the measurement sentinels', () => {
    expect(metafilePaintTexts(ascii('\x1f\x8b gzip wmz'))).toEqual([]);
    expect(metafilePaintTexts(new Uint8Array(0))).toEqual([]);
  });
});
