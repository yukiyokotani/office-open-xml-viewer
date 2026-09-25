import { describe, expect, it } from 'vitest';
import { parseOpenTypeLineMetrics, parseOpenTypeResourceMetrics } from './open-type-metrics.js';

function syntheticSfnt(baseOffset = 0, eastAsianCmap = false): Uint8Array {
  const tableCount = eastAsianCmap ? 4 : 3;
  const directorySize = 12 + tableCount * 16;
  const headOffset = baseOffset + directorySize;
  const hheaOffset = headOffset + 54;
  const os2Offset = hheaOffset + 36;
  const cmapOffset = os2Offset + 78;
  const bytes = new Uint8Array(cmapOffset + (eastAsianCmap ? 40 : 0));
  const view = new DataView(bytes.buffer);
  view.setUint32(baseOffset, 0x00010000);
  view.setUint16(baseOffset + 4, tableCount);

  const record = (index: number, tag: string, offset: number, length: number) => {
    const at = baseOffset + 12 + index * 16;
    for (let i = 0; i < 4; i++) bytes[at + i] = tag.charCodeAt(i);
    view.setUint32(at + 8, offset);
    view.setUint32(at + 12, length);
  };
  record(0, 'head', headOffset, 54);
  record(1, 'hhea', hheaOffset, 36);
  record(2, 'OS/2', os2Offset, 78);
  if (eastAsianCmap) record(3, 'cmap', cmapOffset, 40);

  view.setUint16(headOffset + 18, 2048);
  view.setInt16(hheaOffset + 4, 1802);
  view.setInt16(hheaOffset + 6, -455);
  view.setInt16(hheaOffset + 8, 1024);
  view.setUint16(os2Offset, 4);
  view.setUint16(os2Offset + 62, 0x0080);
  view.setInt16(os2Offset + 68, 1600);
  view.setInt16(os2Offset + 70, -400);
  view.setInt16(os2Offset + 72, 200);
  view.setUint16(os2Offset + 74, 1900);
  view.setUint16(os2Offset + 76, 736);
  if (eastAsianCmap) {
    view.setUint16(cmapOffset + 2, 1);
    view.setUint16(cmapOffset + 4, 3);
    view.setUint16(cmapOffset + 6, 10);
    view.setUint32(cmapOffset + 8, 12);
    view.setUint16(cmapOffset + 12, 12);
    view.setUint32(cmapOffset + 16, 28);
    view.setUint32(cmapOffset + 24, 1);
    view.setUint32(cmapOffset + 28, 0x56fd);
    view.setUint32(cmapOffset + 32, 0x56fd);
    view.setUint32(cmapOffset + 36, 1);
  }
  return bytes;
}

function syntheticSfntWithCmapFormat(
  format: 4 | 12 | 13,
  codePoint = 0x56fd,
): Uint8Array {
  const source = syntheticSfnt(0, true);
  const bytes = new Uint8Array(source.length + 4);
  bytes.set(source);
  const view = new DataView(bytes.buffer);
  const cmapOffset = 12 + 4 * 16 + 54 + 36 + 78;
  view.setUint32(12 + 3 * 16 + 12, 44);
  const subtable = cmapOffset + 12;
  bytes.fill(0, subtable);
  view.setUint16(subtable, format);
  if (format === 4) {
    view.setUint16(subtable + 2, 32);
    view.setUint16(subtable + 6, 4);
    view.setUint16(subtable + 14, codePoint);
    view.setUint16(subtable + 16, 0xffff);
    view.setUint16(subtable + 20, codePoint);
    view.setUint16(subtable + 22, 0xffff);
    view.setInt16(subtable + 24, 1);
    view.setInt16(subtable + 26, 1);
  } else {
    view.setUint32(subtable + 4, 28);
    view.setUint32(subtable + 12, 1);
    view.setUint32(subtable + 16, codePoint);
    view.setUint32(subtable + 20, codePoint);
    view.setUint32(subtable + 24, 1);
  }
  return bytes;
}

describe('parseOpenTypeLineMetrics', () => {
  it('reads a positive OS/2 average character width without inferring glyph advances', () => {
    const bytes = syntheticSfnt();
    const os2Offset = 12 + 3 * 16 + 54 + 36;
    new DataView(bytes.buffer).setInt16(os2Offset + 2, 602);
    expect(parseOpenTypeResourceMetrics(bytes)?.averageCharWidthRatio).toBe(602 / 2048);
    new DataView(bytes.buffer).setInt16(os2Offset + 2, -1);
    expect(parseOpenTypeResourceMetrics(bytes)?.averageCharWidthRatio).toBeUndefined();
  });

  it('reads line metrics from sfnt tables without consulting a family name', () => {
    expect(parseOpenTypeLineMetrics(syntheticSfnt())).toEqual({
      unitsPerEm: 2048,
      hheaAscent: 1802,
      hheaDescent: -455,
      hheaLineGap: 1024,
      typoAscent: 1600,
      typoDescent: -400,
      typoLineGap: 200,
      winAscent: 1900,
      winDescent: 736,
      useTypoMetrics: true,
      farEastCodePage: null,
      hasEastAsianCmap: false,
    });
  });

  it('reads the Word Far East code-page class independently of cmap coverage', () => {
    const source = syntheticSfnt();
    const os2Offset = 12 + 3 * 16 + 54 + 36;
    const withCodePages = new Uint8Array(source.length + 8);
    withCodePages.set(source);
    const view = new DataView(withCodePages.buffer);
    view.setUint32(12 + 2 * 16 + 12, 86);
    for (const bit of [17, 18, 19, 20]) {
      view.setUint32(os2Offset + 78, 1 << bit);
      expect(parseOpenTypeLineMetrics(withCodePages)?.farEastCodePage).toBe(true);
    }
    view.setUint32(os2Offset + 78, 1);
    expect(parseOpenTypeLineMetrics(withCodePages)?.farEastCodePage).toBe(false);
    view.setUint16(os2Offset, 0);
    expect(parseOpenTypeLineMetrics(withCodePages)?.farEastCodePage).toBeNull();
    view.setUint16(os2Offset, 1);
    view.setUint32(12 + 2 * 16 + 12, 82);
    expect(parseOpenTypeLineMetrics(withCodePages)?.farEastCodePage).toBeNull();
  });

  it('detects East Asian glyph coverage from a Unicode cmap instead of a family name', () => {
    expect(parseOpenTypeLineMetrics(syntheticSfnt(0, true))?.hasEastAsianCmap).toBe(true);
    expect(parseOpenTypeLineMetrics(syntheticSfntWithCmapFormat(4))?.hasEastAsianCmap).toBe(true);
    expect(parseOpenTypeLineMetrics(syntheticSfntWithCmapFormat(13))?.hasEastAsianCmap).toBe(true);
  });

  it('does not treat compatibility-width Latin glyphs as East Asian coverage', () => {
    expect(parseOpenTypeLineMetrics(
      syntheticSfntWithCmapFormat(12, 0xff21),
    )?.hasEastAsianCmap).toBe(false);
  });

  it('rejects malformed cmap groups without rejecting otherwise valid line metrics', () => {
    const bytes = syntheticSfntWithCmapFormat(12);
    const cmapOffset = 12 + 4 * 16 + 54 + 36 + 78;
    new DataView(bytes.buffer).setUint32(cmapOffset + 12 + 12, 0xffffffff);
    expect(parseOpenTypeLineMetrics(bytes)?.hasEastAsianCmap).toBe(false);

    const overlapping = syntheticSfntWithCmapFormat(4);
    const format4 = 12 + 4 * 16 + 54 + 36 + 78 + 12;
    const overlappingView = new DataView(overlapping.buffer);
    overlappingView.setUint16(format4 + 16, 0x56fd);
    overlappingView.setUint16(format4 + 22, 0x56fd);
    expect(parseOpenTypeLineMetrics(overlapping)?.hasEastAsianCmap).toBe(false);

    const reversedGroup = syntheticSfntWithCmapFormat(12);
    const format12 = 12 + 4 * 16 + 54 + 36 + 78 + 12;
    const reversedView = new DataView(reversedGroup.buffer);
    reversedView.setUint32(format12 + 16, 0x56fe);
    reversedView.setUint32(format12 + 20, 0x56fd);
    expect(parseOpenTypeLineMetrics(reversedGroup)?.hasEastAsianCmap).toBe(false);
  });

  it('reads a face from a TrueType Collection', () => {
    const sfnt = syntheticSfnt(16);
    const bytes = sfnt.slice();
    const view = new DataView(bytes.buffer);
    bytes.set([0x74, 0x74, 0x63, 0x66], 0);
    view.setUint32(4, 0x00010000);
    view.setUint32(8, 1);
    view.setUint32(12, 16);
    expect(parseOpenTypeLineMetrics(bytes)).toBeNull();
    expect(parseOpenTypeLineMetrics(bytes, 0)?.hheaAscent).toBe(1802);
  });

  it('rejects truncated or structurally invalid fonts', () => {
    expect(parseOpenTypeLineMetrics(new Uint8Array())).toBeNull();
    expect(parseOpenTypeLineMetrics(new Uint8Array([0, 1, 0, 0]))).toBeNull();
    const invalid = syntheticSfnt();
    new DataView(invalid.buffer).setUint16(12 + 8, 0xffff);
    expect(parseOpenTypeLineMetrics(invalid)).toBeNull();
    const duplicate = syntheticSfnt();
    new DataView(duplicate.buffer).setUint32(12 + 2 * 16, 0x68656164);
    expect(parseOpenTypeLineMetrics(duplicate)).toBeNull();
    const invalidUnitsPerEm = syntheticSfnt();
    new DataView(invalidUnitsPerEm.buffer).setUint16(12 + 3 * 16 + 18, 1);
    expect(parseOpenTypeLineMetrics(invalidUnitsPerEm)).toBeNull();
  });
});

function syntheticSfntWithRepeatedCmapRecords(recordCount: number, distinctTables = 1): Uint8Array {
  const tableCount = 4;
  const directorySize = 12 + tableCount * 16;
  const headOffset = directorySize;
  const hheaOffset = headOffset + 54;
  const os2Offset = hheaOffset + 36;
  const cmapOffset = os2Offset + 78;
  const subtableOffset = 4 + recordCount * 8;
  // Use the broadest useful BMP format-4 subtable rather than a one-glyph
  // format-12 record. At the accepted record-count boundary this makes an
  // accidental parse-per-alias implementation do substantial duplicate work
  // and/or duplicate a 65,535-code-point result, while the intended path scans
  // the aliases and parses the shared subtable once.
  const cmapLength = subtableOffset + 32 * distinctTables;
  const bytes = new Uint8Array(cmapOffset + cmapLength);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint16(4, tableCount);
  const record = (index: number, tag: string, offset: number, length: number) => {
    const at = 12 + index * 16;
    for (let i = 0; i < 4; i++) bytes[at + i] = tag.charCodeAt(i);
    view.setUint32(at + 8, offset);
    view.setUint32(at + 12, length);
  };
  record(0, 'head', headOffset, 54);
  record(1, 'hhea', hheaOffset, 36);
  record(2, 'OS/2', os2Offset, 78);
  record(3, 'cmap', cmapOffset, cmapLength);
  view.setUint16(headOffset + 18, 2048);
  view.setInt16(hheaOffset + 4, 1802);
  view.setInt16(hheaOffset + 6, -455);
  view.setInt16(hheaOffset + 8, 0);
  view.setUint16(os2Offset, 4);
  view.setUint16(os2Offset + 62, 0x0080);
  view.setInt16(os2Offset + 68, 1600);
  view.setInt16(os2Offset + 70, -400);
  view.setInt16(os2Offset + 72, 200);
  view.setUint16(os2Offset + 74, 1900);
  view.setUint16(os2Offset + 76, 736);
  view.setUint16(cmapOffset + 2, recordCount);
  for (let index = 0; index < recordCount; index++) {
    const at = cmapOffset + 4 + index * 8;
    view.setUint16(at, 3);
    view.setUint16(at + 2, 10);
    view.setUint32(at + 4, subtableOffset + (index % distinctTables) * 32);
  }
  for (let tableIndex = 0; tableIndex < distinctTables; tableIndex++) {
    const subtable = cmapOffset + subtableOffset + tableIndex * 32;
    view.setUint16(subtable, 4);
    view.setUint16(subtable + 2, 32);
    view.setUint16(subtable + 6, 4);
    view.setUint16(subtable + 14, 0xfffe);
    view.setUint16(subtable + 16, 0xffff);
    view.setUint16(subtable + 20, 0x0000);
    view.setUint16(subtable + 22, 0xffff);
    view.setInt16(subtable + 24, 1);
    view.setInt16(subtable + 26, 1);
  }
  return bytes;
}


describe('opt-in OpenType resource coverage', () => {
  it('reads nonzero glyph coverage from the supported cmap encodings', () => {
    for (const bytes of [syntheticSfnt(0, true), syntheticSfntWithCmapFormat(4),
      syntheticSfntWithCmapFormat(13)]) {
      expect(parseOpenTypeResourceMetrics(bytes)?.unicodeRanges).toContainEqual([0x56fd, 0x56fd]);
    }
  });

  it('rejects variable faces only for caller-resource metric authority', () => {
    const staticFace = syntheticSfnt(0, true);
    const staticView = new DataView(staticFace.buffer);
    const count = staticView.getUint16(4);
    const directoryEnd = 12 + count * 16;
    const variableFace = new Uint8Array(staticFace.length + 16 + 36);
    variableFace.set(staticFace.subarray(0, directoryEnd));
    variableFace.set(staticFace.subarray(directoryEnd), directoryEnd + 16);
    const view = new DataView(variableFace.buffer);
    view.setUint16(4, count + 1);
    for (let index = 0; index < count; index++) {
      const offsetField = 12 + index * 16 + 8;
      view.setUint32(offsetField, staticView.getUint32(offsetField) + 16);
    }
    const fvarOffset = staticFace.length + 16;
    view.setUint32(directoryEnd, 0x66766172); // fvar
    view.setUint32(directoryEnd + 8, fvarOffset);
    view.setUint32(directoryEnd + 12, 36);
    view.setUint16(fvarOffset, 1); // fvar version 1.0
    view.setUint16(fvarOffset + 4, 16); // axes array offset
    view.setUint16(fvarOffset + 8, 1); // one weight axis
    view.setUint16(fvarOffset + 10, 20);
    view.setUint16(fvarOffset + 14, 8); // instance size, no named instances
    view.setUint32(fvarOffset + 16, 0x77676874); // wght
    view.setInt32(fvarOffset + 20, 100 * 65536);
    view.setInt32(fvarOffset + 24, 400 * 65536);
    view.setInt32(fvarOffset + 28, 900 * 65536);
    view.setUint16(fvarOffset + 34, 256); // axis name id
    expect(parseOpenTypeLineMetrics(variableFace)).toEqual(parseOpenTypeLineMetrics(staticFace));
    expect(parseOpenTypeResourceMetrics(variableFace)).toBeNull();
  });

  it('certifies only shared coverage when browser-selectable cmap records disagree', () => {
    const bytes = syntheticSfntWithRepeatedCmapRecords(2, 2);
    const view = new DataView(bytes.buffer);
    const cmapOffset = 12 + 4 * 16 + 54 + 36 + 78;
    // Unicode platform 0/3 and Windows platform 3/1 are both accepted base
    // maps. A has a glyph in the Unicode map but maps to .notdef in Windows.
    view.setUint16(cmapOffset + 4, 0);
    view.setUint16(cmapOffset + 6, 3);
    view.setUint16(cmapOffset + 14, 1);
    for (let index = 0; index < 2; index++) {
      const subtable = cmapOffset + 20 + index * 32;
      view.setUint16(subtable + 14, 0x42);
      view.setUint16(subtable + 20, 0x41);
      view.setInt16(subtable + 24, index === 0 ? -64 : -65);
    }
    expect(parseOpenTypeResourceMetrics(bytes)?.unicodeRanges).toEqual([[0x42, 0x42]]);
    // An unreadable eligible base map cannot simply be omitted from the proof.
    view.setUint16(cmapOffset + 20 + 32, 6);
    expect(parseOpenTypeResourceMetrics(bytes)?.unicodeRanges).toEqual([]);
  });

  it('caps cumulative coverage work across distinct broad subtables', () => {
    expect(parseOpenTypeResourceMetrics(syntheticSfntWithRepeatedCmapRecords(4, 4))?.unicodeRanges)
      .toEqual([[0x0000, 0xfffe]]);
    expect(parseOpenTypeResourceMetrics(syntheticSfntWithRepeatedCmapRecords(5, 5))?.unicodeRanges)
      .toEqual([]);
  });

  it('bounds cmap alias fan-out at the accepted encoding-record boundary', () => {
    expect(parseOpenTypeResourceMetrics(syntheticSfntWithRepeatedCmapRecords(4096))?.unicodeRanges)
      .toEqual([[0x0000, 0xfffe]]);
    expect(parseOpenTypeResourceMetrics(syntheticSfntWithRepeatedCmapRecords(4097))?.unicodeRanges)
      .toEqual([]);
  });
});
