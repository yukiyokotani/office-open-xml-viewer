import { describe, expect, it } from 'vitest';
import { parseOpenTypeLineMetrics } from './open-type-metrics.js';

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
      hasEastAsianCmap: false,
    });
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
