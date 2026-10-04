/** Add one sfnt table without relying on a font parser under test. Fixtures
 * intentionally omit outlines; the FontFace API boundary is separately mocked. */
export function addSfntTable(source: Uint8Array, tag: string, payload: Uint8Array): Uint8Array {
  const old = new DataView(source.buffer, source.byteOffset, source.byteLength);
  const count = old.getUint16(4), end = 12 + count * 16;
  const result = new Uint8Array(source.length + 16 + payload.length);
  result.set(source.subarray(0, end)); result.set(source.subarray(end), end + 16);
  const view = new DataView(result.buffer);
  view.setUint16(4, count + 1);
  for (let i = 0; i < count; i++) view.setUint32(20 + i * 16, old.getUint32(20 + i * 16) + 16);
  for (let i = 0; i < 4; i++) result[end + i] = tag.charCodeAt(i);
  view.setUint32(end + 8, source.length + 16); view.setUint32(end + 12, payload.length);
  result.set(payload, source.length + 16);
  return result;
}
export function withGlyphDomain(source: Uint8Array, glyphCount = 65535): Uint8Array {
  const payload = new Uint8Array(6), view = new DataView(payload.buffer);
  view.setUint32(0, 0x10000); view.setUint16(4, glyphCount);
  return addSfntTable(source, 'maxp', payload);
}
