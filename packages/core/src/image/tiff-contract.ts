/** Optional TIFF codec contract shared by DOCX, XLSX and PPTX. */
export interface TiffRenderer {
  /** Decode one complete classic-TIFF part into a browser-owned bitmap. */
  render(bytes: Uint8Array): Promise<ImageBitmap | null>;
}

/** Content-sniff classic TIFF byte order + version, independent of MIME. */
export function isTiff(bytes: Uint8Array): boolean {
  if (bytes.length < 4) return false;
  const little = bytes[0] === 0x49 && bytes[1] === 0x49;
  const big = bytes[0] === 0x4d && bytes[1] === 0x4d;
  if (!little && !big) return false;
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint16(2, little) === 42;
}

/**
 * Read first-IFD dimensions from a TIFF header prefix. This intentionally owns
 * only the two scalar fields needed by the always-on decode-bomb guard; the
 * optional codec owns full IFD and strip parsing.
 */
export function sniffTiffDimensions(bytes: Uint8Array): { width: number; height: number } | null {
  if (!isTiff(bytes) || bytes.length < 8) return null;
  const little = bytes[0] === 0x49;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const contains = (offset: number, length: number) => Number.isSafeInteger(offset)
    && offset >= 0
    && length >= 0
    && offset <= view.byteLength - length;
  const ifdOffset = view.getUint32(4, little);
  if (!contains(ifdOffset, 2)) return null;
  const count = view.getUint16(ifdOffset, little);
  if (!contains(ifdOffset + 2, count * 12 + 4)) return null;
  let width: number | undefined;
  let height: number | undefined;
  for (let index = 0; index < count; index++) {
    const offset = ifdOffset + 2 + index * 12;
    const tag = view.getUint16(offset, little);
    if (tag !== 256 && tag !== 257) continue;
    const type = view.getUint16(offset + 2, little);
    const valueCount = view.getUint32(offset + 4, little);
    if (valueCount !== 1 || (type !== 1 && type !== 3 && type !== 4)) return null;
    const value = type === 1
      ? view.getUint8(offset + 8)
      : type === 3
        ? view.getUint16(offset + 8, little)
        : view.getUint32(offset + 8, little);
    if (tag === 256) width = value;
    else height = value;
  }
  return width === undefined || height === undefined ? null : { width, height };
}
