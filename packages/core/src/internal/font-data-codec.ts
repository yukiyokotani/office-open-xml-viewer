/** Generic lossless catalogue range codec: unsigned base-128 (gap,length)
 * pairs, with gaps relative to the preceding inclusive endpoint. Generated
 * strings decode only on demand; malformed encodings cannot become coverage.
 * Five bytes bound integer decoding, Unicode bounds prevent amplification. */
export function decodeFontRanges(encoded: string): readonly number[] | undefined {
  let bytes: string;
  try { bytes = atob(encoded); } catch { return undefined; }
  let at = 0, previous = -1;
  const ranges: number[] = [];
  const read = (): number | undefined => {
    let value = 0, factor = 1;
    for (let count = 0; count < 5 && at < bytes.length; count++, factor *= 128) {
      const byte = bytes.charCodeAt(at++); value += (byte & 127) * factor;
      if (value > 0x110000) return undefined;
      if (!(byte & 128)) return value;
    }
    return undefined;
  };
  while (at < bytes.length) {
    const gap = read(), length = read();
    if (gap === undefined || length === undefined) return undefined;
    const lo = previous + gap + 1, hi = lo + length;
    if (hi > 0x10ffff) return undefined;
    ranges.push(lo, hi); previous = hi;
  }
  return Object.freeze(ranges);
}

/** PackBits byte bitmap with implicit trailing zeros. Allocation is bounded by
 * one Unicode bitmap (0x110000 bits); packet lengths are format limits, not
 * font heuristics. The owner keeps these mutable bytes private. */
export function decodeFontBitmap(encoded: string, byteLength: number): Uint8Array | undefined {
  if (!Number.isSafeInteger(byteLength) || byteLength < 0 || byteLength > 0x110000 / 8) return undefined;
  let packed: string;
  try { packed = atob(encoded); } catch { return undefined; }
  const bits = new Uint8Array(byteLength);
  let out = 0;
  for (let at = 0; at < packed.length;) {
    const control = packed.charCodeAt(at++); if (control === 128) continue;
    const count = control < 128 ? control + 1 : 257 - control;
    if (out + count > bits.length) return undefined;
    if (control < 128) {
      if (at + count > packed.length) return undefined;
      for (let end = at + count; at < end;) bits[out++] = packed.charCodeAt(at++);
    } else {
      if (at >= packed.length) return undefined;
      bits.fill(packed.charCodeAt(at++), out, out + count); out += count;
    }
  }
  return bits;
}
