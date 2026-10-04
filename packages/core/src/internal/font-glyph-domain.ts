import type { FontTable } from './font-support-registry.js';
/** Shared maxp domain validation for cmap and the optional shaping audit. The
 * directory has already bounded every table; OpenType maxp versions 0.5/1.0
 * declare numGlyphs at +4. Missing/invalid domain never proves absence. */
export function readFontGlyphCount(view: DataView, tables: ReadonlyMap<number, FontTable>): number | undefined {
  const table = tables.get(0x6d617870);
  if (!table || table.length < 6) return undefined;
  const version = view.getUint32(table.offset);
  if (version !== 0x00005000 && version !== 0x00010000) return undefined;
  return view.getUint16(table.offset + 4) || undefined;
}
