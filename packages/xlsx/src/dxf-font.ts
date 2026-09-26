import type { Dxf, DxfFontToggles } from './types.js';

export const DXF_FONT_TOGGLES = ['bold', 'italic', 'underline', 'strike'] as const satisfies
  readonly (keyof DxfFontToggles)[];

/**
 * The toggle a differential format sets: `true` / `false` as authored, or
 * `undefined` when its `<font>` omits the element and the formatting beneath
 * stays (ECMA-376 §18.8.14-15 dxf, §18.8.2 b CT_BooleanProperty `val` default
 * true). A model without `fontToggles` (older producer) can only say "on".
 */
export function dxfFontToggle(dxf: Dxf, key: keyof DxfFontToggles): boolean | undefined {
  if (!dxf.font) return undefined;
  return dxf.fontToggles ? dxf.fontToggles[key] : dxf.font[key] || undefined;
}
