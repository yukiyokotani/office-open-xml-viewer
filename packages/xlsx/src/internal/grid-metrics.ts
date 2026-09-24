import { PT_TO_PX } from '@silurus/ooxml-core';

/** ECMA-376 §18.3.1.13 fallback when Normal-style MDW is unavailable. */
export const MDW_FALLBACK = 8;

/** Convert an OOXML column width (maximum-digit-width units) to CSS pixels. */
export function colWidthToPx(width: number, mdw: number = MDW_FALLBACK): number {
  return Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw);
}

/** ECMA-376 §18.3.1.81: baseColWidth is a digit count, excluding padding.
 *  At 96 DPI the normative padding is 5 px (two 2 px margins and gridline).
 *  Excel for Mac 16.113.2 instead gave 5 pt of padding and a point-quantized
 *  Normal-font digit width for observed base widths 8 and 10 and Calibri
 *  9–14 pt, Arial 11 pt, and Meiryo UI 11 pt. This is a bounded Mac display
 *  compatibility rule, not an OOXML requirement or a font-family adjustment. */
export function baseColWidthToPx(baseWidth: number, mdw: number, macExcel: boolean): number {
  if (macExcel) {
    const digitPt = Math.round(mdw / PT_TO_PX);
    return Math.round((baseWidth * digitPt + 5) * PT_TO_PX);
  }
  return Math.round(baseWidth * mdw + 5);
}

/** In-memory inverse used by the viewer's resize interaction. */
export function pxToColWidth(px: number, mdw: number = MDW_FALLBACK): number {
  return px / mdw;
}

/** Convert OOXML row height points to CSS pixels at 96 DPI. */
export function rowHeightToPx(heightPt: number): number {
  return Math.round(heightPt * PT_TO_PX);
}

/** In-memory inverse used by the viewer's resize interaction. */
export function pxToRowHeight(px: number): number {
  return px / PT_TO_PX;
}
