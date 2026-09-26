// PivotTable style formatting (ECMA-376 §18.8.40-41, §18.10.1.97).
//
// A PivotTable style is a list of `tableStyleElement`s, each a differential
// format for one structured region of the PivotTable (§18.18.77, whose
// region diagrams define the areas below). Elements apply in the PivotTable
// style element order of §18.8.41; a later element wins per property. Each
// element is a differential format (§18.8.14-15, applied on top of what is
// already there), so a font toggle the element's dxf omits leaves the earlier
// value, and an explicit off (`<b val="0"/>`) turns an earlier on off. A
// region's border edges apply to its outline and its `horizontal` /
// `vertical` edges to its interior rules; an edge whose style is `none`
// (an explicitly cleared edge) overrides an earlier element's edge.
//
// Regions derive from the saved layout: the location offsets
// (§18.10.1.49), the row items (one per body row) and the column items
// (one per data column). Column subheadings need the column header layout,
// which the model does not carry; they are not drawn.
//
// Style option gating follows [MS-XLS] 2.4.273.107 SXAddl_SXCView_SXDTableStyleClient, which
// states the same Excel behavior the XLSX pivotTableStyleInfo booleans
// (§18.10.1.74) name: showRowHeaders applies firstColumn and the row
// subheadings; showColumnHeaders applies headerRow and the column
// subheadings; showLastColumn applies lastColumn; the stripe flags apply
// their stripes. firstHeaderCell, blank rows, subtotals and the grand total
// row are not gated.
import { DXF_FONT_TOGGLES, dxfFontToggle } from './dxf-font.js';
import {
  assertCoordinateRangeArea,
  setCoordinateIndexValue,
  type CoordinateIndexIdentity,
} from './renderer-coordinate-index.js';
import type { BorderEdge, CellFill, Dxf, PivotAxisItem, PivotTableMetadata, Worksheet } from './types.js';

/** The merged PivotTable style format of one cell. */
export interface PivotCellFormat {
  fill?: CellFill;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  top?: BorderEdge;
  bottom?: BorderEdge;
  left?: BorderEdge;
  right?: BorderEdge;
}

/** §18.8.41 PivotTable Style Element Order (later wins). */
const ORDER = [
  'wholeTable',
  'pageFieldLabels',
  'pageFieldValues',
  'firstColumnStripe',
  'secondColumnStripe',
  'firstRowStripe',
  'secondRowStripe',
  'firstColumn',
  'headerRow',
  'firstHeaderCell',
  'firstSubtotalColumn',
  'secondSubtotalColumn',
  'thirdSubtotalColumn',
  'blankRow',
  'firstSubtotalRow',
  'secondSubtotalRow',
  'thirdSubtotalRow',
  'firstColumnSubheading',
  'secondColumnSubheading',
  'thirdColumnSubheading',
  'firstRowSubheading',
  'secondRowSubheading',
  'thirdRowSubheading',
  'lastColumn',
  'totalRow',
] as const;

/** ST_ItemType values that are subtotals (§18.18.43). */
const SUBTOTALS = new Set([
  'default', 'sum', 'countA', 'avg', 'max', 'min', 'product', 'count',
  'stdDev', 'stdDevP', 'var', 'varP',
]);

interface Rect { top: number; bottom: number; left: number; right: number }

/**
 * The subtotal / subheading element level of an axis field at `depth`:
 * [MS-XLS] 2.4.321 TableStyleElement (tseType 0x10-0x12, 0x14-0x19) gives
 * the outermost field `first`, then alternates `second` (odd positions) and
 * `third` (even positions after the first) through the deeper fields.
 */
function level(depth: number): 'first' | 'second' | 'third' {
  if (depth === 0) return 'first';
  return depth % 2 === 1 ? 'second' : 'third';
}

/** Row bands cycling `first` (size a) and `second` (size b) rows. */
function bands(from: number, to: number, a: number, b: number): { first: Rect[]; second: Rect[] } {
  const out = { first: [] as Rect[], second: [] as Rect[] };
  let at = from;
  let odd = true;
  while (at <= to) {
    const size = Math.max(1, odd ? a : b);
    const end = Math.min(to, at + size - 1);
    (odd ? out.first : out.second).push({ top: at, bottom: end, left: 0, right: 0 });
    at = end + 1;
    odd = !odd;
  }
  return out;
}

/** The regions of every element kind for one PivotTable. */
function regions(p: PivotTableMetadata, sizes: Map<string, number>): Map<string, Rect[]> {
  const style = p.style!;
  const { top, bottom, left, right, firstDataRow, firstDataCol } = p.location;
  const bodyTop = top + firstDataRow;
  const dataLeft = left + firstDataCol;
  const out = new Map<string, Rect[]>();
  const add = (kind: string, rect: Rect) => {
    if (rect.top > rect.bottom || rect.left > rect.right) return;
    const list = out.get(kind);
    if (list) list.push(rect);
    else out.set(kind, [rect]);
  };
  add('wholeTable', { top, bottom, left, right });
  if (style.showColumnStripes) {
    const cols = bands(dataLeft, right, sizes.get('firstColumnStripe') ?? 1, sizes.get('secondColumnStripe') ?? 1);
    for (const band of cols.first) add('firstColumnStripe', { top: bodyTop, bottom, left: band.top, right: band.bottom });
    for (const band of cols.second) add('secondColumnStripe', { top: bodyTop, bottom, left: band.top, right: band.bottom });
  }
  if (style.showRowStripes) {
    const rows = bands(bodyTop, bottom, sizes.get('firstRowStripe') ?? 1, sizes.get('secondRowStripe') ?? 1);
    for (const band of rows.first) add('firstRowStripe', { ...band, left, right });
    for (const band of rows.second) add('secondRowStripe', { ...band, left, right });
  }
  if (style.showRowHeaders) add('firstColumn', { top, bottom, left, right: dataLeft - 1 });
  if (style.showColumnHeaders) add('headerRow', { top, bottom: bodyTop - 1, left, right });
  add('firstHeaderCell', { top, bottom: bodyTop - 2, left, right: dataLeft - 1 });
  const leafLevel = Math.max(0, p.rowFields.length - 1);
  (p.rowItems ?? []).forEach((item: PivotAxisItem, index) => {
    const row = bodyTop + index;
    if (row > bottom) return;
    const rect = { top: row, bottom: row, left, right };
    const lvl = level(item.depth);
    if (item.kind === 'blank') add('blankRow', rect);
    else if (item.kind === 'grand') add('totalRow', rect);
    else if (SUBTOTALS.has(item.kind)) add(`${lvl}SubtotalRow`, rect);
    else if (item.kind === 'data' && item.depth < leafLevel && style.showRowHeaders) {
      add(`${lvl}RowSubheading`, rect);
    }
  });
  (p.columnItems ?? []).forEach((item: PivotAxisItem, index) => {
    const col = dataLeft + index;
    if (col > right) return;
    const rect = { top, bottom, left: col, right: col };
    const lvl = level(item.depth);
    if (item.kind === 'grand') {
      if (style.showLastColumn) add('lastColumn', rect);
    } else if (SUBTOTALS.has(item.kind)) add(`${lvl}SubtotalColumn`, rect);
  });
  return out;
}

function apply(target: PivotCellFormat, dxf: Dxf, rect: Rect, row: number, col: number): void {
  if (dxf.fill) target.fill = dxf.fill;
  const font = dxf.font;
  if (font) {
    if (font.color) target.fontColor = font.color;
    for (const key of DXF_FONT_TOGGLES) {
      const value = dxfFontToggle(dxf, key);
      if (value !== undefined) target[key] = value;
    }
  }
  const border = dxf.border;
  if (!border) return;
  const edge = (outer: BorderEdge | null | undefined, inner: BorderEdge | null | undefined, onOutline: boolean) =>
    onOutline ? outer : inner;
  const top = edge(border.top, border.horizontal, row === rect.top);
  const bottom = edge(border.bottom, border.horizontal, row === rect.bottom);
  const left = edge(border.left, border.vertical, col === rect.left);
  const right = edge(border.right, border.vertical, col === rect.right);
  if (top) target.top = top;
  if (bottom) target.bottom = bottom;
  if (left) target.left = left;
  if (right) target.right = right;
}

/** Merged PivotTable style formats by `row:col`. */
export function buildPivotStyleMap(worksheet: Worksheet): Map<string, PivotCellFormat> {
  const map = new Map<string, PivotCellFormat>();
  const identity: CoordinateIndexIdentity = {
    resource: 'worksheet-pivot-style-index',
    operation: 'expand-pivot-style-coordinates',
  };
  for (const pivot of worksheet.pivotTables ?? []) {
    const style = pivot.style;
    if (!style || style.elements.length === 0) continue;
    const elements = new Map(style.elements.map((element) => [element.kind, element]));
    const sizes = new Map(style.elements.map((element) => [element.kind, element.size]));
    assertCoordinateRangeArea(pivot.location, identity);
    const areas = regions(pivot, sizes);
    for (const kind of ORDER) {
      const element = elements.get(kind);
      const rects = areas.get(kind);
      if (!element || !rects) continue;
      for (const rect of rects) {
        for (let row = rect.top; row <= rect.bottom; row++) {
          for (let col = rect.left; col <= rect.right; col++) {
            const key = `${row}:${col}`;
            let format = map.get(key);
            if (!format) {
              format = {};
              setCoordinateIndexValue(map, key, format, identity);
            }
            apply(format, element.dxf, rect, row, col);
          }
        }
      }
    }
  }
  return map;
}
