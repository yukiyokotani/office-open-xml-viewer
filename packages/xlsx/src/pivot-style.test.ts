import { describe, expect, it } from 'vitest';
import { buildPivotStyleMap } from './pivot-style.js';
import type { Dxf, PivotAxisItem, PivotTableMetadata, Worksheet } from './types.js';

const edge = (color: string) => ({ style: 'thin', color });
const dxf = (over: Partial<Dxf>): Dxf => ({ font: null, fill: null, border: null, ...over });
const font = (color: string, bold = false) => ({
  bold, italic: false, underline: false, strike: false, size: 11, color, name: null,
});

function pivot(rowItems: PivotAxisItem[]): PivotTableMetadata {
  return {
    name: 'P',
    cacheId: 1,
    // B2:C8: one header row, six body rows, one label column.
    location: { top: 2, left: 2, bottom: 8, right: 3, firstHeaderRow: 1, firstDataRow: 1, firstDataCol: 1 },
    rowFields: [0, 1],
    columnFields: [],
    pageFields: [],
    dataFields: [],
    status: { state: 'complete' },
    rowItems,
    columnItems: [{ kind: 'data', depth: 0 }],
    style: {
      name: 'S',
      showRowHeaders: true,
      showColumnHeaders: true,
      showRowStripes: true,
      showColumnStripes: false,
      showLastColumn: true,
      elements: [
        { kind: 'wholeTable', size: 1, dxf: dxf({ font: font('#595959') }) },
        {
          kind: 'firstRowStripe',
          size: 1,
          dxf: dxf({
            font: font('#5C7D21', true),
            border: { top: edge('#5C7D21'), bottom: edge('#5C7D21'), left: edge('#5C7D21'), right: edge('#5C7D21') },
          }),
        },
        {
          kind: 'firstRowSubheading',
          size: 1,
          dxf: dxf({
            font: font('#595959'),
            border: { top: null, bottom: null, left: null, right: null },
          }),
        },
        { kind: 'totalRow', size: 1, dxf: dxf({ fill: { patternType: 'solid', fgColor: '#EEEEEE', bgColor: null } }) },
      ],
    },
  } as PivotTableMetadata;
}

describe('PivotTable style regions (ECMA-376 §18.8.41, §18.18.77)', () => {
  it('layers whole table, row stripes, subheadings and the grand total row', () => {
    const worksheet = {
      pivotTables: [pivot([
        { kind: 'data', depth: 0 }, // row 3: subheading (odd stripe row)
        { kind: 'data', depth: 1 }, // row 4: leaf
        { kind: 'data', depth: 1 }, // row 5: leaf, odd stripe
        { kind: 'blank', depth: 0 }, // row 6
        { kind: 'data', depth: 0 }, // row 7: subheading, odd stripe
        { kind: 'grand', depth: 0 }, // row 8
      ])],
    } as unknown as Worksheet;
    const map = buildPivotStyleMap(worksheet);
    // Header row 2 carries only the whole-table colour.
    expect(map.get('2:2')).toEqual({ fontColor: '#595959' });
    // An odd leaf row is boxed across the pivot width.
    expect(map.get('5:2')).toMatchObject({ fontColor: '#5C7D21', bold: true, left: edge('#5C7D21') });
    expect(map.get('5:2')?.right).toBeUndefined();
    expect(map.get('5:3')).toMatchObject({ right: edge('#5C7D21'), top: edge('#5C7D21') });
    // An even row has no stripe.
    expect(map.get('4:2')).toEqual({ fontColor: '#595959' });
    // A subheading on an odd row restores the colour; its null edges leave
    // the stripe box, and the stripe's bold stays (a dxf cannot unset it).
    expect(map.get('3:2')).toMatchObject({ fontColor: '#595959', bold: true });
    // The grand total row takes the total-row fill last.
    expect(map.get('8:3')?.fill?.fgColor).toBe('#EEEEEE');
  });

  it('lets an explicit none edge of a later element clear an earlier one', () => {
    const table = pivot([{ kind: 'data', depth: 0 }, { kind: 'data', depth: 1 }]);
    const none = { style: 'none', color: null };
    table.style!.elements[2] = {
      kind: 'firstRowSubheading',
      size: 1,
      dxf: dxf({ font: font('#595959'), border: { top: none, bottom: none, left: none, right: none } }),
    };
    const map = buildPivotStyleMap({ pivotTables: [table] } as unknown as Worksheet);
    // Row 3 is an odd stripe row and a subheading: the box is cleared.
    expect(map.get('3:2')).toMatchObject({ top: none, left: none, bottom: none });
  });

  it('draws nothing without a style', () => {
    const plain = pivot([{ kind: 'data', depth: 0 }]);
    delete (plain as { style?: unknown }).style;
    expect(buildPivotStyleMap({ pivotTables: [plain] } as unknown as Worksheet).size).toBe(0);
  });
});
