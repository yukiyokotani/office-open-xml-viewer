/**
 * Authored BIFF8 workbook with same-name defined names in two scopes; no
 * private or Office-generated data.
 *
 * Two visible worksheets, `A` and `B`. The globals substream holds, in this
 * order, Lbl #1 `Rate` = 1 local to `A` (itab 1) and Lbl #2 `Rate` = 5 in
 * workbook scope (itab 0) ([MS-XLS] 2.4.150): the local is authored first, so
 * Lbl order alone would let the global shadow it. Each sheet stores A1 = 1 and
 * B1 = 5 and one CondFmt12/CF12 formula rule (2.4.57, 2.4.43; ct 2, template
 * 1) over A1:B1 whose formula is `A1=Rate` (PtgRefN, PtgName, PtgEq), with a
 * solid red DXFPat fill. Sheet `A` names Lbl #1 and sheet `B` Lbl #2, as an
 * unqualified `Rate` binds on each sheet. The rule therefore matches A1 on `A`
 * only when the local wins, and B1 on `B` only when `A`'s local stays out of
 * `B`'s scope.
 */
import { concat, little16, little32 } from './test-fixtures.js';
import { workbookCfb } from './xls-workbook-fixture.js';

const record = (kind: number, data: Uint8Array = new Uint8Array()) => concat(little16(kind), little16(data.length), data);
const bof = (kind: number) => record(0x0809, concat(little16(0x0600), little16(kind), new Uint8Array(12)));

/** Lbl: flags, chKey, cch, cce, reserved, itab, reserved, ASCII name, PtgInt. */
function lbl(itab: number, name: string, value: number): Uint8Array {
  const header = new Uint8Array(15);
  header[3] = name.length;
  header.set(little16(3), 4);
  header.set(little16(itab), 8);
  return record(0x0018, concat(header, Uint8Array.from(name, c => c.charCodeAt(0)), Uint8Array.of(0x1e), little16(value)));
}

function number(col: number, value: number): Uint8Array {
  const data = new Uint8Array(14);
  data.set(little16(col), 2);
  new DataView(data.buffer).setFloat64(6, value, true);
  return record(0x0203, data);
}

/** Ref8U A1:B1 (rwFirst, rwLast, colFirst, colLast). */
const a1b1 = concat(little16(0), little16(0), little16(0), little16(1));
/** FrtRefHeader(U) with fFrtRef = 1 over the rule's bound. */
const frtRef = (rt: number) => concat(little16(rt), little16(1), a1b1);

function conditionalRule(nameIndex: number): Uint8Array {
  const condFmt12 = concat(frtRef(0x0879), little16(1), little16(1 << 1), a1b1, little16(1), a1b1);
  // A1 relative to the bound's top-left, the one-based Lbl index, then `=`.
  const rgce = concat(Uint8Array.of(0x4c), little16(0), little16(0xc000), Uint8Array.of(0x23), little32(nameIndex), Uint8Array.of(0x0b));
  // DXFN12 with only DXFPat: ibitAtrPat and icvBNinch; fls solid, icvFore 10.
  const dxfn = concat(little32((1 << 29) | (1 << 18)), little16(0), little32((1 << 10) | (10 << 16) | (0x41 << 23)));
  const cf12 = concat(
    frtRef(0x087a), Uint8Array.of(2, 0), little16(rgce.length), little16(0),
    little32(dxfn.length), dxfn, rgce,
    little16(0), Uint8Array.of(0), little16(1), little16(1), Uint8Array.of(16), new Uint8Array(16),
  );
  return concat(record(0x0879, condFmt12), record(0x087a, cf12));
}

export function buildXlsScopedNamesFixture(): Uint8Array {
  const bodies = [1, 2].map(nameIndex => concat(
    bof(0x0010), record(0x0081, little16(0)), number(0, 1), number(1, 5), conditionalRule(nameIndex), record(0x000a),
  ));
  const bound = (offset: number, name: string) => record(0x0085, concat(little32(offset), Uint8Array.of(0, 0, 1, 0, name.charCodeAt(0))));
  const globals = (offsets: readonly [number, number]) => concat(
    bof(0x0005), bound(offsets[0], 'A'), bound(offsets[1], 'B'), lbl(1, 'Rate', 1), lbl(0, 'Rate', 5), record(0x000a),
  );
  const size = globals([0, 0]).length;
  return workbookCfb(concat(globals([size, size + bodies[0]!.length]), ...bodies));
}
