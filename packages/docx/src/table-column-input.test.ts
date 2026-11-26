import { describe, expect, it } from 'vitest';
import { tableColumnLayoutInput } from './parser-model.js';
import type { DocTable } from './types.js';

function emptyBorders() {
  return { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null };
}

describe('table column acquisition boundary', () => {
  it('repairs invalid hand-built numeric widths to definitional zero', () => {
    const table = {
      colWidths: [-1, Number.NaN, Number.POSITIVE_INFINITY], rows: [], borders: emptyBorders(),
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left',
    } as unknown as DocTable;

    const result = tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 0, maxWidthPt: 0 }));

    expect(result.gridWidthsPt).toEqual([0, 0, 0]);
    expect(result.gridWidthKeys).toEqual(['0/1', '0/1', '0/1']);
  });

  it('repairs valid exact and over-budget measures with nonfinite geometry to definitional zero', () => {
    const exactNonfinite = `${'1'}${'0'.repeat(400)}pt`;
    const overBudgetNonfinite = `${'9'.repeat(800)}pt`;
    const table = {
      colWidths: [], rows: [], borders: emptyBorders(),
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left',
      __tableLayout: {
        effectiveStyleId: null,
        grid: {
          authored: true,
          columns: [{ width: exactNonfinite }, { width: overBudgetNonfinite }],
          requiredColumnCount: 2,
        },
        preferredWidth: null, layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    const result = tableColumnLayoutInput(table, Number.MAX_VALUE, () => ({
      minWidthPt: 0, maxWidthPt: 0,
    }));

    expect(result.gridWidthsPt).toEqual([0, 0]);
    expect(result.gridWidthKeys).toEqual(['0/1', '0/1']);
  });
  it('preserves the prior finite solver input for a schema-valid large unsigned twips measure', () => {
    const lexicalTwips = '2542686831678384';
    const table = {
      colWidths: [], rows: [], borders: emptyBorders(),
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left',
      __tableLayout: {
        effectiveStyleId: null,
        grid: {
          authored: true,
          columns: [{ width: lexicalTwips }],
          requiredColumnCount: 1,
        },
        preferredWidth: null, layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    const result = tableColumnLayoutInput(table, Number.MAX_VALUE, () => ({
      minWidthPt: 0, maxWidthPt: 0,
    }));

    expect(result.gridWidthsPt).toEqual([Number(lexicalTwips) / 20]);
    expect(result.gridWidthKeys).toEqual(['635671707919596/5']);
  });

  it('normalizes equivalent Transitional and Strict grid measures to exact point identities', () => {
    const table = {
      colWidths: [], rows: [], borders: emptyBorders(),
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left',
      __tableLayout: {
        effectiveStyleId: null,
        grid: {
          authored: true,
          columns: [
            { width: '3' },
            { width: '0.15pt' },
            { width: '1440' },
            { width: '72pt' },
            { width: '1in' },
            { width: '2.54cm' },
            { width: '25.4mm' },
            { width: '6pc' },
            { width: '6pi' },
          ],
          requiredColumnCount: 9,
        },
        preferredWidth: null, layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    const result = tableColumnLayoutInput(table, 200, () => ({
      minWidthPt: 0, maxWidthPt: 0,
    })) as ReturnType<typeof tableColumnLayoutInput> & {
      readonly gridWidthKeys: readonly (string | null)[];
    };

    expect(result.gridWidthsPt.slice(0, 2)).toEqual([0.15, 0.15]);
    expect(result.gridWidthKeys.slice(0, 2)).toEqual(['3/20', '3/20']);
    expect(result.gridWidthKeys.slice(2)).toEqual([
      '72/1', '72/1', '72/1', '72/1', '72/1', '72/1', '72/1',
    ]);
    expect(JSON.parse(JSON.stringify(result)).gridWidthKeys).toEqual(result.gridWidthKeys);
    expect(() => structuredClone(result)).not.toThrow();
  });

  it('accepts the xsd:unsignedLong maximum and rejects unitless overflow before exact arithmetic', () => {
    const table = {
      colWidths: [], rows: [], borders: emptyBorders(),
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left',
      __tableLayout: {
        effectiveStyleId: null,
        grid: {
          authored: true,
          columns: [
            { width: ' 18446744073709551615\t' },
            { width: '\n+18446744073709551615\r' },
            { width: '18446744073709551616' },
            { width: '+18446744073709551616' },
            { width: '9'.repeat(10_000) },
            { width: '-0' },
            { width: ' \t-00\r\n' },
          ],
          requiredColumnCount: 7,
        },
        preferredWidth: null, layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    const result = tableColumnLayoutInput(table, 200, () => ({
      minWidthPt: 0, maxWidthPt: 0,
    }));

    expect(result.gridWidthsPt).toEqual([
      Number('18446744073709551615') / 20,
      Number('18446744073709551615') / 20,
      0,
      0,
      0,
      0,
      0,
    ]);
    expect(result.gridWidthKeys).toEqual([
      '3689348814741910323/4',
      '3689348814741910323/4',
      '0/1',
      '0/1',
      '0/1',
      '0/1',
      '0/1',
    ]);
  });

  it('enforces the ST_PositiveUniversalMeasure lexical grammar for unit-bearing widths', () => {
    const table = {
      colWidths: [], rows: [], borders: emptyBorders(),
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left',
      __tableLayout: {
        effectiveStyleId: null,
        grid: {
          authored: true,
          columns: [
            { width: '0.5pt' },
            { width: '1.0pt' },
            { width: '1e2pt' },
            { width: '.5pt' },
            { width: '1.pt' },
            { width: '+1pt' },
            { width: '-1pt' },
            { width: ' 0.5pt' },
            { width: '0.5pt ' },
          ],
          requiredColumnCount: 9,
        },
        preferredWidth: null, layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    const result = tableColumnLayoutInput(table, 200, () => ({
      minWidthPt: 0, maxWidthPt: 0,
    }));

    expect(result.gridWidthsPt).toEqual([0.5, 1, 0, 0, 0, 0, 0, 0, 0]);
    expect(result.gridWidthKeys).toEqual([
      '1/2', '1/1', '0/1', '0/1', '0/1', '0/1', '0/1', '0/1', '0/1',
    ]);
  });

  it('degrades an over-budget universal magnitude to binary64 geometry with a null identity', () => {
    const table = {
      colWidths: [], rows: [], borders: emptyBorders(),
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left',
      __tableLayout: {
        effectiveStyleId: null,
        grid: {
          authored: true,
          columns: [
            // Millions of redundant fraction digits: over the exact budget. The
            // exact identity is unknown (key null) but the binary64 geometry
            // (72pt) survives.
            { width: `72.${'0'.repeat(1_000_000)}1pt` },
            // A sub-underflow universal magnitude degrades to a zero track.
            { width: `0.${'0'.repeat(1_000_000)}1pt` },
            // A normal exact universal value keeps its exact identity.
            { width: '18pt' },
          ],
          requiredColumnCount: 3,
        },
        preferredWidth: null, layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    const result = tableColumnLayoutInput(table, 200, () => ({
      minWidthPt: 0, maxWidthPt: 0,
    }));

    expect(result.gridWidthKeys).toEqual([null, null, '18/1']);
    expect(result.gridWidthsPt).toEqual([72, 0, 18]);
  });

  it('gives distinct over-budget widths a null identity but their true binary64 width', () => {
    // Two authored widths that differ only past the exact budget both resolve to
    // key: null (identity unknown, NOT asserting equality) yet keep width 72.
    const table = {
      colWidths: [], rows: [], borders: emptyBorders(),
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left',
      __tableLayout: {
        effectiveStyleId: null,
        grid: {
          authored: true,
          columns: [
            { width: `72.${'0'.repeat(800)}1pt` },
            { width: `72.${'0'.repeat(800)}2pt` },
          ],
          requiredColumnCount: 2,
        },
        preferredWidth: null, layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    const result = tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 0, maxWidthPt: 0 }));

    expect(result.gridWidthKeys).toEqual([null, null]);
    expect(result.gridWidthsPt).toEqual([72, 72]);
  });

  it('retains identity at the significant-digit budget and drops it beyond, both keeping geometry', () => {
    const sig768 = `1.${'2'.repeat(767)}pt`; // 768 significant digits
    const sig769 = `1.${'2'.repeat(768)}pt`; // 769 significant digits
    const table = {
      colWidths: [], rows: [], borders: emptyBorders(),
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left',
      __tableLayout: {
        effectiveStyleId: null,
        grid: {
          authored: true,
          columns: [{ width: sig768 }, { width: sig769 }],
          requiredColumnCount: 2,
        },
        preferredWidth: null, layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    const result = tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 0, maxWidthPt: 0 }));

    expect(result.gridWidthKeys![0]).not.toBeNull(); // 768 retained
    expect(result.gridWidthKeys![1]).toBeNull(); // 769 over budget
    expect(result.gridWidthsPt).toEqual([Number('1.' + '2'.repeat(767)), Number('1.' + '2'.repeat(768))]);
  });

  it('retains an in-budget subnormal underflow magnitude as an exact identity', () => {
    const underflow = `0.${'0'.repeat(322)}1pt`; // 1e-323, in the decimal budget
    const table = {
      colWidths: [], rows: [], borders: emptyBorders(),
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left',
      __tableLayout: {
        effectiveStyleId: null,
        grid: { authored: true, columns: [{ width: underflow }], requiredColumnCount: 1 },
        preferredWidth: null, layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    const result = tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 0, maxWidthPt: 0 }));

    expect(result.gridWidthKeys![0]).not.toBeNull();
    expect(result.gridWidthsPt[0]).toBe(1e-323);
  });

  it('normalizes parser-private lexical widths without exposing parser objects to layout', () => {
    const cell = {
      content: [], colSpan: 3, vMerge: null, borders: emptyBorders(),
      background: null, vAlign: 'top', widthPt: null,
      __tableCellLayout: { preferredWidth: { kind: 'pct', value: '2500' }, margins: null },
    };
    const table = {
      colWidths: [36, 0],
      rows: [{
        gridBefore: 1, gridAfter: 1,
        cells: [cell], rowHeight: null, rowHeightRule: 'auto', isHeader: false,
        __tableRowLayout: {
          height: null,
          beforeWidth: { kind: 'pct', value: '15%' },
          afterWidth: { kind: 'dxa', value: '200' },
          cellSpacing: null, exception: null,
        },
      }],
      borders: emptyBorders(),
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
      jc: 'left',
      __tableLayout: {
        effectiveStyleId: 'Synthetic',
        grid: {
          authored: true,
          columns: [{ width: '720' }, { width: null }],
          requiredColumnCount: 5,
        },
        preferredWidth: { kind: 'pct', value: '3750' },
        layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    let intrinsicProbeCount = 0;
    const result = tableColumnLayoutInput(table, 200, () => {
      intrinsicProbeCount += 1;
      return { minWidthPt: 12, maxWidthPt: 30 };
    });

    expect(intrinsicProbeCount).toBe(0);

    expect(result).toEqual({
      layout: 'fixed', availableWidthPt: 200,
      gridWidthsPt: [36, 0, 0, 0, 0],
      // Exact authored grid identity: the '720' twip column is retained as its
      // exact point key (720/20), never re-derived from the IEEE-754 width.
      gridWidthKeys: ['36/1', '0/1', '0/1', '0/1', '0/1'],
      tablePreferredWidthPt: 150,
      tablePreferredWidthRelative: true,
      rows: [{
        // wBefore/wAfter percentages use the page text extents (§17.4.85–86),
        // unlike tcW percentages, which remain relative to final table width.
        before: { columnSpan: 1, preferredWidth: { kind: 'dxa', value: 30 } },
        after: { columnSpan: 1, preferredWidth: { kind: 'dxa', value: 10 } },
        cells: [{
          columnStart: 1, columnSpan: 3,
          preferredWidth: { kind: 'pct', value: 0.5 },
          minContentWidthPt: 0, maxContentWidthPt: 0,
        }],
      }],
    });
  });

  it('uses stable public fields only as a compatibility fallback for hand-built tables', () => {
    const cell = {
      content: [], colSpan: 1, vMerge: null, borders: emptyBorders(),
      background: null, vAlign: 'top', widthPt: 25, widthPct: null,
    };
    const table = {
      colWidths: [40], rows: [{ cells: [cell], gridBefore: 0, gridAfter: 0 }],
      borders: emptyBorders(), cellMarginTop: 0, cellMarginRight: 0,
      cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'autofit',
      widthPct: 2500,
    } as unknown as DocTable;

    expect(tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 10, maxWidthPt: 20 })))
      .toMatchObject({
        layout: 'autofit', gridWidthsPt: [40], tablePreferredWidthPt: 100,
        rows: [{ cells: [{
          columnStart: 0, columnSpan: 1,
          preferredWidth: { kind: 'dxa', value: 25 },
        }] }],
      });
  });

  it('applies first-row tblPrEx fixed layout and width to the whole table in Word mode', () => {
    const table = {
      colWidths: [40],
      rows: [{
        cells: [{
          content: [], colSpan: 1, vMerge: null, borders: emptyBorders(),
          background: null, vAlign: 'top', widthPt: null,
          __tableCellLayout: { preferredWidth: null, margins: null },
        }],
        gridBefore: 0, gridAfter: 0,
        __tableRowLayout: {
          height: null, beforeWidth: null, afterWidth: null, cellSpacing: null,
          exception: {
            preferredWidth: { kind: 'pct', value: '3000' },
            layout: { kind: 'fixed' }, justification: null, indent: null,
            borders: null, cellMargins: null, cellSpacing: null,
          },
        },
      }],
      borders: emptyBorders(), cellMarginTop: 0, cellMarginRight: 0,
      cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'autofit',
      __tableLayout: {
        effectiveStyleId: null,
        grid: { authored: true, columns: [{ width: '800' }], requiredColumnCount: 1 },
        preferredWidth: { kind: 'dxa', value: '2000' },
        layout: { kind: 'autofit' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    expect(tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 10, maxWidthPt: 20 })))
      .toMatchObject({ layout: 'fixed', tablePreferredWidthPt: 120 });
  });

  it.each([
    [{ kind: 'dxa', value: '0' }, 'zero'],
    [{ kind: 'auto', value: '1440' }, 'auto'],
    [{ kind: 'nil', value: '1440' }, 'nil'],
  ] as const)(
    'lets an authored first-row tblPrEx width of %s (%s) shadow the table preferred width',
    (exceptionWidth, _label) => {
      const table = {
        colWidths: [40],
        rows: [{
          cells: [{
            content: [], colSpan: 1, vMerge: null, borders: emptyBorders(),
            background: null, vAlign: 'top', widthPt: null,
            __tableCellLayout: { preferredWidth: null, margins: null },
          }],
          gridBefore: 0, gridAfter: 0,
          __tableRowLayout: {
            height: null, beforeWidth: null, afterWidth: null, cellSpacing: null,
            exception: {
              preferredWidth: exceptionWidth,
              layout: null, justification: null, indent: null,
              borders: null, cellMargins: null, cellSpacing: null,
            },
          },
        }],
        borders: emptyBorders(), cellMarginTop: 0, cellMarginRight: 0,
        cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'fixed',
        __tableLayout: {
          effectiveStyleId: null,
          grid: { authored: true, columns: [{ width: '800' }], requiredColumnCount: 1 },
          preferredWidth: { kind: 'dxa', value: '2000' },
          layout: { kind: 'fixed' }, cellSpacing: null,
        },
      } as unknown as DocTable;

      expect(tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 10, maxWidthPt: 20 })))
        .toMatchObject({ tablePreferredWidthPt: null });
    },
  );

  it('falls back to the parent tblW when a first-row exception omits tblW', () => {
    const table = {
      colWidths: [40],
      rows: [{
        cells: [{
          content: [], colSpan: 1, vMerge: null, borders: emptyBorders(),
          background: null, vAlign: 'top', widthPt: null,
          __tableCellLayout: { preferredWidth: null, margins: null },
        }],
        gridBefore: 0, gridAfter: 0,
        __tableRowLayout: {
          height: null, beforeWidth: null, afterWidth: null, cellSpacing: null,
          exception: {
            preferredWidth: null,
            layout: { kind: 'fixed' }, justification: null, indent: null,
            borders: null, cellMargins: null, cellSpacing: null,
          },
        },
      }],
      borders: emptyBorders(), cellMarginTop: 0, cellMarginRight: 0,
      cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'autofit',
      __tableLayout: {
        effectiveStyleId: null,
        grid: { authored: true, columns: [{ width: '800' }], requiredColumnCount: 1 },
        preferredWidth: { kind: 'dxa', value: '2000' },
        layout: { kind: 'autofit' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    expect(tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 10, maxWidthPt: 20 })))
      .toMatchObject({ layout: 'fixed', tablePreferredWidthPt: 100 });
  });

  it('uses the CT_TblWidth dxa default for an exception width with omitted type', () => {
    const table = {
      colWidths: [40],
      rows: [{
        cells: [{
          content: [], colSpan: 1, vMerge: null, borders: emptyBorders(),
          background: null, vAlign: 'top', widthPt: null,
          __tableCellLayout: { preferredWidth: null, margins: null },
        }],
        gridBefore: 0, gridAfter: 0,
        __tableRowLayout: {
          height: null, beforeWidth: null, afterWidth: null, cellSpacing: null,
          exception: {
            preferredWidth: { kind: null, value: '1440' },
            layout: null, justification: null, indent: null,
            borders: null, cellMargins: null, cellSpacing: null,
          },
        },
      }],
      borders: emptyBorders(), cellMarginTop: 0, cellMarginRight: 0,
      cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'fixed',
      __tableLayout: {
        effectiveStyleId: null,
        grid: { authored: true, columns: [{ width: '800' }], requiredColumnCount: 1 },
        preferredWidth: { kind: 'dxa', value: '2000' },
        layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    expect(tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 10, maxWidthPt: 20 })))
      .toMatchObject({ tablePreferredWidthPt: 72 });
  });

  it('ignores authored gridBefore/gridAfter values which do not fit the table grid', () => {
    const table = {
      colWidths: [20, 40],
      rows: [{
        cells: [{
          content: [], colSpan: 2, vMerge: null, borders: emptyBorders(),
          background: null, vAlign: 'top', widthPt: null,
          __tableCellLayout: { preferredWidth: null, margins: null },
        }],
        gridBefore: 3, gridAfter: 1,
        __tableRowLayout: {
          height: null, beforeWidth: { kind: 'dxa', value: '100' },
          afterWidth: { kind: 'dxa', value: '100' }, cellSpacing: null, exception: null,
        },
      }],
      borders: emptyBorders(), cellMarginTop: 0, cellMarginRight: 0,
      cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', layout: 'fixed',
      __tableLayout: {
        effectiveStyleId: null,
        grid: { authored: true, columns: [{ width: '400' }, { width: '800' }], requiredColumnCount: 2 },
        preferredWidth: null, layout: { kind: 'fixed' }, cellSpacing: null,
      },
    } as unknown as DocTable;

    expect(tableColumnLayoutInput(table, 200, () => ({ minWidthPt: 0, maxWidthPt: 0 })))
      .toMatchObject({
        gridWidthsPt: [20, 40],
        rows: [{ before: null, after: null, cells: [{ columnStart: 0, columnSpan: 2 }] }],
      });
  });
});
