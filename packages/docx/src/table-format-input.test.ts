import { describe, expect, it } from 'vitest';
import { adjacentTableSequenceInput, tableFormatInput } from './parser-model.js';
import type { BodyElement, DocTable } from './types.js';

const noBorders = {
  top: null, bottom: null, left: null, right: null, insideH: null, insideV: null,
};

function cell(cellWire: Record<string, unknown> | undefined = undefined): Record<string, unknown> {
  return {
    content: [], colSpan: 1, vMerge: null, borders: noBorders,
    background: null, vAlign: 'top', widthPt: null,
    ...(cellWire ? { __tableCellLayout: cellWire } : {}),
  };
}

function row(
  rowWire: Record<string, unknown> | undefined,
  cells: Record<string, unknown>[] = [cell()],
  publicHeight: { value: number | null; rule: string } = { value: null, rule: 'auto' },
): Record<string, unknown> {
  return {
    cells, rowHeight: publicHeight.value, rowHeightRule: publicHeight.rule,
    isHeader: false,
    ...(rowWire ? { __tableRowLayout: rowWire } : {}),
  };
}

function table(
  rows: Record<string, unknown>[],
  tableWire: Record<string, unknown> | undefined = undefined,
  bidiVisual = false,
): DocTable {
  return {
    colWidths: [100], rows, borders: noBorders,
    cellMarginTop: 1, cellMarginBottom: 2, cellMarginLeft: 3, cellMarginRight: 4,
    jc: 'left', bidiVisual,
    ...(tableWire ? { __tableLayout: tableWire } : {}),
  } as unknown as DocTable;
}

describe('table format acquisition adapter', () => {
  it('retains effective style, ordinary-flow classification, and positioning atomically', () => {
    const source = table([], {
      effectiveStyleId: 'PositionedStyle',
      ordinaryFlow: false,
      grid: { authored: false, columns: [], requiredColumnCount: 0 },
      preferredWidth: null, layout: null, cellSpacing: null,
    });
    source.tblpPr = {
      leftFromText: 1, rightFromText: 2, topFromText: 3, bottomFromText: 4,
      horzAnchor: 'margin', horzSpecified: true, vertAnchor: 'page',
      tblpX: 5, tblpY: 6, tblpXSpec: 'center', tblpYSpec: 'bottom',
    };

    const input = tableFormatInput(source);

    expect(input).toMatchObject({
      effectiveStyleId: 'PositionedStyle',
      ordinaryFlow: false,
      positioning: {
        leftFromTextPt: 1, rightFromTextPt: 2, topFromTextPt: 3, bottomFromTextPt: 4,
        horzAnchor: 'margin', horzSpecified: true, vertAnchor: 'page',
        xPt: 5, yPt: 6, xAlign: 'center', yAlign: 'bottom',
      },
    });
    source.tblpPr.horzAnchor = 'text';
    expect(input.positioning?.horzAnchor).toBe('margin');
    expect(Object.isFrozen(input.positioning)).toBe(true);

    const ignored = table([], {
      effectiveStyleId: 'OrdinaryStyle', ordinaryFlow: true,
      grid: { authored: false, columns: [], requiredColumnCount: 0 },
      preferredWidth: null, layout: null, cellSpacing: null,
    });
    ignored.tblpPr = { ...source.tblpPr };
    expect(tableFormatInput(ignored)).toMatchObject({
      effectiveStyleId: 'OrdinaryStyle', ordinaryFlow: true, positioning: null,
    });
  });

  it('projects adjacent body values with cached effective table facts', () => {
    const source = table([], {
      effectiveStyleId: 'ProjectedStyle', ordinaryFlow: true,
      grid: { authored: false, columns: [], requiredColumnCount: 0 },
      preferredWidth: null, layout: null, cellSpacing: null,
    }) as DocTable & { type: 'table' };
    source.type = 'table';
    const paragraph = { type: 'paragraph', runs: [] } as unknown as BodyElement;

    const projected = adjacentTableSequenceInput([source, paragraph]);

    expect(projected).toEqual([
      {
        element: source,
        table: { effectiveStyleId: 'ProjectedStyle', ordinaryFlow: true },
      },
      { element: paragraph, table: null },
    ]);
    expect(Object.isFrozen(projected)).toBe(true);
    expect(Object.isFrozen(projected[0]?.table)).toBe(true);
  });

  it('distinguishes omitted hRule from explicit auto and converts twips once', () => {
    const input = tableFormatInput(table([
      row({
        height: { value: '480', rule: 'auto', ruleAuthored: false },
        beforeWidth: null, afterWidth: null, cellSpacing: null, exception: null,
      }),
      row({
        height: { value: '480', rule: 'auto', ruleAuthored: true },
        beforeWidth: null, afterWidth: null, cellSpacing: null, exception: null,
      }),
      row({
        height: { value: '300', rule: 'exact', ruleAuthored: true },
        beforeWidth: null, afterWidth: null, cellSpacing: null, exception: null,
      }),
      row(undefined, [cell()], { value: 18, rule: 'atLeast' }),
      row(undefined, [cell()], { value: 16, rule: 'auto' }),
    ]));

    expect(input.rows.map(({ height }) => height)).toEqual([
      { rule: 'atLeast', valuePt: 24 },
      { rule: 'auto', valuePt: 24 },
      { rule: 'exact', valuePt: 15 },
      { rule: 'atLeast', valuePt: 18 },
      { rule: 'atLeast', valuePt: 16 },
    ]);
  });

  it('resolves cell, exception, and table margins per edge including bidi start/end', () => {
    const exception = {
      preferredWidth: null, layout: null, justification: null, indent: null,
      borders: null, cellSpacing: null,
      cellMargins: {
        top: { kind: 'dxa', value: '100' },
        bottom: null,
        start: { kind: 'dxa', value: '120' },
        end: { kind: 'dxa', value: '140' },
        left: null, right: null,
      },
    };
    const make = (bidi: boolean) => tableFormatInput(table([
      row({
        height: null, beforeWidth: null, afterWidth: null, cellSpacing: null, exception,
      }, [cell({
        preferredWidth: null,
        margins: {
          top: { kind: 'pct', value: '500' },
          bottom: { kind: 'dxa', value: '80' },
          start: { kind: 'dxa', value: '200' },
          end: null, left: null, right: null,
        },
      })]),
    ], undefined, bidi));

    expect(make(false).rows[0]?.cells[0]?.marginsPt).toEqual({
      top: 5, bottom: 4, left: 10, right: 7,
    });
    expect(make(true).rows[0]?.cells[0]?.marginsPt).toEqual({
      top: 5, bottom: 4, left: 7, right: 10,
    });
  });

  it('resolves row spacing by scope and lets Word pct/auto zero shadow lower scopes', () => {
    const tableWire = {
      effectiveStyleId: null,
      grid: { authored: false, columns: [], requiredColumnCount: 1 },
      preferredWidth: null, layout: null,
      cellSpacing: { kind: 'dxa', value: '60' }, cellMargins: null,
    };
    const exception = {
      preferredWidth: null, layout: null, justification: null, indent: null,
      borders: null, cellMargins: null, cellSpacing: { kind: 'dxa', value: '40' },
    };
    const rows = [
      row({ height: null, beforeWidth: null, afterWidth: null,
        cellSpacing: { kind: 'dxa', value: '20' }, exception }),
      row({ height: null, beforeWidth: null, afterWidth: null,
        cellSpacing: { kind: 'pct', value: '500' }, exception }),
      row({ height: null, beforeWidth: null, afterWidth: null,
        cellSpacing: null, exception: null }),
    ];
    expect(tableFormatInput(table(rows, tableWire)).rows.map((item) => item.cellSpacingPt))
      .toEqual([1, 0, 3]);
  });

  it('applies CT_TblWidth defaults and lets nil spacing shadow lower scopes with zero', () => {
    const tableWire = {
      effectiveStyleId: null,
      grid: { authored: false, columns: [], requiredColumnCount: 1 },
      preferredWidth: null, layout: null,
      cellSpacing: { kind: 'dxa', value: '60' }, cellMargins: null,
    };
    const exception = {
      preferredWidth: null, layout: null, justification: null, indent: null,
      borders: null, cellMargins: null, cellSpacing: { kind: 'dxa', value: '40' },
    };
    const rows = [
      row({ height: null, beforeWidth: null, afterWidth: null,
        cellSpacing: { kind: null, value: '20' }, exception }),
      row({ height: null, beforeWidth: null, afterWidth: null,
        cellSpacing: { kind: 'nil', value: '999' }, exception }),
      row({ height: null, beforeWidth: null, afterWidth: null,
        cellSpacing: { kind: null, value: null }, exception }),
      row({ height: null, beforeWidth: null, afterWidth: null,
        cellSpacing: { kind: 'dxa', value: '25%' }, exception }),
    ];

    expect(tableFormatInput(table(rows, tableWire)).rows.map((item) => item.cellSpacingPt))
      .toEqual([1, 0, 0, 0]);
  });

  it('uses the resolved table-style spacing below direct table properties', () => {
    const tableWire = {
      effectiveStyleId: 'Synthetic',
      grid: { authored: false, columns: [], requiredColumnCount: 1 },
      preferredWidth: null, layout: null, cellSpacing: null, cellMargins: null,
    };
    const styled = row({
      height: null, beforeWidth: null, afterWidth: null, cellSpacing: null,
      styleCellSpacing: { kind: null, value: '60' }, exception: null,
    });
    const directTableWire = { ...tableWire, cellSpacing: { kind: null, value: '20' } };

    expect(tableFormatInput(table([styled], tableWire)).rows[0]?.cellSpacingPt).toBe(3);
    expect(tableFormatInput(table([styled], directTableWire)).rows[0]?.cellSpacingPt).toBe(1);
  });

  it('maps table and style start/end margins after bidi direction is known', () => {
    const tableWire = {
      effectiveStyleId: null,
      grid: { authored: false, columns: [], requiredColumnCount: 1 },
      preferredWidth: null, layout: null, cellSpacing: null,
      cellMargins: {
        start: { kind: null, value: '100' },
        end: { kind: 'nil', value: '900' },
      },
    };
    const make = (bidi: boolean) => tableFormatInput(table([
      row({ height: null, beforeWidth: null, afterWidth: null,
        cellSpacing: null,
        styleCellMargins: {
          start: { kind: 'dxa', value: '120' },
          end: { kind: 'dxa', value: '140' },
        },
        exception: null }),
    ], tableWire, bidi)).rows[0]?.cells[0]?.marginsPt;

    expect(make(false)).toEqual({ top: 1, bottom: 2, left: 5, right: 0 });
    expect(make(true)).toEqual({ top: 1, bottom: 2, left: 0, right: 5 });
  });

  it('maps row-conditional style margins below direct table margins', () => {
    const tableWire = {
      effectiveStyleId: 'Synthetic',
      grid: { authored: false, columns: [], requiredColumnCount: 1 },
      preferredWidth: null, layout: null, cellSpacing: null,
      cellMargins: { start: { kind: null, value: '40' } },
    };
    const styledRow = row({
      height: null, beforeWidth: null, afterWidth: null, cellSpacing: null,
      styleCellMargins: {
        start: { kind: null, value: '120' }, end: { kind: null, value: '20' },
      },
      exception: null,
    });
    const make = (bidi: boolean) => tableFormatInput(table(
      [styledRow], tableWire, bidi,
    )).rows[0]?.cells[0]?.marginsPt;

    expect(make(false)).toEqual({ top: 1, bottom: 2, left: 2, right: 1 });
    expect(make(true)).toEqual({ top: 1, bottom: 2, left: 1, right: 2 });
  });

  it('lets percent-sign syntax override a contradictory table margin type', () => {
    const tableWire = {
      effectiveStyleId: 'Synthetic',
      grid: { authored: false, columns: [], requiredColumnCount: 1 },
      preferredWidth: null, layout: null, cellSpacing: null,
      cellMargins: { start: { kind: 'dxa', value: '25%' } },
    };
    const input = tableFormatInput(table([
      row({ height: null, beforeWidth: null, afterWidth: null,
        cellSpacing: null,
        styleCellMargins: { start: { kind: 'dxa', value: '100' } },
        exception: null }),
    ], tableWire));

    expect(input.rows[0]?.cells[0]?.marginsPt.left).toBe(0);
  });

  it('retains authored nil first-row indentation as a zero override', () => {
    const exception = {
      preferredWidth: null, layout: null, justification: null,
      indent: { kind: 'nil', value: '1440' },
      borders: null, cellMargins: null, cellSpacing: null,
    };
    const input = tableFormatInput(table([
      row({ height: null, beforeWidth: null, afterWidth: null,
        cellSpacing: null, exception }),
    ]));

    expect(input.firstRowException).toMatchObject({ indentAuthored: true, indentPt: 0 });
  });

  it('ignores the raw indent type when the lexical value is a percentage', () => {
    const exception = {
      layout: null, justification: null, indent: { kind: 'nil', value: '25%' },
      borders: null, cellMargins: null, cellSpacing: null,
    };
    const input = tableFormatInput(table([
      row({ height: null, beforeWidth: null, afterWidth: null,
        cellSpacing: null, exception }),
    ]));

    expect(input.firstRowException).toMatchObject({
      preferredWidthAuthored: false, indentAuthored: false, indentPt: null,
    });
  });

  it('preserves public per-cell margin fallbacks for hand-built DocTable values', () => {
    const publicCell = {
      ...cell(), marginTop: 9, marginBottom: 8, marginLeft: 7, marginRight: 6,
    };
    const input = tableFormatInput(table([
      row(undefined, [publicCell]),
    ]));

    expect(input.rows[0]?.cells[0]?.marginsPt).toEqual({
      top: 9, bottom: 8, left: 7, right: 6,
    });
  });

  it('returns first-row tblPrEx facts as deeply frozen plain data', () => {
    const top = { width: 1, color: '112233', style: 'single' };
    const exception = {
      preferredWidth: { kind: 'pct', value: '2500' },
      layout: { kind: 'fixed' }, justification: 'center',
      indent: { kind: 'dxa', value: '120' },
      borders: { ...noBorders, top }, cellMargins: null, cellSpacing: null,
    };
    const source = table([
      row({ height: null, beforeWidth: null, afterWidth: null, cellSpacing: null, exception }),
    ]);
    const input = tableFormatInput(source);

    expect(input.firstRowException).toEqual({
      preferredWidthAuthored: true,
      preferredWidth: { kind: 'pct', value: 0.5 },
      layout: 'fixed', justification: 'center', indentAuthored: true, indentPt: 6,
      borders: { ...noBorders, top },
    });
    expect(input.rows[0]?.exception).toBe(input.firstRowException);
    expect(Object.isFrozen(input)).toBe(true);
    expect(Object.isFrozen(input.firstRowException?.borders?.top)).toBe(true);

    top.width = 99;
    expect(tableFormatInput(source)).toBe(input);
    expect(input.firstRowException?.borders?.top?.width).toBe(1);
  });

  it('retains normalized exception facts on later rows for row geometry and paint', () => {
    const bottom = { width: 2, color: null, style: 'double' };
    const laterException = {
      preferredWidth: { kind: 'dxa', value: '1440' },
      layout: { kind: 'autofit' }, justification: 'right',
      indent: { kind: 'pct', value: '500' },
      borders: { ...noBorders, bottom }, cellMargins: null, cellSpacing: null,
    };
    const input = tableFormatInput(table([
      row({ height: null, beforeWidth: null, afterWidth: null,
        cellSpacing: null, exception: null }),
      row({ height: null, beforeWidth: null, afterWidth: null,
        cellSpacing: null, exception: laterException }),
    ]));

    expect(input.firstRowException).toBeNull();
    expect(input.rows[0]?.exception).toBeNull();
    expect(input.rows[1]?.exception).toEqual({
      preferredWidthAuthored: true,
      preferredWidth: { kind: 'dxa', value: 72 },
      layout: 'autofit', justification: 'right', indentAuthored: false, indentPt: null,
      borders: { ...noBorders, bottom },
    });
    expect(Object.isFrozen(input.rows[1]?.exception?.borders?.bottom)).toBe(true);
  });

  it('resolves direct row alignment before the table-property exception', () => {
    const exception = {
      preferredWidth: null, layout: null, justification: 'center', indent: null,
      borders: null, cellMargins: null, cellSpacing: null,
    };
    const input = tableFormatInput(table([
      row({
        height: null, justification: 'end', beforeWidth: null, afterWidth: null,
        cellSpacing: null, exception,
      }),
      row({
        height: null, justification: null, beforeWidth: null, afterWidth: null,
        cellSpacing: null, exception,
      }),
    ]));

    expect(input.rows.map((item) => item.justification)).toEqual(['end', 'center']);
  });
});
