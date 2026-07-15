import { describe, expect, it } from 'vitest';
import {
  effectiveTablePositioning,
  tableAcquisitionInput,
  tableParticipatesInOrdinaryFlow,
} from './parser-model.js';
import type { DocTable } from './types.js';

function tableWithPrivateLayoutWire(): DocTable {
  return {
    colWidths: [36, 0],
    rows: [{
      cells: [{
        content: [], colSpan: 3, vMerge: null,
        borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
        background: null, vAlign: 'top', widthPt: null,
        __tableCellLayout: {
          preferredWidth: { kind: 'pct', value: '2500' },
          margins: {
            start: { kind: 'dxa', value: '100' },
            end: { kind: 'pct', value: '500' },
          },
        },
      }],
      rowHeight: 24, rowHeightRule: 'auto', isHeader: false,
      __tableRowLayout: {
        height: { value: '480', rule: 'auto', ruleAuthored: false },
        justification: 'end',
        beforeWidth: { kind: 'pct', value: '15%' },
        afterWidth: null,
        cellSpacing: { kind: 'dxa', value: '20' },
        exception: null,
      },
    }],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 5.75, cellMarginRight: 5.75,
    jc: 'left',
    __tableLayout: {
      effectiveStyleId: 'SyntheticTableStyle',
      ordinaryFlow: true,
      grid: {
        authored: true,
        columns: [{ width: '720' }, { width: null }],
        requiredColumnCount: 5,
      },
      preferredWidth: { kind: 'pct', value: '3750' },
      layout: { kind: 'fixed' },
      cellSpacing: { kind: 'dxa', value: '40' },
    },
  } as unknown as DocTable;
}

describe('parser-private table acquisition projection', () => {
  it('snapshots plain clone-safe layout facts without widening DocTable', () => {
    const table = tableWithPrivateLayoutWire();
    const input = tableAcquisitionInput(table);

    expect(input).toEqual({
      table: {
        effectiveStyleId: 'SyntheticTableStyle',
        ordinaryFlow: true,
        grid: {
          authored: true,
          columns: [{ width: '720' }, { width: null }],
          requiredColumnCount: 5,
        },
        preferredWidth: { kind: 'pct', value: '3750' },
        layout: { kind: 'fixed' },
        cellSpacing: { kind: 'dxa', value: '40' },
      },
      rows: [{
        row: {
          height: { value: '480', rule: 'auto', ruleAuthored: false },
          justification: 'end',
          beforeWidth: { kind: 'pct', value: '15%' },
          afterWidth: null,
          cellSpacing: { kind: 'dxa', value: '20' },
          exception: null,
        },
        cells: [{
          preferredWidth: { kind: 'pct', value: '2500' },
          margins: {
            start: { kind: 'dxa', value: '100' },
            end: { kind: 'pct', value: '500' },
          },
        }],
      }],
    });
    expect(Object.isFrozen(input)).toBe(true);
    expect(Object.isFrozen(input.table?.grid.columns)).toBe(true);
    expect(Object.isFrozen(input.rows[0]?.cells[0]?.margins)).toBe(true);
  });

  it('caches by table identity and does not retain caller-owned wire objects', () => {
    const table = tableWithPrivateLayoutWire();
    const first = tableAcquisitionInput(table);
    const wire = table as unknown as {
      __tableLayout: { grid: { columns: Array<{ width: string | null }> } };
      rows: Array<{ __tableRowLayout: { height: { value: string } } }>;
    };

    wire.__tableLayout.grid.columns[0]!.width = '9999';
    wire.rows[0]!.__tableRowLayout.height.value = '9999';
    const second = tableAcquisitionInput(table);

    expect(second).toBe(first);
    expect(second.table?.grid.columns[0]?.width).toBe('720');
    expect(second.rows[0]?.row?.height?.value).toBe('480');
  });

  it('projects an equivalent immutable value after a worker structured clone', () => {
    const main = tableAcquisitionInput(tableWithPrivateLayoutWire());
    const worker = tableAcquisitionInput(structuredClone(tableWithPrivateLayoutWire()));
    expect(worker).toEqual(main);
    expect(worker).not.toBe(main);
  });

  it('keeps positional null fallbacks for hand-built public table values', () => {
    const table = tableWithPrivateLayoutWire() as DocTable & Record<string, unknown>;
    delete table.__tableLayout;
    for (const row of table.rows as unknown as Array<Record<string, unknown> & { cells: unknown[] }>) {
      delete row.__tableRowLayout;
      for (const cell of row.cells as Array<Record<string, unknown>>) delete cell.__tableCellLayout;
    }

    expect(tableAcquisitionInput(table)).toEqual({
      table: null,
      rows: [{ row: null, cells: [null] }],
    });
  });

  it('uses retained effective flow status before the lexical tblpPr compatibility field', () => {
    const ignoredPositioning = tableWithPrivateLayoutWire() as DocTable & {
      tblpPr: NonNullable<DocTable['tblpPr']>;
    };
    ignoredPositioning.tblpPr = {
      leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: 'text', horzSpecified: true, vertAnchor: 'margin',
      tblpX: 0, tblpY: 0,
    };
    const effectiveFloating = structuredClone(ignoredPositioning) as typeof ignoredPositioning & {
      __tableLayout: { ordinaryFlow: boolean };
    };
    effectiveFloating.__tableLayout.ordinaryFlow = false;

    expect(tableParticipatesInOrdinaryFlow(ignoredPositioning)).toBe(true);
    expect(tableParticipatesInOrdinaryFlow(effectiveFloating)).toBe(false);
    expect(effectiveTablePositioning(ignoredPositioning)).toBeNull();
    expect(effectiveTablePositioning(effectiveFloating)).toBe(effectiveFloating.tblpPr);
  });

  it('keeps the stable public table fallback deterministic when acquisition facts are absent', () => {
    const ordinary = tableWithPrivateLayoutWire() as DocTable & Record<string, unknown>;
    const floating = structuredClone(ordinary) as DocTable & Record<string, unknown>;
    delete ordinary.__tableLayout;
    delete floating.__tableLayout;
    floating.tblpPr = {
      leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: 'page', horzSpecified: true, vertAnchor: 'page',
      tblpX: 1, tblpY: 1,
    };

    expect(tableParticipatesInOrdinaryFlow(ordinary)).toBe(true);
    expect(tableParticipatesInOrdinaryFlow(floating)).toBe(false);
  });
});
