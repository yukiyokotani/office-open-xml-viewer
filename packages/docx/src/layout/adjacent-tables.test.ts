import { describe, expect, it } from 'vitest';
import {
  normalizeAdjacentTables,
  type AdjacentTableSequenceInput,
} from './adjacent-tables.js';
import type { BodyElement } from '../types.js';

function table(
  effectiveStyleId: string,
  ordinaryFlow: boolean,
  computedWidthPt: number,
): BodyElement {
  return {
    type: 'table',
    colWidths: [computedWidthPt],
    rows: [],
    borders: { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null },
    cellMarginTop: 0,
    cellMarginRight: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    jc: 'left',
    __tableLayout: {
      effectiveStyleId,
      ordinaryFlow,
      grid: { authored: true, columns: [{ width: `${computedWidthPt * 20}` }], requiredColumnCount: 1 },
      preferredWidth: null,
      layout: null,
      cellSpacing: null,
    },
  } as unknown as BodyElement;
}

function paragraph(hidden = false): BodyElement {
  return { type: 'paragraph', hidden } as unknown as BodyElement;
}

function publicTable(): BodyElement {
  const value = table('discarded', true, 42) as BodyElement & Record<string, unknown>;
  delete value.__tableLayout;
  return value;
}

function input(
  element: BodyElement,
  effectiveStyleId: string | null = null,
  ordinaryFlow = true,
): AdjacentTableSequenceInput {
  return { element, table: element.type === 'table' ? { effectiveStyleId, ordinaryFlow } : null };
}

describe('adjacent table normalization (ECMA-376 Part 1 §17.4.37)', () => {
  it('groups directly adjacent in-flow tables by preserved effective style identity', () => {
    const first = table('SameStyle', true, 36);
    const second = table('SameStyle', true, 144);

    expect(normalizeAdjacentTables([
      input(first, 'SameStyle'), input(second, 'SameStyle'),
    ])).toEqual([{
      kind: 'adjacent-table-group',
      effectiveStyleId: 'SameStyle',
      tables: [first, second],
    }]);
  });

  it('keeps different effective style identities distinct even when public geometry matches', () => {
    const first = table('StyleA', true, 72);
    const second = table('StyleB', true, 72);

    expect(normalizeAdjacentTables([input(first, 'StyleA'), input(second, 'StyleB')])).toEqual([
      { kind: 'body-element', element: first },
      { kind: 'body-element', element: second },
    ]);
  });

  it('treats every intervening paragraph as a barrier, including hidden paragraphs', () => {
    const first = table('SameStyle', true, 72);
    const hidden = paragraph(true);
    const second = table('SameStyle', true, 72);

    expect(normalizeAdjacentTables([
      input(first, 'SameStyle'), input(hidden), input(second, 'SameStyle'),
    ])).toEqual([
      { kind: 'body-element', element: first },
      { kind: 'body-element', element: hidden },
      { kind: 'body-element', element: second },
    ]);
  });

  it('does not join across an effectively floating table', () => {
    const first = table('SameStyle', true, 72);
    const floating = table('SameStyle', false, 72);
    const second = table('SameStyle', true, 72);

    expect(normalizeAdjacentTables([
      input(first, 'SameStyle'), input(floating, 'SameStyle', false), input(second, 'SameStyle'),
    ])).toEqual([
      { kind: 'body-element', element: first },
      { kind: 'body-element', element: floating },
      { kind: 'body-element', element: second },
    ]);
  });

  it('keeps hand-built public tables distinct when no parser style identity was preserved', () => {
    const first = publicTable();
    const second = publicTable();

    expect(normalizeAdjacentTables([input(first), input(second)])).toEqual([
      { kind: 'body-element', element: first },
      { kind: 'body-element', element: second },
    ]);
  });
});
