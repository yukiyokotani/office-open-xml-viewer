import {
  tableAcquisitionInput,
  tableParticipatesInOrdinaryFlow,
} from '../parser-model.js';
import type { BodyElement } from '../types.js';

export type TableBodyElement = Extract<BodyElement, { type: 'table' }>;

export type NormalizedBodySequenceEntry =
  | Readonly<{ kind: 'body-element'; element: BodyElement }>
  | Readonly<{
    kind: 'adjacent-table-group';
    effectiveStyleId: string;
    tables: readonly TableBodyElement[];
  }>;

function groupingIdentity(element: BodyElement): string | null {
  if (element.type !== 'table') return null;
  const facts = tableAcquisitionInput(element);
  if (!tableParticipatesInOrdinaryFlow(element)) return null;
  return facts.table?.effectiveStyleId ?? null;
}

/**
 * Normalize the body sequence rule from ECMA-376 Part 1 §17.4.37 without
 * comparing derived table geometry. The standard keys adjacency by table style;
 * [MS-OI29500] 2.1.149(a) keeps an effectively positioned table distinct.
 * Parser-private facts are required so a lexical tblpPr that Word ignores is
 * not mistaken for effective floating placement.
 */
export function normalizeAdjacentTables(
  body: readonly BodyElement[],
): readonly NormalizedBodySequenceEntry[] {
  const result: NormalizedBodySequenceEntry[] = [];
  let pendingStyleId: string | null = null;
  let pendingTables: TableBodyElement[] = [];

  const flush = () => {
    if (pendingTables.length === 1) {
      result.push(Object.freeze({ kind: 'body-element', element: pendingTables[0]! }));
    } else if (pendingTables.length > 1) {
      result.push(Object.freeze({
        kind: 'adjacent-table-group',
        effectiveStyleId: pendingStyleId!,
        tables: Object.freeze(pendingTables.slice()),
      }));
    }
    pendingStyleId = null;
    pendingTables = [];
  };

  for (const element of body) {
    const styleId = groupingIdentity(element);
    if (styleId !== null && element.type === 'table') {
      if (pendingTables.length > 0 && pendingStyleId !== styleId) flush();
      pendingStyleId = styleId;
      pendingTables.push(element);
      continue;
    }

    flush();
    result.push(Object.freeze({ kind: 'body-element', element }));
  }
  flush();

  return Object.freeze(result);
}
