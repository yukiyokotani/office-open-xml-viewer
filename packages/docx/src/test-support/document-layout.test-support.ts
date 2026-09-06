import type { DocumentLayout } from '../layout/types.js';
import type {
  BodyElement,
  DocSettings,
  DocTable,
  DocxDocumentModel,
  SectionProps,
} from '../types.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import type { ResolvedFontMetric } from '@silurus/ooxml-core';

const EMPTY_STORY = Object.freeze({ default: null, first: null, even: null });

export function layoutBodyModel(
  body: readonly BodyElement[],
  section: SectionProps,
  measureContext: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  fontFamilyClasses: Readonly<Record<string, string>> = {},
  settings?: DocSettings,
  fontMetrics?: Readonly<Record<string, ResolvedFontMetric>>,
): DocumentLayout {
  const model: DocxDocumentModel = {
    body: [...body],
    section,
    headers: EMPTY_STORY,
    footers: EMPTY_STORY,
    footnotes: [],
    endnotes: [],
    fontFamilyClasses: { ...fontFamilyClasses },
    ...(settings ? { settings } : {}),
  };
  const services = createLayoutServices(model, { measureContext, fontMetrics });
  return layoutDocument(model, services, { currentDateMs: 0 });
}

/** Exercise the canonical body/table acquisition and return retained row advances. */
export function layoutBodyTableRowAdvances(
  table: DocTable,
  section: SectionProps,
  measureContext: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  settings?: DocSettings,
  fontMetrics?: Readonly<Record<string, ResolvedFontMetric>>,
): readonly number[] {
  const layout = layoutBodyModel(
    [{ type: 'table', ...table } as BodyElement],
    section,
    measureContext,
    {},
    settings,
    fontMetrics,
  );
  const retainedTable = layout.pages
    .flatMap((page) => page.layers.body)
    .find((node) => node.kind === 'table');
  if (!retainedTable || retainedTable.kind !== 'table') {
    throw new Error('Canonical body layout omitted the table');
  }
  return retainedTable.rows.map((row) => row.advancePt);
}
