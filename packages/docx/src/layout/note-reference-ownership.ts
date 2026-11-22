import type { LineLayout, ParagraphLayout, TableLayout } from './types.js';

type NoteReferenceSourceBlock =
  | Readonly<{
      type: 'paragraph';
      runs: readonly Readonly<{
        type: string;
        noteRef?: Readonly<{ kind: 'footnote' | 'endnote'; id: string; customMarkFollows?: boolean }>;
      }>[];
    }>
  | Readonly<{
      type: 'table';
      rows: readonly Readonly<{
        cells: readonly Readonly<{ content: readonly NoteReferenceSourceBlock[] }>[];
      }>[];
    }>
  | Readonly<{ type: string }>;

/** Build default sequential numbering from first-reference order. Unreferenced
 * note-part entries do not consume a displayed number (§17.18.22/.34). */
export function buildNoteNumberMap(
  notes: readonly Readonly<{ id: string }>[] | undefined,
  referenceIds: readonly string[],
  customReferenceIds: ReadonlySet<string> = new Set(),
): Map<string, number> {
  const numbers = new Map<string, number>();
  if (!notes) return numbers;
  const available = new Set(notes.map((note) => note.id));
  let ordinal = 0;
  referenceIds.forEach((id) => {
    // CT_FtnEdnRef/@customMarkFollows: custom notes remain owned/indexed but
    // never consume an automatic ordinal. Zero suppresses in-note *Ref ink.
    if (available.has(id) && !numbers.has(id)) numbers.set(id, customReferenceIds.has(id) ? 0 : ++ordinal);
  });
  return numbers;
}

/** One section's numbering rule for a note kind (§17.11.17-.20). */
export interface SectionNoteNumberingRule {
  readonly format: string;
  readonly start: number;
  readonly restart?: 'continuous' | 'eachSect' | 'eachPage';
}

/** Display values and formats of automatic notes in first-reference order.
 * Each reference takes the rule of the section that holds it (§17.11.5/.11
 * over §17.11.4/.12): the first automatic note starts at numStart, and a
 * section whose rule restarts each section (§17.11.19 eachSect) restarts
 * its first note there. A custom mark (customMarkFollows) keeps the 0
 * sentinel and consumes no value. Top-level sectionBreak elements end
 * sections. */
export function buildSectionedNoteNumbers(
  notes: readonly Readonly<{ id: string }>[] | undefined,
  elements: readonly NoteReferenceSourceBlock[],
  kind: 'footnote' | 'endnote',
  ruleFor: (sectionIndex: number) => SectionNoteNumberingRule,
  customReferenceIds: ReadonlySet<string> = new Set(noteReferenceIdsInDocumentOrder(elements, kind, true)),
  /** §17.11.19 eachPage: the page of each reference; absent before the
   * first pagination iteration, when eachPage numbers continue. */
  referencePage?: (id: string) => number | undefined,
): Readonly<{ values: Map<string, number>; formats: Map<string, string> }> {
  const values = new Map<string, number>();
  const formats = new Map<string, string>();
  if (!notes) return { values, formats };
  const available = new Set(notes.map((note) => note.id));
  let previous: number | undefined;
  let previousSection = -1;
  let previousPage: number | undefined;
  let sectionIndex = 0;
  const visitRuns = (runs: Extract<NoteReferenceSourceBlock, { runs: unknown }>['runs']) => {
    for (const run of runs) {
      const ref = run.type === 'text' ? run.noteRef : undefined;
      if (ref?.kind !== kind || ref.id.length === 0 || !available.has(ref.id) || values.has(ref.id)) continue;
      if (customReferenceIds.has(ref.id)) {
        values.set(ref.id, 0);
        continue;
      }
      const rule = ruleFor(sectionIndex);
      const page = referencePage?.(ref.id);
      const restart = previous === undefined
        || (rule.restart === 'eachSect' && sectionIndex !== previousSection)
        // An eachPage section also restarts where it begins, even on a page
        // it shares with the previous section (word-section-note-numbering).
        || (rule.restart === 'eachPage' && (sectionIndex !== previousSection
          || (page !== undefined && page !== previousPage)));
      const value = restart ? rule.start : previous! + 1;
      values.set(ref.id, value);
      formats.set(ref.id, rule.format);
      previous = value;
      previousSection = sectionIndex;
      previousPage = page;
    }
  };
  const visit = (blocks: readonly NoteReferenceSourceBlock[], topLevel: boolean) => {
    for (const element of blocks) {
      if (element.type === 'paragraph' && 'runs' in element) visitRuns(element.runs);
      else if (element.type === 'table' && 'rows' in element) {
        for (const row of element.rows) for (const cell of row.cells) visit(cell.content, false);
      } else if (topLevel && element.type === 'sectionBreak') sectionIndex += 1;
    }
  };
  visit(elements, true);
  return { values, formats };
}

/** The physical page of each footnote reference of a laid-out document, in
 * page order (the first page a reference appears on), for §17.11.19 eachPage
 * numbering. Body, table-cell and note-story references all count. */
export function footnoteReferencePages(
  layout: Readonly<{ pages: readonly Readonly<{ pageIndex: number; layers: Readonly<{ body: readonly unknown[] }> }>[] }>,
): ReadonlyMap<string, number> {
  const pages = new Map<string, number>();
  const visit = (node: unknown, pageIndex: number): void => {
    if (node === null || typeof node !== 'object') return;
    if (Array.isArray(node)) {
      for (const child of node) visit(child, pageIndex);
      return;
    }
    const record = node as Readonly<Record<string, unknown>>;
    const reference = record.noteReference as Readonly<{ kind?: unknown; id?: unknown }> | undefined;
    if ((record.kind === 'text' || record.kind === 'anchor-host') && reference?.kind === 'footnote'
      && typeof reference.id === 'string' && reference.id.length > 0 && !pages.has(reference.id)) {
      pages.set(reference.id, pageIndex);
    }
    for (const key of ['lines', 'placements', 'rows', 'cells', 'blocks', 'layout', 'story', 'textBoxes', 'drawings']) {
      if (key in record) visit(record[key], pageIndex);
    }
  };
  for (const page of layout.pages) visit(page.layers.body, page.pageIndex);
  return pages;
}

/** Index note-part entries by their OOXML id for story acquisition. */
export function indexNotes<T extends Readonly<{ id: string }>>(
  notes: readonly T[] | undefined,
): Map<string, T> {
  const indexed = new Map<string, T>();
  if (!notes) return indexed;
  for (const note of notes) indexed.set(note.id, note);
  return indexed;
}

/** For each footnote id the main document story references more than once,
 * the index of the last top-level element that references it. A recovery
 * permission is granted per id, so it reaches every occurrence (#2113,
 * `notePublicationPageLimit`). */
export function repeatedFootnoteLastReferenceIndexes(
  elements: readonly NoteReferenceSourceBlock[],
): ReadonlyMap<string, number> {
  const counts = new Map<string, number>();
  const last = new Map<string, number>();
  const visit = (blocks: readonly NoteReferenceSourceBlock[], index: number) => {
    for (const element of blocks) {
      if (element.type === 'paragraph' && 'runs' in element) {
        for (const run of element.runs) {
          if (run.type === 'text' && run.noteRef?.kind === 'footnote' && run.noteRef.id.length > 0) {
            counts.set(run.noteRef.id, (counts.get(run.noteRef.id) ?? 0) + 1);
            last.set(run.noteRef.id, index);
          }
        }
      } else if (element.type === 'table' && 'rows' in element) {
        for (const row of element.rows) for (const cell of row.cells) visit(cell.content, index);
      }
    }
  };
  elements.forEach((element, index) => visit([element], index));
  return new Map([...last].filter(([id]) => (counts.get(id) ?? 0) > 1));
}

/** Collect first reference order from the main document story. */
export function noteReferenceIdsInDocumentOrder(
  elements: readonly NoteReferenceSourceBlock[],
  kind: 'footnote' | 'endnote',
  customOnly = false,
): readonly string[] {
  const ids: string[] = [];
  const seen = new Set<string>();
  for (const element of elements) {
    if (element.type === 'paragraph' && 'runs' in element) {
      for (const run of element.runs) {
        if (run.type !== 'text'
          || run.noteRef?.kind !== kind
          || run.noteRef.id.length === 0
          || (customOnly && run.noteRef.customMarkFollows !== true)
          || seen.has(run.noteRef.id)) continue;
        seen.add(run.noteRef.id);
        ids.push(run.noteRef.id);
      }
    } else if (element.type === 'table' && 'rows' in element) {
      for (const row of element.rows) for (const cell of row.cells) {
        for (const id of noteReferenceIdsInDocumentOrder(cell.content, kind, customOnly)) {
          if (seen.has(id)) continue;
          seen.add(id);
          ids.push(id);
        }
      }
    }
  }
  return Object.freeze(ids);
}

function noteIdsInRetainedLines(
  lines: readonly LineLayout[],
  kind: 'footnote' | 'endnote',
): readonly string[] {
  return Object.freeze([...new Set(lines.flatMap((line) => line.placements.flatMap((placement) => (
    (placement.kind === 'text' || placement.kind === 'anchor-host') && placement.noteReference?.kind === kind
      ? [placement.noteReference.id]
      : []
  ))))]);
}

function tableNoteIds(
  table: TableLayout,
  kind: 'footnote' | 'endnote',
): readonly string[] {
  return table.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks.flatMap(
    (block) => noteIdsInRetainedSlice(block.layout, kind),
  )));
}

function noteIdsInRetainedSlice(
  slice: ParagraphLayout | TableLayout,
  kind: 'footnote' | 'endnote',
): readonly string[] {
  const ids = slice.kind === 'paragraph'
    ? noteIdsInRetainedLines(slice.lines, kind)
    : tableNoteIds(slice, kind);
  return Object.freeze([...new Set(ids)]);
}

export function footnoteIdsInRetainedLines(lines: readonly LineLayout[]): readonly string[] {
  return noteIdsInRetainedLines(lines, 'footnote');
}

/** §17.11.21 / §17.18.34 assign the note to the physical page that paints its
 * reference; after splitting, parser-level paragraph/table identity is too broad. */
export function footnoteIdsInRetainedSlice(
  slice: ParagraphLayout | TableLayout,
): readonly string[] {
  return noteIdsInRetainedSlice(slice, 'footnote');
}

/** §17.18.22: unreferenced endnotes are not displayed. */
export function endnoteIdsInRetainedSlice(
  slice: ParagraphLayout | TableLayout,
): readonly string[] {
  return noteIdsInRetainedSlice(slice, 'endnote');
}
