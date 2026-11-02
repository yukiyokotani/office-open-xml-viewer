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

/** Index note-part entries by their OOXML id for story acquisition. */
export function indexNotes<T extends Readonly<{ id: string }>>(
  notes: readonly T[] | undefined,
): Map<string, T> {
  const indexed = new Map<string, T>();
  if (!notes) return indexed;
  for (const note of notes) indexed.set(note.id, note);
  return indexed;
}

/** Footnote ids the main document story references more than once. A
 * recovery permission is granted per id, so it reaches every occurrence
 * (#2113, `notePublicationPageLimit`). */
export function repeatedFootnoteReferenceIds(elements: readonly NoteReferenceSourceBlock[]): readonly string[] {
  const counts = new Map<string, number>();
  const visit = (blocks: readonly NoteReferenceSourceBlock[]) => {
    for (const element of blocks) {
      if (element.type === 'paragraph' && 'runs' in element) {
        for (const run of element.runs) {
          if (run.type === 'text' && run.noteRef?.kind === 'footnote' && run.noteRef.id.length > 0) {
            counts.set(run.noteRef.id, (counts.get(run.noteRef.id) ?? 0) + 1);
          }
        }
      } else if (element.type === 'table' && 'rows' in element) {
        for (const row of element.rows) for (const cell of row.cells) visit(cell.content);
      }
    }
  };
  visit(elements);
  return Object.freeze([...counts].filter(([, count]) => count > 1).map(([id]) => id));
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
