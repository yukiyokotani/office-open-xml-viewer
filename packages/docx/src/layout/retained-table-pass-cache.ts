import type { BodyAcquisitionState } from './acquisition-context.js';
import { paragraphAcquisitionCacheOf } from './runtime-state.js';
import type { RetainedTableAcquisition } from './table-acquisition.js';
import type { TableLayoutSource } from './table-source-acquisition.js';
import type { LayoutServices } from './types.js';

/**
 * Cross-pass memo for top-level body table acquisitions.
 *
 * One pagination execution can run several complete body passes over the same
 * source (anchor-destination convergence and the
 * WORD_LATER_ANCHOR_EARLIER_LINE_WRAP host-carry pass). Each pass opens a new
 * body layout session, whose per-session `retainedTablesBySourceIndex` starts
 * empty, so every table was acquired again from scratch: column intrinsic
 * widths, every cell paragraph, the table layout and its snapshots.
 *
 * `acquireRetainedTable` is a deterministic function of its table source, the
 * inline extent and the acquisition state it reads. The memo key therefore
 * classifies EVERY {@link BodyAcquisitionState} field (the `Record` below is
 * exhaustive, so a new field fails type checking until it is classified):
 *
 * - `value`: page/location/field facts. They are part of the key by value, so
 *   a pass that acquires the table at a different page, column or geometry
 *   misses and acquires afresh exactly as before.
 * - `identity`: document-scoped objects shared by every pass of one execution
 *   (measurement context, services, settings). A different object misses.
 *   Section contexts are plain data rebuilt per page, so they key by value.
 * - `session`: values each session rebuilds identically from the same sealed
 *   source and options (closures, note-number maps, the session's own table
 *   record map). They cannot differ between passes of one execution. The
 *   complete text-box story acquirer takes its section and page from the
 *   owner state (both keyed here); from the session's body state it reads
 *   only document- and session-constant facts, never its cursor or floats.
 * - `replaced`: fields table acquisition never reads from the outer state
 *   (cell states replace `floats`, `floatParaSeq`, `y` and the page-start
 *   prescan set; `frozenAnchorFrames` is consulted only by body paragraphs).
 *
 * Only acquisitions whose graph is reusable across pages are stored (the same
 * proof the within-session reuse relies on: no PAGE-field text and no anchored
 * drawings), and the markup-view author colour resolver bypasses the memo.
 * The memo lives on the pagination services view, the scope of the paragraph
 * acquisition memo, so it is fresh for every document layout or variant.
 * Retained cell layouts require a stable bidi provider for that pagination
 * execution. After changing the provider, start a new layout; local ordering
 * invalidation does not revise whole retained tables or within-pass records.
 */
type FieldClass = 'value' | 'identity' | 'session' | 'replaced';

const BODY_ACQUISITION_STATE_FIELDS: Readonly<Record<keyof BodyAcquisitionState, FieldClass>> = {
  ctx: 'identity',
  verticalGlyphMeasurement: 'identity',
  acquisitionInputs: 'identity',
  contentX: 'value',
  contentW: 'value',
  y: 'replaced',
  pageH: 'value',
  pageIndex: 'value',
  totalPages: 'value',
  displayPageNumber: 'value',
  pageNumberFormat: 'value',
  marginLeft: 'value',
  marginRight: 'value',
  marginTop: 'value',
  marginBottom: 'value',
  pageWidth: 'value',
  layoutSettings: 'identity',
  sectionLayout: 'value',
  storyContext: 'identity',
  docEastAsian: 'value',
  fontFamilyClasses: 'identity',
  resolvedLocalFonts: 'identity',
  layoutServices: 'identity',
  retainedTableAcquisition: 'session',
  acquireCompleteTextBoxStory: 'session',
  cellHostFlowPageTranslationPt: 'value',
  cellParagraphAnchorReferenceDeltaPt: 'value',
  textBoxStoryHostFrames: 'value',
  retainedTablesBySourceIndex: 'session',
  uprightPhysicalTable: 'value',
  kinsoku: 'identity',
  defaultTabPt: 'value',
  currentDateMs: 'value',
  showTrackedChanges: 'value',
  revisionAuthorColor: 'session',
  noteNumbers: 'session',
  noteNumbering: 'value',
  noteReferenceNumber: 'value',
  containerShading: 'value',
  floats: 'replaced',
  floatParaSeq: 'replaced',
  pageAnchorPrescanned: 'replaced',
  frozenAnchorFrames: 'replaced',
  verticalCJK: 'value',
  verticalAllRotated: 'value',
  verticalPhys: 'value',
};

const FIELD_NAMES = Object.freeze(
  Object.keys(BODY_ACQUISITION_STATE_FIELDS).sort(),
) as readonly (keyof BodyAcquisitionState)[];

/** Retained acquisitions per table source; each keeps its two most recent keys. */
const memos = new WeakMap<LayoutServices, WeakMap<TableLayoutSource, Map<string, RetainedTableAcquisition>>>();
const RETAINED_KEYS_PER_TABLE = 2;

/** The exact memo key for acquiring `sourceIndex` at `contentWidthPt` from
 * `state`, or undefined when the acquisition must not be memoized. */
export function retainedTablePassKey(
  state: BodyAcquisitionState,
  sourceIndex: number,
  contentWidthPt: number,
): string | undefined {
  const services = state.layoutServices;
  const cache = services ? paragraphAcquisitionCacheOf(services) : undefined;
  if (!services || !cache || state.revisionAuthorColor !== undefined) return undefined;
  const parts: unknown[] = [sourceIndex, contentWidthPt];
  for (const field of FIELD_NAMES) {
    const kind = BODY_ACQUISITION_STATE_FIELDS[field];
    if (kind === 'session' || kind === 'replaced') continue;
    const value = state[field];
    if (kind === 'identity') {
      parts.push(value === undefined || value === null
        ? null
        : cache.objectIdentity(value as object));
    } else {
      parts.push(value === undefined ? null : value);
    }
  }
  // Values are plain data (numbers, strings, booleans and small frames).
  return JSON.stringify(parts);
}

export function retainedTableFromPassMemo(
  services: LayoutServices,
  table: TableLayoutSource,
  key: string,
): RetainedTableAcquisition | undefined {
  return memos.get(services)?.get(table)?.get(key);
}

export function retainTableInPassMemo(
  services: LayoutServices,
  table: TableLayoutSource,
  key: string,
  acquisition: RetainedTableAcquisition,
): void {
  let byTable = memos.get(services);
  if (!byTable) {
    byTable = new WeakMap();
    memos.set(services, byTable);
  }
  let byKey = byTable.get(table);
  if (!byKey) {
    byKey = new Map();
    byTable.set(table, byKey);
  }
  byKey.delete(key);
  byKey.set(key, acquisition);
  while (byKey.size > RETAINED_KEYS_PER_TABLE) byKey.delete(byKey.keys().next().value!);
}
