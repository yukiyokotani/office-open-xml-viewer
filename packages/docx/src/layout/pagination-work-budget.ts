import type { LayoutSourceStore } from './layout-source-store.js';
import type { SourceRef } from './types.js';
import { sourceKey } from './source-key.js';
import { LayoutInvariantError } from './diagnostics.js';
import { MAX_BODY_LAYOUT_PAGES } from './resource-budgets.js';

export const PARAGRAPH_ACQUISITION_MISS_BUDGET = 25_000;
export const MAX_PAGINATION_ACQUISITION_MISSES = 1_000_000;

export interface PaginationWorkBudget {
  readonly limit: number;
  readonly misses: number;
  noteMiss(): void;
  admitPageStories(owner: string, pageIndex: number, sources: readonly (SourceRef | null)[]): void;
}

const saturatingAdd = (left: number, right: number): number =>
  Math.min(MAX_PAGINATION_ACQUISITION_MISSES, left + right);

interface StoryWork {
  readonly story: SourceRef['story'];
  readonly own: number;
  readonly children: readonly SourceRef[];
  total?: number;
}

/** Count each repository definition once, independently of where it is used.
 * Source occurrences (including aliased paragraph objects) each cost work;
 * object identity must not deduplicate them. Complete text-box content is a
 * detached SourceRef, whereas compatibility text is inline. Follow both through
 * arbitrary cells/boxes, not just root runs. These are accounting rules for all
 * accepted source graphs, not a conformance claim for nested txbxContent
 * (ECMA-376 Part 1 §20.4.2.38 restricts that content).
 *
 * The iterative postorder memoizes transitive work per story root. Repeated
 * references count their occurrences without expanding a shared descendant
 * graph repeatedly. Unit arithmetic and each local traversal saturate at HARD;
 * the memo and postorder stack are linear in source definitions. A cycle of
 * detached references is invalid input, not a source of recurring credit. */
function sourceWorkCounter(source: LayoutSourceStore) {
  const definitions = new Map<string, StoryWork>();
  const local = (ref: SourceRef): StoryWork => {
    const key = sourceKey(ref);
    const prior = definitions.get(key);
    if (prior) return prior;
    let own = 1;
    let visited = 0;
    const children: SourceRef[] = [];
    const pending: object[] = [source.blocks.storyRoot(ref)];
    while (pending.length && own < MAX_PAGINATION_ACQUISITION_MISSES) {
      if (++visited >= MAX_PAGINATION_ACQUISITION_MISSES) {
        own = MAX_PAGINATION_ACQUISITION_MISSES;
        break;
      }
      const record = pending.pop()! as Record<string, unknown>;
      if (record.type === 'paragraph') own = saturatingAdd(own, 1);
      if (Array.isArray(record.cells)) own = saturatingAdd(own, record.cells.length);
      if (record.kind === 'compatibility' && Array.isArray(record.paragraphs)) {
        own = saturatingAdd(own, saturatingAdd(1, record.paragraphs.length));
      }
      if (record.kind === 'complete' && record.source) {
        children.push(record.source as SourceRef);
      }
      for (const child of Object.values(record)) {
        if (child && typeof child === 'object') {
          if (pending.length >= MAX_PAGINATION_ACQUISITION_MISSES) {
            own = MAX_PAGINATION_ACQUISITION_MISSES;
            break;
          }
          pending.push(child);
        }
      }
    }
    const result: StoryWork = { story: ref.story, own, children };
    definitions.set(key, result);
    return result;
  };
  const total = (ref: SourceRef): number => {
    const root = local(ref);
    const active = new Set<StoryWork>([root]);
    const pending = [{ entry: root, index: 0, units: root.own }];
    while (pending.length) {
      const frame = pending[pending.length - 1]!;
      if (frame.entry.total !== undefined || frame.index === frame.entry.children.length
        || frame.units === MAX_PAGINATION_ACQUISITION_MISSES) {
        frame.entry.total ??= frame.units;
        active.delete(frame.entry);
        pending.pop();
        continue;
      }
      const child = local(frame.entry.children[frame.index]!);
      if (child.total !== undefined) {
        frame.units = saturatingAdd(frame.units, child.total);
        frame.index++;
      } else {
        if (active.has(child)) throw new TypeError('Cyclic complete text-box story source');
        active.add(child);
        pending.push({ entry: child, index: 0, units: child.own });
      }
    }
    return root.total!;
  };
  return { local, total, definitions };
}

/** Library resource policy, not an OOXML rule: allow 32 acquisitions per source
 * unit for reserve/field/anchor solves, plus 32 per selected page-story unit.
 * This is an operational allowance rather than a promise that every document
 * below the page ceiling converges. A single execution can never exceed one
 * million misses, including nested solves and all service views:
 * min(1_000_000, max(25_000, 32 * (sourceUnits + pageStoryUnits))). Source work
 * is counted once; each selected page/owner/root earns its full recursive
 * text-box work once, never once per pass or miss. Numbered notes also receive
 * page credits: page-dependent continued notes can reacquire a whole source
 * on each destination. Separators share the same boundary and deduplication.
 * Root-level deduplication permits separate header/footer acquisition and
 * changed first/even/default selection. The ledger
 * retains neither layouts nor source roots after counting. */
export function createPaginationWorkBudget(source?: LayoutSourceStore): PaginationWorkBudget {
  const weights = new Map<string, number>();
  let units = 0;
  if (source) {
    const work = sourceWorkCounter(source);
    units = work.local({ story: 'body', storyInstance: 'body', path: [] }).own;
    for (const kind of ['footnote', 'endnote'] as const) {
      for (const note of kind === 'footnote' ? source.blocks.footnotes : source.blocks.endnotes) {
        const ref: SourceRef = { story: kind, storyInstance: note.id, path: [] };
        units = saturatingAdd(units, work.local(ref).own);
        weights.set(sourceKey(ref), work.total(ref));
      }
    }
    for (const ref of source.blocks.reservedNoteRoots) {
      units = saturatingAdd(units, work.local(ref).own);
      weights.set(sourceKey(ref), work.total(ref));
    }
    // Body and note text boxes are document-wide definitions, not repeated
    // page stories. Their source paragraphs are acquired through the same
    // ledger. Markup includes all source revisions here, even when hidden;
    // comments only project retained geometry and acquire no layout stories.
    for (const ref of source.blocks.sources) {
      if (ref.story === 'textbox') work.local({ ...ref, path: [] });
    }
    const sections = [source.bodyLayoutInput.initialSection,
      ...source.bodyLayoutInput.sequence.flatMap(entry => entry.kind === 'begin-section' ? [entry.section] : [])];
    for (const section of sections) {
      for (const ref of [...Object.values(section.headers), ...Object.values(section.footers)]) {
        if (ref && !weights.has(sourceKey(ref))) weights.set(sourceKey(ref), work.total(ref));
      }
    }
    // Producer page objects are page stories of their one page.
    for (const object of source.bodyLayoutInput.pageObjects ?? []) {
      if (!weights.has(sourceKey(object.source))) {
        weights.set(sourceKey(object.source), work.total(object.source));
      }
    }
    // Sum local definitions, not transitive totals: shared descendants earn U
    // once, while each selected page story earns its entire closure in W.
    // Include empty detached roots discovered through complete SourceRefs too.
    for (const entry of work.definitions.values()) {
      for (const child of entry.children) work.local(child);
      if (entry.story === 'textbox') units = saturatingAdd(units, entry.own);
    }
  }
  let limit = Math.min(MAX_PAGINATION_ACQUISITION_MISSES,
    Math.max(PARAGRAPH_ACQUISITION_MISS_BUDGET, 32 * units));
  let misses = 0;
  let exhausted = false;
  let pageStoryUnits = 0;
  // Retain each source-owned owner string once, without serializing/copying it
  // into every page credit. Positive credits reach HARD after finitely many
  // admissions; pages with no source work retain no ledger entry. Each page
  // retains only its source-owned selected roots, including slot changes.
  const admitted = new Map<string, Map<number, Set<string>>>();
  return Object.freeze({
    get limit() { return limit; },
    get misses() { return misses; },
    noteMiss() {
      if (exhausted || ++misses > limit) {
        // Provisional publication can catch a failed acquisition. That must
        // never let later page admission revive an already exhausted ledger.
        exhausted = true;
        throw new LayoutInvariantError('NON_CONVERGENCE',
          `layout acquisition exceeded the operational miss budget ${limit}`);
      }
    },
    admitPageStories(owner: string, pageIndex: number, sources: readonly (SourceRef | null)[]) {
      if (exhausted || limit === MAX_PAGINATION_ACQUISITION_MISSES) return;
      if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= MAX_BODY_LAYOUT_PAGES) return;
      let pages = admitted.get(owner);
      let roots = pages?.get(pageIndex);
      let work = 0;
      for (const ref of sources) {
        if (!ref) continue;
        const key = sourceKey(ref);
        const weight = weights.get(key) ?? 0;
        if (weight === 0 || roots?.has(key)) continue;
        if (!pages) admitted.set(owner, pages = new Map());
        if (!roots) pages.set(pageIndex, roots = new Set());
        roots.add(key);
        work = saturatingAdd(work, weight);
      }
      if (work === 0) return;
      pageStoryUnits = saturatingAdd(pageStoryUnits, work);
      limit = Math.min(MAX_PAGINATION_ACQUISITION_MISSES,
        Math.max(PARAGRAPH_ACQUISITION_MISS_BUDGET, 32 * (units + pageStoryUnits)));
      if (limit === MAX_PAGINATION_ACQUISITION_MISSES) admitted.clear();
    },
  });
}
