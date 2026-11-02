import { paginateBodySteps } from './body-paginator.js';
import { layoutDocumentProgressively } from './progressive.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import { createHash } from 'node:crypto';
import { describe, expect, it, vi } from 'vitest';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import { noteContinuationModel } from '../testing/note-continuation-model.js';
import type { BodyElement, DocParagraph, DocxDocumentModel } from '../types.js';
import type { DocumentLayout, NoteLayout, PaintNode } from './types.js';
import { endnoteIdsInRetainedSlice, footnoteIdsInRetainedSlice } from './note-reference-ownership.js';
import * as tableLayout from './table.js';
import * as tableAcquisition from './table-acquisition.js';
import * as budgets from './resource-budgets.js';
import * as runtimeState from './runtime-state.js';
import * as noteColumns from './note-column-acquisition.js';
import * as noteBlocks from './note-block-acquisition.js';
import * as pageFactory from './page-factory.js';
import * as bodyPagination from './body-pagination.js';
import { NoteRecoveryProgress } from './note-recovery-progress.js';
import { NOTE_RECOVERY_WORK_MULTIPLIER, NoteOverflowPolicy, NoteOverflowRecoveryRequired } from './note-overflow-policy.js';

type Kind = 'footnote' | 'endnote';
const context = {
  font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
  measureText: (text: string) => ({ width: [...text].length * 3,
    actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }),
} as unknown as CanvasRenderingContext2D;

function paragraph(text: string, height = 10): DocParagraph & { type: 'paragraph' } {
  const base = noteContinuationModel().footnotes![0]!.content[0] as DocParagraph;
  const run = base.runs[0];
  if (run?.type !== 'text') throw new Error('Expected text fixture');
  return { ...base, type: 'paragraph', widowControl: false, spaceBefore: 0, spaceAfter: 0,
    lineSpacing: { value: height, rule: 'exact', explicit: true },
    runs: [{ ...run, text, noteRef: undefined, vertAlign: null }] };
}
function table(content: DocParagraph[], header = false): Extract<BodyElement, { type: 'table' }> {
  const row = (items: DocParagraph[], isHeader = false) => ({
    cells: [{ content: items, colSpan: 1, vMerge: null, borders: {}, background: null,
      vAlign: 'top', widthPt: 40 }],
    rowHeight: null, rowHeightRule: 'auto', isHeader, cantSplit: false,
  });
  return { type: 'table', colWidths: [40], rows: [
    ...(header ? [row([paragraph('header;')], true)] : []), row(content),
  ], borders: {}, cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
  jc: 'left', layout: 'fixed', tblpPr: null } as unknown as Extract<BodyElement, { type: 'table' }>;
}
function document(kind: Kind, content: BodyElement[], mode?: number, columns = 2): DocxDocumentModel {
  const model = noteContinuationModel();
  model.settings = { compatibilityMode: mode };
  model.section.columns = columns === 1 ? null : { count: columns, spacePt: 20, equalWidth: true, sep: false, cols: [] };
  const reference = paragraph('ref;');
  const run = reference.runs[0];
  if (run?.type !== 'text') throw new Error('Expected text fixture');
  run.noteRef = { kind, id: '7' };
  model.body = [reference as BodyElement];
  model.footnotes = kind === 'footnote' ? [{ id: '7', content }] : [];
  model.endnotes = kind === 'endnote' ? [{ id: '7', content }] : [];
  return model;
}
const render = (model: DocxDocumentModel, continuation = true) => layoutDocument(model,
  createLayoutServices(model, { measureContext: context, allowFootnoteContinuation: continuation }), { currentDateMs: 0 });
const notes = (layout: DocumentLayout): NoteLayout[] => layout.pages.flatMap(page => page.layers.notes)
  .filter((node): node is NoteLayout => node.kind === 'note');
const text = (block: PaintNode): string => blockText(block, false);
function blockText(block: PaintNode, sourceOnly: boolean): string {
  if (block.kind === 'paragraph') return block.lines.flatMap(line => line.placements)
    .flatMap(run => run.kind === 'text' ? [run.text] : []).join('');
  if (block.kind === 'table') return block.rows
    .filter(row => !sourceOnly || !('ownership' in row) || row.ownership !== 'repeated-header')
    .flatMap(row => row.cells)
    .flatMap(cell => cell.blocks.map(block => blockText(block.layout, sourceOnly))).join('');
  return '';
}
/** Every committed reference must retain its complete source exactly once,
 * including continuations on later pages. Repeated table headers are decoration. */
function expectConserved(layout: DocumentLayout, expected: Readonly<Record<string, string>>, kind: Kind = 'footnote') {
  const references = layout.pages.flatMap(page => page.layers.body.flatMap(block =>
    block.kind === 'paragraph' || block.kind === 'table'
      ? (kind === 'footnote' ? footnoteIdsInRetainedSlice : endnoteIdsInRetainedSlice)(block) : []));
  expect([...new Set(references)].sort()).toEqual(Object.keys(expected).sort());
  for (const id of references) {
    expect(notes(layout).filter(note => note.source.story === kind && note.source.storyInstance === id)
      .flatMap(note => note.story.blocks.map(block => blockText(block, true))).join('')).toBe(expected[id]);
  }
}

function independentOversizedNotes(count: number) {
  const model = document('footnote', [], undefined, 1);
  const reference = model.body[0] as DocParagraph;
  const run = reference.runs[0]!;
  const expected: Record<string, string> = {};
  model.footnotes = Array.from({ length: count }, (_, index) => {
    const id = String(index + 7);
    const content = Array.from({ length: 2 }, (_, i) => paragraph(`n${id}-${i};`, i === 0 ? 120 : 10));
    expected[id] = content.map((_, i) => `n${id}-${i};`).join('');
    return { id, content };
  });
  model.body = model.footnotes.map(note => ({ ...reference,
    runs: [{ ...run, noteRef: { kind: 'footnote', id: note.id } }],
  })) as BodyElement[];
  return { model, expected };
}

function measuredRecovery(model: DocxDocumentModel, continuation: boolean, verify: (layout: DocumentLayout) => void) {
  let work = 0;
  let sourceUnits = 0;
  let passes = 1;
  const fieldView = runtimeState.createFieldAcquisitionServicesView;
  const fieldSpy = vi.spyOn(runtimeState, 'createFieldAcquisitionServicesView').mockImplementation((...args) => {
    passes++;
    return fieldView(...args);
  });
  const original = budgets.createFootnoteAcquisitionWorkBudget;
  const spy = vi.spyOn(budgets, 'createFootnoteAcquisitionWorkBudget').mockImplementation(() => {
    const budget = original();
    return { ...budget, charge(source) {
      const size = budget.sourceUnits(source);
      if (Array.isArray(source)) sourceUnits = Math.max(sourceUnits, size);
      work += size;
      budget.charge(source);
    } };
  });
  const started = performance.now();
  let outcome = 'complete';
  let layout: DocumentLayout | undefined;
  try { layout = render(model, continuation); }
  catch (error) {
    // These are v0.89's failures for these exact synthetic source classes.
    outcome = model.footnotes!.length ? 'FootnoteAdmissionOverflowError' : 'NoteCapacityExceededError';
    expect(error).toMatchObject({ name: outcome });
  } finally { spy.mockRestore(); fieldSpy.mockRestore(); }
  const ms = Math.round(performance.now() - started);
  if (layout) verify(layout);
  // The note-local 32 sweeps plus ordinary pre-rejection acquisitions in at
  // most four field passes (two whole acquisitions/pass) are all linear in
  // this note, independently of the global 32M execution ceiling.
  expect(work).toBeLessThanOrEqual((NOTE_RECOVERY_WORK_MULTIPLIER + 8) * (sourceUnits + 1));
  return { work, ms, outcome, passes };
}

function assertLinearRecovery(results: readonly { work: number }[]) {
  const current = results.at(-1)!, previous = results.at(-2);
  // Doubling source costs 2x, with slack for destination/pass boundaries and
  // label-width changes; quadratic reacquisition approaches 4x. The original
  // large-N matrices (64-512 tables, 512-2048 fields) needed a fourth field
  // pass and an extra label digit: 2 * 4/3 * 16/15 < 2.9. The old paths
  // exceeded 3.3x and 3.6x respectively. Keep that work bound at CI-sized N.
  if (previous) expect(current.work).toBeLessThanOrEqual(2.9 * previous.work);
}

// SHA-256 of the complete JSON.stringify(DocumentLayout), acquired from an
// isolated clean origin/v0.89 (cb78d8a60119ae78a3e03b126612dafe2f2d8d8a)
// source tree with these exact synthetic models and measurement service. This
// covers IDs, source cuts, fields, ordering and all geometry, not a summary.
const ordinaryLayouts = [
  {"kind": "footnote", "continuation": false, "shape": "unequal-tables", "digest": "4fefc479c58b523bb907037b248a52172a3879836517a4542926004a96180806"},
  {"kind": "footnote", "continuation": false, "shape": "paragraph-balance", "digest": "1066b4d932565e4dcd66c3d37fd1afeb6968ecb3818172f5aa3e1cb3ac752ecd"},
  {"kind": "footnote", "continuation": false, "shape": "fitting-table", "digest": "6299910a1ea7dd2c1895da21471592e7022bc86971de6c8670e99c54d6f05109"},
  {"kind": "footnote", "continuation": true, "shape": "unequal-tables", "digest": "4fefc479c58b523bb907037b248a52172a3879836517a4542926004a96180806"},
  {"kind": "footnote", "continuation": true, "shape": "paragraph-balance", "digest": "1066b4d932565e4dcd66c3d37fd1afeb6968ecb3818172f5aa3e1cb3ac752ecd"},
  {"kind": "footnote", "continuation": true, "shape": "fitting-table", "digest": "6299910a1ea7dd2c1895da21471592e7022bc86971de6c8670e99c54d6f05109"},
  {"kind": "footnote", "continuation": true, "shape": "text-continuation", "digest": "7d20682eb418137707eebde5a71146602ae65ceee30c6921caaf2578ae379341"},
  {"kind": "endnote", "continuation": false, "shape": "unequal-tables", "digest": "8b4911e12ccda128c4c2f3d41540e30c4f9efa80f6fc50771e2a6e10547f2c89"},
  {"kind": "endnote", "continuation": false, "shape": "paragraph-balance", "digest": "a35e290ffdc3624825268111ea1b5ef5527e57b9c233fd675894ccabe1262aab"},
  {"kind": "endnote", "continuation": false, "shape": "fitting-table", "digest": "4bf15267e202da14e518c5c4c8a0a8b2dcf2b3003ba417660efda8dbe0700146"},
  {"kind": "endnote", "continuation": false, "shape": "text-continuation", "digest": "934a6aa5d2f953d51fd253b82a9e94057dc6062f549c887c2bfeed49ee80c8c3"},
  {"kind": "endnote", "continuation": false, "shape": "three-heads", "digest": "c927ddc7a87d54ebbdf1ae499d68bc6e558df1f36f355a056ec74966507826ab"},
  {"kind": "endnote", "continuation": true, "shape": "unequal-tables", "digest": "8b4911e12ccda128c4c2f3d41540e30c4f9efa80f6fc50771e2a6e10547f2c89"},
  {"kind": "endnote", "continuation": true, "shape": "paragraph-balance", "digest": "a35e290ffdc3624825268111ea1b5ef5527e57b9c233fd675894ccabe1262aab"},
  {"kind": "endnote", "continuation": true, "shape": "fitting-table", "digest": "4bf15267e202da14e518c5c4c8a0a8b2dcf2b3003ba417660efda8dbe0700146"},
  {"kind": "endnote", "continuation": true, "shape": "text-continuation", "digest": "934a6aa5d2f953d51fd253b82a9e94057dc6062f549c887c2bfeed49ee80c8c3"},
  {"kind": "endnote", "continuation": true, "shape": "three-heads", "digest": "c927ddc7a87d54ebbdf1ae499d68bc6e558df1f36f355a056ec74966507826ab"},
] as const;

describe('reviewed note admission regressions', () => {
  it('keeps consecutive oversized table notes clear of their reference body', () => {
    const model = document('footnote', [], undefined, 1);
    model.section.pageHeight = 400;
    model.section.headerDistance = 0;
    model.headers.default = { body: [paragraph('header;', 120) as BodyElement] };
    // Each note is 16 * (6pt label + 12pt table) plus a 6pt band =
    // 294pt. The header leaves 270pt. The seed pass fits whole notes in
    // 380pt, then reserve convergence splits the first note, moving the
    // second reference beyond the seed's page set. That unseen fresh page
    // must use its real header band instead of zero reserve.
    const expected: Record<string, string> = {};
    model.footnotes = ['7', '8'].map(id => {
      const content = Array.from({ length: 16 }, (_, index) => {
        const grid = table([paragraph(`${id}a${index};`, 12)]);
        grid.colWidths = [40, 40];
        grid.rows[0]!.cells.push({ ...grid.rows[0]!.cells[0]!, content: [paragraph(`${id}b${index};`, 12)] });
        return [paragraph(`${id}label${index};`, 6), grid];
      }).flat();
      expected[id] = Array.from({ length: 16 }, (_, index) =>
        `${id}label${index};${id}a${index};${id}b${index};`).join('');
      return { id, content };
    });
    model.body = ['7', '8'].map((id, index) => {
      const reference = paragraph(`body${id};`, 14);
      reference.pageBreakBefore = index > 0;
      const run = reference.runs[0];
      if (run?.type !== 'text') throw new Error('Expected reference text');
      reference.runs = [{ ...run, noteRef: { kind: 'footnote', id } },
        ...paragraph('following;').runs];
      return reference;
    }) as BodyElement[];
    // layoutDocument runs the same invariant validation/freezing as the
    // browser; an overlapping retained result must fail here.
    const layout = render(model);
    expectConserved(layout, expected);
    for (const id of ['7', '8']) expect(notes(layout).filter(note => note.source.storyInstance === id).length)
      .toBeGreaterThan(1);
    expect(layout.pages.flatMap(page => page.layers.body.map(text)).join('')).toBe('1following;2following;');
  });

  it('debits completed-page copies against the recovered tail allowance', () => {
    // The original 8192-page tail reached the allowance. Count charged copies
    // directly so a missing debit is also detectable with a short tail.
    const count = 16;
    const model = document('footnote', Array.from({ length: count }, () => paragraph('x;', 120)), undefined, 1);
    let originalFailure: Error | undefined;
    let reads = 0, limit = 0, chargedCopies = 0;
    let copying = false, exhaustCopy = false;
    let readsAtExhaustion: number | undefined;
    const copy = NoteRecoveryProgress.prototype.chargePageSequenceCopy;
    const copies = vi.spyOn(NoteRecoveryProgress.prototype, 'chargePageSequenceCopy').mockImplementation(function (this: NoteRecoveryProgress, count) {
      copying = true;
      try { return copy.call(this, count); } finally { copying = false; }
    });
    const recover = NoteOverflowPolicy.prototype.recover;
    const recovery = vi.spyOn(NoteOverflowPolicy.prototype, 'recover').mockImplementation(function (this: NoteOverflowPolicy, error) {
      if (error instanceof NoteOverflowRecoveryRequired) originalFailure ??= error.failure;
      return recover.call(this, error);
    });
    const charge = NoteOverflowPolicy.prototype.charge;
    const ledger = vi.spyOn(NoteOverflowPolicy.prototype, 'charge').mockImplementation(function (this: NoteOverflowPolicy, kind, id, sourceUnits, units) {
      if (this.allows(kind, [id])) {
        limit = NOTE_RECOVERY_WORK_MULTIPLIER * (sourceUnits + 1);
        if (copying) {
          chargedCopies += units;
          if (exhaustCopy) {
            readsAtExhaustion = reads;
            // Simulate an allowance already consumed by earlier recovery
            // trials, without constructing thousands of completed pages.
            charge.call(this, kind, id, sourceUnits, limit);
          }
        }
      }
      return charge.call(this, kind, id, sourceUnits, units);
    });
    const commit = bodyPagination.commitPageFlowTransition;
    const pages = vi.spyOn(bodyPagination, 'commitPageFlowTransition').mockImplementation((state, ...args) => {
      // Observe real completed-page reads at the reducer boundary. This also
      // permits a future persistent sequence to finish with linear work.
      const observed = new Proxy(state.pages, { get(target, key, receiver) {
        if (typeof key === 'string' && /^(0|[1-9][0-9]*)$/.test(key)) reads++;
        return Reflect.get(target, key, receiver);
      } });
      return commit({ ...state, pages: observed }, ...args);
    });
    try {
      let result: DocumentLayout | undefined;
      try { result = render(model); }
      catch (error) {
        expect(originalFailure).toMatchObject({ name: 'FootnoteAdmissionOverflowError' });
        expect(error).toBe(originalFailure);
      }
      if (result) expectConserved(result, { '7': 'x;'.repeat(count) });
      expect(limit).toBeGreaterThan(0);
      expect(reads).toBeGreaterThan(count);
      expect(reads).toBeLessThanOrEqual(chargedCopies);
      expect(reads).toBeLessThanOrEqual(limit);

      originalFailure = undefined;
      exhaustCopy = true;
      let failure: unknown;
      try { render(model); } catch (error) { failure = error; }
      expect(originalFailure).toMatchObject({ name: 'FootnoteAdmissionOverflowError' });
      expect(failure).toBe(originalFailure);
      expect(readsAtExhaustion).toBeDefined();
      expect(reads).toBe(readsAtExhaustion);
    } finally { pages.mockRestore(); ledger.mockRestore(); recovery.mockRestore(); copies.mockRestore(); }
  });

  it('preserves the captured rejection when recovery finalization throws', () => {
    let originalFailure: Error | undefined;
    const recover = NoteOverflowPolicy.prototype.recover;
    const policy = vi.spyOn(NoteOverflowPolicy.prototype, 'recover').mockImplementation(function (this: NoteOverflowPolicy, error) {
      if (error instanceof NoteOverflowRecoveryRequired) originalFailure ??= error.failure;
      return recover.call(this, error);
    });
    const finalize = pageFactory.finalizeLayoutPage;
    let injected = 0;
    const finalizer = vi.spyOn(pageFactory, 'finalizeLayoutPage').mockImplementation((...args) => {
      if (originalFailure) { injected++; throw new RangeError('Injected finalization failure'); }
      return finalize(...args);
    });
    try {
      let failure: unknown;
      try { render(document('footnote', [paragraph('tall;', 120), paragraph('tail;')], 15), false); }
      catch (error) { failure = error; }
      expect(injected).toBe(1);
      expect(originalFailure).toMatchObject({ name: 'FootnoteAdmissionOverflowError' });
      expect(failure).toBe(originalFailure);
    } finally { finalizer.mockRestore(); policy.mockRestore(); }
  });

  it('preserves admission geometry when recovered columns fail finalization', () => {
    const model = document('footnote', Array.from({ length: 9 }, () => paragraph('n;')), 11, 3);
    model.section.columns = { count: 3, spacePt: 10, equalWidth: false, sep: false,
      cols: [40, 60, 80].map(widthPt => ({ widthPt, spacePt: 10 })) };
    model.body = [paragraph('before;'), paragraph('before;'), ...model.body] as BodyElement[];
    expect(() => render(model, false)).toThrow(expect.objectContaining({
      name: 'FootnoteAdmissionOverflowError', reservePt: 96, admissionChargePt: 106, freshPageExtentPt: 80,
    }));
  });

  it('charges every repeated table materialization before an inner recovery loop can fan out', () => {
    let originalFailure: Error | undefined;
    let attempts = 0;
    const recover = NoteOverflowPolicy.prototype.recover;
    const policy = vi.spyOn(NoteOverflowPolicy.prototype, 'recover').mockImplementation(function (this: NoteOverflowPolicy, error) {
      if (error instanceof NoteOverflowRecoveryRequired) originalFailure ??= error.failure;
      return recover.call(this, error);
    });
    const fragment = tableAcquisition.layoutTableFragment;
    const materialize = vi.spyOn(tableAcquisition, 'layoutTableFragment').mockImplementation((...args) => {
      if (!originalFailure) return fragment(...args);
      // A future inner solver can repeat the same immutable acquisition on
      // one destination without changing a page or source cursor. Cache hits
      // in the size memo must still pay for every real materialization.
      for (let trial = 0; trial < 100_000; trial++) { attempts++; fragment(...args); }
      throw new Error('Recovery failed to bound table materialization');
    });
    try {
      let failure: unknown;
      try { render(document('footnote', [table([paragraph('tall;', 120)])], 15), false); }
      catch (error) { failure = error; }
      expect(originalFailure).toMatchObject({ name: 'FootnoteAdmissionOverflowError' });
      expect(failure).toBe(originalFailure);
      expect(attempts).toBeGreaterThan(1);
      expect(attempts).toBeLessThan(100);
    } finally { materialize.mockRestore(); policy.mockRestore(); }
  });

  it('bounds repeated-header materialization by the note-local work allowance', () => {
    const results: { n: number; rows: number; work: number; limit: number; ms: number }[] = [];
    // 64 repeated headers already exhaust the note-local allowance; the
    // original 64-512 sweep repeated the same bound at much higher cost.
    for (const n of [64]) {
      const grid = table([paragraph('h;', 1)]);
      grid.rows = Array.from({ length: 2 * n }, (_, index) => ({ ...grid.rows[0]!,
        isHeader: index < n, cells: [{ ...grid.rows[0]!.cells[0]!,
          content: [paragraph(index < n ? 'h;' : 'r;', index < n ? 1 : 120)] }],
      }));
      const model = document('footnote', [grid], 15);
      model.section.columns!.spacePt = 10;
      let originalFailure: Error | undefined;
      let rows = 0, work = 0, limit = 0;
      const recover = NoteOverflowPolicy.prototype.recover;
      const policy = vi.spyOn(NoteOverflowPolicy.prototype, 'recover').mockImplementation(function (this: NoteOverflowPolicy, error) {
        if (error instanceof NoteOverflowRecoveryRequired) originalFailure ??= error.failure;
        return recover.call(this, error);
      });
      const charge = NoteOverflowPolicy.prototype.charge;
      const counter = vi.spyOn(NoteOverflowPolicy.prototype, 'charge').mockImplementation(function (this: NoteOverflowPolicy, kind, id, sourceUnits, units) {
        charge.call(this, kind, id, sourceUnits, units);
        if (this.allows(kind, [id])) { work += units; limit = NOTE_RECOVERY_WORK_MULTIPLIER * (sourceUnits + 1); }
      });
      const layoutTable = tableLayout.layoutTable;
      const materialize = vi.spyOn(tableLayout, 'layoutTable').mockImplementation((...args) => {
        const result = layoutTable(...args);
        if (originalFailure) rows += args[0].rows.length;
        return result;
      });
      const started = performance.now();
      try {
        let failure: unknown;
        try { render(model, false); } catch (error) { failure = error; }
        expect(originalFailure).toMatchObject({ name: 'FootnoteAdmissionOverflowError' });
        expect(failure).toBe(originalFailure);
        expect(work).toBeLessThanOrEqual(limit);
        // A fixed allowance bounds actual materialization, including copies
        // of every header on failed candidates, independently of wall time.
        expect(rows).toBeLessThan(32 * 2 * n);
        results.push({ n, rows, work, limit, ms: Math.round(performance.now() - started) });
      } finally { materialize.mockRestore(); counter.mockRestore(); policy.mockRestore(); }
    }
    process.stdout.write(JSON.stringify({ shape: 'repeated-headers', results }) + '\n');
  });

  it.each((['footnote', 'endnote'] as const).flatMap(kind => [1, 2].flatMap(columns =>
    ['repeated-source', 'empty-fragment'].map(failureMode => ({ kind, columns, failureMode })))))(
    'fails with the original error for $failureMode recovery ($kind, $columns columns)', ({ kind, columns, failureMode }) => {
      const original = noteColumns.acquireNoteColumnPartition;
      let head: ReturnType<typeof original>;
      let rejected = 0;
      const spy = vi.spyOn(noteColumns, 'acquireNoteColumnPartition').mockImplementation((...args) => {
        if (failureMode === 'repeated-source' && head && args[2]) {
          rejected++;
          // Returning already consumed content at the same source cursor
          // must fail even if its translated geometry/ink looks nonempty.
          return { ...head, nextCursor: args[2] };
        }
        const result = original(...args);
        if (failureMode === 'empty-fragment' && result && args[2]) {
          rejected++;
          // A cursor claiming progress cannot substitute for committed
          // content. A fresh page/band must reject this separator-only plan.
          return { ...result, fragment: { ...result.fragment, story: { ...result.fragment.story, blocks: [] } } };
        }
        if (result?.nextCursor) head = result;
        return result;
      });
      try {
        const model = document(kind, [table([paragraph('tall;', 120)]),
          ...Array.from({ length: 20 }, (_, i) => paragraph(`tail${i};`))], 14, columns);
        expect(() => render(model)).toThrow(expect.objectContaining({
          name: kind === 'footnote' ? 'FootnoteAdmissionOverflowError' : 'NoteCapacityExceededError',
        }));
        expect(rejected).toBeGreaterThan(0);
        expect(rejected).toBeLessThanOrEqual(2);
      } finally { spy.mockRestore(); }
    });

  it.each(['footnote', 'endnote'] as const)('preserves the original %s failure if destination reflow is unsupported', kind => {
    const original = noteBlocks.noteBlockAcquirer;
    let rejected = 0;
    const spy = vi.spyOn(noteBlocks, 'noteBlockAcquirer').mockImplementation(note => {
      const acquire = original(note);
      return acquire && ((block, start) => {
        if (start?.paragraph.boundary) { rejected++; throw new Error('Unsupported recovery source cursor'); }
        return acquire(block, start);
      });
    });
    try {
      const p = paragraph('');
      const run = p.runs[0]!;
      if (run.type !== 'text') throw new Error('Expected text fixture');
      p.runs = Array.from({ length: 240 }, () => [
        { ...run, text: 'x=' },
        { ...run, type: 'field' as const, fieldType: 'page', instruction: 'PAGE', fallbackText: '1' },
        { ...run, text: '; ' },
      ]).flat();
      expect(() => render(document(kind, [table([paragraph('tall;', 120)]), p], undefined, 1)))
        .toThrow(expect.objectContaining({
          name: kind === 'footnote' ? 'FootnoteAdmissionOverflowError' : 'NoteCapacityExceededError',
        }));
      expect(rejected).toBe(1);
    } finally { spy.mockRestore(); }
  });

  it('retains a table reference beside its recovered reserve in three columns', () => {
    const model = document('footnote', [paragraph('before;'), paragraph('tall;', 120), paragraph('after;')], 15, 3);
    model.body = [table([model.body[0] as DocParagraph])];
    const layout = render(model, false);
    expectConserved(layout, { '7': 'before;tall;after;' });
    expect(layout.pages.flatMap(page => page.layers.body).map(text).join('')).toBe('1');
    expect(layout.pages.length).toBeLessThanOrEqual(4);
  });

  it.each(['footnote', 'endnote'] as const)('resumes PAGE paragraph source after wrapping changes in a %s', kind => {
    const p = paragraph('');
    const run = p.runs[0]!;
    if (run.type !== 'text') throw new Error('Expected text fixture');
    p.runs = Array.from({ length: 60 }, (_, i) => [
      { ...run, text: `f${i}=` },
      { ...run, type: 'field' as const, fieldType: 'page', instruction: 'PAGE', fallbackText: '1' },
      { ...run, text: '; ' },
    ]).flat();
    const model = document(kind, [table([paragraph('tall;', 120)]), p], undefined, 1);
    model.section.pageNumType = { start: 8, fmt: 'decimal' };
    const layout = render(model);
    const fragments = layout.pages.flatMap(page => page.layers.notes.flatMap(note => {
      if (note.kind !== 'note') return [];
      return note.story.blocks.map(text);
    }));
    const combined = fragments.join('');
    expect(combined.replace(/f(\d+)=\d+; /g, 'f$1;')).toBe('tall;' + Array.from({ length: 60 }, (_, i) => `f${i};`).join(''));
    // A field may start on the last line of a page: verify each field glyph's
    // ownership, independently of whether its surrounding label also fits.
    let fields = 0;
    for (const page of layout.pages) for (const note of page.layers.notes) {
      if (note.kind !== 'note') continue;
      for (const block of note.story.blocks) if (block.kind === 'paragraph') {
        for (const line of block.lines) for (const placement of line.placements) {
          if (placement.kind !== 'text') continue;
          for (const owner of placement.sourceRuns ?? []) {
            if (owner.sourceRunIndex % 3 !== 1) continue;
            const start = Math.max(placement.range.start, owner.range.start) - placement.range.start;
            const end = Math.min(placement.range.end, owner.range.end) - placement.range.start;
            if (end > start) { fields++; expect(placement.text.slice(start, end)).toBe(String(page.pageIndex + 8)); }
          }
        }
      }
    }
    expect(fields).toBe(60);
    expect(layout.pages.length).toBeGreaterThan(2);
  });

  // The largest case still crosses the two-digit page-count boundary.
  it.each([16, 32, 64, 128])('bounds NUMPAGES paragraph acquisition linearly for %i fields', count => {
    const p = paragraph('');
    const run = p.runs[0]!;
    if (run.type !== 'text') throw new Error('Expected text fixture');
    p.runs = Array.from({ length: count }, () => [
      { ...run, text: 'x=' },
      { ...run, type: 'field' as const, fieldType: 'numPages', instruction: 'NUMPAGES', fallbackText: '1' },
      { ...run, text: '; ' },
    ]).flat();
    let work = 0;
    const original = budgets.createFootnoteAcquisitionWorkBudget;
    const spy = vi.spyOn(budgets, 'createFootnoteAcquisitionWorkBudget').mockImplementation(() => {
      const budget = original();
      return { ...budget, charge(source) { work += budget.sourceUnits(source); budget.charge(source); } };
    });
    try {
      const model = document('footnote', [table([paragraph('tall;', 120)]), p], undefined, 1);
      // A 40pt text band reaches two-digit NUMPAGES with only 128 fields.
      model.section.pageWidth = 60;
      const layout = render(model);
      if (count === 128) expect(layout.pages.length).toBeGreaterThanOrEqual(10);
      expectConserved(layout, { '7': 'tall;' + `x=${layout.pages.length}; `.repeat(count) });
      // The fixture contributes 170 source units/field plus <=262 fixed
      // units. Allow four source/block acquisitions in each of up to three
      // page-count passes (including the two-digit NUMPAGES width change):
      // 12*(170N+262) < 2500N for N>=16. Destination pages must add no full
      // paragraph acquisitions; the allowance is independent of page count.
      expect(work).toBeLessThanOrEqual(2500 * count);
    } finally { spy.mockRestore(); }
  });

  it("keeps independent recovery allowances and the execution's first rejection across groups", () => {
    const policy = new NoteOverflowPolicy();
    const first = new Error('first admission');
    policy.recover(new NoteOverflowRecoveryRequired('footnote', ['small'], first));
    policy.recover(new NoteOverflowRecoveryRequired('endnote', ['large'], new Error('later admission')));
    expect(policy.failure('endnote', ['large'])).toBe(first);
    policy.charge('footnote', 'small', 1, NOTE_RECOVERY_WORK_MULTIPLIER * 2);
    policy.charge('endnote', 'large', 1000, NOTE_RECOVERY_WORK_MULTIPLIER * 1001);
    // The large note buys no work for the small one. Neither another group's
    // permission nor another source-size lookup resets earlier debits.
    expect(() => policy.charge('footnote', 'small', 1, 1)).toThrow(first);
    expect(() => policy.charge('endnote', 'large', 1000, 1)).toThrow(first);
  });

  it('resumes a bounded field paragraph inside long original text runs', () => {
    // Each original run exceeds a full 74pt continuation band, so both
    // sides of PAGE resume inside a run rather than just at run boundaries.
    const before = 'word '.repeat(96) + 'page=';
    const after = ';' + ' tail'.repeat(96);
    const p = paragraph(before);
    const run = p.runs[0]!;
    if (run.type !== 'text') throw new Error('Expected text fixture');
    p.runs.push({ ...run, type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '1' },
      { ...run, text: after });
    const model = document('footnote', [table([paragraph('tall;', 120)]), p], 15, 1);
    model.section.pageNumType = { start: 99, fmt: 'decimal' };
    const layout = render(model);
    const owner = layout.pages.find(page => page.layers.notes.some(note => note.kind === 'note'
      && note.story.blocks.some(block => text(block).includes('page='))))!;
    expectConserved(layout, { '7': `tall;${before}${owner.pageIndex + 99}${after}` });
    expect(layout.pages.length).toBeGreaterThan(3);
  });

  it.each(['footnote', 'endnote'] as const)('stops repeated %s refresh trials at the note-local work allowance', kind => {
    const p = paragraph('page=');
    const run = p.runs[0]!;
    if (run.type !== 'text') throw new Error('Expected text fixture');
    p.runs.push({ ...run, type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '1' });
    let originalFailure: Error | undefined;
    let attempts = 0;
    const recover = NoteOverflowPolicy.prototype.recover;
    const policy = vi.spyOn(NoteOverflowPolicy.prototype, 'recover').mockImplementation(function (this: NoteOverflowPolicy, error) {
      if (error instanceof NoteOverflowRecoveryRequired) originalFailure ??= error.failure;
      return recover.call(this, error);
    });
    const original = noteBlocks.noteBlockAcquirer;
    const refresh = vi.spyOn(noteBlocks, 'noteBlockAcquirer').mockImplementation(note => {
      const acquire = original(note);
      return acquire && ((block, start, capacity, maximumLines) => {
        if (block.kind !== 'paragraph') return acquire(block, start, capacity, maximumLines);
        // Model a future reserve solver repeatedly trying a field block on
        // one destination. Each capacity misses the latest-only cache, and
        // no progress/physical-page guard can see this inner acquisition loop.
        for (let trial = 0; trial < 1000; trial++) {
          attempts++;
          acquire(block, start, (capacity ?? 80) / (trial + 1), maximumLines);
        }
        throw new Error('Recovery did not stop acquisition fan-out');
      });
    });
    try {
      let failure: unknown;
      try { render(document(kind, [table([paragraph('tall;', 120)]), p], 15, 1)); }
      catch (error) { failure = error; }
      expect(originalFailure).toBeDefined();
      expect(failure).toBe(originalFailure);
      expect(attempts).toBeGreaterThan(1);
      expect(attempts).toBeLessThan(100);
    } finally { refresh.mockRestore(); policy.mockRestore(); }
  });

  it.each(['page', 'numPages'] as const)('bounds continuation-off %s table recovery under doubling', fieldType => {
    const results: { n: number; work: number; ms: number; outcome: string; passes: number }[] = [];
    // Three small doublings exercise recovery without making runner speed
    // part of the contract; the work counter distinguishes 2x from 4x growth.
    for (const n of [8, 16, 32, 64]) {
      const content = Array.from({ length: n }, (_, i) => {
        const p = paragraph(`n${i}=`);
        const run = p.runs[0]!;
        if (run.type !== 'text') throw new Error('Expected text fixture');
        p.runs.push({ ...run, type: 'field', fieldType, instruction: fieldType.toUpperCase(), fallbackText: '1' },
          { ...run, type: 'text', text: ';' });
        return table([p]);
      });
      const model = document('footnote', content, 15, 2);
      model.section.pageNumType = { start: 99, fmt: 'decimal' };
      const result = measuredRecovery(model, false, layout => {
        const retained = layout.pages.flatMap(page => page.layers.notes.flatMap(note => {
          if (note.kind !== 'note') return [];
          const value = fieldType === 'page' ? page.pageIndex + 99 : layout.pages.length;
          const fragments = note.story.blocks.map(text).join('');
          for (const match of fragments.matchAll(/n\d+=(\d+);/g)) expect(Number(match[1])).toBe(value);
          return [...fragments.matchAll(/n(\d+)=\d+;/g)].map(match => Number(match[1]));
        }));
        expect(retained).toEqual(Array.from({ length: n }, (_, i) => i));
      });
      results.push({ n, ...result });
      assertLinearRecovery(results);
    }
    process.stdout.write(JSON.stringify({ shape: 'tables', fieldType, results }) + '\n');
  });

  it.each(['footnote', 'endnote'] as const)('bounds a long PAGE paragraph in %s recovery under doubling', kind => {
    const results: { n: number; work: number; ms: number; outcome: string; passes: number }[] = [];
    // The oversized table forces recovery even at the smallest field count;
    // later doublings carry the paragraph across destination pages.
    for (const n of [16, 32, 64, 128]) {
      const p = paragraph('');
      const run = p.runs[0]!;
      if (run.type !== 'text') throw new Error('Expected text fixture');
      p.runs = Array.from({ length: n }, (_, i) => [
        { ...run, type: 'text' as const, text: `f${i}=` },
        { ...run, type: 'field' as const, fieldType: 'page', instruction: 'PAGE', fallbackText: '1' },
        { ...run, type: 'text' as const, text: '; ' },
      ]).flat();
      const model = document(kind, [table([paragraph('tall;', 120)]), p], 15, 1);
      model.section.pageNumType = { start: 99, fmt: 'decimal' };
      results.push({ n, ...measuredRecovery(model, true, layout => {
        const combined = notes(layout).flatMap(note => note.story.blocks.map(text)).join('');
        expect(combined.replace(/f(\d+)=\d+; /g, 'f$1;')).toBe('tall;' + Array.from({ length: n }, (_, i) => `f${i};`).join(''));
        let fields = 0;
        for (const page of layout.pages) for (const note of page.layers.notes) {
          if (note.kind !== 'note') continue;
          for (const block of note.story.blocks) if (block.kind === 'paragraph') {
            for (const line of block.lines) for (const placement of line.placements) {
              if (placement.kind !== 'text') continue;
              for (const owner of placement.sourceRuns ?? []) {
                if (owner.sourceRunIndex % 3 !== 1) continue;
                const start = Math.max(placement.range.start, owner.range.start) - placement.range.start;
                const end = Math.min(placement.range.end, owner.range.end) - placement.range.start;
                if (end > start) { fields++; expect(placement.text.slice(start, end)).toBe(String(page.pageIndex + 99)); }
              }
            }
          }
        }
        expect(fields).toBe(n);
      }) });
      assertLinearRecovery(results);
    }
    process.stdout.write(JSON.stringify({ shape: 'paragraph', kind, results }) + '\n');
  });

  it.each(['footnote', 'endnote'] as const)('continues oversized %s tables across unequal columns', kind => {
    for (const singleTable of [false, true]) {
      const content = Array.from({ length: 20 }, (_, i) => paragraph(`n${i};`));
      const model = document(kind, singleTable ? [table(content)] : content.flatMap(p => [paragraph('label;'), table([p])]), 15);
      model.section.columns = { count: 2, spacePt: 20, equalWidth: false, sep: false,
        cols: [{ widthPt: 60, spacePt: 20 }, { widthPt: 100, spacePt: 0 }] };
      const layout = render(model);
      expectConserved(layout, { '7': content.map((_, i) => `${singleTable ? '' : 'label;'}n${i};`).join('') }, kind);
      expect(layout.pages.length).toBeLessThanOrEqual(6);
    }
  });

  it.each(['footnote', 'endnote'] as const)('keeps twenty PAGE-field rows atomic in an oversized %s', kind => {
    const rows = Array.from({ length: 20 }, (_, i) => {
      const p = paragraph(`r${i};`);
      const run = p.runs[0];
      if (run?.type !== 'text') throw new Error('Expected text fixture');
      p.runs.push({ ...run, type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '1' });
      return p;
    });
    const t = table([rows[0]!]) as Extract<BodyElement, { type: 'table' }>;
    t.rows = rows.map(p => ({ ...t.rows[0]!, cells: [{ ...t.rows[0]!.cells[0]!, content: [p] }] }));
    const model = document(kind, [t], undefined, 1);
    // Initial acquisition sees page 1; actual ownership after relocation is
    // page 2. Atomic overflow must reacquire there, not reuse PAGE=1 rows.
    model.body = [...Array.from({ length: 7 }, () => paragraph('before;')), ...model.body] as BodyElement[];
    const layout = render(model);
    expect(layout.pages.length).toBeLessThanOrEqual(3);
    expect(notes(layout)).toHaveLength(1);
    const page = layout.pages.find(page => page.layers.notes.some(note => note.kind === 'note'))!;
    expectConserved(layout, { '7': rows.map((_, i) => `r${i};${page.pageIndex + 1}`).join('') }, kind);
    const retained = notes(layout)[0]!.story.blocks[0]!;
    expect(retained.kind === 'table' ? retained.rows.length : 0).toBe(20);
  });

  it.each(['footnote', 'endnote'] as const)('retains oversized table and nested-table units for %s', kind => {
    for (const nested of [false, true]) {
      const content = Array.from({ length: 12 }, (_, i) => paragraph(`n${i};`, i === 3 ? 120 : 10));
      const t = table(content) as Extract<BodyElement, { type: 'table' }>;
      const block = nested ? table([paragraph('unused;')]) as Extract<BodyElement, { type: 'table' }> : t;
      if (nested) block.rows[0]!.cells[0]!.content = [t];
      const layout = render(document(kind, [block], 15));
      expectConserved(layout, { '7': content.map((_, i) => `n${i};`).join('') }, kind);
      expect(layout.pages.length).toBeLessThanOrEqual(6);
    }
  });

  // Multiple completed pages precede each later retry; three doublings
  // distinguish active-page replay from replaying the entire growing prefix.
  it.each([2, 4, 8, 16])('replays only the active page for %i independent oversized notes', count => {
    const { model, expected } = independentOversizedNotes(count);
    const source = layoutSourceStore(model);
    const steps = paginateBodySteps(source.bodyLayoutInput,
      createLayoutServices(source, { measureContext: context, allowFootnoteContinuation: true }), { currentDateMs: 0 });
    let yielded = 0;
    let step = steps.next();
    while (!step.done) {
      // Two-page notes need admission/retry/continuation work in each
      // convergence pass. Seven steps/note leaves boundary slack, but
      // rejects replay of all prior pages by the later doublings.
      expect(++yielded).toBeLessThanOrEqual(7 * count);
      step = steps.next();
    }
    expectConserved(step.value, expected);
    expect(step.value.pages.length).toBeLessThanOrEqual(8 * count);
  });

  it('replays active anchor journals for independent oversized notes in progressive layout', async () => {
    // Twelve references plus the anchor exceed the progressive preview's
    // 12-entry gate while keeping independent retries and journal replay cheap.
    const count = 12;
    const { model, expected } = independentOversizedNotes(count);
    const floating = table([paragraph('anchor;')]) as Extract<BodyElement, { type: 'table' }>;
    floating.tblpPr = { leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: 'margin', horzSpecified: true, vertAnchor: 'page', tblpX: 0, tblpY: 0 };
    model.body.unshift(floating);
    const source = layoutSourceStore(model);
    let yields = 0;
    const previews: DocumentLayout[] = [];
    const layout = await layoutDocumentProgressively(source.bodyLayoutInput,
      createLayoutServices(source, { measureContext: context, allowFootnoteContinuation: true }),
      { currentDateMs: 0 }, {
        onPreview: preview => previews.push(preview.layout),
        scheduler: { sliceMs: 0, yieldToHost: async () => { yields++; } },
      });
    expectConserved(layout, expected);
    expect(layout.pages.length).toBeLessThanOrEqual(8 * count);
    expect(yields).toBeLessThanOrEqual(100 * count);
    // Every reference sits on page 0 and each note recovers, so #2113 may
    // withhold every publication; whatever is published stays final.
    const counts = previews.map(preview => preview.pages.length);
    expect(counts).toEqual([...new Set(counts)].sort((a, b) => a - b));
    for (const preview of previews) preview.pages.forEach((page, index) => {
      expect(page.layers.body).toEqual(layout.pages[index]!.layers.body);
      expect(page.layers.notes).toEqual(layout.pages[index]!.layers.notes);
    });
  });

  it.each([1, 2])('recovers an oversized footnote tail while endnotes open a page (%i columns)', columns => {
    const model = document('footnote', [paragraph('before;'), paragraph('tall;', 120), paragraph('after;')], columns === 2 ? 15 : undefined, columns);
    const reference = model.body[0] as DocParagraph;
    const run = reference.runs[0]!;
    reference.runs.push({ ...run, noteRef: { kind: 'endnote', id: '8' } } as DocParagraph['runs'][number]);
    model.endnotes = [{ id: '8', content: Array.from({ length: 20 }, (_, i) => paragraph(`e${i};`)) }];
    const layout = render(model);
    expectConserved(layout, { '7': 'before;tall;after;' });
    expectConserved(layout, { '8': Array.from({ length: 20 }, (_, i) => `e${i};`).join('') }, 'endnote');
    expect(layout.pages.length).toBeLessThanOrEqual(8);
  });

  it('keeps progressive publications growing across a late oversized-note retry', async () => {
    const model = document('footnote', [paragraph('tall;', 120), paragraph('tail;')], undefined, 1);
    model.body = [...Array.from({ length: 40 }, () => paragraph('before;')), ...model.body] as BodyElement[];
    const source = layoutSourceStore(model);
    const counts: number[] = [];
    const layout = await layoutDocumentProgressively(source.bodyLayoutInput,
      createLayoutServices(source, { measureContext: context, allowFootnoteContinuation: true }),
      { currentDateMs: 0 }, {
        onPreview: preview => counts.push(preview.layout.pages.length),
        scheduler: { sliceMs: 0, yieldToHost: async () => undefined },
      });
    expect(counts.length).toBeGreaterThanOrEqual(2);
    expect(counts).toEqual([...new Set(counts)].sort((a, b) => a - b));
    expect(layout).toEqual(render(model));
    expectConserved(layout, { '7': 'tall;tail;' });
    expect(layout.pages.length).toBeLessThanOrEqual(8);
  });

  it.each([false, true])('makes progress for three 40+10pt heads (continuation: %s)', continuation => {
    const model = document('footnote', [], 15);
    const reference = model.body[0] as DocParagraph;
    const run = reference.runs[0]!;
    model.footnotes = ['7', '8', '9'].map(id => ({ id,
      content: [paragraph(`head-${id};`, 40), paragraph(`tail-${id};`)] as BodyElement[] }));
    reference.runs = model.footnotes.map(note => ({ ...run, noteRef: { kind: 'footnote', id: note.id } })) as DocParagraph['runs'];
    const layout = render(model, continuation);
    expectConserved(layout, { '7': 'head-7;tail-7;', '8': 'head-8;tail-8;', '9': 'head-9;tail-9;' });
    expect(layout.pages.length).toBeLessThanOrEqual(4);
  });

  it.each(ordinaryLayouts)('preserves v0.89 bytes ($kind, $shape, continuation $continuation)', ({ kind, shape, continuation, digest }) => {
    const content = shape === 'unequal-tables'
      ? [table([paragraph('a;')]), table([paragraph('b;')]), table([paragraph('c;')])]
      : shape === 'fitting-table' ? [table([paragraph('a;'), paragraph('b;')])]
      : Array.from({ length: shape === 'text-continuation' ? 20 : 3 }, (_, i) => paragraph(`n${i};`)) as BodyElement[];
    const model = document(kind, content, 15);
    if (shape === 'unequal-tables') model.section.columns = { count: 2, spacePt: 20, equalWidth: false, sep: false,
      cols: [{ widthPt: 60, spacePt: 20 }, { widthPt: 100, spacePt: 0 }] };
    let expected = shape === 'unequal-tables' ? 'a;b;c;' : shape === 'fitting-table' ? 'a;b;'
      : Array.from({ length: shape === 'text-continuation' ? 20 : 3 }, (_, i) => `n${i};`).join('');
    if (shape === 'three-heads') {
      const reference = model.body[0] as DocParagraph;
      const run = reference.runs[0]!;
      model.endnotes = ['7', '8', '9'].map(id => ({ id,
        content: [paragraph(`head-${id};`, 40), paragraph(`tail-${id};`)] as BodyElement[] }));
      reference.runs = model.endnotes.map(note => ({ ...run, noteRef: { kind, id: note.id } })) as DocParagraph['runs'];
      expected = '';
    }
    const layout = render(model, continuation);
    expectConserved(layout, shape === 'three-heads'
      ? { '7': 'head-7;tail-7;', '8': 'head-8;tail-8;', '9': 'head-9;tail-9;' } : { '7': expected }, kind);
    expect(layout.pages.length).toBeLessThanOrEqual(2);
    expect(createHash('sha256').update(JSON.stringify(layout)).digest('hex')).toBe(digest);
  });

  it.each(['footnote', 'endnote'] as const)('bounds overflow balancing for a twelve-paragraph %s', kind => {
    const content = Array.from({ length: 12 }, (_, i) => paragraph(`n${i};`, i === 3 ? 120 : 10));
    const layout = render(document(kind, content as BodyElement[], 15));
    expectConserved(layout, { '7': Array.from({ length: 12 }, (_, i) => `n${i};`).join('') }, kind);
    expect(layout.pages.length).toBeLessThanOrEqual(5);
  });

  it.each([false, true])('retains all three oversized notes after rejected admission (notice: %s)', notice => {
    const model = document('footnote', [], 15);
    const reference = model.body[0] as DocParagraph;
    const run = reference.runs[0]!;
    model.footnotes = ['7', '8', '9'].map(id => ({ id,
      content: [paragraph(`tall-${id};`, 120), paragraph(`tail-${id};`)] as BodyElement[] }));
    reference.runs = model.footnotes.map(note => ({ ...run, noteRef: { kind: 'footnote', id: note.id } })) as DocParagraph['runs'];
    if (notice) {
      (model as unknown as Record<string, unknown>).__noteLayoutSettings = {
        reservedSeparators: { footnote: { separator: {}, continuationSeparator: {}, continuationNotice: {
          paragraph: { paragraph: { ...paragraph(''), runs: [] }, contentMark: { run: paragraph('').runs[0] } },
        } } },
      };
    }
    const layout = render(model);
    expectConserved(layout, { '7': 'tall-7;tail-7;', '8': 'tall-8;tail-8;', '9': 'tall-9;tail-9;' });
    if (notice) {
      const first = layout.pages[0]!.layers.notes.filter((node): node is NoteLayout => node.kind === 'note');
      expect(first.filter(note => note.trailing)).toHaveLength(1);
      expect(first.at(-1)?.trailing?.role).toBe('continuationNotice');
    }
  });

  it('uses trial capacity for a five-paragraph note with continuation disabled', () => {
    const content = Array.from({ length: 5 }, (_, i) => paragraph(`n${i};`));
    const model = document('footnote', content as BodyElement[], 15);
    model.body = [...Array.from({ length: 6 }, () => paragraph('before;')), ...model.body,
      ...Array.from({ length: 6 }, () => paragraph('after;'))] as BodyElement[];
    const layout = render(model, false);
    expectConserved(layout, { '7': 'n0;n1;n2;n3;n4;' });
    expect(layout.pages).toHaveLength(2);
  });

  it('preserves the ordinary 2+1 note-column balance with continuation disabled', () => {
    const model = document('footnote', [paragraph('n0;'), paragraph('n1;'), paragraph('n2;')] as BodyElement[], 15);
    const layout = render(model, false);
    expectConserved(layout, { '7': 'n0;n1;n2;' });
    expect(notes(layout).map(note => note.story.blocks.map(text).join(''))).toEqual(['n0;n1;', 'n2;']);
    expect(layout.pages).toHaveLength(1);
  });

  it.each((['footnote', 'endnote'] as const).flatMap(kind => [undefined, 14, 15].flatMap(mode =>
    [1, 2].map(columns => ({ kind, mode, columns })))))(
    'reaches an oversized cell unit behind a repeating header ($kind, mode $mode, $columns columns)',
    ({ kind, mode, columns }) => {
      const model = document(kind, [table([paragraph('before;'), paragraph('tall;', 120), paragraph('after;')], true)], mode, columns);
      const layout = render(model);
      expectConserved(layout, { '7': 'header;before;tall;after;' }, kind);
      const tall = notes(layout).find(note => note.story.blocks.some(block => text(block).includes('tall;')))!;
      // The body table policy retains repeated headers with a fresh-band
      // oversized unit, allowing its ink to overflow instead of losing either.
      expect(tall.story.blocks.map(text).join('')).toBe('header;tall;');
    });

  it.each(['footnote', 'endnote'] as const)('balances complete cell cursors for a thirty-paragraph %s row', kind => {
    const content = Array.from({ length: 30 }, (_, i) => paragraph(`n${i};`));
    const layout = render(document(kind, [table(content)], 15));
    expectConserved(layout, { '7': content.map(p => p.runs[0]!.type === 'text' ? p.runs[0]!.text : '').join('') }, kind);
    for (const page of layout.pages) for (const note of page.layers.notes) {
      if (note.kind !== 'note') continue;
      for (const block of note.story.blocks) if (block.kind === 'table') for (const row of block.rows) {
        expect(row).toMatchObject({ physicalPageIndex: page.pageIndex, displayPageNumber: page.pageIndex + 1 });
      }
    }
    expect(layout.pages.length).toBeLessThanOrEqual(3);
    expect(notes(layout)[0]!.story.blocks.map(text).join('')).toContain('n1;');
  });

  it('keeps continued tall-row recovery block reads linear under source doubling', () => {
    const results: { n: number; reads: number; ms: number }[] = [];
    for (const n of [16, 32, 64, 128]) {
      let reads = 0;
      const original = tableAcquisition.acquireRetainedTable;
      const spy = vi.spyOn(tableAcquisition, 'acquireRetainedTable').mockImplementation((...args) => {
        const acquired = original(...args);
        return { ...acquired, input: { ...acquired.input, rows: acquired.input.rows.map(row => ({
          ...row, cells: row.cells.map(cell => ({ ...cell,
            blocks: new Proxy([...cell.blocks], { get(target, key, receiver) {
              if (typeof key === 'string' && /^\d+$/.test(key)) reads++;
              return Reflect.get(target, key, receiver);
            } }),
          })),
        })) } };
      });
      try {
        const content = Array.from({ length: n }, (_, i) => paragraph(`p${i};`));
        const started = performance.now();
        const layout = render(document('footnote', [table(content)], 15), false);
        const ms = Math.round(performance.now() - started);
        expectConserved(layout, { '7': content.map((_, i) => `p${i};`).join('') });
        results.push({ n, reads, ms });
      } finally { spy.mockRestore(); }
    }
    process.stdout.write(JSON.stringify({ shape: 'continued-tall-row', results }) + '\n');
    for (let i = 1; i < results.length; i++) {
      // Labels all fit one line; page/column geometry and the two convergence
      // passes stay fixed. Doubling doubles source visits and band selection,
      // plus a partial boundary band. The original 512-4096-block sweep read
      // 7858/15684/31336/62642 blocks. Keep its 2.2x bound at small N;
      // the old full-row scan on every cut approaches 4x.
      expect(results[i]!.reads).toBeLessThanOrEqual(2.2 * results[i - 1]!.reads);
    }
  });

  it.each(['rowNeedsPageOrigins', 'retainedTableAcquisitionIsReusableAcrossPages'] as const)(
    'charges repeated %s scans to the note-local recovery allowance', predicate => {
      let originalFailure: Error | undefined;
      let attempts = 0;
      const recover = NoteOverflowPolicy.prototype.recover;
      const policy = vi.spyOn(NoteOverflowPolicy.prototype, 'recover').mockImplementation(function (this: NoteOverflowPolicy, error) {
        if (error instanceof NoteOverflowRecoveryRequired) originalFailure ??= error.failure;
        return recover.call(this, error);
      });
      const rowPredicate = tableAcquisition.rowNeedsPageOrigins;
      const reusePredicate = tableAcquisition.retainedTableAcquisitionIsReusableAcrossPages;
      const repeatScan = (scan: () => boolean) => {
        // Model a future unbounded inner pagination loop. Real production
        // scans must debit before work even without source/page progress.
        for (let i = 0; i < 100_000; i++) { attempts++; scan(); }
        throw new Error('Recovery did not stop source scans');
      };
      const scan = predicate === 'rowNeedsPageOrigins'
        ? vi.spyOn(tableAcquisition, predicate).mockImplementation((source, row, charge) => {
            if (!charge) return rowPredicate(source, row);
            // A fresh immutable occurrence forces a real predicate cache miss.
            return repeatScan(() => rowPredicate(source, { ...row }, charge));
          })
        : vi.spyOn(tableAcquisition, predicate).mockImplementation((source, charge) => {
            if (!charge) return reusePredicate(source);
            return repeatScan(() => reusePredicate({ ...source }, charge));
          });
      try {
        let failure: unknown;
        try {
          render(document('footnote', [table(Array.from({ length: 30 }, (_, i) => paragraph(`p${i};`)))], 15), false);
        } catch (error) { failure = error; }
        expect(originalFailure).toMatchObject({ name: 'FootnoteAdmissionOverflowError' });
        expect(failure).toBe(originalFailure);
        expect(attempts).toBeGreaterThan(1);
        expect(attempts).toBeLessThan(100_000);
      } finally { scan.mockRestore(); policy.mockRestore(); }
    });

  it.each([8, 16, 32, 64])('bounds whole-story layoutTable work linearly for %i independent tables', count => {
    const spy = vi.spyOn(tableLayout, 'layoutTable');
    try {
      const content = Array.from({ length: count }, (_, i) => table([paragraph(`n${i};`, 20)]));
      const layout = render(document('footnote', content, undefined, 1));
      expectConserved(layout, { '7': Array.from({ length: count }, (_, i) => `n${i};`).join('') });
      expect(spy.mock.calls.length).toBeLessThanOrEqual(12 * count);
    } finally { spy.mockRestore(); }
  });

  it.each([8, 16, 32, 64])('bounds acquisition with a trailing PAGE field for %i tables', count => {
    const field = paragraph('page=');
    const run = field.runs[0]!;
    if (run.type !== 'text') throw new Error('Expected text fixture');
    field.runs.push({ ...run, type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '1' });
    const spy = vi.spyOn(tableLayout, 'layoutTable');
    try {
      const content = Array.from({ length: count }, (_, i) => table([paragraph(`n${i};`, 20)]));
      const layout = render(document('footnote', [...content, field], undefined, 1));
      const destination = layout.pages.find(page => page.layers.notes.some(note =>
        note.kind === 'note' && note.story.blocks.some(block => text(block).includes('page='))))!;
      expectConserved(layout, { '7': content.map((_, i) => `n${i};`).join('') + `page=${destination.pageIndex + 1}` });
      // Ordinary admission, the destination-complete recovery head, and the
      // shared continuation epoch each need measurement + placement in the
      // initial and settled page-count passes: 3*2*2 = 12N, independent of
      // how many continuation pages the source requires.
      expect(spy.mock.calls.length).toBeLessThanOrEqual(12 * count);
    } finally { spy.mockRestore(); }
  });

  it.each([2, 4, 8, 16])('indexes nested parent cells once for %i outer rows', count => {
    const outer = table([paragraph('unused;')]) as Extract<BodyElement, { type: 'table' }>;
    outer.rows = Array.from({ length: count }, (_, i) => ({ ...outer.rows[0]!, cells: [{
      ...outer.rows[0]!.cells[0]!,
      content: [table(Array.from({ length: 12 }, (_, j) => paragraph(`r${i}p${j};`)))],
    }] }));
    let reads = 0;
    const original = tableLayout.layoutTable;
    const spy = vi.spyOn(tableLayout, 'layoutTable').mockImplementation((...args) => {
      const result = original(...args);
      if (args[0].rows.length !== count) return result;
      return { ...result, layout: { ...result.layout, rows: result.layout.rows.map(row => ({
        ...row, get cells() { reads += 1; return row.cells; },
      })) } };
    });
    try {
      const layout = render(document('footnote', [outer], undefined, 1));
      expectConserved(layout, { '7': Array.from({ length: count }, (_, i) =>
        Array.from({ length: 12 }, (_, j) => `r${i}p${j};`).join('')).join('') });
      // Budget one row read per paragraph unit (12) and eight setup/boundary
      // reads in each of two page-count passes: 2*(12+8)N. Parent lookup must
      // add one indexing traversal, not a full-table scan per unit trial.
      expect(reads).toBeLessThanOrEqual(40 * count);
    } finally { spy.mockRestore(); }
  });

  it.each(['footnote', 'endnote'] as const)('resolves interleaved field blocks on each %s destination', kind => {
    const content = Array.from({ length: 40 }, (_, i) => {
      const p = paragraph(`p${i}=`);
      const run = p.runs[0]!;
      if (run.type !== 'text') throw new Error('Expected text fixture');
      p.runs.push({ ...run, type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '1' },
        { ...run, text: ';' });
      // Exercise both root paragraphs and atomic page-dependent tables.
      return i % 2 ? p : table([p]);
    });
    const model = document(kind, content, undefined, 1);
    model.section.pageNumType = { start: 9, fmt: 'decimal' };
    const layout = render(model);
    const retained = layout.pages.flatMap(page => page.layers.notes.flatMap(note => {
      if (note.kind !== 'note') return [];
      const fragments = note.story.blocks.map(text).join('');
      for (const match of fragments.matchAll(/p\d+=(\d+);/g)) expect(Number(match[1])).toBe(page.pageIndex + 9);
      return [...fragments.matchAll(/p(\d+)=\d+;/g)].map(match => Number(match[1]));
    }));
    expect(retained).toEqual(Array.from({ length: 40 }, (_, i) => i));
    expect(layout.pages.length).toBeGreaterThan(2);
  });
});

// #2113: progressive publications are prefixes of the final layout. Recovery
// grants an oversized note its permission at the rejecting page, and later
// passes apply it to every page of that note, so a page holding an unfinished
// note's reference or fragment must wait.
describe('progressive publication around note recovery (#2113)', () => {
  const continuingFootnoteModel = (oversized: boolean, anchoredTable: boolean) => {
    const content = Array.from({ length: 30 }, (_, i) => paragraph(`n${i};`, 24));
    if (oversized) content.push(paragraph('tall;', 800), paragraph('tail;', 24));
    const model = document('footnote', content, 15, 2);
    const body: BodyElement[] = [model.body[0]!,
      ...Array.from({ length: 129 }, (_, i) => paragraph(`b${i};`, 24) as BodyElement)];
    if (anchoredTable) {
      const floating = table([paragraph('anchor;')]);
      floating.tblpPr = { leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
        horzAnchor: 'margin', horzSpecified: true, vertAnchor: 'page', tblpX: 0, tblpY: 0 };
      body.splice(40, 0, floating);
    }
    model.body = body;
    return model;
  };
  const publications = async (model: DocxDocumentModel) => {
    const source = layoutSourceStore(model);
    const previews: DocumentLayout[] = [];
    const layout = await layoutDocumentProgressively(source.bodyLayoutInput,
      createLayoutServices(source, { measureContext: context, allowFootnoteContinuation: true }),
      { currentDateMs: 0 }, {
        onPreview: preview => previews.push(preview.layout),
        scheduler: { sliceMs: 0, yieldToHost: async () => undefined },
      });
    return { previews, layout };
  };
  // Every published page keeps the final page's body and note geometry.
  const expectPublishedPagesFinal = (previews: readonly DocumentLayout[], layout: DocumentLayout) => {
    for (const preview of previews) {
      preview.pages.forEach((page, index) => {
        const final = layout.pages[index]!;
        expect(page.pageIndex).toBe(final.pageIndex);
        expect(page.layers.body).toEqual(final.layers.body);
        expect(page.layers.notes).toEqual(final.layers.notes);
      });
    }
  };

  it('never revises a published page when an oversized footnote recovers late', async () => {
    const model = continuingFootnoteModel(true, true);
    const { previews, layout } = await publications(model);
    const blocking = render(model);
    expect(layout).toEqual(blocking);
    expect(previews.length).toBeGreaterThan(0);
    expectPublishedPagesFinal(previews, layout);
    const counts = previews.map(preview => preview.pages.length);
    expect(counts).toEqual([...new Set(counts)].sort((a, b) => a - b));
  });

  it('never publishes a page a recovered reference can move back to', async () => {
    // The opening pass defers the reference past page 1; after recovery the
    // final layout admits it, and the note's first fragments, on page 1.
    const content: BodyElement[] = [...Array.from({ length: 18 }, (_, i) => paragraph(`n${i};`, 24)),
      table([paragraph('tall;', 160), paragraph('tail;', 24)])];
    const model = document('footnote', content, 15, 2);
    const reference = model.body[0]!;
    const body: BodyElement[] = Array.from({ length: 80 }, (_, i) => paragraph(`b${i};`, 24) as BodyElement);
    body.splice(8, 0, reference);
    const floating = table([paragraph('anchor;')]);
    floating.tblpPr = { leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: 'margin', horzSpecified: true, vertAnchor: 'page', tblpX: 0, tblpY: 0 };
    body.splice(12, 0, floating);
    model.body = body;
    const { previews, layout } = await publications(model);
    expect(layout).toEqual(render(model));
    expectPublishedPagesFinal(previews, layout);
  });

  it('keeps the footnote bound in a mode-15 document with endnotes', async () => {
    const model = continuingFootnoteModel(true, true);
    const last = model.body.at(-1) as DocParagraph;
    const run = last.runs[0]!;
    last.runs.push({ ...run, text: 'e;', noteRef: { kind: 'endnote', id: '8' } } as DocParagraph['runs'][number]);
    model.endnotes = [{ id: '8', content: [paragraph('end;')] }];
    const { previews, layout } = await publications(model);
    expect(layout).toEqual(render(model));
    expectPublishedPagesFinal(previews, layout);
  });

  it('keeps an earlier occurrence unsettled while its footnote is referenced again', async () => {
    // A permission is per footnote id: recovery at the second reference also
    // changes the first occurrence's note area.
    const model = document('footnote', [paragraph('a;', 50), paragraph('b;', 10)], 15, 2);
    const reference = model.body[0]! as DocParagraph;
    const filler = (prefix: string) => Array.from({ length: 60 }, (_, i) => paragraph(`${prefix}${i};`) as BodyElement);
    const columns = model.section.columns;
    const body: BodyElement[] = [reference, ...filler('s'), {
      type: 'sectionBreak', kind: 'nextPage', columns,
      geom: { pageWidth: 200, pageHeight: 100, marginTop: 10, marginRight: 10, marginBottom: 10,
        marginLeft: 10, headerDistance: 0, footerDistance: 0 },
    } as BodyElement, structuredClone(reference), ...filler('t')];
    const floating = table([paragraph('anchor;')]);
    floating.tblpPr = { leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: 'margin', horzSpecified: true, vertAnchor: 'page', tblpX: 0, tblpY: 0 };
    body.splice(25, 0, floating);
    model.body = body;
    model.section = { ...model.section, marginTop: 30, marginBottom: 30 };
    const { previews, layout } = await publications(model);
    expect(layout).toEqual(render(model));
    expectPublishedPagesFinal(previews, layout);
  });

  it('only delays publication of an ordinary continuing footnote', async () => {
    const model = continuingFootnoteModel(false, false);
    const { previews, layout } = await publications(model);
    expect(layout).toEqual(render(model));
    expect(previews.length).toBeGreaterThan(0);
    expectPublishedPagesFinal(previews, layout);
  });
});
