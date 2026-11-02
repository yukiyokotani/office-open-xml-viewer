import { describe, expect, it, vi } from 'vitest';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import { noteContinuationModel } from '../testing/note-continuation-model.js';
import type { BodyElement, DocParagraph, DocxDocumentModel } from '../types.js';
import type { DocumentLayout, NoteLayout, PaintNode } from './types.js';
import { endnoteIdsInRetainedSlice, footnoteIdsInRetainedSlice } from './note-reference-ownership.js';
import * as tableLayout from './table.js';

type Kind = 'footnote' | 'endnote';
const context = {
  font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
  measureText: (text: string) => ({ width: [...text].length * 3,
    actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }),
} as unknown as CanvasRenderingContext2D;

function paragraph(text: string, height = 10): DocParagraph {
  const base = noteContinuationModel().footnotes![0]!.content[0] as DocParagraph;
  const run = base.runs[0];
  if (run?.type !== 'text') throw new Error('Expected text fixture');
  return { ...base, widowControl: false, spaceBefore: 0, spaceAfter: 0,
    lineSpacing: { value: height, rule: 'exact', explicit: true },
    runs: [{ ...run, text, noteRef: undefined, vertAlign: null }] };
}
function table(content: DocParagraph[], header = false): BodyElement {
  const row = (items: DocParagraph[], isHeader = false) => ({
    cells: [{ content: items, colSpan: 1, vMerge: null, borders: {}, background: null,
      vAlign: 'top', widthPt: 40 }],
    rowHeight: null, rowHeightRule: 'auto', isHeader, cantSplit: false,
  });
  return { type: 'table', colWidths: [40], rows: [
    ...(header ? [row([paragraph('header;')], true)] : []), row(content),
  ], borders: {}, cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
  jc: 'left', layout: 'fixed', tblpPr: null } as unknown as BodyElement;
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

describe('reviewed note admission regressions', () => {
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

  it.each([40, 80, 160])('bounds whole-story layoutTable work linearly for %i independent tables', count => {
    const spy = vi.spyOn(tableLayout, 'layoutTable');
    try {
      const content = Array.from({ length: count }, (_, i) => table([paragraph(`n${i};`, 20)]));
      const layout = render(document('footnote', content, undefined, 1));
      expectConserved(layout, { '7': Array.from({ length: count }, (_, i) => `n${i};`).join('') });
      expect(spy.mock.calls.length).toBeLessThanOrEqual(12 * count);
    } finally { spy.mockRestore(); }
  });
});
