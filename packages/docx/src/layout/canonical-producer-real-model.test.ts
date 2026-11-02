import { describe, expect, it, vi } from 'vitest';
import * as paginationWorkBudget from './pagination-work-budget.js';
import { bodyLayoutKernelOf, createParagraphAcquisitionCacheServicesView,
  createFieldAcquisitionServicesView, paragraphAcquisitionCacheOf, type PageFieldAcquisitionContext } from './runtime-state.js';
import { createPaginationWorkBudget, MAX_PAGINATION_ACQUISITION_MISSES } from './pagination-work-budget.js';
import { buildBookmarkPageMap } from '../bookmark-nav.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocumentProgressively } from './progressive.js';
import { layoutFingerprint } from './invariants.js';
import { layoutDocument } from '../document-layout.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import { layoutDocumentInputAsync } from './document.js';
import type { BodyElement, CellElement, DocParagraph, DocxDocumentModel, SectionProps } from '../types.js';
import type { LayoutServices, NoteLayout } from './types.js';

function measureContext(): CanvasRenderingContext2D {
  return {
    font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * 5,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function paragraph(): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text: '1', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'serif',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: 'super',
      hyperlink: null, noteRef: { kind: 'footnote', id: '7' },
    }],
    defaultFontSize: 10, defaultFontFamily: 'serif', widowControl: false,
  } as unknown as DocParagraph;
}

function ordinaryParagraph(text: string): DocParagraph {
  const result = paragraph();
  const run = result.runs[0];
  if (run?.type === 'text') {
    run.text = text;
    run.noteRef = undefined;
    run.vertAlign = null;
  }
  return result;
}

function columnStoryDocument(body: DocParagraph[], footnotes: { id: string; content: DocParagraph[] }[], nextLines = 8): DocxDocumentModel {
  return {
    section: {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5,
      titlePage: true, evenAndOddHeaders: false, sectionStart: 'nextPage',
      columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
    },
    body,
    headers: {
      first: { body: [ordinaryParagraph('first')] },
      default: { body: Array.from({ length: nextLines }, () => ordinaryParagraph('next')) }, even: null,
    },
    footers: { default: null, first: null, even: null },
    footnotes, endnotes: [], fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

function ordinaryBodyParagraph(text: string): BodyElement {
  return ordinaryParagraph(text) as unknown as BodyElement;
}

function endnoteReferenceParagraph(id: string): BodyElement {
  const result = paragraph();
  const run = result.runs[0];
  if (run?.type === 'text') {
    run.noteRef = { kind: 'endnote', id };
  }
  return result as unknown as BodyElement;
}

function noteMarkerParagraph(kind: 'footnote' | 'endnote'): DocParagraph {
  const result = ordinaryParagraph('');
  const run = result.runs[0];
  if (run?.type === 'text') {
    run.noteRef = { kind, id: '' };
    run.vertAlign = 'super';
  }
  return result;
}

function pageOwnedWrappingParagraph(): BodyElement {
  const result = ordinaryParagraph('');
  result.runs = [{
    type: 'shape',
    widthPt: 30,
    heightPt: 20,
    anchorXPt: 0,
    anchorYPt: 0,
    anchorXFromMargin: true,
    anchorYFromPara: false,
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: { fillType: 'solid', color: 'FFFFFF' },
    stroke: null,
    wrapMode: 'square',
    wrapSide: 'bothSides',
    distTop: 0,
    distBottom: 0,
    distLeft: 0,
    distRight: 0,
  }] as DocParagraph['runs'];
  return result as unknown as BodyElement;
}

function fittedTextPathParagraph(): BodyElement {
  const result = ordinaryParagraph('');
  result.runs = [{
    type: 'shape',
    widthPt: 30,
    heightPt: 20,
    anchorXPt: 0,
    anchorYPt: 0,
    anchorXFromMargin: false,
    anchorYFromPara: true,
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: { fillType: 'solid', color: 'D9D9D9' },
    stroke: null,
    textPath: {
      string: 'DRAFT',
      fontFamily: 'Arial',
      bold: false,
      italic: false,
      textPathOk: true,
      on: true,
      fitShape: true,
      fitPath: true,
      trim: false,
      xScale: false,
      fontSizePt: 12,
    },
  }] as unknown as DocParagraph['runs'];
  return result as unknown as BodyElement;
}

function sectionBreak(): BodyElement {
  return {
    type: 'sectionBreak', kind: 'continuous', columns: null,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    titlePage: false,
  } as unknown as BodyElement;
}

function sectionBoundaryModel(
  before: BodyElement,
  startType: 'continuous' | 'nextPage',
): DocxDocumentModel {
  const section = {
    pageWidth: 200, pageHeight: 100,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 5, footerDistance: 5, titlePage: false,
    evenAndOddHeaders: false, sectionStart: startType, columns: null,
  } as SectionProps;
  return {
    section,
    body: [before, sectionBreak(), ordinaryBodyParagraph('after')],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [], endnotes: [], fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

function floatingTable(): BodyElement {
  return {
    type: 'table', colWidths: [40],
    rows: [{
      cells: [{
        content: [ordinaryParagraph('1')], colSpan: 1, vMerge: null,
        borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
        background: null, vAlign: 'top', widthPt: 40,
      }],
      rowHeight: null, rowHeightRule: 'auto', isHeader: false,
    }],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed', overlap: 'overlap',
    tblpPr: {
      leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: 'text', horzSpecified: true, vertAnchor: 'text',
      tblpX: 20, tblpY: 10,
    },
  } as unknown as BodyElement;
}

function ordinaryTable(
  justification: 'left' | 'center' | 'right',
  indentPt: number,
  bidiVisual: boolean,
): BodyElement {
  const source = floatingTable() as Extract<BodyElement, { type: 'table' }>;
  return {
    ...source,
    jc: justification,
    tblInd: indentPt,
    bidiVisual,
    tblpPr: null,
  } as unknown as BodyElement;
}

function fragmentLineAdvancesPt(fragment: Extract<ReturnType<typeof layoutDocument>['pages'][number]['layers']['body'][number], { kind: 'paragraph' }>): number {
  return fragment.lines.reduce((sum, line) => sum + line.advancePt, 0);
}

describe('canonical producer with a real document model', () => {
  it.each([
    { name: 'shared restart selects odd', start: 2, shared: true, titlePage: false, defaultLines: 1, evenLines: 6 },
    { name: 'shared restart selects even', start: 3, shared: true, titlePage: false, defaultLines: 6, evenLines: 1 },
    { name: 'empty shared region preserves restart', start: 3, shared: false, titlePage: false, defaultLines: 1, evenLines: 6 },
    { name: 'shared restart selects first story', start: 2, shared: true, titlePage: true, defaultLines: 6, evenLines: 6 },
    { name: 'empty shared region selects first story', start: 2, shared: false, titlePage: true, defaultLines: 6, evenLines: 6 },
  ].flatMap(test => [false, true].map(continuation => ({ ...test, continuation }))))(
    'opens continuous sections with final page context: $name (continuation $continuation)',
    ({ start, shared, titlePage, defaultLines, evenLines, continuation }) => {
      const model = columnStoryDocument([], [{ id: '7', content: [ordinaryParagraph('note')] }], defaultLines);
      model.section.columns = null;
      model.section.sectionStart = 'continuous';
      model.section.pageNumType = { start, fmt: 'decimal' };
      model.section.titlePage = titlePage;
      model.section.evenAndOddHeaders = true;
      model.headers.even = { body: Array.from({ length: evenLines }, () => ordinaryBodyParagraph('even')) };
      const endingSection = {
        type: 'sectionBreak', kind: 'nextPage', columns: null,
        geom: { ...model.section, sectionStart: 'nextPage', pageNumType: null },
        pageNumType: null, titlePage: false,
        headers: { default: { body: Array.from({ length: 4 }, () => ordinaryParagraph('outhead')) },
          first: null, even: null },
        footers: model.footers,
      } as unknown as BodyElement;
      const table = ordinaryTable('left', 0, false) as Extract<BodyElement, { type: 'table' }>;
      table.rows[0]!.cantSplit = true;
      table.rows[0]!.cells[0]!.content = [{ ...paragraph(), type: 'paragraph' }];
      model.body = [...Array.from({ length: shared ? 1 : 4 }, () => ordinaryBodyParagraph('out')),
        endingSection, ...(shared ? [ordinaryBodyParagraph('in')] : []), table];
      const layout = layoutDocument(model, createLayoutServices(model, {
        measureContext: measureContext(), allowFootnoteContinuation: continuation,
      }), { currentDateMs: 0 });
      expect(layout.pages.map(page => page.pageNumber.displayNumber)).toEqual([1, start + (shared ? 1 : 0)]);
      expect(layout.pages[1]!.layers.body[0]).toMatchObject({
        kind: 'table', flowBounds: { xPt: 10, yPt: 15, heightPt: 10 },
      });
      expect(layout.pages.map(page => page.layers.notes.map(note => note.advancePt))).toEqual([[], [16]]);
      expect(layout.pages[1]!.layers.header).toHaveLength(1);
    },
  );

  // Different first/default and odd/even stories reserve different body bands
  // (§17.6.11, §17.10.1/.6). An unused destination must not reject this page.
  it.each([
    { band: 'header', slot: 'default', nextLines: 6, bodyKind: 'paragraph', continuation: false },
    { band: 'header', slot: 'default', nextLines: 6, bodyKind: 'paragraph', continuation: true },
    { band: 'header', slot: 'default', nextLines: 10, bodyKind: 'paragraph', continuation: false },
    { band: 'header', slot: 'default', nextLines: 10, bodyKind: 'paragraph', continuation: true },
    { band: 'footer', slot: 'even', nextLines: 6, bodyKind: 'paragraph', continuation: false },
    { band: 'footer', slot: 'even', nextLines: 6, bodyKind: 'paragraph', continuation: true },
    { band: 'header', slot: 'default', nextLines: 6, bodyKind: 'frame', continuation: false },
    { band: 'header', slot: 'default', nextLines: 6, bodyKind: 'frame', continuation: true },
    { band: 'footer', slot: 'even', nextLines: 6, bodyKind: 'table', continuation: false },
    { band: 'footer', slot: 'even', nextLines: 6, bodyKind: 'table', continuation: true },
  ] as const)('keeps a fitting $bodyKind note with a taller unused $slot $band ($nextLines lines, continuation $continuation)', ({
    band, slot, nextLines, bodyKind, continuation,
  }) => {
    const reference = paragraph();
    if (bodyKind === 'frame') reference.framePr = {
      dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
      hRule: 'auto', hSpace: 0, vSpace: 0, w: 40, h: 10, x: 0, y: 0,
    };
    const table = ordinaryTable('left', 0, false) as Extract<BodyElement, { type: 'table' }>;
    table.rows[0]!.cells[0]!.content = [{ ...reference, type: 'paragraph' }];
    const stories = {
      first: { body: [ordinaryParagraph('first')] },
      default: null,
      even: null,
      [slot]: { body: Array.from({ length: nextLines }, () => ordinaryParagraph('next')) },
    };
    const emptyStories = { default: null, first: null, even: null };
    const model = {
      section: {
        pageWidth: 200, pageHeight: 100,
        marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
        headerDistance: 5, footerDistance: 5,
        titlePage: true, evenAndOddHeaders: slot === 'even',
        sectionStart: 'nextPage', columns: null,
      },
      body: [bodyKind === 'table' ? table : reference],
      headers: band === 'header' ? stories : emptyStories,
      footers: band === 'footer' ? stories : emptyStories,
      footnotes: [{ id: '7', content: [ordinaryParagraph('note')] }],
      endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: continuation,
    });
    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(1);
    const page = layout.pages[0]!;
    expect(page.layers.body).toHaveLength(1);
    expect(page.layers.body[0]?.ordinaryFlow).toBe(bodyKind !== 'frame');
    expect(page.layers.body[0]?.flowBounds.yPt).toBe(band === 'header' ? 15 : 10);
    expect(page.layers.notes).toHaveLength(1);
    expect(page.layers.notes[0]).toMatchObject({
      kind: 'note', advancePt: 16, source: { storyInstance: '7' },
      flowDomainId: 'notes:page:0',
    });
  });

  it.each([false, true])('validates successive current-page notes with a narrower unused header (continuation %s)', continuation => {
    const second = paragraph();
    const run = second.runs[0];
    if (run?.type === 'text') run.noteRef = { kind: 'footnote', id: '8' };
    const model = {
      section: {
        pageWidth: 200, pageHeight: 100,
        marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
        headerDistance: 5, footerDistance: 5,
        titlePage: true, evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
      },
      body: [paragraph(), second],
      headers: {
        first: { body: [ordinaryParagraph('first')] },
        default: { body: Array.from({ length: 6 }, () => ordinaryParagraph('next')) }, even: null,
      },
      footers: { default: null, first: null, even: null },
      footnotes: ['7', '8'].map(id => ({ id, content: [ordinaryParagraph('note')] })),
      endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: continuation,
    }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(1);
    expect(layout.pages[0]?.layers.notes.map(note => note.advancePt)).toEqual([16, 10]);
    expect(layout.pages[0]?.layers.notes.map(note => note.source.storyInstance)).toEqual(['7', '8']);
  });

  it.each([
    { nextLines: 5, continuation: false }, { nextLines: 5, continuation: true },
    { nextLines: 6, continuation: false }, { nextLines: 6, continuation: true },
  ])('uses the destination header when a note actually relocates ($nextLines lines, continuation $continuation)', ({
    nextLines, continuation,
  }) => {
    const model = {
      section: {
        pageWidth: 200, pageHeight: 100,
        marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
        headerDistance: 5, footerDistance: 5,
        titlePage: true, evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
      },
      body: [...Array.from({ length: 6 }, () => ordinaryParagraph('before')), paragraph()],
      headers: {
        first: { body: [ordinaryParagraph('first')] },
        default: { body: Array.from({ length: nextLines }, () => ordinaryParagraph('next')) }, even: null,
      },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [ordinaryParagraph('note')] }],
      endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const render = () => layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: continuation,
    }), { currentDateMs: 0 });
    // Only 15pt remains on page one. The 26pt reference + note fits a 35pt
    // destination band. With only 25pt there, retain the note below the
    // reference and let its ink overflow instead of rejecting the document.
    const layout = render();
    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map(page => page.layers.body.length)).toEqual([6, 1]);
    expect(layout.pages[0]?.layers.notes).toEqual([]);
    expect(layout.pages[1]?.layers.body[0]?.flowBounds.yPt).toBe(nextLines === 6 ? 65 : 55);
    expect(layout.pages[1]?.layers.notes[0]).toMatchObject({
      advancePt: 16, flowDomainId: 'notes:page:1', source: { storyInstance: '7' },
    });
  });

  it.each([
    { bodyKind: 'table', continuation: false, prefixLines: 5, nextLines: 8 },
    { bodyKind: 'table', continuation: true, prefixLines: 5, nextLines: 8 },
    { bodyKind: 'frame', continuation: false, prefixLines: 5, nextLines: 8 },
    { bodyKind: 'frame', continuation: true, prefixLines: 5, nextLines: 8 },
    { bodyKind: 'table', continuation: false, prefixLines: 6, nextLines: 5 },
    { bodyKind: 'frame', continuation: true, prefixLines: 6, nextLines: 5 },
  ] as const)('relocates a note-bearing $bodyKind past $prefixLines lines to its actual region (continuation $continuation)', ({ bodyKind, continuation, prefixLines, nextLines }) => {
    const reference = paragraph();
    if (bodyKind === 'frame') reference.framePr = {
      dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
      hRule: 'auto', hSpace: 0, vSpace: 0, w: 40, h: 10, x: 0, y: 0,
    };
    const table = ordinaryTable('left', 0, false) as Extract<BodyElement, { type: 'table' }>;
    table.rows[0]!.cells[0]!.content = [{ ...reference, type: 'paragraph' }];
    const model = {
      section: {
        pageWidth: 200, pageHeight: 100,
        marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
        headerDistance: 5, footerDistance: 5,
        titlePage: true, evenAndOddHeaders: false, sectionStart: 'nextPage',
        columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
      },
      body: [...Array.from({ length: prefixLines }, () => ordinaryParagraph('before')), bodyKind === 'table' ? table : reference],
      headers: {
        first: { body: [ordinaryParagraph('first')] },
        default: { body: Array.from({ length: nextLines }, () => ordinaryParagraph('next')) }, even: null,
      },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [ordinaryParagraph('note')] }],
      endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: continuation,
    }), { currentDateMs: 0 });
    // After 50pt, the 16pt note clears the first column and the body unit
    // moves to column two. After 60pt, that band would clip committed text:
    // another column cannot help, so the unit and note need a physical page.
    const nextPage = prefixLines === 6;
    expect(layout.pages).toHaveLength(nextPage ? 2 : 1);
    const page = layout.pages.at(-1)!;
    expect(page.layers.body).toHaveLength(nextPage ? 1 : 6);
    expect(page.layers.body.at(-1)).toMatchObject({
      kind: bodyKind === 'table' ? 'table' : 'paragraph',
      flowBounds: { xPt: nextPage ? 10 : 110, yPt: nextPage ? 55 : 15 },
    });
    expect(page.layers.notes[0]).toMatchObject({ advancePt: 16, flowDomainId: `notes:page:${nextPage ? 1 : 0}` });
  });

  it.each([6, 7])('keeps a whole ordinary note after a column move with a %s-line unused header', (nextLines) => {
    const model = columnStoryDocument(
      [...Array.from({ length: 4 }, () => ordinaryParagraph('before')), paragraph()],
      [{ id: '7', content: [ordinaryParagraph('note'), ordinaryParagraph('note')] }], nextLines,
    );
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: true,
    }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(1);
    expect(layout.pages[0]!.layers.body.at(-1)).toMatchObject({
      flowBounds: { xPt: 110, yPt: 15 },
    });
    expect(layout.pages[0]!.layers.notes).toHaveLength(1);
    expect(layout.pages[0]!.layers.notes[0]).toMatchObject({ advancePt: 26, flowDomainId: 'notes:page:0' });
  });

  it('replans a continuing ordinary note in the next column before choosing a physical page', () => {
    const model = columnStoryDocument(
      [...Array.from({ length: 5 }, () => ordinaryParagraph('before')), paragraph()],
      [{ id: '7', content: [ordinaryParagraph('note'), ordinaryParagraph('note')] }], 5,
    );
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: true,
    }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(2);
    expect(layout.pages[0]!.layers.body.at(-1)).toMatchObject({ flowBounds: { xPt: 110, yPt: 15 } });
    expect(layout.pages.map(page => page.layers.notes.map(note => note.advancePt))).toEqual([[16], [16]]);
  });

  it('replans keepNext notes in another column before a whole-note physical-page preflight', () => {
    const reference = paragraph();
    reference.keepNext = true;
    reference.runs.push(...ordinaryParagraph(' after').runs);
    const model = columnStoryDocument(
      [...Array.from({ length: 5 }, () => ordinaryParagraph('before')), reference, ordinaryParagraph('last')],
      [{ id: '7', content: [ordinaryParagraph('note'), ordinaryParagraph('note')] }], 3,
    );
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: true,
    }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(2);
    expect(layout.pages[0]!.layers.body).toHaveLength(7);
    expect(layout.pages[0]!.layers.body[5]).toMatchObject({ flowBounds: { xPt: 110, yPt: 15 } });
    expect(layout.pages[1]!.layers.body).toHaveLength(0);
    expect(layout.pages.map(page => page.layers.notes.map(note => note.advancePt))).toEqual([[16], [16]]);
  });

  it('keeps mandatory later reference notes whole when their paragraph moves to another column', () => {
    const reference = paragraph();
    const later = paragraph();
    if (later.runs[0]?.type === 'text') later.runs[0].noteRef = { kind: 'footnote', id: '8' };
    reference.keepLines = true;
    reference.runs.push({ type: 'break', breakType: 'line' }, ...later.runs);
    const model = columnStoryDocument(
      [...Array.from({ length: 3 }, () => ordinaryParagraph('before')), reference],
      [{ id: '7', content: [ordinaryParagraph('first note')] },
        { id: '8', content: [ordinaryParagraph('second note'), ordinaryParagraph('second note')] }],
    );
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: true,
    }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(1);
    expect(layout.pages[0]!.layers.body.at(-1)).toMatchObject({
      flowBounds: { xPt: 110, yPt: 15 }, lines: [{}, {}],
    });
    expect(layout.pages[0]!.layers.notes.map(note => note.advancePt)).toEqual([16, 20]);
  });

  it.each([3, 4])('retains committed notes and earlier-column clearance after %s lines', (prefixLines) => {
    const earlier = paragraph();
    if (earlier.runs[0]?.type === 'text') earlier.runs[0].noteRef = { kind: 'footnote', id: '8' };
    const needsPhysicalPage = prefixLines === 4;
    const model = columnStoryDocument(
      [earlier, ...Array.from({ length: prefixLines - 1 }, () => ordinaryParagraph('before')), paragraph()],
      [{ id: '8', content: [ordinaryParagraph('earlier note')] },
        { id: '7', content: [ordinaryParagraph('later note'), ordinaryParagraph('later note')] }],
      needsPhysicalPage ? 3 : 8,
    );
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: false,
    }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(needsPhysicalPage ? 2 : 1);
    expect(layout.pages.at(-1)!.layers.body.at(-1)).toMatchObject({
      flowBounds: { xPt: needsPhysicalPage ? 10 : 110, yPt: needsPhysicalPage ? 35 : 15 },
    });
    expect(layout.pages.flatMap(page => page.layers.notes.map(note => note.advancePt)))
      .toEqual(needsPhysicalPage ? [16, 26] : [16, 20]);
  });

  it.each(['keepLines', 'keepNext'] as const)('keeps a %s group in the next column despite a smaller unused physical page', (policy) => {
    const first = ordinaryParagraph('first');
    first[policy] = true;
    const last = ordinaryParagraph('last');
    if (policy === 'keepLines') first.runs.push({ type: 'break', breakType: 'line' }, ...last.runs);
    const group = policy === 'keepNext' ? [first, last] : [first];
    const model = columnStoryDocument(
      [...Array.from({ length: 6 }, () => ordinaryParagraph('before')), ...group], [],
    );
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(),
    }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(1);
    expect(layout.pages[0]!.layers.body[6]).toMatchObject({ flowBounds: { xPt: 110, yPt: 15 } });
    expect(layout.pages[0]!.layers.body).toHaveLength(6 + group.length);
  });

  it.each([
    { bodyKind: 'frame', columns: 2, continuation: false },
    { bodyKind: 'frame', columns: 2, continuation: true },
    { bodyKind: 'table', columns: 3, continuation: false },
    { bodyKind: 'table', columns: 3, continuation: true },
    { bodyKind: 'table', columns: 2, continuation: false },
    { bodyKind: 'table', columns: 2, continuation: true },
  ] as const)('moves a whole $bodyKind and its distinct note past $columns note-shrunken columns (continuation $continuation)', ({
    bodyKind, columns, continuation,
  }) => {
    const reference = paragraph();
    if (reference.runs[0]?.type === 'text') reference.runs[0].noteRef = { kind: 'footnote', id: '8' };
    if (bodyKind === 'frame') reference.framePr = {
      dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
      hRule: 'exact', hSpace: 0, vSpace: 0, w: 40, h: 50, x: 0, y: 0,
    };
    else {
      reference.keepLines = true;
      reference.lineSpacing = { value: 10, rule: 'exact', explicit: true };
      for (let line = 1; line < 5; line++) reference.runs.push(
        { type: 'break', breakType: 'line' }, ...ordinaryParagraph('row').runs,
      );
    }
    const table = ordinaryTable('left', 0, false) as Extract<BodyElement, { type: 'table' }>;
    table.rows[0]!.cantSplit = true;
    table.rows[0]!.cells[0]!.content = [{ ...reference, type: 'paragraph' }];
    const model = columnStoryDocument([paragraph(), reference],
      ['7', '8'].map(id => ({ id, content: [ordinaryParagraph('note')] })), 1);
    model.section.columns = { count: columns, spacePt: columns === 3 ? 10 : 20,
      equalWidth: true, sep: false, cols: [] };
    if (bodyKind === 'table') model.body[1] = table;
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: continuation,
    }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map(page => page.layers.body.length)).toEqual([1, 1]);
    expect(layout.pages.map(page => page.layers.notes.map(note => ({
      id: note.source.storyInstance, advance: note.advancePt,
    })))).toEqual([[{ id: '7', advance: 16 }], [{ id: '8', advance: 16 }]]);
    expect(layout.pages[1]!.layers.body[0]).toMatchObject({
      kind: bodyKind === 'frame' ? 'paragraph' : 'table',
      flowBounds: { xPt: 10, yPt: 15 },
    });
  });

  it.each([
    { policy: 'keepLines', prefixLines: 1, continuation: false },
    { policy: 'keepLines', prefixLines: 1, continuation: true },
    { policy: 'keepLines', prefixLines: 4, continuation: false },
    { policy: 'keepLines', prefixLines: 4, continuation: true },
    { policy: 'keepNext', prefixLines: 1, continuation: false },
    { policy: 'keepNext', prefixLines: 1, continuation: true },
  ] as const)('preserves $policy after $prefixLines lines with committed notes in earlier columns (continuation $continuation)', ({
    policy, prefixLines, continuation,
  }) => {
    const group = Array.from({ length: 6 }, (_, index) => ordinaryParagraph(`line${index}`));
    group.forEach((item, index) => {
      item.lineSpacing = { value: 10, rule: 'exact', explicit: true };
      if (policy === 'keepNext') item.keepNext = index < 5;
    });
    if (policy === 'keepLines') {
      group[0]!.keepLines = true;
      group[0]!.runs = group.flatMap((item, index) => [
        ...(index > 0 ? [{ type: 'break' as const, breakType: 'line' as const }] : []), ...item.runs,
      ]);
      group.splice(1);
    }
    const model = columnStoryDocument([
      paragraph(), ...Array.from({ length: prefixLines - 1 }, () => ordinaryParagraph('before')), ...group,
    ], [{ id: '7', content: [ordinaryParagraph('note')] }], 1);
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: continuation,
    }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map(page => page.layers.body.length)).toEqual([prefixLines, group.length]);
    expect(layout.pages[1]!.layers.body[0]).toMatchObject({ flowBounds: { xPt: 10, yPt: 15 } });
    expect(layout.pages[1]!.layers.body.flatMap(block => block.kind === 'paragraph' ? block.lines : []))
      .toHaveLength(6);
    expect(layout.pages.map(page => page.layers.notes.map(note => note.advancePt))).toEqual([[16], []]);
  });

  it.each([false, true])('observes an empty opened page before retrying a cantSplit row (continuation %s)', continuation => {
    const table = ordinaryTable('left', 0, false) as Extract<BodyElement, { type: 'table' }>;
    table.rows[0]!.cantSplit = true;
    table.rows[0]!.cells[0]!.content = [{ ...paragraph(), type: 'paragraph' }];
    const model = columnStoryDocument(Array.from({ length: 5 }, () => ordinaryParagraph('before')),
      [{ id: '7', content: [ordinaryParagraph('note')] }], 8);
    model.section.columns = null;
    model.body.push(table);
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: continuation,
    }), { currentDateMs: 0 });
    // Main's fragmenter clips this indivisible row to its 9pt remaining frame
    // and retains the empty physical continuation. Opening measurement must
    // reach that same settled result, rather than defer it for 10,000 pages.
    expect(layout.pages.map(page => page.layers.body.length)).toEqual([6, 0]);
    expect(layout.pages[0]!.layers.body.at(-1)).toMatchObject({
      kind: 'table', flowBounds: { xPt: 10, yPt: 65, heightPt: 9 },
    });
    expect(layout.pages.map(page => page.layers.notes.map(note => note.advancePt))).toEqual([[16], []]);
  });

  it.each([false, true])('retains an impossible row and its note at the observed fresh opening (continuation %s)', continuation => {
    const table = ordinaryTable('left', 0, false) as Extract<BodyElement, { type: 'table' }>;
    table.rows[0]!.cantSplit = true;
    table.rows[0]!.cells[0]!.content = [{ ...paragraph(), type: 'paragraph' }];
    table.rows.push(structuredClone(table.rows[0]!));
    const model = columnStoryDocument(Array.from({ length: 4 }, () => ordinaryParagraph('before')),
      [{ id: '7', content: [ordinaryParagraph('note')] }], 6);
    model.section.columns = null;
    model.body.push(table);
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: continuation,
    }), { currentDateMs: 0 });
    expect(layout.pages.flatMap(page => page.layers.body).flatMap(block => block.kind === 'table'
      ? block.rows : [])).toHaveLength(2);
    // Each row references the same note; a repeated reference on another
    // physical page owns a fresh occurrence, with no reused descendant ids.
    for (const page of layout.pages) {
      if (page.layers.body.some(block => block.kind === 'table')) {
        expect(page.layers.notes).toHaveLength(1);
        expect(page.layers.notes[0]!.source.storyInstance).toBe('7');
      }
    }
  });

  it('observes independent empty openings across more than sixteen section boundaries', () => {
    const model = columnStoryDocument([], [{ id: '7', content: [ordinaryParagraph('note')] }], 8);
    model.section.columns = null;
    for (let section = 0; section < 17; section++) {
      if (section > 0) model.body.push({
        type: 'sectionBreak', kind: 'nextPage', columns: null,
        headers: model.headers, footers: model.footers, titlePage: true,
      } as BodyElement);
      const table = ordinaryTable('left', 0, false) as Extract<BodyElement, { type: 'table' }>;
      table.rows[0]!.cantSplit = true;
      table.rows[0]!.cells[0]!.content = [{ ...paragraph(), type: 'paragraph' }];
      model.body.push(...Array.from({ length: 5 }, () => ordinaryBodyParagraph('before')), table);
    }
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(),
    }), { currentDateMs: 0 });
    expect(layout.pages.map(page => page.layers.body.length)).toEqual(Array.from({ length: 17 }, () => [6, 0]).flat());
    expect(layout.pages.filter(page => page.layers.body.length > 0).every(page =>
      page.layers.body.at(-1)?.flowBounds.heightPt === 9 && page.layers.notes[0]?.advancePt === 16)).toBe(true);
  });

  it('retains fitted path text without interrupting surrounding body blocks', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section,
      body: [
        ordinaryBodyParagraph('before'),
        fittedTextPathParagraph(),
        ordinaryBodyParagraph('after'),
      ],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const paragraphs = layout.pages.flatMap((page) => page.layers.body)
      .filter((node) => node.kind === 'paragraph');
    const text = paragraphs.flatMap((paragraph) => paragraph.lines)
      .flatMap((line) => line.placements)
      .filter((placement) => placement.kind === 'text')
      .map((placement) => placement.text);
    const fitted = paragraphs.find((paragraph) => paragraph.source.path[0] === 1);

    expect(text).toEqual(expect.arrayContaining(['before', 'after']));
    expect(fitted?.drawings[0]?.commands).toMatchObject([{ kind: 'watermark-text', text: 'DRAFT' }]);
    expect(layout.diagnostics).toEqual([]);
  });

  it('lays out document-end notes through the retained story engine', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section,
      body: [endnoteReferenceParagraph('2')],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [],
      endnotes: [
        { id: '1', content: [ordinaryParagraph('unreferenced')] },
        { id: '2', content: [noteMarkerParagraph('endnote')] },
      ],
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const page = layout.pages[0]!;
    const endnote = page.layers.notes.find((node) => node.source.story === 'endnote');

    expect(layout.diagnostics).toEqual([]);
    expect(endnote).toMatchObject({
      kind: 'note',
      source: { story: 'endnote', storyInstance: '2' },
      flowDomainId: 'endnotes:page:0',
    });
    expect(endnote!.flowBounds.yPt).toBeGreaterThanOrEqual(
      page.layers.body[0]!.flowBounds.yPt + page.layers.body[0]!.flowBounds.heightPt,
    );
    expect(page.readingOrder).toEqual([page.layers.body[0]!.id, endnote!.id]);
    if (endnote?.kind !== 'note') throw new Error('Expected retained endnote');
    const storyParagraph = endnote.story.blocks[0];
    expect(storyParagraph?.kind).toBe('paragraph');
    if (storyParagraph?.kind !== 'paragraph') throw new Error('Expected retained endnote paragraph');
    expect(storyParagraph.lines.flatMap((line) => line.placements)
      .find((placement) => placement.kind === 'text' && placement.noteReference))
      .toMatchObject({ kind: 'text', text: '1' });
  });

  it('retains the entire document-end note when it needs additional pages', () => {
    const section = {
      pageWidth: 80, pageHeight: 50,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section,
      body: [endnoteReferenceParagraph('1')],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [],
      endnotes: [{ id: '1', content: [ordinaryParagraph('endnote '.repeat(80))] }],
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.diagnostics).toEqual([]);
    expect(layout.pages.length).toBeGreaterThan(1);
    const painted = layout.pages.flatMap(page => page.layers.notes)
      .flatMap(note => note.kind === 'note' ? note.story.blocks : [])
      .flatMap(block => block.kind === 'paragraph' ? block.lines : [])
      .flatMap(line => line.placements.flatMap(placement => placement.kind === 'text' ? [placement.text] : []))
      .join('');
    expect(painted.replace(/\s/g, '')).toBe('endnote'.repeat(80));
  });

  it('surfaces unsupported authored note positions instead of silently treating them as defaults', () => {
    const section = {
      pageWidth: 200, pageHeight: 140,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section,
      body: [paragraph() as unknown as BodyElement, endnoteReferenceParagraph('1')],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [ordinaryParagraph('footnote')] }],
      endnotes: [{ id: '1', content: [ordinaryParagraph('endnote')] }],
      fontFamilyClasses: {},
      __noteLayoutSettings: {
        footnotePosition: 'beneathText',
        endnotePosition: 'sectEnd',
      },
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages.flatMap((page) => page.layers.notes).map((note) => note.source.story))
      .toEqual(['endnote', 'footnote']);
    expect(layout.diagnostics.map((diagnostic) => diagnostic.message)).toEqual([
      expect.stringContaining('Unsupported footnote position "beneathText"'),
      expect.stringContaining('Unsupported endnote position "sectEnd"'),
    ]);
  });

  it('retains parser-model bookmark starts through production pagination', () => {
    const anchored = ordinaryParagraph('destination');
    anchored.bookmarks = ['destination', 'alias'];
    const model = sectionBoundaryModel(
      anchored as unknown as BodyElement,
      'nextPage',
    );
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages[0]?.bookmarkStarts.map(({ name }) => name))
      .toEqual(['destination', 'alias']);
    expect([...buildBookmarkPageMap(layout)]).toEqual([
      ['destination', 0],
      ['alias', 0],
    ]);
  });

  it('derives bookmark page metadata after header stories join the retained graph', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const headerDestination = ordinaryParagraph('header destination');
    headerDestination.bookmarks = ['header-destination'];
    const model = {
      section,
      body: [ordinaryBodyParagraph('body')],
      headers: {
        default: { body: [headerDestination as unknown as BodyElement] },
        first: null,
        even: null,
      },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages[0]?.bookmarkStarts).toEqual([
      expect.objectContaining({
        name: 'header-destination',
        sectionOccurrenceId: layout.pages[0]?.sectionOccurrenceId,
      }),
    ]);
    expect([...buildBookmarkPageMap(layout)]).toEqual([['header-destination', 0]]);
  });

  it('composes compatible continuous sections as disjoint regions on one physical page', () => {
    const model = sectionBoundaryModel(ordinaryBodyParagraph('before'), 'continuous');
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const page = layout.pages[0]!;
    const [outgoing, incoming] = page.sectionRegions;

    expect(layout.pages).toHaveLength(1);
    expect(page.sectionRegions).toHaveLength(2);
    expect(outgoing!.blockEndPt).toBe(incoming!.blockStartPt);
    expect(outgoing!.blockEndPt).toBeGreaterThan(outgoing!.blockStartPt);
    expect(page.layers.body.map((node) => node.flowDomainId)).toEqual([
      outgoing!.flowDomainIds[0], incoming!.flowDomainIds[0],
    ]);
  });

  it('retains a floating table in an empty outgoing region without charging incoming flow', () => {
    const model = sectionBoundaryModel(floatingTable(), 'continuous');
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const page = layout.pages[0]!;
    const [outgoing, incoming] = page.sectionRegions;
    const [floating, follower] = page.layers.body;

    expect(layout.pages).toHaveLength(1);
    expect(page.sectionRegions).toHaveLength(2);
    expect(outgoing!.blockStartPt).toBe(outgoing!.blockEndPt);
    expect(outgoing!.blockEndPt).toBe(incoming!.blockStartPt);
    expect(floating).toMatchObject({
      kind: 'table', ordinaryFlow: false, flowDomainId: outgoing!.flowDomainIds[0],
    });
    expect(follower).toMatchObject({
      kind: 'paragraph', ordinaryFlow: true, flowDomainId: incoming!.flowDomainIds[0],
      flowBounds: { yPt: incoming!.blockStartPt },
    });
    expect(page.flowDomains.find((domain) => domain.id === outgoing!.flowDomainIds[0]))
      .toMatchObject({ logicalBounds: { heightPt: 0 } });
  });

  it('keeps a non-continuous control on separate physical pages', () => {
    const model = sectionBoundaryModel(floatingTable(), 'nextPage');
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map((page) => page.sectionRegions.length)).toEqual([1, 1]);
    expect(layout.pages[0]!.layers.body[0]).toMatchObject({
      kind: 'table', ordinaryFlow: false,
      flowDomainId: layout.pages[0]!.sectionRegions[0]!.flowDomainIds[0],
    });
    expect(layout.pages[1]!.layers.body[0]).toMatchObject({
      kind: 'paragraph', ordinaryFlow: true,
      flowDomainId: layout.pages[1]!.sectionRegions[0]!.flowDomainIds[0],
    });
  });

  it('advances a single-column nextColumn section to its own next-page geometry', () => {
    const outgoing = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5,
    };
    const incoming = {
      pageWidth: 300, pageHeight: 160,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 8, footerDistance: 8, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextColumn', columns: null,
    } as SectionProps;
    const endingSection = {
      type: 'sectionBreak',
      kind: 'nextPage',
      geom: outgoing,
      columns: null,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      titlePage: false,
    } as unknown as BodyElement;
    const model = {
      section: incoming,
      body: [
        ordinaryBodyParagraph('before'),
        endingSection,
        ordinaryBodyParagraph('after'),
      ],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(2);
    expect(layout.pages.map((page) => ({
      widthPt: page.geometry.widthPt,
      heightPt: page.geometry.heightPt,
    }))).toEqual([
      { widthPt: 200, heightPt: 100 },
      { widthPt: 300, heightPt: 160 },
    ]);
    expect(layout.pages.map((page) => page.sectionRegions.length)).toEqual([1, 1]);
  });

  it('retains page-owned wrap authority across a same-page nextColumn section cutover', () => {
    const columns = { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] };
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextColumn', columns,
    } as SectionProps;
    const endingSection = {
      type: 'sectionBreak',
      kind: 'nextPage',
      geom: { ...section, sectionStart: 'nextPage' },
      columns,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      titlePage: false,
    } as unknown as BodyElement;
    const model = {
      section,
      body: [ordinaryBodyParagraph('before'), endingSection, pageOwnedWrappingParagraph()],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(1);
    expect(layout.pages[0]!.sectionRegions).toHaveLength(2);
    expect(layout.pages[0]!.layers.body.some((node) =>
      node.kind === 'paragraph' && node.drawings.length > 0)).toBe(true);
  });

  it.each([
    ['LTR left positive', 'left', 12, false, 22],
    ['LTR left negative', 'left', -12, false, -2],
    ['LTR center positive', 'center', 12, false, 92],
    ['LTR center negative', 'center', -12, false, 68],
    ['LTR right positive', 'right', 12, false, 162],
    ['LTR right negative', 'right', -12, false, 138],
    ['RTL left positive', 'left', 12, true, 138],
    ['RTL left negative', 'left', -12, true, 162],
    ['RTL center positive', 'center', 12, true, 68],
    ['RTL center negative', 'center', -12, true, 92],
    ['RTL right positive', 'right', 12, true, -2],
    ['RTL right negative', 'right', -12, true, 22],
  ] as const)(
    'projects parser-owned tblInd placement for %s',
    (_name, justification, indentPt, bidiVisual, expectedXPt) => {
      const section = {
        pageWidth: 200, pageHeight: 100,
        marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
        headerDistance: 5, footerDistance: 5, titlePage: false,
        evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
      } as SectionProps;
      const model = {
        section, body: [ordinaryTable(justification, indentPt, bidiVisual)],
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        footnotes: [], endnotes: [], fontFamilyClasses: {},
      } as unknown as DocxDocumentModel;
      const services = createLayoutServices(model, { measureContext: measureContext() });

      const retained = layoutDocument(model, services, { currentDateMs: 0 })
        .pages[0]?.layers.body[0];

      expect(retained?.kind).toBe('table');
      if (retained?.kind !== 'table') throw new Error('expected retained table');
      expect(retained.flowBounds.xPt).toBe(expectedXPt);
      expect(retained.rows[0]?.flowBounds.xPt).toBe(expectedXPt);
    },
  );

  it('retains the note-reference id on the destination-page text placement', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section, body: [paragraph()],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [paragraph()] }],
      endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const placement = layout.pages[0]!.layers.body
      .filter((node) => node.kind === 'paragraph')
      .flatMap((node) => node.kind === 'paragraph' ? node.lines : [])
      .flatMap((line) => line.placements)
      .find((candidate) => candidate.kind === 'text');

    expect(placement).toMatchObject({ noteReference: { kind: 'footnote', id: '7' } });
  });

  it('removes only the explicitly empty footnote separator rule', () => {
    const section = {
      pageWidth: 200, pageHeight: 140,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const base = {
      section, body: [paragraph()],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [ordinaryParagraph('note')] }],
      endnotes: [], fontFamilyClasses: {},
    };
    const render = (separator?: string) => {
      const model = {
        ...base, __noteLayoutSettings: { footnoteSeparator: separator },
      } as unknown as DocxDocumentModel;
      const services = createLayoutServices(model, { measureContext: measureContext() });
      return layoutDocument(model, services, { currentDateMs: 0 })
        .pages.flatMap((page) => page.layers.notes)[0];
    };
    const standard = render();
    const empty = render('none');
    expect(standard?.kind).toBe('note');
    expect(empty?.kind).toBe('note');
    if (standard?.kind !== 'note' || empty?.kind !== 'note') {
      throw new Error('Expected retained footnote');
    }
    expect(standard?.separator).toHaveLength(1);
    expect(empty?.separator).toHaveLength(0);
    expect(empty?.flowBounds).toEqual(standard?.flowBounds);
  });

  it('uses the authored continuation rule width on a retained note fragment', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const render = (style?: 'short' | 'full' | 'none', ordinary?: 'short' | 'none', count = 8, longBody = false, leading = 0) => {
      const bodyParagraph = paragraph();
      if (longBody) bodyParagraph.runs = [...bodyParagraph.runs, ...ordinaryParagraph('body '.repeat(100)).runs];
      const model = {
        section, body: [...Array.from({ length: leading }, () => ordinaryParagraph('body')), bodyParagraph],
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        footnotes: [{ id: '7', content: Array.from({ length: count }, () => ordinaryParagraph('note')) }],
        endnotes: [], fontFamilyClasses: {},
        __noteLayoutSettings: { footnoteContinuationSeparator: style, footnoteSeparator: ordinary },
      } as unknown as DocxDocumentModel;
      const services = createLayoutServices(model, {
        measureContext: measureContext(), allowFootnoteContinuation: true,
      });
      return layoutDocument(model, services, { currentDateMs: 0 }).pages
        .flatMap((page) => page.layers.notes)
        .filter((note) => note.kind === 'note');
    };
    const relocated = render('full', undefined, 2, false, 6);
    expect(relocated[0]?.flowDomainId).toBe('notes:page:1');
    expect(relocated[0]?.story.blocks.every(block =>
      block.flowDomainId === 'notes:page:1:footnote:7')).toBe(true);
    const short = render('short', 'short');
    const full = render('full');
    expect(short.length).toBeGreaterThan(1);
    expect(full).toHaveLength(short.length);
    const width = (note: (typeof short)[number]) => {
      const edge = note.separator[0];
      return edge ? edge.to.xPt - edge.from.xPt : 0;
    };
    expect(width(short[0]!)).toBeCloseTo(width(full[0]!), 5);
    expect(width(short[1]!)).toBeCloseTo(60, 5);
    expect(width(full[1]!)).toBeCloseTo(180, 5);
    // A missing story retains the full-width continuation fallback; it must
    // not inherit the ordinary Short rule or become an explicit empty rule.
    const absent = render();
    expect(absent.map(width)).toEqual(full.map(width));
    const emptyFirst = render('full', 'none');
    expect(width(emptyFirst[0]!)).toBe(0);
    expect(width(emptyFirst[1]!)).toBeCloseTo(180, 5);
    const emptyContinuation = render('none');
    expect(width(emptyContinuation[0]!)).toBeCloseTo(60, 5);
    expect(width(emptyContinuation[1]!)).toBe(0);
    expect(emptyContinuation.map(note => note.advancePt)).toEqual(full.map(note => note.advancePt));
    // A note taller than a physical page must be acquired before partitioning.
    const long = render('full', undefined, 30);
    expect(long.length).toBeGreaterThan(2);
    expect(long.flatMap(note => note.story.blocks).filter(block => block.kind === 'paragraph')).toHaveLength(30);
    const continuedBody = render('full', undefined, 30, true);
    expect(continuedBody.length).toBeGreaterThan(2);
    expect(continuedBody.flatMap(note => note.story.blocks).filter(block => block.kind === 'paragraph')).toHaveLength(30);
  });

  it.each([{ prefix: 0, rejects: true }, { prefix: 35, rejects: false }])(
    'validates the source cut after a page-field reflow (prefix $prefix)', ({ prefix, rejects }) => {
    const note = ordinaryParagraph('x'.repeat(prefix));
    note.runs = [
      ...note.runs,
      { ...ordinaryParagraph('field').runs[0], type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '1' },
      ...ordinaryParagraph(' tail'.repeat(80)).runs,
    ] as DocParagraph['runs'];
    const model = {
      section: { pageWidth: 200, pageHeight: 100, marginTop: 10, marginBottom: 10,
        marginLeft: 10, marginRight: 10, headerDistance: 5, footerDistance: 5,
        titlePage: false, evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
        pageNumType: { start: 9, format: 'decimal' } },
      body: [paragraph()], headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [note] }], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: true,
    });
    // PAGE 9 -> 10 moves a word across the old line-six cut. Resuming at the
    // same shaped line would paint 81 copies of "tail" from a source of 80.
    if (rejects) {
      expect(() => layoutDocument(model, services, { currentDateMs: 0 }))
        .toThrow('Footnote continuation source cut changed during reflow');
    } else {
      const layout = layoutDocument(model, services, { currentDateMs: 0 });
      const text = layout.pages.flatMap(page => page.layers.notes).flatMap(note =>
        note.kind === 'note' ? note.story.blocks.flatMap(block => block.kind === 'paragraph'
          ? block.lines.flatMap(line => line.placements.flatMap(placement => placement.kind === 'text'
            ? [placement.text] : [])) : []) : []).join('');
      expect(text).toBe('x'.repeat(prefix) + '9' + ' tail'.repeat(80));
    }
  });

  it('rejects an over-budget note before shaping its source text', () => {
    const model = { section: {
      pageWidth: 200, pageHeight: 100, marginTop: 10, marginBottom: 10,
      marginLeft: 10, marginRight: 10, headerDistance: 5, footerDistance: 5,
      titlePage: false, evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    }, body: [paragraph()], headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '7', content: [ordinaryParagraph('x'.repeat(1_000_001))] }],
      endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const context = measureContext();
    const measure = context.measureText;
    context.measureText = text => {
      if (text.length > 100) throw new Error('Over-budget source reached shaping');
      return measure(text);
    };
    const services = createLayoutServices(model, {
      measureContext: context, allowFootnoteContinuation: true,
    });
    expect(() => layoutDocument(model, services, { currentDateMs: 0 }))
      .toThrow('Footnote acquisition source budget exceeded');
  });

  it('retains a frame paragraph as an out-of-flow placed occurrence', () => {
    const framed = paragraph();
    framed.framePr = {
      dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
      hRule: 'auto', hSpace: 0, vSpace: 0, w: 40, h: 20, x: 15, y: 5,
    };
    const following = paragraph();
    (following.runs[0] as { noteRef?: unknown }).noteRef = undefined;
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section, body: [framed, following],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const body = layout.pages[0]!.layers.body;

    const frame = body[0]!;
    const follower = body[1]!;
    expect(frame.kind).toBe('paragraph');
    if (frame.kind !== 'paragraph') throw new Error('expected retained frame paragraph');
    expect(frame.ordinaryFlow).toBe(false);
    expect(frame.advancePt).toBe(0);
    expect(frame.flowBounds.heightPt).toBeCloseTo(
      frame.spacing.beforePt
        + fragmentLineAdvancesPt(frame)
        + frame.spacing.afterPt,
      6,
    );
    expect(frame.flowBounds).toMatchObject({ xPt: 25, yPt: 15 });
    expect(follower.flowBounds.yPt).toBe(10);
  });

  it.each([{ kind: 'fixed', grow: false }, { kind: 'autofit', grow: false }, { kind: 'autofit', grow: true }])('centers a complete RTL floating frame independently of acquired origin: $kind/$grow', ({ kind, grow }) => {
    const results = [-10, 0, 10].map((indentPt) => {
      const table = floatingTable() as Extract<BodyElement, { type: 'table' }>;
      table.colWidths = [20, 40];
      table.tblInd = indentPt;
      table.bidiVisual = true;
      table.layout = kind;
      table.widthPt = grow ? undefined : 60;
      const cell = table.rows[0]!.cells[0]!;
      table.rows = [{ ...table.rows[0]!, cells: [
        { ...cell, widthPt: grow ? null : 20, noWrap: grow, content: [{ ...ordinaryParagraph(grow ? 'ABCDEFGHIJKLMNOPQRSTUVWXYZABCDEFGHIJKLMNOPQRSTUVWXYZ' : 'A'), type: 'paragraph' } as CellElement] },
        { ...cell, widthPt: grow ? null : 40, content: [{ ...ordinaryParagraph('B'), type: 'paragraph' } as CellElement] },
      ] }, { ...table.rows[0]!, cells: [{ ...cell, widthPt: grow ? null : 60, colSpan: 2, content: [{ ...ordinaryParagraph('C'), type: 'paragraph' } as CellElement] }] }];
      table.tblpPr = { leftFromText: 5, rightFromText: 10, topFromText: 2, bottomFromText: 4,
        horzAnchor: 'margin', horzSpecified: true, vertAnchor: 'page', tblpX: 0, tblpXSpec: 'center', tblpY: 20 };
      const model = {
        section: { pageWidth: 200, pageHeight: 100, marginTop: 10, marginRight: 20, marginBottom: 10,
          marginLeft: 10, headerDistance: 5, footerDistance: 5, titlePage: false, evenAndOddHeaders: false,
          sectionStart: 'nextPage', columns: null },
        body: [table, ordinaryBodyParagraph('following')],
        headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
        footnotes: [], endnotes: [], fontFamilyClasses: {},
      } as unknown as DocxDocumentModel;
      const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), { currentDateMs: 0 });
      const frame = layout.pages[0]!.layers.body[0]!;
      if (frame.kind !== 'table') throw new Error('expected floating frame');
      expect(frame.ordinaryFlow).toBe(false);
      if (grow) expect(frame.flowBounds.widthPt).toBeGreaterThan(60);
      expect(frame.flowBounds.xPt + frame.flowBounds.widthPt / 2).toBeCloseTo(95);
      expect(frame.rows[0]!.cells[0]!.flowBounds.xPt).toBeGreaterThan(frame.rows[0]!.cells[1]!.flowBounds.xPt);
      expect(layout.pages[0]!.layers.body[1]!.flowBounds.yPt).toBe(10);
      return { frame: frame.flowBounds, rows: frame.rows.map((row) => ({ bounds: row.flowBounds,
        cells: row.cells.map((cell) => ({ bounds: cell.flowBounds, clips: cell.clipBounds,
          blocks: cell.blocks })) })), borders: frame.borders };
    });
    expect(results[0]).toEqual(results[1]);
    expect(results[2]).toEqual(results[1]);
  });

  it('retains an effective positioned table without charging ordinary flow', () => {
    const following = paragraph();
    (following.runs[0] as { noteRef?: unknown }).noteRef = undefined;
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const model = {
      section, body: [floatingTable(), following],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(model, { measureContext: measureContext() });

    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    const body = layout.pages[0]!.layers.body;

    const floating = body[0]!;
    const follower = body[1]!;
    expect(floating.kind).toBe('table');
    if (floating.kind !== 'table') throw new Error('expected retained floating table');
    expect(floating.ordinaryFlow).toBe(false);
    expect(floating.advancePt).toBeCloseTo(
      floating.rows.reduce((sum, row) => sum + row.advancePt, 0),
      6,
    );
    expect(floating.advancePt).toBeGreaterThan(0);
    expect(floating.flowBounds).toMatchObject({ xPt: 30, yPt: 20 });
    expect(follower.flowBounds.yPt).toBe(10);
  });

  it.each([
    { name: 'default top distance with an empty paragraph', topFromText: 0, emptyParagraph: true },
    { name: 'positive top distance without an empty paragraph', topFromText: 20, emptyParagraph: false },
  ])('keeps preceding text lines outside a page-positioned floating table: $name', ({
    topFromText, emptyParagraph,
  }) => {
    const section = {
      pageWidth: 200, pageHeight: 120,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const table = floatingTable() as Extract<BodyElement, { type: 'table' }>;
    table.colWidths = [180];
    table.rows[0]!.cells[0]!.widthPt = 180;
    table.tblpPr = {
      ...table.tblpPr!,
      horzAnchor: 'margin', vertAnchor: 'page', tblpX: 0, tblpY: 25,
      topFromText,
    };
    const tableIndex = emptyParagraph ? 2 : 1;
    const model = {
      section,
      body: [ordinaryBodyParagraph('before '.repeat(12)),
        ...(emptyParagraph ? [ordinaryBodyParagraph('')] : []), table],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), {
      currentDateMs: 0,
    });
    const body = layout.pages[0]!.layers.body;
    const heading = body.find((node) => node.kind === 'paragraph' && node.source.path[0] === 0);
    const placedTable = body.find((node) => node.kind === 'table'
      && node.source.path[0] === tableIndex);
    expect(heading?.kind).toBe('paragraph');
    expect(placedTable?.kind).toBe('table');
    if (heading?.kind !== 'paragraph' || placedTable?.kind !== 'table') return;
    expect(placedTable.flowBounds.yPt).toBe(25);
    expect(heading.lines.length).toBeGreaterThan(1);
    expect(heading.lines.every((line) => (
      line.bounds.yPt + line.bounds.heightPt <= placedTable.flowBounds.yPt
      || line.bounds.yPt >= placedTable.flowBounds.yPt + placedTable.flowBounds.heightPt
    ))).toBe(true);
    if (topFromText > 0) {
      expect(heading.lines.every((line) => line.bounds.yPt >= placedTable.flowBounds.yPt
        + placedTable.flowBounds.heightPt)).toBe(true);
    }
  });

  it('does not apply a later section page table to the preceding section', () => {
    const section = {
      pageWidth: 200, pageHeight: 120,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const table = floatingTable() as Extract<BodyElement, { type: 'table' }>;
    table.colWidths = [180];
    table.rows[0]!.cells[0]!.widthPt = 180;
    table.tblpPr = {
      ...table.tblpPr!,
      horzAnchor: 'margin', vertAnchor: 'page', tblpX: 0, tblpY: 25,
    };
    const breakMark = {
      ...sectionBreak(), kind: 'nextPage',
      geom: { ...section, sectionStart: 'nextPage' },
    } as unknown as BodyElement;
    const before = ordinaryBodyParagraph('before '.repeat(12));
    const model = (body: BodyElement[]) => ({
      section, body,
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    }) as unknown as DocxDocumentModel;
    const control = model([before]);
    const withLaterTable = model([before, breakMark, table]);
    const controlLayout = layoutDocument(control,
      createLayoutServices(control, { measureContext: measureContext() }), { currentDateMs: 0 });
    const candidateLayout = layoutDocument(withLaterTable,
      createLayoutServices(withLaterTable, { measureContext: measureContext() }), { currentDateMs: 0 });
    const firstParagraph = (layout: ReturnType<typeof layoutDocument>) => layout.pages[0]!
      .layers.body.find((node) => node.kind === 'paragraph' && node.source.path[0] === 0);
    const controlParagraph = firstParagraph(controlLayout);
    const candidateParagraph = firstParagraph(candidateLayout);
    expect(controlParagraph?.kind).toBe('paragraph');
    expect(candidateParagraph?.kind).toBe('paragraph');
    if (controlParagraph?.kind !== 'paragraph' || candidateParagraph?.kind !== 'paragraph') return;
    expect(candidateParagraph.lines.map((line) => line.bounds.yPt))
      .toEqual(controlParagraph.lines.map((line) => line.bounds.yPt));
    expect(candidateLayout.pages[1]!.layers.body.some((node) => node.kind === 'table')).toBe(true);
  });

  it('keeps a continued floating table on its accepted page fragments', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const table = floatingTable() as Extract<BodyElement, { type: 'table' }>;
    table.colWidths = [180];
    table.rows = Array.from({ length: 8 }, (_, index) => ({
      ...table.rows[0]!,
      cells: [{
        ...table.rows[0]!.cells[0]!,
        widthPt: 180,
        content: [{ type: 'paragraph' as const, ...ordinaryParagraph(`row ${index} `.repeat(14)) }],
      }],
    }));
    table.tblpPr = {
      ...table.tblpPr!,
      horzAnchor: 'margin', vertAnchor: 'page', tblpX: 0, tblpY: 25,
    };
    const model = {
      section, body: [ordinaryBodyParagraph('before '.repeat(8)), table],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), {
      currentDateMs: 0,
    });
    const fragments = layout.pages.flatMap((page) => page.layers.body
      .filter((node) => node.kind === 'table' && node.source.path[0] === 1)
      .map((node) => ({ pageIndex: page.pageIndex, node })));
    expect(fragments.length).toBeGreaterThan(1);
    expect(fragments[0]!.pageIndex).toBe(1);
    expect(fragments[0]!.node.sectionFlowOwnership).toBe('page');
    expect(fragments.slice(1).every(({ node }) => node.sectionFlowOwnership !== 'page')).toBe(true);
    expect(new Set(fragments.map(({ pageIndex }) => pageIndex)).size).toBe(fragments.length);
  });

  it('rechecks page-owned table placement after a visible header changes the body band', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 30, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const table = floatingTable() as Extract<BodyElement, { type: 'table' }>;
    table.colWidths = [180];
    table.rows[0]!.cells[0]!.widthPt = 180;
    table.tblpPr = {
      ...table.tblpPr!,
      horzAnchor: 'margin', vertAnchor: 'page', tblpX: 0, tblpY: 65,
    };
    const model = {
      section,
      body: [ordinaryBodyParagraph('before '.repeat(20)), table],
      headers: { default: { body: [ordinaryParagraph('Header')] }, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), {
      currentDateMs: 0,
    });
    const headingPages = layout.pages.filter((page) => page.layers.body.some((node) =>
      node.kind === 'paragraph' && node.source.path[0] === 0));
    const tablePage = layout.pages.find((page) => page.layers.body.some((node) =>
      node.kind === 'table' && node.source.path[0] === 1));
    expect(headingPages.length).toBeGreaterThan(0);
    expect(tablePage).toBeDefined();
    expect(layout.pages[0]!.geometry.contentTopPt).toBeGreaterThan(section.marginTop);
    expect(layout.pages.every((page) => page.layers.body.length > 0)).toBe(true);
    expect(tablePage!.pageIndex).toBeGreaterThanOrEqual(headingPages.at(-1)!.pageIndex);
  });

  it('does not skip a page when the table source naturally follows multi-page text', () => {
    const section = {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    } as SectionProps;
    const table = floatingTable() as Extract<BodyElement, { type: 'table' }>;
    table.tblpPr = {
      ...table.tblpPr!,
      horzAnchor: 'margin', vertAnchor: 'page', tblpX: 0, tblpY: 25,
    };
    const model = {
      section,
      body: [ordinaryBodyParagraph('before '.repeat(150)), table],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [], endnotes: [], fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), {
      currentDateMs: 0,
    });
    const tablePage = layout.pages.find((page) => page.layers.body.some((node) =>
      node.kind === 'table' && node.source.path[0] === 1));
    const lastTextPage = [...layout.pages].reverse().find((page) => page.layers.body.some((node) =>
      node.kind === 'paragraph' && node.source.path[0] === 0));
    expect(layout.pages.length).toBeGreaterThanOrEqual(3);
    expect(layout.pages.every((page) => page.layers.body.length > 0)).toBe(true);
    expect(tablePage).toBeDefined();
    expect(lastTextPage).toBeDefined();
    expect(tablePage!.pageIndex).toBeGreaterThanOrEqual(lastTextPage!.pageIndex);
  });
});

/** Measurement whose vertical metrics follow the requested CSS font size, so
 * a participant's own size is observable in the acquired line box. */
function sizedMeasureContext(): CanvasRenderingContext2D {
  const context = {
    font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
    measureText(text: string) {
      const px = Number(/([0-9.]+)px/.exec(context.font)?.[1] ?? 10);
      return {
        width: [...text].length * px / 2,
        actualBoundingBoxAscent: px * 0.8, actualBoundingBoxDescent: px * 0.2,
        fontBoundingBoxAscent: px * 0.8, fontBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  };
  return context as unknown as CanvasRenderingContext2D;
}

/** Reserved separator wire, shaped as a producer emits it. */
const reservedRun = (fontSize: number, color: string | null = null) => ({
  text: '', bold: false, italic: false, underline: false, strikethrough: false, fontSize, color,
  fontFamily: 'serif', isLink: false, background: null, vertAlign: null, hyperlink: null,
});
const reservedParagraph = (markPt: number, overrides: Record<string, unknown> = {}) => ({
  alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0,
  lineSpacing: null, numbering: null, tabStops: [], runs: [], defaultFontSize: markPt,
  defaultFontFamily: 'serif', widowControl: false,
  paragraphMarkFontFacts: { fontFamily: 'serif', fontSize: markPt }, ...overrides,
});
function reservedRule(mark: 'short' | 'full', options: Readonly<{
  controlPt?: number; markPt?: number; color?: string; paragraph?: Record<string, unknown>;
}> = {}) {
  const markPt = options.markPt ?? 10;
  return {
    rule: { mark, control: { run: reservedRun(options.controlPt ?? 10, options.color ?? null) } },
    paragraph: {
      paragraph: reservedParagraph(markPt, options.paragraph), contentMark: { run: reservedRun(markPt) },
    },
  };
}
const reservedNotice = () => ({
  paragraph: { paragraph: reservedParagraph(10), contentMark: { run: reservedRun(10) } },
});
const reservedEmpty = () => ({});

function reservedNotePages(
  stories: Readonly<{ separator: unknown; continuationSeparator: unknown; continuationNotice: unknown }>,
  options: Readonly<{ notes?: number; continuation?: boolean }> = {},
): NoteLayout[][] {
  const model = {
    section: {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    },
    body: [paragraph()],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [{ id: '7', content: Array.from({ length: options.notes ?? 1 }, () => ordinaryParagraph('note')) }],
    endnotes: [], fontFamilyClasses: {},
    __noteLayoutSettings: {
      footnoteSeparator: 'short', footnoteContinuationSeparator: 'full',
      reservedSeparators: { footnote: stories },
    },
  } as unknown as DocxDocumentModel;
  const services = createLayoutServices(model, {
    measureContext: sizedMeasureContext(), allowFootnoteContinuation: options.continuation === true,
  });
  return layoutDocument(model, services, { currentDateMs: 0 }).pages.map((page) =>
    page.layers.notes.filter((node): node is NoteLayout => node.kind === 'note'));
}

/** Ordinary text paragraph at an explicit size, optionally referencing a note. */
function sizedParagraph(text: string, fontSize: number, noteId?: string): DocParagraph {
  const result = ordinaryParagraph(text);
  result.defaultFontSize = fontSize;
  result.runs = [
    ...result.runs.map((run) => ({ ...run, fontSize })),
    ...(noteId ? [{ ...paragraph().runs[0]!, fontSize, noteRef: { kind: 'footnote' as const, id: noteId } }] : []),
  ] as DocParagraph['runs'];
  return result;
}

/** A long note continues from page 1 while a later body line on that page
 * references a short second note that still fits beneath it. */
function mixedNoteLayout(notice: unknown, options: Readonly<{ fillers?: number; longPt?: number }> = {}) {
  const model = {
    section: {
      pageWidth: 200, pageHeight: 100,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 5, footerDistance: 5, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    },
    body: [
      sizedParagraph('a', 4, '7'),
      ...Array.from({ length: options.fillers ?? 0 }, () => sizedParagraph('b', 4)),
      sizedParagraph('c', 4, '8'),
    ],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [
      { id: '7', content: Array.from({ length: 30 }, (_, index) => sizedParagraph(`long-${index};`, options.longPt ?? 20)) },
      { id: '8', content: [sizedParagraph('short;', 4)] },
    ],
    endnotes: [], fontFamilyClasses: {},
    __noteLayoutSettings: {
      footnoteSeparator: 'short', footnoteContinuationSeparator: 'full',
      reservedSeparators: { footnote: {
        separator: reservedRule('short'), continuationSeparator: reservedRule('full'), continuationNotice: notice,
      } },
    },
  } as unknown as DocxDocumentModel;
  return { model, services: () => createLayoutServices(model, {
    measureContext: sizedMeasureContext(), allowFootnoteContinuation: true,
  }) };
}

describe('reserved note separator consumer', () => {
  const stories = (separator: unknown) => ({
    separator, continuationSeparator: reservedRule('full'), continuationNotice: reservedEmpty(),
  });
  const leadingOf = (separator: unknown) => {
    const note = reservedNotePages(stories(separator))[0]![0]!;
    if (!note.leading) throw new Error('Expected a retained reserved separator occurrence');
    return { note, leading: note.leading };
  };

  it.each([
    // Fresh 180pt region: body 20 + reserved separator 10 + 13 lines of note 1
    // + the later note's 1-line minimum + ONE continuation notice 10 = 180.
    { pageHeight: 200, heads: [13, 1], bandPt: 160 },
    // Fresh 60pt region, exact fit: body 20 + separator 10 + one line of each
    // note + one notice = 60. Both notes then continue together on every
    // following page under the continuation separator and one notice.
    { pageHeight: 80, heads: [1, 1], bandPt: 40 },
  ])('reserves the single reserved notice once for a widow unit with two continuing references (page $pageHeight)', ({
    pageHeight, heads, bandPt,
  }) => {
    // A widow-controlled two-line paragraph references a 20-line note from
    // each line, so both lines are one admission unit. Exact 10pt lines.
    const exact = (text: string) => ({
      ...ordinaryParagraph(text), lineSpacing: { value: 10, rule: 'exact' as const, explicit: true },
    });
    const reference = (id: string) => ({ ...paragraph().runs[0]!, vertAlign: null, noteRef: { kind: 'footnote' as const, id } });
    const body = { ...exact('B01'), widowControl: true };
    body.runs = [
      reference('1'), ...body.runs, { type: 'break', breakType: 'line' }, reference('2'), ...exact('B02').runs,
    ] as DocParagraph['runs'];
    // Each note is one 20-line paragraph, so continuation resumes inside a
    // paragraph through its line cursor.
    const noteLines = (id: string) => {
      const note = exact(`N${id}-1;`);
      note.runs = Array.from({ length: 20 }, (_, index) => [
        ...(index > 0 ? [{ type: 'break', breakType: 'line' }] : []),
        ...exact(`N${id}-${index + 1};`).runs,
      ]).flat() as DocParagraph['runs'];
      return [note];
    };
    const lineCount = (note: NoteLayout) => note.story.blocks
      .reduce((sum, block) => sum + (block.kind === 'paragraph' ? block.lines.length : 0), 0);
    const notice = reservedNotice();
    const exactMark = { lineSpacing: { value: 10, rule: 'exact', explicit: true } };
    const model = {
      ...mixedNoteLayout(reservedEmpty()).model,
      section: { ...mixedNoteLayout(reservedEmpty()).model.section, pageHeight },
      body: [body],
      footnotes: [{ id: '1', content: noteLines('1') }, { id: '2', content: noteLines('2') }],
      __noteLayoutSettings: { reservedSeparators: { footnote: {
        separator: reservedRule('short', { paragraph: exactMark }),
        continuationSeparator: reservedRule('full', { paragraph: exactMark }),
        continuationNotice: { ...notice, paragraph: { ...notice.paragraph,
          paragraph: reservedParagraph(10, exactMark) } },
      } } },
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: sizedMeasureContext(), allowFootnoteContinuation: true,
    }), { currentDateMs: 0 });
    const pages = layout.pages.map((page) => page.layers.notes
      .filter((node): node is NoteLayout => node.kind === 'note'));
    expect(layout.pages[0]!.layers.body.flatMap((block) => block.kind === 'paragraph' ? block.lines : []))
      .toHaveLength(2);
    const first = pages[0]!;
    expect(first.map((note) => note.source.storyInstance)).toEqual(['1', '2']);
    expect(first.map(lineCount)).toEqual(heads);
    expect(first.map((note) => note.leading !== undefined)).toEqual([true, false]);
    expect(first.filter((note) => note.trailing)).toHaveLength(1);
    expect(first.reduce((sum, note) => sum + note.advancePt, 0)).toBe(bandPt);
    // Every source line of both notes is retained once across the pages, and
    // each page that still continues a note carries exactly one notice.
    for (const id of ['1', '2']) {
      const text = pages.flat().filter((note) => note.source.storyInstance === id)
        .flatMap((note) => note.story.blocks).flatMap((block) => block.kind === 'paragraph'
          ? block.lines.flatMap((line) => line.placements.flatMap((placement) =>
            placement.kind === 'text' ? [placement.text] : [])) : []).join('');
      expect(text).toBe(Array.from({ length: 20 }, (_, index) => `N${id}-${index + 1};`).join(''));
    }
    pages.forEach((notes, index) => {
      expect(notes.filter((note) => note.trailing)).toHaveLength(index + 1 < pages.length ? 1 : 0);
      expect(notes.filter((note) => note.leading)).toHaveLength(1);
    });
  });

  it('charges a reserved notice in the reference-line note plan', () => {
    // This is the reserved-story producer wire boundary, not a normal DOCX parser
    // fixture or an Office pagination oracle. Exact authored line heights
    // make separator, note, notice, and body ownership independently visible.
    const exact = (text: string, height = 10) => ({
      ...ordinaryParagraph(text), lineSpacing: { value: height, rule: 'exact' as const, explicit: true },
    });
    const body = exact('B01');
    body.runs = [
      { ...paragraph().runs[0]!, vertAlign: null }, ...body.runs,
      { type: 'break', breakType: 'line' }, ...exact('B02').runs,
    ] as DocParagraph['runs'];
    const layout = (remainingTwips: number, withNotice: boolean, continuation = true) => {
      const notice = reservedNotice();
      const measuredNotice = { ...notice, paragraph: { ...notice.paragraph,
        paragraph: reservedParagraph(10, { lineSpacing: { value: 10, rule: 'exact', explicit: true } }),
      } };
      const model = {
        ...mixedNoteLayout(reservedEmpty()).model,
        section: { ...mixedNoteLayout(reservedEmpty()).model.section, pageHeight: 160 + remainingTwips / 20 },
        body: [exact('P', 140), body],
        footnotes: [{ id: '7', content: Array.from({ length: 4 }, (_, index) => exact(`N${index + 1};`)) }],
        __noteLayoutSettings: { reservedSeparators: { footnote: {
          separator: reservedRule('short', { paragraph: { lineSpacing: { value: 10, rule: 'exact', explicit: true } } }),
          continuationSeparator: reservedRule('full', { paragraph: { lineSpacing: { value: 10, rule: 'exact', explicit: true } } }),
          continuationNotice: withNotice ? measuredNotice : reservedEmpty(),
        } } },
      } as unknown as DocxDocumentModel;
      return layoutDocument(model, createLayoutServices(model, {
        measureContext: sizedMeasureContext(), allowFootnoteContinuation: continuation,
      }), { currentDateMs: 0 });
    };
    const notes = (value: ReturnType<typeof layout>) => value.pages.map(page => page.layers.notes
      .filter((node): node is NoteLayout => node.kind === 'note'));
    const bodyText = (value: ReturnType<typeof layout>) => value.pages.map(page => page.layers.body
      .flatMap(block => block.kind === 'paragraph' ? block.lines.flatMap(line => line.placements
        .flatMap(placement => placement.kind === 'text' ? [placement.text] : [])) : []).join(''));
    // 50pt remains after the 140pt line. The note plan is taken at the B01
    // reference line (capacity 40) and B02 is admitted only against it.
    const first = layout(1000, true);
    expect(bodyText(first)).toEqual(['P1B01', 'B02']);
    const note = notes(first)[0]![0]!;
    expect(note.advancePt).toBe(40); // separator10 + two note lines + notice10
    expect(note.story.blocks).toHaveLength(2);
    expect(note.trailing?.advancePt).toBe(10);
    expect(notes(first).at(-1)!.at(-1)!.trailing).toBeUndefined();
    const text = notes(first).flatMap(page => page.flatMap(note => note.story.blocks))
      .flatMap(block => block.kind === 'paragraph' ? block.lines.flatMap(line => line.placements
        .flatMap(placement => placement.kind === 'text' ? [placement.text] : [])) : []).join('');
    expect(text).toBe('N1;N2;N3;N4;');
    // One twip less: the notice-reserved partition keeps one real line.
    const lower = notes(layout(999, true))[0]![0]!;
    expect(lower.advancePt).toBe(30);
    expect(lower.story.blocks).toHaveLength(1);
    expect(lower.trailing?.advancePt).toBe(10);
    // Without a notice paragraph the same capacity holds three real lines.
    const withoutNotice = notes(layout(1000, false))[0]![0]!;
    expect(withoutNotice.story.blocks).toHaveLength(3);
    expect(withoutNotice.trailing).toBeUndefined();
    expect(layout(1000, true, false).pages[0]!.layers.notes).toEqual([]);
  });

  it('consumes control, content-mark and paragraph properties through the paragraph pipeline', () => {
    const { note, leading } = leadingOf(reservedRule('short'));
    // The reserved occurrence owns the rule; the modern scalar band is unused.
    expect(note.separator).toEqual([]);
    expect(note.advancePt).toBeCloseTo(leading.advancePt + note.story.advancePt);
    expect(note.story.flowBounds.yPt).toBeCloseTo(leading.flowBounds.yPt + leading.advancePt);
    expect(leading.source).toEqual({ story: 'footnote', storyInstance: 'reserved:separator', path: [] });
    // Both characters are zero-advance, inkless metric participants.
    expect(leading.paragraph?.lines.flatMap((line) => line.placements.map((placement) => placement.kind)))
      .toEqual(['anchor-host', 'anchor-host']);
    // Each participant's own size reaches the shared line box.
    expect(leadingOf(reservedRule('short', { controlPt: 30 })).leading.advancePt)
      .toBeGreaterThan(leading.advancePt + 5);
    expect(leadingOf(reservedRule('short', { markPt: 30 })).leading.advancePt)
      .toBeGreaterThan(leading.advancePt + 5);
    // Authored paragraph spacing is real flow, charged once.
    expect(leadingOf(reservedRule('short', { paragraph: { spaceAfter: 12 } })).leading.advancePt)
      .toBeCloseTo(leading.advancePt + 12);
    // Indentation moves the paragraph's own placements; the rule keeps the
    // existing W/3 Short span from the main-story start.
    const indented = leadingOf(reservedRule('short', { paragraph: { indentLeft: 20 } })).leading;
    expect(indented.paragraph?.lines[0]?.placements[0]?.bounds?.xPt)
      .toBeCloseTo((leading.paragraph?.lines[0]?.placements[0]?.bounds?.xPt ?? 0) + 20);
    expect(indented.rule?.segment).toEqual(leading.rule?.segment);
    expect(leading.rule).toMatchObject({
      mark: 'short',
      source: { storyInstance: 'reserved:separator', path: [0, 0] },
      segment: { color: '#000000', widthPt: 0.5 },
    });
    expect(leading.rule!.segment.to.xPt - leading.rule!.segment.from.xPt).toBeCloseTo(60);
    const line = leading.paragraph!.lines[0]!;
    expect(leading.rule!.segment.from.yPt).toBeCloseTo(line.bounds.yPt + line.bounds.heightPt / 2);
    // The control's own colour is the rule ink.
    expect(leadingOf(reservedRule('full', { color: 'FF0000' })).leading.rule?.segment)
      .toMatchObject({ color: '#FF0000' });
  });

  it('owns one role-selected rule per page and a notice only on continuing fragments', () => {
    const render = (notice: unknown) => reservedNotePages({
      separator: reservedRule('short'),
      // An explicit Short continuation mark is kept independently of its role.
      continuationSeparator: reservedRule('short'),
      continuationNotice: notice,
    }, { notes: 8, continuation: true });
    const pages = render(reservedNotice());
    const notes = pages.flat();
    expect(notes.length).toBeGreaterThan(1);
    for (const page of pages) {
      expect(page.filter((note) => note.leading !== undefined).length).toBeLessThanOrEqual(1);
    }
    expect(notes[0]!.leading?.source.storyInstance).toBe('reserved:separator');
    expect(notes[1]!.leading?.source.storyInstance).toBe('reserved:continuation-separator');
    expect(notes[1]!.leading?.rule?.mark).toBe('short');
    // Source conservation across every fragment.
    expect(notes.flatMap((note) => note.story.blocks).filter((block) => block.kind === 'paragraph'))
      .toHaveLength(8);
    notes.forEach((note, index) => {
      const continues = index < notes.length - 1;
      expect(note.trailing !== undefined).toBe(continues);
      if (!note.trailing) return;
      expect(note.trailing.source.storyInstance).toBe('reserved:continuation-notice');
      expect(note.trailing.flowBounds.yPt + note.trailing.advancePt)
        .toBeCloseTo(note.flowBounds.yPt + note.advancePt);
      expect(note.advancePt).toBeCloseTo(
        (note.leading?.advancePt ?? 0) + note.story.advancePt + note.trailing.advancePt);
    });
    // The notice reserve can only shorten the admitted prefix.
    const admitted = (layout: NoteLayout[][]) => layout.flat()[0]!.story.blocks.length;
    expect(admitted(pages)).toBeLessThanOrEqual(admitted(render(reservedEmpty())));
    expect(render(reservedEmpty()).flat().every((note) => note.trailing === undefined)).toBe(true);
    // An intact note never carries a notice.
    const intact = reservedNotePages({ ...stories(reservedRule('short')), continuationNotice: reservedNotice() },
      { continuation: true });
    expect(intact.flat()).toHaveLength(1);
    expect(intact[0]![0]!.trailing).toBeUndefined();
  });
  it('measures a visible paragraph-only notice through its own content-mark properties', () => {
    const notice = (mark: Record<string, unknown> | null, overrides: Record<string, unknown> = {}) => ({
      paragraph: {
        paragraph: reservedParagraph(10, overrides),
        contentMark: mark ? { run: { ...reservedRun(10), ...mark } } : {},
      },
    });
    const trailingOf = (continuationNotice: unknown) => reservedNotePages({
      separator: reservedRule('short'), continuationSeparator: reservedRule('full'), continuationNotice,
    }, { notes: 8, continuation: true }).flat()[0]!.trailing;
    const plain = trailingOf(notice({}));
    if (!plain) throw new Error('Expected a reserved continuation notice');
    // Existing library line-box composition for this controlled fixture (not
    // an ECMA-376 §17.3.2.24 or Office formula): the raise adds to the mark's
    // own selected-face ascent instead of replacing it.
    expect(trailingOf(notice({ position: 6 }))?.advancePt).toBeCloseTo(plain.advancePt + 6);
    // §17.3.2.42 superscript reaches the mark through the existing
    // effective-size policy; no coefficient is introduced here.
    expect(trailingOf(notice({ vertAlign: 'super' }))?.advancePt).not.toBeCloseTo(plain.advancePt);
    // The retained participant carries its selected-face sides, not zeros.
    expect(plain.paragraph!.lines[0]!.placements[0]).toMatchObject({
      sourceMetrics: { ascentPt: 8, descentPt: 2 },
    });
    // The mark stays text-free and zero-advance.
    const placements = plain.paragraph!.lines.flatMap((line) => line.placements);
    expect(placements.filter((placement) => placement.kind === 'text')).toEqual([]);
    expect(placements.every((placement) => placement.bounds?.widthPt === 0)).toBe(true);
    // A vanished mark and an empty story own no notice occurrence.
    expect(trailingOf(notice(null, { markVanish: true }))).toBeUndefined();
    expect(trailingOf(reservedEmpty())).toBeUndefined();
  });

  it('composes a raised rule control or mark with the other participant unmasked', () => {
    const raised = (control: Record<string, unknown>, mark: Record<string, unknown>) => {
      const story = reservedRule('short', { controlPt: 14, markPt: 9 }) as {
        rule: { control: { run: Record<string, unknown> } };
        paragraph: { contentMark: { run: Record<string, unknown> } };
      };
      Object.assign(story.rule.control.run, control);
      Object.assign(story.paragraph.contentMark.run, mark);
      return leadingOf(story).leading;
    };
    const plain = raised({}, {});
    // Existing library composition: the 14pt control owns the line box, so a
    // 6pt raise on it adds 6pt; the same raise on the 9pt mark adds only its
    // excess over the control's ascent.
    expect(raised({ position: 6 }, {}).advancePt).toBeCloseTo(plain.advancePt + 6);
    expect(raised({}, { position: 6 }).advancePt).toBeCloseTo(plain.advancePt + 2);
    // One rule ink, no text placements, zero inline extent.
    const leading = raised({ position: 6 }, {});
    expect(leading.rule?.segment.widthPt).toBe(0.5);
    const placements = leading.paragraph!.lines.flatMap((line) => line.placements);
    expect(placements.every((placement) => placement.kind === 'anchor-host'
      && placement.bounds?.widthPt === 0)).toBe(true);
  });

  it('acquires a small-caps rule control through its transformed probe at full metric size', () => {
    const story = reservedRule('short') as { rule: { control: { run: Record<string, unknown> } } };
    story.rule.control.run.smallCaps = true;
    // Small caps transform the probe's display text; acquisition must still
    // succeed and keep the existing full-size metric policy.
    expect(leadingOf(story).leading.advancePt).toBeCloseTo(leadingOf(reservedRule('short')).leading.advancePt);
  });

  it('keeps one band-final notice when a later short note joins a continuing page', async () => {
    const notes = (layout: ReturnType<typeof layoutDocument>) => layout.pages.map((page) =>
      page.layers.notes.filter((node): node is NoteLayout => node.kind === 'note'));
    const { model, services } = mixedNoteLayout(reservedNotice());
    const layout = layoutDocument(model, services(), { currentDateMs: 0 });
    const pages = notes(layout);
    const first = pages[0]!;
    // Both references stay on page 1, in reference order.
    expect(first.map((note) => note.source.storyInstance)).toEqual(['7', '8']);
    const [continuing, later] = first as [NoteLayout, NoteLayout];
    // §17.18.33: the one notice ends the band and is charged exactly once.
    expect(continuing.trailing).toBeUndefined();
    expect(continuing.advancePt).toBeCloseTo((continuing.leading?.advancePt ?? 0) + continuing.story.advancePt);
    expect(later.trailing?.role).toBe('continuationNotice');
    expect(later.advancePt).toBeCloseTo(later.story.advancePt + later.trailing!.advancePt);
    expect(later.trailing!.flowBounds.yPt + later.trailing!.advancePt)
      .toBeCloseTo(later.flowBounds.yPt + later.advancePt);
    expect(later.trailing!.paragraph?.flowDomainId).toBe(later.flowDomainId);
    expect(first.flatMap((note) => note.trailing ? [note] : [])).toHaveLength(1);
    // Every source row of the continued note is retained once, in order.
    const text = pages.flat().filter((note) => note.source.storyInstance === '7')
      .flatMap((note) => note.story.blocks).flatMap((block) => block.kind === 'paragraph'
        ? block.lines.flatMap((line) => line.placements.flatMap((placement) =>
          placement.kind === 'text' ? [placement.text] : [])) : []).join('');
    expect(text).toBe(Array.from({ length: 30 }, (_, index) => `long-${index};`).join(''));
    expect(pages.flat().filter((note) => note.source.storyInstance === '8')).toHaveLength(1);
    // Each later page that still continues carries its own single notice.
    pages.slice(1).forEach((page, index) => {
      expect(page.filter((note) => note.trailing !== undefined))
        .toHaveLength(index + 2 < pages.length ? 1 : 0);
    });
    // The sliced driver retains the same ownership.
    const slicedSource = layoutSourceStore(mixedNoteLayout(reservedNotice()).model);
    const sliced = await layoutDocumentInputAsync(slicedSource.bodyLayoutInput,
      createLayoutServices(slicedSource, {
        measureContext: sizedMeasureContext(), allowFootnoteContinuation: true,
      }), { currentDateMs: 0 }, { sliceMs: 0, yieldToHost: async () => undefined });
    expect(sliced).toEqual(layout);
    // Without a notice paragraph nothing is reserved or painted for it; both
    // notes are still retained once.
    const empty = mixedNoteLayout(reservedEmpty());
    const plain = notes(layoutDocument(empty.model, empty.services(), { currentDateMs: 0 })).flat();
    expect(plain.every((note) => note.trailing === undefined)).toBe(true);
    expect(plain.filter((note) => note.source.storyInstance === '8')).toHaveLength(1);
  });
});

describe('over-page footnote content retention', () => {
  it.each([
    { kind: 'footnote', shape: 'rows', mode: undefined, columns: 1 },
    { kind: 'footnote', shape: 'row', mode: undefined, columns: 1 },
    { kind: 'footnote', shape: 'splittable-row', mode: undefined, columns: 1 },
    { kind: 'footnote', shape: 'mixed-row', mode: undefined, columns: 1 },
    { kind: 'footnote', shape: 'line', mode: undefined, columns: 1 },
    { kind: 'footnote', shape: 'frame', mode: undefined, columns: 1 },
    { kind: 'footnote', shape: 'band-table', mode: undefined, columns: 1 },
    { kind: 'footnote', shape: 'frame-ref', mode: undefined, columns: 1 },
    { kind: 'footnote', shape: 'table-ref', mode: undefined, columns: 1 },
    { kind: 'footnote', shape: 'boundary-line', mode: 15, columns: 2 },
    { kind: 'footnote', shape: 'rows', mode: 14, columns: 2 },
    { kind: 'footnote', shape: 'line', mode: 14, columns: 2 },
    { kind: 'footnote', shape: 'rows', mode: 15, columns: 2 },
    { kind: 'footnote', shape: 'line', mode: 15, columns: 2 },
    { kind: 'endnote', shape: 'rows', mode: undefined, columns: 1 },
    { kind: 'endnote', shape: 'line', mode: undefined, columns: 1 },
    { kind: 'endnote', shape: 'rows', mode: 15, columns: 2 },
    { kind: 'endnote', shape: 'line', mode: 15, columns: 2 },
    { kind: 'footnote', shape: 'whole', mode: undefined, columns: 1 },
    { kind: 'footnote', shape: 'whole', mode: 15, columns: 2 },
  ] as const)('retains oversized $kind $shape (mode $mode, $columns columns)', ({ kind, shape, mode, columns }) => {
    const exact = (text: string, height: number) => ({ ...ordinaryParagraph(text),
      lineSpacing: { value: height, rule: 'exact' as const, explicit: true },
    });
    const labels = shape === 'rows' ? Array.from({ length: 4 }, (_, i) => `r${i};`)
      : shape === 'splittable-row' || shape === 'mixed-row' ? Array.from({ length: 12 }, (_, i) => `l${i};`)
      : shape === 'row' ? ['huge;'] : ['oversized;'];
    const table = ordinaryTable('left', 0, false) as Extract<BodyElement, { type: 'table' }>;
    table.rows = labels.map(text => ({ ...table.rows[0]!, cantSplit: true,
      cells: [{ ...table.rows[0]!.cells[0]!, content: [{ ...exact(text, shape === 'rows' ? 50 : 120), type: 'paragraph' as const }] }],
    }));
    if (shape === 'splittable-row' || shape === 'mixed-row') table.rows = [{ ...table.rows[0]!, cantSplit: false,
      cells: [{ ...table.rows[0]!.cells[0]!, content: labels.map((text, index) =>
        ({ ...exact(text, shape === 'mixed-row' && index === 3 ? 120 : 20), type: 'paragraph' as const })) }],
    }];
    if (shape === 'band-table') table.rows[0]!.cells[0]!.content.push(floatingTable() as CellElement, { ...ordinaryParagraph('anchor;'), type: 'paragraph' });
    const tall = exact(labels[0]!, shape === 'boundary-line' ? 74 : 120);
    if (shape === 'frame') tall.framePr = {
      dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
      hRule: 'auto', hSpace: 0, vSpace: 0, w: 40, h: 120, x: 0, y: 0,
    };
    const suffix = shape === 'frame' ? Array.from({ length: 10 }, (_, i) => `after${i};`) : ['suffix;'];
    const tail = suffix.map(ordinaryBodyParagraph);
    const content: BodyElement[] = shape === 'line' || shape === 'frame' || shape === 'whole' || shape === 'boundary-line'
      ? [tall as BodyElement, ...tail]
      : [table, ...tail];
    const model = columnStoryDocument([kind === 'endnote' ? endnoteReferenceParagraph('7') as DocParagraph : paragraph()], [], 0);
    if (shape === 'frame-ref') {
      const reference = model.body[0] as DocParagraph;
      reference.framePr = {
        dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
        hRule: 'auto', hSpace: 0, vSpace: 0, w: 40, h: 10, x: 0, y: 0,
      };
    }
    if (shape === 'table-ref') {
      const referenceTable = ordinaryTable('left', 0, false) as Extract<BodyElement, { type: 'table' }>;
      referenceTable.rows[0]!.cells[0]!.content = [model.body[0] as CellElement];
      model.body = [referenceTable];
    }
    if (shape === 'frame-ref' || shape === 'table-ref') model.body.unshift(ordinaryBodyParagraph('before;'));
    if (columns === 1) model.section.columns = null;
    model.settings = { compatibilityMode: mode };
    model.headers = { default: null, first: null, even: null };
    model.footnotes = kind === 'footnote' ? [{ id: '7', content }] : [];
    model.endnotes = kind === 'endnote' ? [{ id: '7', content }] : [];
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: shape !== 'whole',
    }), { currentDateMs: 0 });
    const notes = layout.pages.flatMap(page => page.layers.notes)
      .filter((node): node is NoteLayout => node.kind === 'note' && node.source.story === kind);
    const textIn = (blocks: readonly import('./types.js').PaintNode[]): string => blocks.map(block => {
      if (block.kind === 'paragraph') return block.lines.flatMap(line => line.placements)
        .flatMap(placement => placement.kind === 'text' ? [placement.text] : []).join('');
      if (block.kind === 'table') return block.rows.flatMap(row => row.cells
        .map(cell => textIn(cell.blocks.map(block => block.layout)))).join('');
      return '';
    }).join('');
    // A positioned child in the band-dependent case is inspected separately;
    // it is page-owned, rather than a second ordinary table-cell block.
    expect(notes.map(note => textIn(note.story.blocks)).join('')).toBe(labels.join('') + (shape === 'band-table' ? 'anchor;' : '') + suffix.join(''));
    if (shape === 'mixed-row') {
      const oversized = notes.find(note => textIn(note.story.blocks).includes('l3;'))!;
      expect(textIn(oversized.story.blocks)).toBe('l3;');
      expect(notes.indexOf(oversized)).toBeLessThan(notes.length - 1);
    }
    if (shape === 'splittable-row') expect(notes.length).toBeGreaterThan(1);
    if (shape === 'rows') {
      expect(notes.length).toBeGreaterThan(1);
      expect(notes.flatMap(note => note.story.blocks).flatMap(block => block.kind === 'table'
        ? block.rows.map(row => row.source.path) : [])).toHaveLength(4);
    }
    if (shape === 'band-table') {
      expect(notes.some(note => note.story.bandDependent)).toBe(true);
      const positioned = notes.flatMap(note => note.story.blocks.flatMap(block => block.kind === 'table'
        ? block.resolvedFloatingTables ?? [] : []));
      expect(positioned).toHaveLength(1);
      expect(textIn(positioned.map(placement => placement.child))).toBe('1');
    }
    if (shape === 'frame-ref' || shape === 'table-ref') {
      expect(layout.pages[0]!.layers.notes).toHaveLength(0);
    }
    if (shape !== 'rows' && shape !== 'splittable-row' && shape !== 'mixed-row' && shape !== 'boundary-line') {
      const oversized = notes.find(note => note.story.blocks.some(block => block.source.path[0] === 0))!;
      expect(oversized.inkBounds.yPt + oversized.inkBounds.heightPt).toBeGreaterThan(100);
      expect(oversized.flowBounds.yPt).toBeGreaterThanOrEqual(10);
    }
  }, 30_000);

  it('leaves a preceding label on the reference page when the following table does not fit', () => {
    const content = Array.from({ length: 8 }, (_, index) => {
      const label = ordinaryParagraph(`label-${index};`);
      label.lineSpacing = { value: 6, rule: 'exact', explicit: true };
      const cell = ordinaryParagraph(`cell-${index};`);
      cell.lineSpacing = { value: 10, rule: 'exact', explicit: true };
      const table = ordinaryTable('left', 0, false) as Extract<BodyElement, { type: 'table' }>;
      table.rows[0]!.cells[0]!.content = [{ ...cell, type: 'paragraph' }];
      return [{ ...label, type: 'paragraph' as const }, table];
    }).flat();
    const reference = paragraph();
    reference.lineSpacing = { value: 4, rule: 'exact', explicit: true };
    const model = columnStoryDocument([reference], [], 0);
    model.section.columns = null;
    model.headers = { default: null, first: null, even: null };
    model.footnotes = [{ id: '7', content }];
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: measureContext(), allowFootnoteContinuation: true,
    }), { currentDateMs: 0 });
    const notes = layout.pages.flatMap(page => page.layers.notes)
      .filter((node): node is NoteLayout => node.kind === 'note');
    // 80pt band: 4pt reference + 6pt separator + four (6+10)pt
    // label/table pairs + one 6pt label. The fifth table resumes next page.
    expect(notes[0]!.story.blocks.map(block => block.kind)).toEqual([
      'paragraph', 'table', 'paragraph', 'table', 'paragraph', 'table', 'paragraph', 'table', 'paragraph',
    ]);
    expect(notes[1]!.story.blocks[0]?.kind).toBe('table');
    expect(notes.flatMap(note => note.story.blocks).map(block => block.source.path[0]))
      .toEqual(Array.from({ length: 16 }, (_, index) => index));
  });

  it.each(['paragraphs', 'tables'] as const)(
    'continues a footnote taller than a physical page made of %s without losing content',
    kind => {
      const labels = Array.from({ length: 20 }, (_, index) => `note-${index};`);
      const content = labels.map(text => {
        if (kind === 'paragraphs') return ordinaryBodyParagraph(text);
        const table = ordinaryTable('left', 0, false) as Extract<BodyElement, { type: 'table' }>;
        table.rows[0]!.cells[0]!.content = [{ ...ordinaryParagraph(text), type: 'paragraph' }];
        return table;
      });
      const model = columnStoryDocument([paragraph()], [], 0);
      model.section.columns = null;
      model.headers = { default: null, first: null, even: null };
      model.footnotes = [{ id: '7', content }];
      const layout = layoutDocument(model, createLayoutServices(model, {
        measureContext: measureContext(), allowFootnoteContinuation: true,
      }), { currentDateMs: 0 });
      expect(layout.pages.length).toBeGreaterThan(2);
      const notes = layout.pages.flatMap(page => page.layers.notes)
        .filter((node): node is NoteLayout => node.kind === 'note');
      const textIn = (blocks: readonly import('./types.js').PaintNode[]): string => blocks.map(block => {
        if (block.kind === 'paragraph') return block.lines.flatMap(line => line.placements)
          .flatMap(placement => placement.kind === 'text' ? [placement.text] : []).join('');
        if (block.kind === 'table') return block.rows.map(row => row.cells
          .map(cell => textIn(cell.blocks.map(block => block.layout))).join('')).join('');
        return '';
      }).join('');
      expect(notes.map(note => textIn(note.story.blocks)).join('')).toBe(labels.join(''));
      expect(layout.pages[0]!.layers.body.some(block => block.kind === 'paragraph'
        && block.lines.some(line => line.placements.some(placement => placement.kind === 'text'
          && placement.noteReference?.id === '7')))).toBe(true);
      expect(layout.pages[0]!.layers.notes).toHaveLength(1);
      for (const page of layout.pages) {
        expect(page.layers.notes).toHaveLength(1);
        const note = page.layers.notes[0]!;
        expect(note.flowBounds.yPt + note.advancePt).toBeLessThanOrEqual(90);
      }
    },
  );
});

describe('multiple pending footnote continuations', () => {
  it('preserves deferred cursors when an earlier continuation fills the page band', async () => {
    const base = mixedNoteLayout(reservedNotice());
    const model = {
      ...base.model,
      footnotes: base.model.footnotes!.map((note) => note.id === '8'
        ? { ...note, content: Array.from({ length: 30 }, (_, index) => sizedParagraph(`second-${index};`, 4)) }
        : note),
    } as DocxDocumentModel;
    const services = () => createLayoutServices(model, {
      measureContext: sizedMeasureContext(), allowFootnoteContinuation: true,
    });
    const layout = layoutDocument(model, services(), { currentDateMs: 0 });
    const pages = layout.pages.map((page) => page.layers.notes
      .filter((node): node is NoteLayout => node.kind === 'note'));
    // Both notes start with their references before their cursors are pending.
    expect(pages[0]!.map((note) => note.source.storyInstance)).toEqual(['7', '8']);
    for (const [id, prefix] of [['7', 'long'], ['8', 'second']]) {
      const text = pages.flat().filter((note) => note.source.storyInstance === id)
        .flatMap((note) => note.story.blocks).flatMap((block) => block.kind === 'paragraph'
          ? block.lines.flatMap((line) => line.placements.flatMap((placement) =>
            placement.kind === 'text' ? [placement.text] : [])) : []).join('');
      expect(text).toBe(Array.from({ length: 30 }, (_, index) => `${prefix}-${index};`).join(''));
    }
    // Every continuing page charges exactly one band-final notice. A deferred
    // note contributes no empty fragment or duplicate leading/notice reserve.
    pages.slice(0, -1).forEach((notes) => {
      expect(notes.filter((note) => note.trailing)).toHaveLength(1);
      expect(notes.at(-1)!.trailing).toBeDefined();
    });
    expect(pages.at(-1)!.every((note) => !note.trailing)).toBe(true);
    expect(pages.flat().every((note) => note.story.blocks.length > 0)).toBe(true);
    const source = layoutSourceStore(model);
    const sliced = await layoutDocumentInputAsync(source.bodyLayoutInput, createLayoutServices(source, {
      measureContext: sizedMeasureContext(), allowFootnoteContinuation: true,
    }), { currentDateMs: 0 }, { sliceMs: 0, yieldToHost: async () => undefined });
    expect(sliced).toEqual(layout);
  });
});

// Issue #1806: the endnote case is a passing control; the footnote case
// exercises the same stack arithmetic after relocation to the page bottom.
it.each([
  { kind: 'endnote' as const, count: 50, linePt: 10.3 },
  { kind: 'endnote' as const, count: 20, linePt: 25.056 },
  { kind: 'footnote' as const, count: 5, linePt: 25.056 },
])('keeps $count fractional $kind advances inside their composed domain', ({ kind, count, linePt }) => {
  const ids = Array.from({ length: count }, (_, index) => String(index + 1));
  const body = ordinaryParagraph('references');
  body.runs = ids.flatMap(id => {
    const reference = endnoteReferenceParagraph(id) as unknown as DocParagraph;
    const run = reference.runs[0];
    if (run?.type === 'text') run.noteRef = { kind, id };
    return reference.runs;
  });
  const note = ordinaryParagraph('note');
  note.lineSpacing = { value: linePt, rule: 'exact', explicit: true };
  const notes = ids.map(id => ({ id, content: [note] }));
  const model = {
    section: {
      pageWidth: 612, pageHeight: 792,
      marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36, titlePage: false,
      evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null,
    },
    body: [body], headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: kind === 'footnote' ? notes : [],
    endnotes: kind === 'endnote' ? notes : [], fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
  const layout = layoutDocument(model,
    createLayoutServices(model, { measureContext: measureContext() }), { currentDateMs: 0 });
  expect(layout.pages[0]?.layers.notes).toHaveLength(ids.length);
});

function storyCacheFixture(kind: 'header' | 'footer' = 'header') {
  const header = ordinaryParagraph('');
  header.runs = [{ type: 'field', fieldType: 'numPages', instruction: 'NUMPAGES',
    fallbackText: '?', fontSize: 10, fontFamily: 'serif', bold: false, italic: false,
    underline: false, strikethrough: false, color: null, background: null, vertAlign: null,
  }] as unknown as DocParagraph['runs'];
  const model = {
    section: { pageWidth: 200, pageHeight: 100, marginTop: 10, marginRight: 10,
      marginBottom: 10, marginLeft: 10, headerDistance: 5, footerDistance: 5,
      titlePage: false, evenAndOddHeaders: false, sectionStart: 'nextPage', columns: null },
    body: [ordinaryBodyParagraph('body')],
    headers: { default: { body: [header] }, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [], endnotes: [], fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
  if (kind === 'footer') {
    model.footers.default = model.headers.default;
    model.headers.default = null;
  }
  const source = layoutSourceStore(model);
  const services = createParagraphAcquisitionCacheServicesView(createLayoutServices(source, { measureContext: measureContext() }));
  const section = source.bodyLayoutInput.initialSection.context;
  const open = (totalPages: number, scope: LayoutServices = services, options = { currentDateMs: 0 }, destination?: PageFieldAcquisitionContext) => bodyLayoutKernelOf(scope)!.openBodyLayoutSession({
    source: source.bodyLayoutInput.source, section,
    initialLocation: { pageIndex: 0, columnIndex: 0, flowDomainId: 'body', section,
      cursorPt: { xPt: 10, yPt: 10 },
      availableBounds: { xPt: 10, yPt: 10, widthPt: 180, heightPt: 80 } },
  }, createFieldAcquisitionServicesView(scope, { totalPages, resolveDestinationPage: () => destination }), options);
  const request = (pageIndex: number) => ({
    source: (kind === 'header' ? source.bodyLayoutInput.initialSection.headers : source.bodyLayoutInput.initialSection.footers).default!, pageIndex, section,
    pageNumber: { displayNumber: pageIndex + 1, format: 'decimal' as const },
    container: { id: `story:${kind}:page:${pageIndex}`, kind,
      bounds: { xPt: 10, yPt: 0, widthPt: 180, heightPt: 100 } },
  });
  return { model, source, services, open, request };
}

it.each(['header', 'footer'] as const)('reuses complete %s stories across passes without paragraph acquisition and invalidates NUMPAGES', kind => {
  const { services, open, request } = storyCacheFixture(kind);
  const first = open(3);
  const stories = [0, 1, 2].map(page => first.layoutStory!(request(page)));
  const budget = paragraphAcquisitionCacheOf(services)!.workBudget;
  const before = budget.misses;
  expect(before).toBeGreaterThan(0);
  const repeat = open(3);
  [0, 1, 2].forEach(page => expect(repeat.layoutStory!(request(page))).toBe(stories[page]));
  expect(budget.misses).toBe(before);
  const changed = open(12).layoutStory!(request(0));
  expect(changed).not.toBe(stories[0]);
  expect(budget.misses).toBeGreaterThan(before);
  const text = (story: NonNullable<typeof changed>) => story.blocks.flatMap(block => block.kind === 'paragraph'
    ? block.lines.flatMap(line => line.placements.flatMap(p => p.kind === 'text' ? [p.text] : [])) : []).join('');
  expect(text(stories[0]!)).toBe('3');
  expect(text(changed)).toBe('12');
  const nextScope = createParagraphAcquisitionCacheServicesView(services);
  expect(open(3, nextScope).layoutStory!(request(0))).not.toBe(stories[0]);
  expect(paragraphAcquisitionCacheOf(nextScope)!.workBudget.misses).toBeGreaterThan(0);
});

it('scales source work once, bounds aggregate misses and resets the next pagination ledger', () => {
  const { source, services } = storyCacheFixture();
  // Count source work without laying out 40,000 paragraphs. Reusing immutable
  // paragraph definitions still represents 40,000 distinct source occurrences.
  const body = Object.freeze(Array(40_000).fill(source.blocks.body[0]));
  const larger = { ...source, blocks: { ...source.blocks, body,
    storyRoot: (ref: Parameters<typeof source.blocks.storyRoot>[0]) =>
      ref.story === 'body' ? body : source.blocks.storyRoot(ref),
  } };
  const capped = createPaginationWorkBudget(larger);
  expect(capped.limit).toBe(MAX_PAGINATION_ACQUISITION_MISSES);
  for (let i = 0; i < MAX_PAGINATION_ACQUISITION_MISSES; i++) capped.noteMiss();
  expect(() => capped.noteMiss()).toThrow(/operational miss budget 1000000/);
  const retry = createParagraphAcquisitionCacheServicesView(services);
  expect(paragraphAcquisitionCacheOf(retry)!.workBudget.misses).toBe(0);
});

it('credits a selected page/owner once rather than once per convergence pass or miss', () => {
  const { model } = storyCacheFixture();
  const source = layoutSourceStore({ ...model, body: Array.from({ length: 1000 }, () => ordinaryBodyParagraph('x')),
    footers: { default: { body: [ordinaryBodyParagraph('footer')] }, first: null, even: null },
  });
  const scoped = createParagraphAcquisitionCacheServicesView(createLayoutServices(source, { measureContext: measureContext() }));
  const cache = paragraphAcquisitionCacheOf(scoped)!;
  const budget = cache.workBudget;
  const before = budget.limit;
  expect(before).toBeGreaterThan(25_000);
  expect(before).toBeLessThan(MAX_PAGINATION_ACQUISITION_MISSES);
  // Exercise the runtime debit used by paragraph and nested-story acquisitions,
  // past the former fixed ceiling, without retaining a 25,001-paragraph layout.
  for (let i = 0; i < 25_001; i++) cache.noteMiss();
  const ref = source.bodyLayoutInput.initialSection.headers.default;
  budget.admitPageStories('owner', 0, [ref]);
  const headerCredit = budget.limit;
  // Header and footer can arrive separately, and a later pass can select a
  // different slot. Deduplicate roots, not the first call for this page/owner.
  const footer = source.bodyLayoutInput.initialSection.footers.default;
  budget.admitPageStories('owner', 0, [footer]);
  const credited = budget.limit;
  expect(credited).toBeGreaterThan(headerCredit);
  expect(headerCredit).toBeGreaterThan(before);
  for (let pass = 0; pass < 10; pass++) {
    budget.admitPageStories('owner', 0, [ref, footer, ref]);
    budget.noteMiss();
  }
  expect(budget.limit).toBe(credited);
  budget.admitPageStories('owner', 1, [ref]);
  expect(budget.limit).toBeGreaterThan(credited);
  const finalLimit = budget.limit;
  while (budget.misses < finalLimit) cache.noteMiss();
  expect(() => cache.noteMiss()).toThrow(/operational miss budget/);
  budget.admitPageStories('owner', 2, [ref]);
  expect(budget.limit).toBe(finalLimit);
  expect(() => cache.noteMiss()).toThrow(/operational miss budget/);
});

it.each(['section', 'container', 'band', 'frames', 'page-number', 'destination-number', 'destination-format', 'options'] as const)(
  'invalidates a page story when its %s context changes', axis => {
    const { services, open, request } = storyCacheFixture();
    const original = request(0);
    const destination = { pageIndex: 0, displayPageNumber: 1, pageNumberFormat: 'decimal' as const };
    const first = open(3, services, { currentDateMs: 0 }, destination).layoutStory!(original);
    const changed = axis === 'section' ? { ...original, section: { ...original.section, verticalAlignment: 'center' } }
      : axis === 'container' ? { ...original, container: { ...original.container,
        bounds: { ...original.container.bounds, widthPt: 90 } } }
      : axis === 'band' ? { ...original, bandTranslationPt: { xPt: 0, yPt: 0 } }
      : axis === 'frames' ? { ...original, pageFrames: { page: original.container.bounds, margin: original.container.bounds } }
      : axis === 'page-number' ? { ...original, pageNumber: { displayNumber: 12, format: 'upperRoman' as const } }
      : original;
    const result = open(3, services, { currentDateMs: axis === 'options' ? 1 : 0 },
      axis === 'destination-number' ? { ...destination, displayPageNumber: 12 }
        : axis === 'destination-format' ? { ...destination, pageNumberFormat: 'upperRoman' } : destination,
    ).layoutStory!(changed);
    expect(result).not.toBe(first);
  },
);

it.each(['reserved-note', 'textbox'] as const)('counts detached %s definitions in the pagination allowance', kind => {
  const { source } = storyCacheFixture();
  const ref = { story: kind === 'textbox' ? 'textbox' as const : 'footnote' as const,
    storyInstance: 'detached-work-root', path: [] };
  const root = Object.freeze(Array(1000).fill(source.blocks.body[0]));
  const extended = { ...source, blocks: { ...source.blocks,
    reservedNoteRoots: kind === 'reserved-note' ? [ref] : source.blocks.reservedNoteRoots,
    sources: kind === 'textbox' ? [...source.blocks.sources, { ...ref, path: [0] }] : source.blocks.sources,
    storyRoot: (request: typeof ref) => request.storyInstance === ref.storyInstance
      ? root : source.blocks.storyRoot(request),
  } };
  expect(createPaginationWorkBudget(extended).limit).toBeGreaterThan(createPaginationWorkBudget(source).limit);
});

it('admits page-story work before progressive acquisition crosses the initial allowance', async () => {
  const { model } = storyCacheFixture();
  const source = layoutSourceStore({ ...model,
    // §17.6.11 signed top margin allows overlap, isolating work admission
    // from reserve geometry. All four header paragraphs fit in the page band.
    section: { ...model.section, marginTop: -10 },
    body: Array.from({ length: 14 }, (_, index) => ({ ...ordinaryBodyParagraph('body'), pageBreakBefore: index > 0 })),
    headers: { default: { body: Array.from({ length: 4 }, () => ordinaryBodyParagraph('h')) }, first: null, even: null },
    // Unreferenced definitions cost no layout. They lift the source allowance
    // above its 25,000 floor so even a small selected story earns more credit.
    endnotes: Array.from({ length: 800 }, (_, index) => ({ id: String(index), content: [] })),
  });
  const budgets: { budget: ReturnType<typeof createPaginationWorkBudget>; initialLimit: number }[] = [];
  const createBudget = createPaginationWorkBudget;
  const factory = vi.spyOn(paginationWorkBudget, 'createPaginationWorkBudget').mockImplementation(source => {
    const budget = createBudget(source);
    const initialLimit = budget.limit;
    // Seed earlier work through the internal factory seam, leaving 16 misses:
    // the 14 body paragraphs fit, but their four-paragraph header needs credit.
    // Keep the real ledger and failure latch: admission must precede preview
    // acquisition, not repair exhaustion later.
    for (let i = 0; i < initialLimit - 16; i++) budget.noteMiss();
    budgets.push({ budget, initialLimit });
    return budget;
  });
  try {
    const open = () => createLayoutServices(source, { measureContext: measureContext() });
    const blocking = layoutDocument(source, open(), { currentDateMs: 0 });
    expect(blocking.pages).toHaveLength(14);
    let lastPreviewPages = 0;
    let previewMisses = 0;
    const progressive = await layoutDocumentProgressively(source.bodyLayoutInput, open(), { currentDateMs: 0 }, {
      onPreview: preview => {
        lastPreviewPages = preview.layout.pages.length;
        previewMisses = budgets[1]!.budget.misses;
      },
    });
    expect(budgets).toHaveLength(2);
    expect(lastPreviewPages).toBe(7);
    expect(previewMisses).toBeGreaterThan(budgets[1]!.initialLimit);
    expect(layoutFingerprint(progressive)).toBe(layoutFingerprint(blocking));
  } finally {
    factory.mockRestore();
  }
});

it.each(['footnote', 'endnote'] as const)('credits repeated whole %s acquisition before descendant debits', kind => {
  const { model } = storyCacheFixture();
  const source = layoutSourceStore({ ...model,
    section: { ...model.section, pageHeight: 100 },
    headers: { default: null, first: null, even: null },
    [kind === 'footnote' ? 'footnotes' : 'endnotes']: [
      // Lift the source allowance above its floor without acquiring unused notes.
      ...Array.from({ length: 800 }, (_, index) => ({ id: `unused-${index}`, content: [] })),
      { id: '7', content: Array.from({ length: 4 }, () => ordinaryBodyParagraph('n')) },
    ],
  });
  const services = createParagraphAcquisitionCacheServicesView(createLayoutServices(source, { measureContext: measureContext() }));
  const section = source.bodyLayoutInput.initialSection.context;
  const session = bodyLayoutKernelOf(services)!.openBodyLayoutSession({
    source: source.bodyLayoutInput.source, section,
    initialLocation: { pageIndex: 0, columnIndex: 0, flowDomainId: 'body', section,
      cursorPt: { xPt: 10, yPt: 10 }, availableBounds: { xPt: 10, yPt: 10, widthPt: 180, heightPt: 80 } },
  }, services, { currentDateMs: 0 });
  const budget = paragraphAcquisitionCacheOf(services)!.workBudget;
  const initialLimit = budget.limit;
  expect(initialLimit).toBeGreaterThan(25_000);
  // Two remaining misses cannot acquire even the first four-paragraph note
  // unless its page credit is admitted before descendant paragraph debits.
  for (let i = 0; i < initialLimit - 2; i++) budget.noteMiss();
  for (let pageIndex = 0; pageIndex < 3; pageIndex++) {
    const before = budget.limit;
    const story = session.layoutStory!({ source: { story: kind, storyInstance: '7', path: [] },
      pageIndex, section, container: { id: `${kind}:${pageIndex}`, kind,
        bounds: { xPt: 10, yPt: 0, widthPt: 180, heightPt: 100 } } });
    expect(story.blocks).toHaveLength(4);
    expect(budget.limit).toBeGreaterThan(before);
  }
  expect(budget.misses).toBeGreaterThan(initialLimit);
  expect(budget.limit).toBeGreaterThan(initialLimit);
});
