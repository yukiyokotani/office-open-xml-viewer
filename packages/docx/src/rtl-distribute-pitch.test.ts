import { expect, it } from 'vitest';
import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService } from './layout/text.js';
import { createLayoutServices } from './layout-runtime.js';
import { createLayoutServicesRuntimeView } from './layout/runtime-state.js';
import { layoutDocument } from './document-layout.js';
import { makeRecordingCanvas } from './test-support/recording-canvas.test-support.js';
import type { TextPlacement, ParagraphLayout } from './layout/types.js';
import type { DocxDocumentModel, DocParagraph, DocRun, DocTable } from './types.js';

const identity = 'office-local:pitch-proof';
function paragraph(word = 'Xray', ascii: Partial<DocRun> = {}, atEnd = false): DocParagraph {
  const run = (text: string, rtl: boolean) => ({ type: 'text', text, rtl,
    fontFamily: 'Pitch Face', fontFamilyEastAsia: 'Pitch Face', fontSize: 10,
    bold: false, italic: false, underline: false, strikethrough: false });
  return { alignment: 'distribute', bidi: true, wordWrap: true,
    indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0,
    lineSpacing: null, numbering: null, tabStops: [], widowControl: false,
    runs: atEnd ? [run('كتاب قلم', true), { ...run(' ' + word, false), ...ascii }]
      : [run('كتاب', true), { ...run(' ' + word + ' ', false), ...ascii }, run('قلم', true)],
  } as DocParagraph;
}
function acquire(p = paragraph(), mode = 14, conflict = false, missingTrackedInk = false, contextual = false, withCell = false) {
  const model = { body: [{ type: 'paragraph', ...p }], settings: { compatibilityMode: mode },
    section: { pageWidth: 240, pageHeight: 400, marginLeft: 0, marginRight: 0,
      marginTop: 0, marginBottom: 0, headerDistance: 0, footerDistance: 0,
      titlePage: false, evenAndOddHeaders: false },
    headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
  } as DocxDocumentModel;
  if (withCell) {
    const borders = { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null };
    const table = { type: 'table', layout: 'fixed', colWidths: [240], widthPt: 240, borders, jc: 'left',
      cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
      rows: [{ rowHeight: null, rowHeightRule: 'auto', isHeader: false,
        cells: [{ content: [{ type: 'paragraph', ...p }], colSpan: 1, vMerge: null, borders,
          background: null, vAlign: 'top', widthPt: 240, widthPct: null,
          marginTop: 0, marginBottom: 0, marginLeft: 0, marginRight: 0 }] }] } as unknown as DocTable & { type: 'table' };
    model.body.push(table);
  }
  const { canvas } = makeRecordingCanvas();
  let trackedRequests = 0, ownershipRequests = 0;
  const metric = { family: 'Pitch Loaded', requestedFamily: 'Pitch Face', sourceIdentity: identity,
    averageCharWidthRatio: 0.5, lineHeightRatio: 1, unicodeRanges: [[0, 0x10ffff]] as [number, number][] };
  const text = createTextLayoutService({
    fonts: createFontResolver([{ requestedFamily: 'Pitch Face', resolvedFamily: 'Pitch Loaded',
      source: 'local', resourceIdentity: identity, weight: 400, style: 'normal', coverageProbeFamily: 'Pitch Zero' }]),
    fontMetrics: { metric, ...(conflict ? { conflict: { ...metric, averageCharWidthRatio: 0.6 } } : {}) },
    measurer: { fingerprint: 'pitch-independent-advances', ownsGlyphs: () => { ownershipRequests++; return true; }, sameGlyphRaster: () => true,
      measure: request => { if (request.letterSpacingPt) trackedRequests++; return ({ advancePt: [...request.text].reduce((sum, c) => sum + (c === ' ' ? 3 : 10), 0)
          + request.text.length * (request.letterSpacingPt ?? 0)
          - (contextual && request.text.startsWith('كت') ? 12 : 0),
        ascentPt: 8, descentPt: 2, inkBounds: missingTrackedInk && request.letterSpacingPt ? undefined : { xMinPt: 0,
          xMaxPt: [...request.text].reduce((sum, c) => sum + (c === ' ' ? 0 : 10), 0)
              + Math.max(0, request.text.trimEnd().length - 1) * (request.letterSpacingPt ?? 0),
          ascentPt: request.letterSpacingPt ? 11 : 8, descentPt: request.letterSpacingPt ? 4 : 2 },
        horizontalInkBoundsAreTight: !(missingTrackedInk && request.letterSpacingPt) }); },
    },
  });
  const services = createLayoutServicesRuntimeView(createLayoutServices(model, { measureContext: canvas.getContext('2d') }), { text });
  const layout = layoutDocument(model, services);
  const published = layout.pages[0]!.layers.body.find(n => n.kind === 'paragraph') as ParagraphLayout;
  expect(published.lines).toHaveLength(1);
  const source = p.runs.map(r => 'text' in r ? r.text : '').join('');
  const pieces = published.lines[0]!.placements.filter(piece => piece.kind === 'text');
  const logical = [...pieces].sort((a, b) => a.range.start - b.range.start);
  expect(logical.map(piece => piece.text).join('')).toBe(source);
  let cursor = 0;
  for (const piece of logical) { expect(piece.range.start).toBe(cursor); cursor = piece.range.end; }
  expect(cursor).toBe(source.length);
  const placementPens = (target: ParagraphLayout, start: number, length: number) => Array.from({ length }, (_, index) => {
    const offset = start + index;
    for (const piece of target.lines.flatMap(line => line.placements).filter((p): p is TextPlacement => p.kind === 'text')) for (const cluster of piece.clusters) {
      if (cluster.range.start === offset && cluster.range.end === offset + 1) return piece.origin.xPt + cluster.offset.xPt;
    }
    throw Error('Missing source-owned glyph');
  });
  const pens = (start: number, length: number) => placementPens(published, start, length);
  const table = layout.pages[0]!.layers.body.find(node => node.kind === 'table');
  const cell = table?.kind === 'table' ? table.rows[0]!.cells[0]!.blocks.find(block => block.layout.kind === 'paragraph')?.layout : undefined;
  return { layout, published, pens, cellPens: (start: number, length: number) => {
    if (cell?.kind !== 'paragraph') throw Error('Missing cell');
    return placementPens(cell, start, length);
  }, trackedRequests, ownershipRequests, warm: () => layoutDocument(model, services) };
}

it('allocates source spaces and ordinary glyph-following pitch in one RTL published line', () => {
  // Natural nonSPACE=110, two selected xAvg cells=5 each, six source
  // opportunities =>20pt. Arabic words remain intact, not character-spread.
  const { published, pens } = acquire();
  expect(pens(5, 4)).toEqual([55, 85, 115, 145]);
  expect(published.lines[0]!.placements.filter((p): p is TextPlacement => p.kind === 'text' && p.text === 'Xray')
    .flatMap(p => p.paintOps).map(op => op.text).join('')).toBe('Xray');
  const word = published.lines[0]!.placements.find((p): p is TextPlacement => p.kind === 'text' && p.text === 'Xray')!;
  expect(word.paintOps[0]!.letterSpacingPt).toBe(20);
  expect(word.paintOps[0]!.inkBounds?.xMaxPt).toBe(100);
  // Actual tracked ink may exceed the unchanged8/2 font box. Both published
  // ink views must describe that SAME receipt, not stale acquisition metadata.
  expect(word.paintOps[0]!.inkBounds?.ascentPt).toBe(11);
  expect(word.paintOps[0]!.inkBounds?.descentPt).toBe(4);
  expect(word.paintOps[0]!.blockAxisInkBounds).toEqual({ startPt: -11, endPt: 4 });
  expect(word.advancePt).toBe(100); // native aggregate includes a trailing pitch; this does not.

});

it('counts digits and punctuation generically but gives true EOF no following cut', () => {
  // Four ordinary ASCII glyphs, but only three following cuts at EOF.
  // (240−110−2×5)/5=24; no literal original eight-opportunity constant.
  // The ordinary-authored leading SPACE stays before the LTR word:5+24=29.
  expect(acquire(paragraph('R0-m', {}, true)).pens(9, 4)).toEqual([29, 63, 97, 131]);
  // A true-EOF single glyph has ZERO ordinary cuts: only two SPACE cells
  // divide the remaining150pt; no minimum ASCII length or trailing pitch.
  const one = acquire(paragraph('A', {}, true));
  expect(one.pens(9, 1)).toEqual([80]);
  const glyph = one.published.lines[0]!.placements.find((p): p is TextPlacement => p.kind === 'text' && p.text === 'A')!;
  expect(glyph.advancePt).toBe(10);
  expect(glyph.paintOps[0]!.letterSpacingPt).toBe(0);
});

it('preserves the old complete presentation for excluded source, profile and unresolved pitch', () => {
  const negative = acquire(paragraph('Xray', { cs: true }));
  const negativePens = negative.pens(5, 4);
  expect(negativePens.slice(1).map((pen, index) => pen - negativePens[index]!)).toEqual([10, 10, 10]);
  expect(negative.warm()).toEqual(negative.layout);
  for (const options of [{ p: paragraph('Xray', { rtl: true }) }, { mode: 15 }, { conflict: true }]) {
    const result = acquire(options.p ?? paragraph(), options.mode ?? 14, options.conflict ?? false);
    const pens = result.pens(5, 4);
    expect(pens.slice(1).map((pen, index) => pen - pens[index]!)).toEqual([10, 10, 10]);
    if (options.conflict) expect(pens).toEqual([30, 40, 50, 60]);
  }
});

it('declines an RTL-owned seam and late invalid tracked ink as complete old presentations', () => {
  const opposite = paragraph();
  const [a, b, c] = opposite.runs;
  const rtlSpace = { ...opposite, runs: [{ ...a, text: 'كتاب ' }, { ...b, text: 'Xray ' }, c] } as DocParagraph;
  expect(acquire(rtlSpace).pens(5, 4)).toEqual([30, 40, 50, 60]);
  expect(acquire(paragraph(), 14, false, true).pens(5, 4)).toEqual([30, 40, 50, 60]);
});

it('publishes the same complete projection on a warm acquisition without mutating the cached native geometry', () => {
  const result = acquire();
  expect(result.warm()).toEqual(result.layout);
  expect(result.pens(5, 4)).toEqual([55, 85, 115, 145]);
});

it('retains a negative contextual Arabic contribution in the visible pitch budget', () => {
  // Joining changes the second Arabic prefix delta from10 to−2. Signed V=98,
  // two q=5 cells and six opportunities give22; clamp-to-zero would differ.
  expect(acquire(paragraph(), 14, false, false, true).pens(5, 4)).toEqual([57, 89, 121, 153]);
});

it('does not carry a root projection into the same source paragraph in a table cell', () => {
  const result = acquire(paragraph(), 14, false, false, false, true);
  expect(result.pens(5, 4)).toEqual([55, 85, 115, 145]);
  expect(result.cellPens(5, 4)).toEqual([30, 40, 50, 60]);
});

it('declines excessive optional request metadata before new native tracking or glyph proofs', () => {
  const p = paragraph();
  const largeUnusedSlot = 'unused'.repeat(500_000);
  const source = { ...p, runs: p.runs.map(run => ({ ...run, fontFamilyEastAsia: largeUnusedSlot })) } as DocParagraph;
  const result = acquire(source);
  expect(result.pens(5, 4)).toEqual([30, 40, 50, 60]);
  expect(result.trackedRequests).toBe(0);
  expect(result.ownershipRequests).toBe(0);
});
