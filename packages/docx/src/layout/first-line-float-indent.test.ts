/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { documentBytes, measureContext, matrix, fittingMatrix } from '../test-support/tab-fitting.test-support.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

function layoutParagraph(bytes: Uint8Array, advancePt = 5) {
  const archive = new DocxArchive(bytes);
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(advancePt) }), { currentDateMs: 0 });
  const paragraph = layout.pages[0]?.layers.body.find((node) => node.kind === 'paragraph');
  if (paragraph?.kind !== 'paragraph') throw new Error('Missing paragraph');
  return paragraph;
}

describe('first-line indents beside floats through the DOCX parser', () => {
  // Independent Word positions for the valid-image 84-case matrix. The
  // control has a 24 pt cell; this deterministic measurer has a 20 pt cell,
  // so leading/trailing/center positions differ by 4/0/2 pt in RTL (0/4/2 LTR).
  // No-float ordinary tabs retain main's geometry. Five no-float RTL positional
  // cases instead use Word: margin-left, both centers, and both right targets.
  const floatStarts = [
    [552, 108, 484, 40], [566, 352, 240, 26],
    [532, 108, 484, 60], [566, 332, 260, 26],
    [542, 108, 484, 50], [566, 342, 250, 26],
    [542, 108, 484, 50], [566, 342, 250, 26],
    [272, 72, 520, 320], [272, 108, 484, 320],
    [396, 196, 396, 196], [396, 214, 378, 196],
    [520, 320, 272, 72], [520, 320, 272, 72],
  ] as const;

  it.each(matrix)('$kind $alignment ($count), rtl=$rtl, float=$float matches Word geometry', (entry) => {
    const relativeTo = 'relativeTo' in entry ? entry.relativeTo as 'margin' | 'indent' : undefined;
    const paragraph = layoutParagraph(documentBytes('w:left="720" w:hanging="720"', entry.rtl,
      entry.float === 'right' ? 268 : 0, {
        alignment: entry.alignment, count: entry.count,
        text: entry.alignment === 'decimal' ? '12.3' : 'word',
        positional: relativeTo !== undefined, relativeTo, noFloat: entry.float === 'none',
      }));
    const texts = paragraph.lines.flatMap((line) => line.placements).filter((node) => node.kind === 'text');
    expect(texts).toHaveLength(1);
    if (entry.float === 'none') {
      const xPt = entry.rtl ? entry.rtlX : entry.ltrX;
      const widthPt = entry.rtl ? entry.rtlWidth : entry.ltrWidth;
      expect(paragraph.lines.map((line) => line.bounds)).toEqual([
        { xPt: entry.rtl ? xPt : 72, yPt: 72, widthPt, heightPt: 12 },
      ]);
      expect(texts.map((text) => text.bounds)).toEqual([
        { xPt, yPt: 72, widthPt: 20, heightPt: 12 },
      ]);
    } else {
      const index = (entry.rtl ? 2 : 0) + (entry.float === 'right' ? 1 : 0);
      const expectedX = floatStarts[entry.tabIndex][index];
      // WORD_TAB_DISPLACED_READING_FRAME (mode absent): the displaced cell keeps
      // Word's position past the float window or the margin. One residual is
      // unexplained: a second tab whose pen already lies past the right margin
      // advances 14 pt in Word at 10 and 20 pt (tabs-floats-overflow-font20),
      // while the library resolves the automatic 36 pt grid (10 pt further).
      const pastMarginSecondTab = entry.count === 2 && (entry.rtl ? entry.float === 'right' : entry.float === 'left');
      expect(texts.map((text) => text.bounds)).toEqual([
        { xPt: expectedX + (pastMarginSecondTab ? (entry.rtl ? -10 : 10) : 0), yPt: 72, widthPt: 20, heightPt: 12 },
      ]);
    }
  });

  // Latin and CJK text in left-to-right runs: outside the measured
  // WORD_ALIGNED_TAB_CELL_FIT_BACK bidi rows (w:rtl runs), the last start gap
  // stays provisional and the bidi walk shrinks it as the cell fills the line.
  it.each([
    { name: 'two tabs with hanging indent and CJK', indent: 'w:left="720" w:hanging="720"', count: 2, stop: 5600, text: '漢'.repeat(50) },
    { name: 'three tabs and Latin words', indent: '', count: 3, stop: 720, text: 'word '.repeat(15).trim() },
  ])('preserves in-band no-float RTL fitting: $name', ({ indent, count, stop, text }) => {
    const paragraph = layoutParagraph(documentBytes(indent, true, 0,
      { alignment: 'left', count, stop, text, noFloat: true }));
    expect(paragraph.lines).toHaveLength(1);
    const texts = paragraph.lines[0].placements.filter((node) => node.kind === 'text');
    expect(texts.map((node) => node.text).join('')).toBe(text);
    for (const node of texts) {
      expect(node.bounds.xPt).toBeGreaterThanOrEqual(72);
      expect(node.bounds.xPt + node.bounds.widthPt).toBeLessThanOrEqual(540);
    }
    if (count === 2) {
      expect(texts[0].bounds.xPt).toBe(72);
      expect(texts.reduce((width, node) => width + node.bounds.widthPt, 0)).toBe(250);
    }
  });

  it.each([
    { name: 'short trailing cell', stop: 11000, cells: ['title', '12.3'], xPt: 72 },
    { name: 'long interleaved CJK cell', stop: 2400, cells: ['aaa', '漢'.repeat(90), 'b'], xPt: 67 },
  ])('preserves contained RTL tabs with a negative trailing indent: $name', ({ stop, cells, xPt }) => {
    const paragraph = layoutParagraph(documentBytes('w:right="-720"', true, 0,
      { alignment: 'left', count: cells.length - 1, stop, text: cells.join(''), cells, noFloat: true }));
    expect(paragraph.lines).toHaveLength(1);
    const texts = paragraph.lines[0].placements.filter((node) => node.kind === 'text');
    expect(texts[0].bounds.xPt).toBe(xPt);
    for (const node of texts) {
      expect(node.bounds.xPt).toBeGreaterThanOrEqual(36);
      expect(node.bounds.xPt + node.bounds.widthPt).toBeLessThanOrEqual(540);
    }
  });

  // ECMA-376 §17.3.1.37 stops are independent of the paragraph's right indent;
  // without a float, a fitting ordinary aligned cell keeps that allocation on
  // its line (sample-1 footer regression).
  it('aligns a fitting ordinary cell at its stop past the right indent', () => {
    const paragraph = layoutParagraph(documentBytes('w:right="7200"', false, 0,
      { alignment: 'right', count: 1, text: 'word '.repeat(10).trim(), noFloat: true }));
    expect(paragraph.lines).toHaveLength(1);
    const texts = paragraph.lines[0].placements.filter((node) => node.kind === 'text');
    const last = texts.at(-1);
    expect(last ? last.bounds.xPt + last.bounds.widthPt : undefined).toBe(352);
  });

  // Positional tabs keep the #1675 containment policy; their Word placement
  // past a right indent is an open follow-up (see margin-tab-allocation.test).
  it('wraps a margin ptab cell wider than the paragraph band', () => {
    const paragraph = layoutParagraph(documentBytes('w:right="7200"', false, 0,
      { alignment: 'right', count: 1, text: 'word '.repeat(10).trim(), positional: true, noFloat: true }));
    expect(paragraph.lines.length).toBeGreaterThan(1);
    for (const line of paragraph.lines) {
      for (const text of line.placements) {
        if (text.kind === 'text') {
          expect(text.bounds.xPt).toBeGreaterThanOrEqual(72);
          const trailingSpaceWidth = (text.text.length - text.text.trimEnd().length) * 5;
          expect(text.bounds.xPt + text.bounds.widthPt - trailingSpaceWidth).toBeLessThanOrEqual(180);
        }
      }
    }
  });

  it.each([false, true])('wraps a no-float aligned cell wider than its text margin, positional=%s', (positional) => {
    const paragraph = layoutParagraph(documentBytes('w:right="7200"', false, 0,
      { alignment: 'right', count: 1, text: 'word '.repeat(100).trim(), positional, noFloat: true }));
    expect(paragraph.lines.length).toBeGreaterThan(1);
    for (const line of paragraph.lines) {
      for (const text of line.placements) {
        if (text.kind === 'text') {
          expect(text.bounds.xPt).toBeGreaterThanOrEqual(72);
          const trailingSpaceWidth = (text.text.length - text.text.trimEnd().length) * 5;
          expect(text.bounds.xPt + text.bounds.widthPt - trailingSpaceWidth).toBeLessThanOrEqual(180);
        }
      }
    }
  });

  // The margin extension is a line band only where no exclusion intersects
  // it: a float inside the right indent keeps #1675's indent-band fitting.
  it.each([false, true].flatMap((rtl) => [false, true].map((positional) => ({ rtl, positional }))))(
    'keeps aligned cells out of a float inside the right indent, rtl=$rtl, positional=$positional', ({ rtl, positional }) => {
      const paragraph = layoutParagraph(documentBytes('w:right="7200"', rtl, 268,
        { alignment: 'right', count: 1, text: 'word', positional }));
      const texts = paragraph.lines.flatMap((line) => line.placements).filter((node) => node.kind === 'text');
      expect(texts.map((text) => text.text).join('')).toBe('word');
      for (const text of texts) {
        // Float occupies x=340–540pt, y=72–172pt.
        const overlapsFloat = text.bounds.xPt + text.bounds.widthPt > 340 && text.bounds.yPt < 172;
        expect(overlapsFloat).toBe(false);
        expect(text.bounds.xPt).toBeGreaterThanOrEqual(72);
      }
    });

  it.each(['left', 'center', 'right', 'both', 'distribute'].flatMap((jc) => [false, true].flatMap((rtl) =>
    [false, true].map((positional) => ({ jc, rtl, positional })))))(
    'keeps a margin-allocated line inside its band under jc=$jc, rtl=$rtl, positional=$positional', ({ jc, rtl, positional }) => {
      const paragraph = layoutParagraph(documentBytes(`w:right="7200"/><w:jc w:val="${jc}"`, rtl, 0,
        { alignment: 'right', count: 1, text: 'word', positional, noFloat: true, prefix: 'prefix' }));
      const texts = paragraph.lines.flatMap((line) => line.placements).filter((node) => node.kind === 'text');
      expect(texts.map((text) => text.text).sort()).toEqual(['prefix', 'word']);
      for (const text of texts) {
        expect(text.bounds.xPt).toBeGreaterThanOrEqual(72);
        expect(text.bounds.xPt + text.bounds.widthPt).toBeLessThanOrEqual(540);
      }
      if (!rtl && !positional) {
        // An ordinary right stop at 352pt extends the band to the margin,
        // leaving 188pt that center/right alignment distribute as usual.
        expect(paragraph.lines).toHaveLength(1);
        const prefix = texts.find((text) => text.text === 'prefix');
        const word = texts.find((text) => text.text === 'word');
        const shift = jc === 'center' ? 94 : jc === 'right' ? 188 : 0;
        if (jc !== 'distribute') {
          expect(prefix?.bounds.xPt).toBe(72 + shift);
          expect((word?.bounds.xPt ?? 0) + (word?.bounds.widthPt ?? 0)).toBe(352 + shift);
        }
      }
    });

  it('wraps content after an automatic tab in a narrowed float window', () => {
    const paragraph = layoutParagraph(documentBytes('w:left="720" w:hanging="720"', false, 268,
      { alignment: 'left', count: 1, text: 'word '.repeat(100).trim(), automatic: true }), 6);
    expect(paragraph.lines).toHaveLength(12);
    for (const line of paragraph.lines) {
      for (const placement of line.placements) {
        if (placement.kind === 'text') expect(placement.bounds.xPt + placement.bounds.widthPt).toBeLessThanOrEqual(540);
      }
    }
  });

  it.each(['margin', 'indent'].flatMap((relativeTo) => [false, true].map((long) => ({ relativeTo: relativeTo as 'margin' | 'indent', long }))))('moves an RTL positional tab past its $relativeTo target to the next line, long=$long', ({ relativeTo, long }) => {
    const paragraph = layoutParagraph(documentBytes('', true, 0,
      { alignment: 'left', count: 1, text: long ? 'word '.repeat(100).trim() : 'word', positional: true, relativeTo, noFloat: true, prefix: 'prefix' }), 6);
    const text = paragraph.lines.flatMap((line) => line.placements).find((node) => node.kind === 'text' && node.text.startsWith('word'));
    expect(text?.kind === 'text' ? text.bounds.yPt : undefined).toBe(84);
    if (long) expect(paragraph.lines.length).toBeGreaterThan(2);
    else expect(paragraph.lines).toHaveLength(2);
  });


  it.each(fittingMatrix)('$script after $kind $alignment ($count), rtl=$rtl, float=$float retains normal fitting', (entry) => {
    const relativeTo = 'relativeTo' in entry ? entry.relativeTo as 'margin' | 'indent' : undefined;
    const paragraph = layoutParagraph(documentBytes('w:left="720" w:hanging="720"', entry.rtl,
      entry.float === 'right' ? 268 : 0, {
        alignment: entry.alignment, count: entry.count, text: entry.content,
        positional: relativeTo !== undefined, relativeTo, noFloat: entry.float === 'none',
        automatic: entry.kind === 'automatic',
      }));
    expect(paragraph.lines.length).toBeGreaterThan(1);
    const textLines = paragraph.lines.map((line) => line.placements.filter((node) => node.kind === 'text'));
    expect(textLines.flat().map((text) => text.text).join('').replace(/\s/g, ''))
      .toBe(entry.content.replace(/\s/g, ''));
    for (const texts of textLines) {
      for (const text of texts) {
        if (text.text.trim().length === 0) continue;
        // Trailing spaces contribute advance but no ink, and normally hang
        // outside the fitted band. Check the visible text extent instead.
        // A right-to-left line's end is its left edge, where its line-end
        // space hangs (WORD_RTL_LINE_END_SPACE_HANGS).
        const trailingSpaceWidth = (text.text.length - text.text.trimEnd().length) * 5;
        const left = text.bounds.xPt + (entry.rtl ? trailingSpaceWidth : 0);
        const right = text.bounds.xPt + text.bounds.widthPt - (entry.rtl ? 0 : trailingSpaceWidth);
        expect(left).toBeGreaterThanOrEqual(72);
        expect(right).toBeLessThanOrEqual(540);
        if (text.bounds.yPt < 172 && entry.float !== 'none') {
          if (entry.float === 'left') expect(left).toBeGreaterThanOrEqual(272);
          else expect(right).toBeLessThanOrEqual(340);
        }
      }
    }
  });

  it.each([
    { positional: false, relativeTo: 'indent' as const, firstCount: 93, firstX: 72 },
    { positional: true, relativeTo: 'indent' as const, firstCount: 86, firstX: 74 },
    { positional: true, relativeTo: 'margin' as const, firstCount: 86, firstX: 74 },
  ])('preserves in-band no-float RTL CJK fitting, positional=$positional, reference=$relativeTo', ({ positional, relativeTo, firstCount, firstX }) => {
    const paragraph = layoutParagraph(documentBytes('w:left="720" w:hanging="720"', true, 0,
      { alignment: 'left', count: 1, text: '漢'.repeat(100), positional, relativeTo, noFloat: true }));
    const texts = paragraph.lines.map((line) => line.placements.filter((node) => node.kind === 'text'));
    expect(texts.map((line) => line.map((text) => text.text).join('')))
      .toEqual(['漢'.repeat(firstCount), '漢'.repeat(100 - firstCount)]);
    expect(texts[0][0].bounds.xPt).toBe(firstX);
  });

  it('splits CJK at legal opportunities after an in-band automatic tab', () => {
    const paragraph = layoutParagraph(documentBytes('w:left="720" w:hanging="720"', false, 268,
      { alignment: 'left', count: 1, text: '漢'.repeat(50), automatic: true }));
    expect(paragraph.lines).toHaveLength(2);
    const first = paragraph.lines[0].placements.filter((node) => node.kind === 'text');
    expect(first.map((text) => text.text).join('')).toBe('漢'.repeat(46));
    const last = first.at(-1);
    expect(last ? last.bounds.xPt + last.bounds.widthPt : undefined).toBe(338);
  });

  // tabs-floats-overflow-font20 (Word, mode absent): the first tab keeps its
  // float-displaced stop past the window (WORD_TAB_DISPLACED_READING_FRAME), so
  // its cell starts past the margin. Word then advances the second tab 14 pt
  // at 10 and 20 pt (word at 566 pt), an unexplained amount; the library uses
  // the automatic 36 pt grid (576 pt), a known 10 pt residual.
  it.each([10, 20])('overflows a displaced tab cell past the margin at %s pt font size', (fontSizePt) => {
    const paragraph = layoutParagraph(documentBytes('w:left="720" w:hanging="720"', false, 0,
      { alignment: 'left', count: 2, text: 'word', fontSizePt }), fontSizePt * 0.6);
    expect(paragraph.lines).toHaveLength(1);
    const text = paragraph.lines.flatMap((line) => line.placements).find((node) => node.kind === 'text');
    expect(text?.bounds.xPt).toBe(576);
    expect(text?.bounds.widthPt).toBe(fontSizePt * 2.4);
  });

  it.each([
    { name: 'hanging beside a left float', indent: 'w:left="720" w:hanging="720"', rtl: false, floatX: 0, expectedStart: 272, continuationStart: 272 },
    { name: 'positive first-line indent beside a left float', indent: 'w:firstLine="720"', rtl: false, floatX: 0, expectedStart: 308, continuationStart: 272 },
    { name: 'hanging beyond a nonblocking left float', indent: 'w:left="4800" w:hanging="360"', rtl: false, floatX: 0, expectedStart: 294, continuationStart: 312 },
    { name: 'hanging beside a right float in RTL', indent: 'w:left="720" w:hanging="720"', rtl: true, floatX: 268, expectedEnd: 340 },
  ])('$name', ({ indent, rtl, floatX, expectedStart, expectedEnd, continuationStart }) => {
    const paragraph = layoutParagraph(documentBytes(indent, rtl, floatX));
    const line = paragraph.lines[0];
    expect(line?.bounds.yPt).toBe(72);
    // A correction only to placement would leave the breaker's old, widened
    // hanging-line budget and let long text cross the opposite window edge.
    expect(paragraph.lines.length).toBeGreaterThan(1);
    expect(line?.bounds.widthPt).toBeLessThanOrEqual(268);
    if (expectedStart !== undefined) expect(line?.bounds.xPt).toBeCloseTo(expectedStart, 8);
    if (expectedEnd !== undefined) expect((line?.bounds.xPt ?? 0) + (line?.bounds.widthPt ?? 0)).toBeCloseTo(expectedEnd, 8);
    if (continuationStart !== undefined) expect(paragraph.lines[1]?.bounds.xPt).toBeCloseTo(continuationStart, 8);
  });
});
