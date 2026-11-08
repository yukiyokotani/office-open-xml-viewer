/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { documentBytes, measureContext } from '../test-support/tab-fitting.test-support.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';
import { resolveDefaultTabPt } from '../line-breaker/tabs.js';
import { SegmentQueue } from '../line-breaker/segment-queue.js';
import { textRunsForPage } from '../text-run-projection.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

// Every character advances 5 pt; the text band runs from 72 to 540 pt.
function paragraphDocumentLayout(bytes: Uint8Array, advancePt = 5) {
  const archive = new DocxArchive(bytes);
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
  return layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(advancePt) }), { currentDateMs: 0 });
}

function paragraphLines(bytes: Uint8Array, advancePt = 5) {
  const layout = paragraphDocumentLayout(bytes, advancePt);
  const paragraph = layout.pages[0]?.layers.body.find((node) => node.kind === 'paragraph');
  if (paragraph?.kind !== 'paragraph') throw new Error('Missing paragraph');
  return paragraph.lines;
}

function lineTexts(bytes: Uint8Array, advancePt = 5) {
  return paragraphLines(bytes, advancePt).map((line) => line.placements.flatMap((node) => node.kind === 'text'
    ? [{ text: node.text, left: node.bounds.xPt, right: node.bounds.xPt + node.bounds.widthPt }] : []));
}

describe('WORD_TAB_GLUED_TO_FOLLOWING_TEXT', () => {
  const word = 'x'.repeat(50); // 250 pt: does not fit behind a 280 pt stop (188 pt left)

  it.each([false, true])('moves the tab with a word that does not fit behind its stop, rtl=%s', (rtl) => {
    // As measured: a Latin word in a left-to-right paragraph, a Hebrew word
    // in a w:rtl run in a bidi paragraph.
    const letter = rtl ? '\u05d0' : 'x';
    const lines = lineTexts(documentBytes('', rtl, 0, {
      alignment: 'left', count: 1, stop: 5600, prefix: 'label', text: letter.repeat(50), noFloat: true, runRtl: rtl,
    }));
    expect(lines.map((line) => line.map((text) => text.text).join(''))).toEqual(['label', letter.repeat(37), letter.repeat(13)]);
    // Line 2 starts at the stop again: the tab travelled with the word, which
    // is then broken by characters behind it.
    const cell = lines[1]![0]!;
    if (rtl) expect(cell.right).toBe(540 - 280);
    else expect(cell.left).toBe(72 + 280);
  });

  it('leaves a left-to-right word in a bidi paragraph out of the rules', () => {
    // Unmeasured direction mix (the #1672 r4 controls show such tabs ending
    // their line): the tab keeps the earlier provisional gap, which the bidi
    // walk shrinks so the word stays on the label's line.
    const lines = lineTexts(documentBytes('', true, 0,
      { alignment: 'left', count: 1, stop: 5600, prefix: 'label', text: word, noFloat: true }));
    expect(lines.map((line) => line.map((text) => text.text).join(''))).toEqual([`${word}label`]); // visual order
  });

  // Outside the measured scope (one plain run) the tab keeps the earlier
  // behaviour: it ends line 1 and the word starts line 2 at the margin.
  it('leaves a word continued across a formatting seam out of the rule', () => {
    const text = 'a'.repeat(10) + 'b'.repeat(40);
    const lines = lineTexts(documentBytes('', false, 0,
      { alignment: 'left', count: 1, stop: 5600, prefix: 'label', text, boldFrom: 10, noFloat: true }));
    expect(lines.map((line) => line.map((piece) => piece.text).join(''))).toEqual(['label', text]);
    expect(lines[1]![0]!.left).toBe(72);
  });

  it.each([false, true])('never overflows the margin with a joined word behind a leading tab, rtl=%s', (rtl) => {
    // 36 regular characters fill the 180 pt behind a 288 pt stop exactly; a
    // joined bold continuation follows.
    const text = 'a'.repeat(36) + 'b'.repeat(14);
    const lines = lineTexts(documentBytes('', rtl, 0,
      { alignment: 'left', count: 1, stop: 5760, text, boldFrom: 36, noFloat: true }));
    for (const piece of lines.flat()) {
      expect(piece.left).toBeGreaterThanOrEqual(72);
      expect(piece.right).toBeLessThanOrEqual(540);
    }
    expect(lines.map((line) => line.map((piece) => piece.text).join('')).join('')).toBe(text);
  });

  it('keeps a legal discretionary hyphen after the tab as a break opportunity', () => {
    // A 10-character prefix and its hyphen fit behind the 280 pt stop.
    const lines = lineTexts(documentBytes('', false, 0, {
      alignment: 'left', count: 1, stop: 5600, prefix: 'label', text: 'W', noFloat: true,
      textXml: `<w:t>${'a'.repeat(10)}</w:t><w:softHyphen/><w:t>${'b'.repeat(40)}</w:t>`,
    }));
    expect(lines[0]!.find((piece) => piece.text.startsWith('a'))?.left).toBe(72 + 280);
  });

  it('breaks the word behind a tab that starts the line', () => {
    const lines = lineTexts(documentBytes('', false, 0,
      { alignment: 'left', count: 1, stop: 5600, text: word, noFloat: true }));
    expect(lines.map((line) => line.map((text) => text.text).join(''))).toEqual(['x'.repeat(37), 'x'.repeat(13)]);
    expect(lines[0]![0]!.left).toBe(352);
  });

  it('lets a space after the tab end the line', () => {
    const lines = lineTexts(documentBytes('', false, 0,
      { alignment: 'left', count: 1, stop: 5600, prefix: 'label', text: ` ${word}`, noFloat: true }));
    expect(lines.map((line) => line.map((text) => text.text).join('').trim())).toEqual(['label', word]);
    expect(lines[1]![0]!.left).toBe(72);
  });
});

describe('WORD_ALIGNED_TAB_CELL_FIT_BACK', () => {
  it('slides a centred cell that would overflow back to the line end', () => {
    // Centred at 400 pt the 300 pt cell would end at 550 pt, past the 468 pt line.
    const lines = lineTexts(documentBytes('', false, 0,
      { alignment: 'center', count: 1, stop: 8000, prefix: 'ab', text: 'x'.repeat(60), noFloat: true }));
    expect(lines).toHaveLength(1);
    expect(lines[0]!.at(-1)).toEqual({ text: 'x'.repeat(60), left: 240, right: 540 });
  });

  it.each([
    // Right float x 340-540 pt: the 400 pt right stop lies past the window end.
    { name: 'right stop past a trailing float', floatX: 268, stop: 8000, mode15: 320, legacy: 452 },
    // Left float x 72-272 pt: mode 15 keeps the margin-relative stop (352 pt);
    // before it the stop moves with the 200 pt displaced line start.
    { name: 'right stop beside a leading float', floatX: 0, stop: 5600, mode15: 332, legacy: 532 },
  ])('$name: limits the cell to the window from mode 15 only', ({ floatX, stop, mode15, legacy }) => {
    for (const [compatibilityMode, expected] of [[15, mode15], [undefined, legacy]] as const) {
      const lines = lineTexts(documentBytes('', false, floatX,
        { alignment: 'right', count: 1, stop, prefix: 'ab', text: 'word', compatibilityMode }));
      expect(lines[0]!.at(-1), `mode ${compatibilityMode}`).toEqual({ text: 'word', left: expected, right: expected + 20 });
    }
  });
});

describe('WORD_ALIGNED_TAB_CELL_FIT_BACK in a bidi paragraph', () => {
  // A 10 pt label, a centre stop at 400 pt and a 295 pt Hebrew cell with
  // spaces: centred, the cell would pass the line end, so it is fitted back.
  const cell = Array.from({ length: 12 }, () => 'אבגד').join(' ');

  it.each([
    // Unjustified, the gap takes the room the line leaves (468 - 10 - 295).
    { jc: undefined, gap: 163 },
    // Justified, the gap stays empty and the word spaces take the slack
    // (blk-0898-tabs-rtl D2 bidi rows, as left-to-right).
    { jc: 'distribute', gap: 0 },
  ])('jc $jc: the tab gap and the cell reaching the line end', ({ jc, gap }) => {
    const lines = paragraphLines(documentBytes('', true, 0,
      { alignment: 'center', count: 1, stop: 8000, prefix: 'ab', text: cell, noFloat: true, runRtl: true, ...(jc ? { jc } : {}) }));
    expect(lines).toHaveLength(1);
    const tab = lines[0]!.placements.find((node) => node.kind === 'tab');
    expect(tab?.bounds?.widthPt).toBeCloseTo(gap);
    const texts = lines[0]!.placements.filter((node) => node.kind === 'text');
    expect(Math.min(...texts.map((node) => node.bounds.xPt))).toBeCloseTo(72);
  });
});

describe('WORD_ALIGNED_TAB_CELL_FIT_BACK on a kashida-justified line', () => {
  it('keeps the gap empty when the slack goes into kashida elongation', () => {
    // A lowKashida line ended by a break: 'ab', a centre stop at 400 pt and
    // 60 Arabic beh (300 pt), fitted back with no gap.
    const beh = '\u0628'.repeat(60);
    const lines = paragraphLines(documentBytes('', true, 0, {
      alignment: 'center', count: 1, stop: 8000, prefix: 'ab', text: 'W', noFloat: true, jc: 'lowKashida', runRtl: true,
      textXml: `<w:t>${beh}</w:t><w:br/><w:t>x</w:t>`,
    }));
    const tab = lines[0]!.placements.find((node) => node.kind === 'tab');
    expect(tab?.bounds?.widthPt).toBeCloseTo(0);
  });
});

describe('an aligned cell ended by a later tab', () => {
  it('keeps its stop-aligned gap, so no fitted-back gap moves the later stop', () => {
    // A 110 pt prefix, a 295 pt cell centred on 400 pt and a left stop at
    // 432 pt. The elastic fit-back was measured only for a cell that ends its
    // line, so this cell keeps its stop-aligned gap and wraps there; the next
    // line's tab then resolves its own stop (the 400 pt centre stop).
    const lines = lineTexts(documentBytes('', false, 0, {
      alignment: 'center', count: 1, stop: 8000, text: '', noFloat: true,
      extraStops: [{ alignment: 'left', pos: 8640 }],
      cells: ['p'.repeat(22), Array.from({ length: 12 }, () => 'xxxx').join(' '), 'zz'],
    }));
    expect(lines).toHaveLength(2);
    expect(lines[0]![1]!.left).toBe(72 + 400 - 295 / 2);
    expect(lines[1]!.at(-1)).toEqual({ text: 'zz', left: 72 + 395, right: 72 + 405 });
  });
});

describe('tab rules outside their measured scope keep the earlier layout', () => {
  it('keeps the leading tab-only line before a long decimal cell without a separator', () => {
    const lines = paragraphLines(documentBytes('w:left="720" w:hanging="720"', false, 0,
      { alignment: 'decimal', count: 2, stop: 5600, text: 'a'.repeat(100), noFloat: true }), 7.2);
    expect(lines).toHaveLength(3);
    expect(lines[0]!.placements.some((node) => node.kind === 'text')).toBe(false);
  });

  it.each(['left', 'decimal'])('two %s tabs then a Latin word in a bidi paragraph end their line', (alignment) => {
    // #1672 r4 controls (fit-049, fit-229): Word ends the line with the tabs
    // and starts the unbreakable left-to-right word at the next line start.
    const lines = paragraphLines(documentBytes('w:left="720" w:hanging="720"', true, 0,
      { alignment, count: 2, stop: 5600, text: 'a'.repeat(100), noFloat: true }));
    expect(lines[0]!.placements.some((node) => node.kind === 'text')).toBe(false);
    const second = lines[1]!.placements.filter((node) => node.kind === 'text');
    expect(Math.max(...second.map((node) => node.bounds.xPt + node.bounds.widthPt))).toBe(540 - 36);
  });

  it('does not glue a tab to its word on a float-narrowed line', () => {
    // 7.2 pt glyphs (Courier New 12, as fit-006). A left float narrows the
    // line; the displaced stop lies past the window and the long cell keeps the
    // window containment: the words stay beside the float (272 pt) instead of
    // the tab carrying them below it.
    const lines = lineTexts(documentBytes('w:left="720" w:hanging="720"', false, 0,
      { alignment: 'left', count: 1, stop: 5600, text: 'word '.repeat(40).trim() }), 7.2);
    expect(lines.find((line) => line.length > 0)![0]!.left).toBe(272);
  });

  it('keeps the window containment for an empty cell before a long one', () => {
    // Two tabs beside a left float, then a long CJK cell (fit-037): only short
    // paragraph rests were measured overflowing the float window.
    const lines = paragraphLines(documentBytes('', false, 0,
      { alignment: 'left', count: 2, stop: 5600, text: '\u6f22'.repeat(100) }));
    for (const node of lines.flatMap((line) => line.placements)) {
      if (node.kind === 'tab') expect(node.bounds!.xPt + node.bounds!.widthPt).toBeLessThanOrEqual(540);
    }
  });
});

describe('tab rule work stays linear', () => {
  it('measures the paragraph rest once for consecutive float-displaced tabs', () => {
    // Before mode 15, beside a 200 pt left float: N tabs at a 280 pt left
    // stop, then 'x'. Each displaced tab asks whether the rest of the
    // paragraph fits one line; the queue is walked once overall.
    const visits = (count: number) => {
      const original = SegmentQueue.prototype[Symbol.iterator];
      let yielded = 0;
      SegmentQueue.prototype[Symbol.iterator] = function* counted(this: SegmentQueue) {
        for (const item of original.call(this)) {
          yielded += 1;
          yield item;
        }
      };
      try {
        paragraphLines(documentBytes('', false, 0, { alignment: 'left', count, stop: 5600, text: 'x' }));
      } finally {
        SegmentQueue.prototype[Symbol.iterator] = original;
      }
      return yielded;
    };
    expect(visits(800) / visits(200)).toBeLessThan(8);
  });
});

describe('WORD_ALIGNED_TAB_CELL_FIT_BACK needs a real decimal separator', () => {
  it('keeps the earlier gap and leader for a bidi decimal cell without one', () => {
    // A 468 pt bidi line, 'ab', a dot-leader decimal stop at 400 pt and a
    // Hebrew cell ending in "1234": only the numeric fallback point exists.
    const cell = `${'\u05d0\u05d1\u05d2\u05d3 '.repeat(11)}1234`;
    const lines = paragraphLines(documentBytes('', true, 0, {
      alignment: 'decimal', count: 1, stop: 8000, prefix: 'ab', text: cell, noFloat: true, runRtl: true, jc: 'distribute',
    }));
    const tab = lines[0]!.placements.find((node) => node.kind === 'tab');
    expect(tab?.bounds?.widthPt).toBeGreaterThan(100);
  });
});

describe('resolveDefaultTabPt', () => {
  it.each([5e-324, 1e-310, 0, -36, Number.NaN])('resolves %s, which cannot form a grid, to 36 pt', (value) => {
    expect(resolveDefaultTabPt({ defaultTabStop: value } as never)).toBe(36);
  });

  it.each([1e-300, 1e-14, 0.01, 0.05, 18])('keeps the authored interval %s', (value) => {
    expect(resolveDefaultTabPt({ defaultTabStop: value } as never)).toBe(value);
  });
});

describe('source-owned RTL tab-cell justification', () => {
  it('gives equal source spaces equal visible expansion without moving the label seam', () => {
    // A trailing-aligned cell contains homogeneous Hebrew visible islands. Equal 5 pt glyph metrics make every source U+0020 the same width;
    // justification may enlarge those gaps, not move their allocation to a
    // different source seam. The long cell soft-wraps, so `both` justifies it.
    const text = Array.from({ length: 20 }, () => 'אבגד').join(' ');
    const [line] = paragraphLines(documentBytes('', true, 0, {
      alignment: 'right', count: 1, stop: 7200, prefix: 'ab', text,
      noFloat: true, runRtl: true, jc: 'both', compatibilityMode: 15,
    }));
    const words = line!.placements.filter((node) => node.kind === 'text')
      .flatMap(node => node.paintOps.filter(op => /\S/.test(op.text)).map(op => ({
        text: op.text.trim(),
        left: node.origin.xPt + op.offset.xPt,
        right: node.origin.xPt + op.offset.xPt + [...op.text.trim()].length * 5,
      }))).sort((a, b) => a.left - b.left);
    const label = words.find(word => word.text === 'ab')!;
    const cell = words.filter(word => word !== label);
    const gaps = cell.slice(1).map((word, index) => word.left - cell[index]!.right);
    expect(cell.length).toBeGreaterThan(2);
    expect(cell.at(-1)!.right).toBeCloseTo(label.left);
    for (const gap of gaps) expect(gap).toBeCloseTo(gaps[0]!);
  });

  it('retains the numeric anchor while expanding every source-space gap in a justified tab cell', () => {
    // An interior numeric CS keeps its numeric class under authored RTL.
    // The already acquired numeric word stays whole; source-space projection
    // must not change #2070's independent stop/fit-back acquisition.
    const text = `12.5 ${Array.from({ length: 20 }, () => 'אבגד').join(' ')}`;
    const all = paragraphLines(documentBytes('', true, 0, {
      alignment: 'decimal', count: 1, stop: 7200, prefix: 'ab', text,
      noFloat: true, runRtl: true, jc: 'both', compatibilityMode: 15,
    }));
    const numeric = all[0]!.placements.find(p => p.kind === 'text' && p.text.trim() === '12.5');
    if (numeric?.kind !== 'text') throw new Error('Missing decimal source');
    const period = numeric.clusters.find(c => c.range.start === numeric.range.start + 2)!;
    // Existing justified fit-back keeps the full terminal cell against its
    // label, not re-tabbed to the ordinary stop after spreading. The2glyph
    // label ends at540; the4glyph number abuts it and its period follows2glyphs.
    expect(numeric.origin.xPt + period.offset.xPt).toBeCloseTo(540 - 2 * 5 - 4 * 5 + 2 * 5);
    // Equal acquired source spaces own equal expansion, including the
    // numeric/word seam; a numeric CS is not a separate RTL opportunity.
    const words = all[0]!.placements.filter(node => node.kind === 'text')
      .flatMap(node => node.paintOps.filter(op => /\S/.test(op.text)).map(op => ({
        text: op.text.trim(), left: node.origin.xPt + op.offset.xPt,
        right: node.origin.xPt + op.offset.xPt + op.text.trim().length * 5,
      }))).filter(word => word.text !== 'ab').sort((a, b) => a.left - b.left);
    const gaps = words.slice(1).map((word, index) => word.left - words[index]!.right);
    expect(gaps.length).toBeGreaterThan(2);
    for (const gap of gaps) expect(gap).toBeCloseTo(gaps[0]!);
    const source = all.flatMap(line => line.placements).sort((a, b) => a.range.start - b.range.start);
    let cursor = 0;
    for (const placement of source) {
      expect(placement.range.start).toBe(cursor);
      cursor = placement.range.end;
    }
    expect(source.map(p => p.kind === 'text' ? p.text : p.kind === 'tab' ? '\t' : '').join(''))
      .toBe(`ab\t${text}`);
  });

  it('retains each consecutive source space after a supplementary character in the projected box', () => {
    // Source UTF16 positions differ from code-point cuts after the emoji.
    // Four visible glyphs per word, with two U+0020 between words: 80 pt
    // natural in a 468 pt band. The four spaces each receive (468-80)/4.
    const text = 'אב😀ד  הוזח  טיכל';
    const layout = paragraphDocumentLayout(documentBytes('', true, 0, {
      alignment: 'left', count: 0, text, noFloat: true, runRtl: true,
      jc: 'distribute', compatibilityMode: 15,
    }));
    const paragraph = layout.pages[0]!.layers.body.find(node => node.kind === 'paragraph');
    if (paragraph?.kind !== 'paragraph') throw new Error('Missing paragraph');
    expect(paragraph.lines).toHaveLength(1);
    const placements = paragraph.lines[0]!.placements.filter(node => node.kind === 'text');
    const sourceOrder = [...placements].sort((a, b) => a.range.start - b.range.start);
    expect(sourceOrder.map(node => node.text).join('')).toBe(text);
    let cursor = 0;
    let spaces = 0;
    for (const placement of sourceOrder) {
      expect(placement.range.start).toBe(cursor);
      let clusterCursor = placement.range.start;
      let advancePt = 0;
      for (const cluster of placement.clusters) {
        expect(cluster.range.start).toBe(clusterCursor);
        const source = text.slice(cluster.range.start, cluster.range.end);
        if (source === ' ') {
          expect(cluster.advancePt).toBeCloseTo(102);
          spaces += 1;
        }
        expect(cluster.advancePt).toBeGreaterThanOrEqual(0);
        advancePt += cluster.advancePt;
        clusterCursor = cluster.range.end;
      }
      expect(clusterCursor).toBe(placement.range.end);
      expect(advancePt).toBeCloseTo(placement.advancePt);
      expect(placement.bounds.widthPt).toBeCloseTo(advancePt);
      cursor = placement.range.end;
      // The same source intervals, including zero-ink whitespace, reach paint.
      let painted = placement.range.start;
      for (const operation of placement.paintOps) {
        expect(operation.range.start).toBe(painted);
        painted = operation.range.end;
      }
      expect(painted).toBe(placement.range.end);
    }
    expect(cursor).toBe(text.length);
    expect(spaces).toBe(4);
    const words = [...text.matchAll(/\S+/gu)].map(match => {
      const start = match.index;
      const end = start + match[0].length;
      const intervals = placements.flatMap(placement => placement.paintOps
        .filter(operation => operation.range.start >= start && operation.range.end <= end)
        .map(operation => ({
          left: placement.origin.xPt + operation.offset.xPt,
          width: placement.clusters.filter(cluster => cluster.range.start >= operation.range.start
            && cluster.range.end <= operation.range.end).reduce((sum, cluster) => sum + cluster.advancePt, 0),
        })));
      return {
        left: Math.min(...intervals.map(interval => interval.left)),
        right: Math.max(...intervals.map(interval => interval.left + interval.width)),
      };
    }).sort((a, b) => a.left - b.left);
    expect(words[0]!.left).toBeCloseTo(72);
    expect(words.at(-1)!.right).toBeCloseTo(540);
    expect(words[1]!.left - words[0]!.right).toBeCloseTo(204);
    expect(words[2]!.left - words[1]!.right).toBeCloseTo(204);
    const projected = textRunsForPage(layout, 0, { scale: 2 });
    expect(projected.map(run => run.text).sort()).toEqual(placements.map(node => node.text).sort());
    expect(projected.reduce((sum, run) => sum + run.w, 0)).toBeCloseTo(468 * 2);
  });
});
