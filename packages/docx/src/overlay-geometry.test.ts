import { afterEach, describe, expect, it, vi } from 'vitest';

import { resolveDocxCommentThreads } from './comments.js';
import { buildDocxHighlightLayer } from './find-highlight-layer.js';
import {
  clipRunRect, overlayBoxes, overlayRectsExamined, placeRunRect, runAdvanceRects, runGlyphRects, runOverlayRects, runTextVisible,
} from './overlay-geometry.js';
import { DocxFindController } from './find.js';
import { textRunSourceIndexForDocument } from './layout/text-index.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { collectLayoutSourceCommentRanges, collectLayoutSourceCommentRangesIfPresent } from './comments.js';
import { buildDocxCommentMargin } from './comment-margin.js';
import { installDom, makeEl } from './scroll-viewer-test-dom.js';
import { reviewProjectionIndexForDocument } from './layout/text-index.js';
import { buildDocxTextLayer } from './text-layer.js';
import { textRunsForPage } from './text-run-projection.js';
import {
  fakeElement, inlinePictureRun, intersect, layerBoxes, pageBoxOf, paintedDocument, PICTURE_PART, withFakeDocument,
  type FakeElement, type PageBox, type PaintedGlyph,
} from './test-support/painted-glyphs.test-support.js';
import type { DocxTextRunInfo } from './types.js';

afterEach(() => {
  vi.unstubAllGlobals();
});

/**
 * The overlay invariants (overlay-geometry.ts), checked against what is
 * painted: generated paragraphs are laid out, painted on a recording canvas
 * that reports the box of every drawn character under its clip, and every
 * overlay rect is compared with those boxes. Each generated page draws a
 * character at most once (spaces aside), so a drawn character identifies
 * its source character.
 */

/** Deterministic generator (mulberry32). */
function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

const POOLS = {
  latin: [...'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'],
  digit: [...'0123456789'],
  hebrew: [...'אבגדהוזחטיכלמנסעפצקרשת'],
  arabic: [...'بتثجحخسشصضطظعغفقكلمنهي'],
  cjk: Array.from({ length: 200 }, (_, index) => String.fromCodePoint(0x4e00 + index * 7)),
} as const;
type Script = keyof typeof POOLS;

interface Case {
  readonly name: string;
  readonly body: string;
  readonly vertical: boolean;
  readonly section: string;
  /** Source run indices of each paragraph a comment covers: `[paragraph, first, end)`. */
  readonly commented: readonly (readonly [number, number, number])[];
}

const TRACKINGS = [-110, -30, 0, 0, 40, 160] as const;

/** One generated document: a few paragraphs of runs, each with its own
 * script, tracking and fitting, in one of the page arrangements. */
function generate(seed: number): Case {
  const next = random(seed);
  const pick = <T>(values: readonly T[]): T => values[Math.floor(next() * values.length)]!;
  const pools: Record<Script, string[]> = {
    latin: [...POOLS.latin], digit: [...POOLS.digit], hebrew: [...POOLS.hebrew], arabic: [...POOLS.arabic], cjk: [...POOLS.cjk],
  };
  const draw = (script: Script, count: number): string => {
    let text = '';
    for (let index = 0; index < count && pools[script].length > 0; index += 1) {
      text += pools[script].splice(Math.floor(next() * pools[script].length), 1)[0]!;
    }
    return text;
  };
  const arrangement = pick(['plain', 'plain', 'rtl', 'mixed', 'kashida', 'grid', 'vertical', 'tate-chu-yoko', 'exact-row', 'cell'] as const);
  const vertical = arrangement === 'vertical' || arrangement === 'tate-chu-yoko';
  const section = arrangement === 'grid' ? '<w:docGrid w:type="snapToChars" w:linePitch="360" w:charSpace="2048"/>' : '';
  const commented: [number, number, number][] = [];
  const paragraphs: string[] = [];
  const count = 1 + Math.floor(next() * 3);
  for (let paragraph = 0; paragraph < count; paragraph += 1) {
    const rtl = arrangement === 'rtl' || arrangement === 'kashida' || (arrangement === 'mixed' && next() < 0.5);
    const alignment = arrangement === 'kashida' ? 'lowKashida'
      : pick(['left', 'left', 'both', 'distribute', 'center', 'right'] as const);
    const runs: string[] = [];
    const runCount = 1 + Math.floor(next() * 4);
    for (let run = 0; run < runCount; run += 1) {
      const script: Script = arrangement === 'kashida' ? 'arabic'
        : arrangement === 'vertical' ? pick(['cjk', 'cjk', 'latin'] as const)
          : arrangement === 'tate-chu-yoko' ? pick(['cjk', 'digit'] as const)
            : arrangement === 'grid' ? pick(['cjk', 'latin'] as const)
              : arrangement === 'rtl' ? pick(['hebrew', 'hebrew', 'digit'] as const)
                : arrangement === 'mixed' ? pick(['hebrew', 'latin', 'digit'] as const)
                  : pick(['latin', 'latin', 'cjk', 'digit'] as const);
      const tateChuYoko = arrangement === 'tate-chu-yoko' && script === 'digit';
      const length = tateChuYoko ? 2 : 1 + Math.floor(next() * 5);
      let text = draw(script, length);
      if (text === '') continue;
      // Words, so justification and distribution have gaps to widen.
      if (!tateChuYoko && next() < 0.5) text += ' ';
      const tracking = arrangement === 'kashida' ? 0 : pick(TRACKINGS);
      const fitted = !vertical && !rtl && arrangement !== 'grid' && script === 'latin' && next() < 0.2;
      const scale = !vertical && next() < 0.15 ? pick([66, 150]) : 100;
      // Off the line's baseline, or at another size.
      const raised = next() < 0.2 ? pick(['<w:position w:val="8"/>', '<w:position w:val="-6"/>',
        '<w:vertAlign w:val="superscript"/>', '<w:vertAlign w:val="subscript"/>']) : '';
      const size = !tateChuYoko && arrangement !== 'grid' && next() < 0.15 ? pick([16, 36]) : 24;
      runs.push('<w:r><w:rPr>'
        + (script === 'hebrew' || script === 'arabic' ? '<w:rtl/>' : '')
        + (scale !== 100 ? `<w:w w:val="${scale}"/>` : '')
        + (tracking !== 0 ? `<w:spacing w:val="${tracking}"/>` : '')
        + raised + `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>`
        + (fitted ? `<w:fitText w:val="${pick([400, 1200])}" w:id="${seed * 100 + paragraph * 10 + run + 1}"/>` : '')
        + (tateChuYoko ? '<w:eastAsianLayout w:id="1" w:vert="1"/>' : '')
        + `</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`);
    }
    if (runs.length === 0) continue;
    if (runs.length > 1 && next() < 0.7) {
      const first = Math.floor(next() * runs.length);
      commented.push([paragraphs.length, first, first + 1 + Math.floor(next() * (runs.length - first))]);
    }
    // A narrow measure wraps the paragraph, so justified lines exist.
    const narrow = alignment === 'both' || alignment === 'distribute' || arrangement === 'kashida';
    const exact = arrangement === 'exact-row' ? `<w:spacing w:before="0" w:after="0" w:line="${pick([240, 360, 800])}" w:lineRule="exact"/>` : '';
    // In a cell, some paragraphs start before the cell's edge.
    const outdent = (arrangement === 'cell' || arrangement === 'exact-row') && next() < 0.5 ? pick([-70, -150, -290]) : 0;
    const indent = narrow ? `<w:ind w:right="${vertical ? 9000 : 7200}"/>`
      : outdent !== 0 ? `<w:ind w:${rtl ? 'right' : 'left'}="${outdent}"/>` : '';
    paragraphs.push(`<w:p><w:pPr>${rtl ? '<w:bidi/>' : ''}${exact}<w:jc w:val="${alignment}"/>${indent}</w:pPr>`
      + `${runs.join('')}</w:p>`);
  }
  const content = paragraphs.join('');
  const body = arrangement === 'exact-row' || arrangement === 'cell'
    ? '<w:tbl><w:tblPr><w:tblW w:w="2400" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>'
      + `<w:tr>${arrangement === 'exact-row' ? `<w:trPr><w:trHeight w:val="${pick([200, 300, 500])}" w:hRule="exact"/></w:trPr>` : ''}`
      + `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${content || '<w:p/>'}</w:tc></w:tr></w:tbl><w:p/>`
    : content;
  return { name: `${arrangement} #${seed}`, body, vertical, section, commented };
}

const EPSILON = 1e-6;
const sameBox = (left: PageBox, right: PageBox): boolean => Math.abs(left.left - right.left) < EPSILON
  && Math.abs(left.top - right.top) < EPSILON && Math.abs(left.right - right.right) < EPSILON
  && Math.abs(left.bottom - right.bottom) < EPSILON;
const show = (box: PageBox | null): string => (box
  ? `[${box.left.toFixed(3)}, ${box.top.toFixed(3)} .. ${box.right.toFixed(3)}, ${box.bottom.toFixed(3)}]` : 'nothing');

/** Whether two sets of boxes cover the same region of the page. */
function regionDifference(left: readonly PageBox[], right: readonly PageBox[]): string | null {
  const snap = (value: number): number => Math.round(value / EPSILON) * EPSILON;
  const xs = [...new Set([...left, ...right].flatMap((box) => [snap(box.left), snap(box.right)]))].sort((a, b) => a - b);
  const ys = [...new Set([...left, ...right].flatMap((box) => [snap(box.top), snap(box.bottom)]))].sort((a, b) => a - b);
  const covers = (boxes: readonly PageBox[], x: number, y: number): boolean => boxes.some((box) => (
    snap(box.left) <= x && x < snap(box.right) && snap(box.top) <= y && y < snap(box.bottom)));
  for (let column = 0; column < xs.length - 1; column += 1) {
    if (xs[column + 1]! - xs[column]! <= 2 * EPSILON) continue;
    for (let row = 0; row < ys.length - 1; row += 1) {
      if (ys[row + 1]! - ys[row]! <= 2 * EPSILON) continue;
      const x = (xs[column]! + xs[column + 1]!) / 2;
      const y = (ys[row]! + ys[row + 1]!) / 2;
      if (covers(left, x, y) !== covers(right, x, y)) {
        return `(${x.toFixed(3)}, ${y.toFixed(3)}) is ${covers(left, x, y) ? 'marked but not painted' : 'painted but not marked'}`;
      }
    }
  }
  return null;
}

/** Every violation of the invariants on one generated page. */
function violations(testCase: Case, runs: readonly DocxTextRunInfo[], glyphs: readonly PaintedGlyph[]): string[] {
  const found: string[] = [];
  const painted = new Map<string, PaintedGlyph[]>();
  for (const glyph of glyphs) painted.set(glyph.character, [...(painted.get(glyph.character) ?? []), glyph]);
  const visibleOf = (character: string): PageBox | null | undefined => {
    const drawn = painted.get(character);
    if (!drawn || drawn.length !== 1) return undefined;
    return intersect(drawn[0]!.box, drawn[0]!.clip);
  };
  for (const run of runs) {
    for (let index = 0; index < run.text.length; index += 1) {
      const character = run.text[index]!;
      if (character === ' ') continue;
      const expected = visibleOf(character);
      if (expected === undefined) {
        found.push(`${JSON.stringify(character)} of ${JSON.stringify(run.text)} is drawn ${painted.get(character)?.length ?? 0} times`);
        continue;
      }
      // Invariants 1, 3 and 4: the rects that mark a character are its
      // painted glyph inside its clips, exactly.
      const rects = runOverlayRects(run, index, index + 1).map((rect) => pageBoxOf(placeRunRect(run, rect)));
      const marked = rects.length === 1 ? rects[0]! : null;
      if (rects.length > 1 || (marked === null) !== (expected === null) || (marked && expected && !sameBox(marked, expected))) {
        found.push(`${JSON.stringify(character)} of ${JSON.stringify(run.text)}: marked ${rects.map(show).join(' ') || 'nothing'}, painted ${show(expected)}`);
      }
      if (runTextVisible(run, index, index + 1) !== (expected !== null)) {
        found.push(`${JSON.stringify(character)} of ${JSON.stringify(run.text)}: visibility ${runTextVisible(run, index, index + 1)}, painted ${show(expected)}`);
      }
    }
    // Invariant 2: advance rects run from pen to pen without overlap.
    if (!run.eastAsianVert) {
      const advances = runAdvanceRects(run).sort((left, right) => left.x - right.x);
      for (let index = 1; index < advances.length; index += 1) {
        const previous = advances[index - 1]!;
        if (advances[index]!.x < previous.x + previous.width - EPSILON) {
          found.push(`advances of ${JSON.stringify(run.text)} overlap at ${advances[index]!.x.toFixed(3)}`);
        }
      }
    }
  }
  // Invariants 4 and 5: a comment over some runs of a paragraph marks
  // their painted glyphs inside the clips and the advance after each
  // (tracking, justification, elongation), and nothing else.
  const paragraphs = [...new Set(runs.map((run) => JSON.stringify(run.source?.path)))];
  for (const [paragraph, first, end] of testCase.commented) {
    const anchored = runs.filter((run) => JSON.stringify(run.source?.path) === paragraphs[paragraph]
      && run.sourceRunIndex !== undefined && run.sourceRunIndex >= first && run.sourceRunIndex < end);
    const expected = anchored.flatMap((run) => [...run.text].flatMap((character, index) => {
      // The advance box of the character, over the extent its glyph has
      // across the run.
      const glyph = runGlyphRects(run, index, index + 1)[0];
      const advance = runAdvanceRects(run, index, index + 1)
        .flatMap((rect) => clipRunRect(run, glyph ? { ...rect, y: glyph.y, height: glyph.height } : rect))
        .map((rect) => pageBoxOf(placeRunRect(run, rect)));
      if (character !== ' ') {
        // A glyph its clips hide whole marks nothing, its advance included.
        const box = visibleOf(character);
        return box ? [box, ...advance] : [];
      }
      // A space is drawn many times: its own rect stands for it.
      const own = runOverlayRects(run, index, index + 1).map((rect) => pageBoxOf(placeRunRect(run, rect)));
      return own.length > 0 || !glyph ? [...own, ...advance] : [];
    }));
    const boxes = overlayBoxes(anchored).map(pageBoxOf);
    const difference = regionDifference(boxes, expected);
    if (difference) found.push(`comment on runs [${first}, ${end}) of paragraph ${paragraph}: ${difference}`);
    // Joined boxes are disjoint: nothing is tinted twice.
    for (const [index, box] of boxes.entries()) {
      if (boxes.slice(index + 1).some((other) => intersect(box, other) !== null)) found.push(`comment boxes overlap at ${show(box)}`);
    }
  }
  return found;
}

describe('overlay rects are the painted glyphs inside their clips', () => {
  const seeds = Array.from({ length: 300 }, (_, index) => index + 1);
  it.each(seeds.map((seed) => [generate(seed).name, seed] as const))('%s', async (_name, seed) => {
    const testCase = generate(seed);
    const { runs, glyphs } = await paintedDocument(testCase.body, { vertical: testCase.vertical, section: testCase.section });
    expect(runs.length).toBeGreaterThan(0);
    expect(violations(testCase, runs, glyphs)).toEqual([]);
  });
});

/** The find highlight boxes of `text` on a painted page, as the viewer's
 * layer draws them. */
function findBoxes(runs: DocxTextRunInfo[], text: string): PageBox[] {
  const runIndex = runs.findIndex((run) => run.text.includes(text));
  const start = runs[runIndex]!.text.indexOf(text);
  return withFakeDocument(() => {
    const layer = fakeElement();
    buildDocxHighlightLayer(layer as never, runs, [{ active: false, slices: [{ runIndex, start, end: start + text.length }] }],
      612, 792, () => () => { throw new Error('a projected run is never measured'); });
    return layerBoxes(layer, 612, 792);
  });
}

/** Where the characters of `text` are painted inside their clips. */
function paintedBoxes(glyphs: readonly PaintedGlyph[], text: string): PageBox[] {
  return [...text].flatMap((character) => glyphs.filter((glyph) => glyph.character === character)
    .flatMap((glyph) => intersect(glyph.box, glyph.clip) ?? []));
}

const exactRowCell = (rowTwips: number, paragraph: string): string => (
  '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
  + '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
  + `<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="${rowTwips}" w:hRule="exact"/></w:trPr>`
  + `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${paragraph}</w:tc></w:tr></w:tbl><w:p/>`);

describe('find highlights are the painted glyphs inside their clips', () => {
  it('gives no highlight to a line painted below an exact-height row', async () => {
    const { runs, glyphs } = await paintedDocument(exactRowCell(240,
      '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Visible</w:t><w:br/><w:t>Secret</w:t></w:r></w:p>'));
    expect(paintedBoxes(glyphs, 'Vs')).toHaveLength(2);
    expect(paintedBoxes(glyphs, 'S')).toEqual([]);
    expect(findBoxes(runs, 'Secret')).toEqual([]);
    expect(regionDifference(findBoxes(runs, 'Visible'), paintedBoxes(glyphs, 'Visble'))).toBeNull();
  });

  it('covers only the part of the font box a partially visible exact-height row shows', async () => {
    // 12pt text on a 40pt exact line in a 25pt exact row: the font box
    // starts below the row top and is cut at the row bottom.
    const { runs, glyphs } = await paintedDocument(exactRowCell(500,
      '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="800" w:lineRule="exact"/></w:pPr>'
      + '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Sample</w:t></w:r></w:p>'));
    const painted = paintedBoxes(glyphs, 'Sample');
    const found = findBoxes(runs, 'Sample');
    expect(found).toHaveLength(1);
    expect(regionDifference(found, painted)).toBeNull();
    // Blank leading above the font box is not highlighted.
    const run = runs.find((candidate) => candidate.text === 'Sample')!;
    expect(found[0]!.top).toBeGreaterThan(run.y + 1);
    expect(found[0]!.bottom - found[0]!.top).toBeLessThan(12);
  });

  it('centres an upright vertical glyph where its operation paints it, whatever its tracking', async () => {
    const { runs, glyphs } = await paintedDocument(
      '<w:p><w:r><w:rPr><w:spacing w:val="160"/><w:sz w:val="24"/></w:rPr><w:t>漢字</w:t></w:r></w:p>', { vertical: true });
    for (const character of '漢字') {
      const [found] = findBoxes(runs, character);
      const [painted] = paintedBoxes(glyphs, character);
      expect(found).toBeDefined();
      expect(regionDifference([found!], [painted!])).toBeNull();
    }
  });

  it('gives a glyph painted in its own operation the width it is drawn at alone', async () => {
    // Condensed past its advance, the `o` floors and is painted alone; shaped
    // with its neighbours it is kerned, alone it is not.
    const { runs, glyphs } = await paintedDocument(
      '<w:p><w:r><w:rPr><w:kern w:val="2"/><w:spacing w:val="-110"/><w:sz w:val="24"/></w:rPr><w:t>ToT</w:t></w:r></w:p>',
      { kerning: { To: 0.07, oT: 0.07 } });
    const [painted] = paintedBoxes(glyphs, 'o');
    const [found] = findBoxes(runs, 'o');
    expect(painted!.right - painted!.left).toBeCloseTo(6, 9);
    expect(regionDifference([found!], [painted!])).toBeNull();
  });
});

describe('comment highlights are the painted glyphs of the commented runs', () => {
  it('leaves uncommented digits between two parts of a right-to-left run untinted', async () => {
    const run = (text: string, rtl: boolean) => `<w:r><w:rPr>${rtl ? '<w:rtl/>' : ''}<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>${text}</w:t></w:r>`;
    const { runs, glyphs } = await paintedDocument(
      `<w:p><w:pPr><w:bidi/></w:pPr>${run('אב', true)}${run('12', false)}${run('34', false)}${run('גד', true)}</w:p>`);
    const source = runs[0]!.source!;
    const anchor = {
      commentId: '1', source, startRunIndex: 0, endRunIndex: 2,
      reference: { source, runIndex: 2, affinity: 'preceding' as const },
    };
    const [thread] = resolveDocxCommentThreads([{ id: '1', author: 'A', text: 'c' }], [anchor], runs);
    const rects = thread!.anchors[0]!.rects.map(pageBoxOf);
    expect(regionDifference(rects, paintedBoxes(glyphs, 'אב12'))).toBeNull();
    for (const other of paintedBoxes(glyphs, '34גד')) {
      expect(rects.some((rect) => intersect(rect, other) !== null)).toBe(false);
    }
  });
});

describe('a comment highlight covers everything in its range of runs and nothing outside it', () => {
  const run = (text: string, properties = '') => `<w:r><w:rPr>${properties}<w:sz w:val="24"/></w:rPr>${text}</w:r>`;
  const text = (value: string) => `<w:t xml:space="preserve">${value}</w:t>`;
  const commentRects = (runs: DocxTextRunInfo[], first: number, end: number): PageBox[] => {
    const source = runs[0]!.source!;
    const [thread] = resolveDocxCommentThreads([{ id: '1', author: 'A', text: 'c' }], [{
      commentId: '1', source, startRunIndex: first, endRunIndex: end,
      reference: { source, runIndex: end, affinity: 'preceding' as const },
    }], runs);
    return thread!.anchors[0]!.rects.map(pageBoxOf);
  };
  const body = `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="2880"/></w:tabs></w:pPr>${run(text('ab'))}${run('<w:tab/>')}${run(text('cd'))}</w:p>`;

  it('tints a tab between two commented runs with the text around it', async () => {
    const { runs, glyphs } = await paintedDocument(body);
    const [a] = paintedBoxes(glyphs, 'a');
    const [d] = paintedBoxes(glyphs, 'd');
    const rects = commentRects(runs, 0, 3);
    expect(rects).toHaveLength(1);
    for (const [side, expected] of [['left', a!.left], ['top', a!.top], ['right', d!.right], ['bottom', d!.bottom]] as const) {
      expect(rects[0]![side]).toBeCloseTo(expected, 6);
    }
    // The tab stop is 144pt from the text edge: the tab is most of the tint.
    expect(d!.left - a!.right).toBeGreaterThan(100);
  });

  it('leaves a tab outside the commented runs untinted', async () => {
    const { runs, glyphs } = await paintedDocument(body);
    // Two comments, one on each text run: neither covers the tab between.
    const before = commentRects(runs, 0, 1);
    const after = commentRects(runs, 2, 3);
    expect(regionDifference(before, paintedBoxes(glyphs, 'ab'))).toBeNull();
    expect(regionDifference(after, paintedBoxes(glyphs, 'cd'))).toBeNull();
  });

  it('tints the tracking and the widened spaces inside the range, and none of the text after it', async () => {
    const { runs, glyphs } = await paintedDocument(
      `<w:p><w:pPr><w:jc w:val="both"/><w:ind w:right="7600"/></w:pPr>${run(text('ab cd '), '<w:spacing w:val="60"/>')}`
      + `${run(text('ef gh ij kl mn op'))}</w:p>`);
    const rects = commentRects(runs, 0, 1);
    const [a] = paintedBoxes(glyphs, 'a');
    const [d] = paintedBoxes(glyphs, 'd');
    const [e] = paintedBoxes(glyphs, 'e');
    // One unbroken tint from the first glyph past the last one of the range.
    expect(rects).toHaveLength(1);
    expect(rects[0]!.left).toBeCloseTo(a!.left, 6);
    expect(rects[0]!.right).toBeGreaterThanOrEqual(d!.right + 3 - 1e-6);
    expect(rects[0]!.right).toBeLessThanOrEqual(e!.left + 1e-6);
  });
});

describe('text an exact-height row hides', () => {
  const hidden = exactRowCell(240,
    '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Visible</w:t><w:br/><w:t>Secret</w:t></w:r></w:p>');

  it('is no find match, while the visible line is', async () => {
    const { runs } = await paintedDocument(hidden);
    const find = new DocxFindController(() => 1, async () => runs);
    expect(await find.find('Secret')).toEqual([]);
    expect((await find.find('Visible')).map((match) => match.text)).toEqual(['Visible']);
  });

  it('is not selectable in the text layer', async () => {
    const { runs } = await paintedDocument(hidden);
    const spans: FakeElement[] = [];
    const collect = (element: FakeElement): void => {
      if (element.children.length === 0 && (element as { textContent?: string }).textContent) spans.push(element);
      element.children.forEach(collect);
    };
    withFakeDocument(() => {
      const layer = fakeElement();
      buildDocxTextLayer(layer as never, runs, 612, 792);
      collect(layer);
    });
    const text = (span: FakeElement): string => (span as { textContent?: string }).textContent ?? '';
    const of = (value: string) => spans.filter((span) => (
      span.dataset.runIndex === String(runs.findIndex((run) => run.text === value))));
    const secret = of('Secret');
    const visible = of('Visible');
    expect(secret.map(text).join('')).toBe('Secret');
    expect(visible.map(text).join('')).toBe('Visible');
    for (const span of secret) {
      expect(span.dataset.clippedAway).toBe('');
      expect(span.style.cssText).toContain('user-select:none');
    }
    for (const span of visible) {
      expect(span.dataset.clippedAway).toBeUndefined();
      expect(span.style.cssText).not.toContain('user-select:none');
    }
  });
});

describe('a comment on inline content that is not text', () => {
  const sized = '<w:sz w:val="24"/>';
  const letters = (value: string) => `<w:r><w:rPr>${sized}</w:rPr><w:t>${value}</w:t></w:r>`;
  const rectsOf = (runs: DocxTextRunInfo[], first: number, end: number): PageBox[] => {
    const source = runs[0]!.source!;
    const [thread] = resolveDocxCommentThreads([{ id: '1', author: 'A', text: 'c' }], [{
      commentId: '1', source, startRunIndex: first, endRunIndex: end,
      reference: { source, runIndex: end, affinity: 'preceding' as const },
    }], runs);
    return thread?.anchors[0]?.rects.map(pageBoxOf) ?? [];
  };

  it('tints an inline picture of the range over the box it is painted in, with the text around it', async () => {
    const { runs, glyphs, pictures } = await paintedDocument(
      `<w:p>${letters('ab')}${inlinePictureRun(40, 30)}${letters('cd')}</w:p>`, PICTURE_PART);
    expect(pictures).toHaveLength(1);
    const picture = intersect(pictures[0]!.box, pictures[0]!.clip)!;
    expect(picture.right - picture.left).toBeCloseTo(40, 6);
    expect(regionDifference(rectsOf(runs, 0, 3), [...paintedBoxes(glyphs, 'abcd'), picture])).toBeNull();
    // A comment on the picture alone tints the picture alone.
    expect(regionDifference(rectsOf(runs, 1, 2), [picture])).toBeNull();
  });

  it('leaves an inline picture outside the range untinted', async () => {
    const { runs, glyphs } = await paintedDocument(
      `<w:p>${letters('ab')}${inlinePictureRun(40, 30)}${letters('cd')}</w:p>`, PICTURE_PART);
    expect(regionDifference(rectsOf(runs, 0, 1), paintedBoxes(glyphs, 'ab'))).toBeNull();
    expect(regionDifference(rectsOf(runs, 2, 3), paintedBoxes(glyphs, 'cd'))).toBeNull();
  });

  it('tints a raised picture only where its exact line shows it', async () => {
    // A 30pt picture raised 5pt on an exact 24pt line: the line cuts it.
    const { runs, pictures } = await paintedDocument(
      '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/></w:pPr>'
      + `${inlinePictureRun(40, 30, '<w:position w:val="10"/>')}</w:p>`, PICTURE_PART);
    expect(pictures[0]!.clip).not.toBeNull();
    const shown = intersect(pictures[0]!.box, pictures[0]!.clip)!;
    expect(shown.bottom - shown.top).toBeLessThan(30);
    expect(regionDifference(rectsOf(runs, 0, 1), [shown])).toBeNull();
  });

  it('marks nothing for a picture its line hides whole, and none of the text after it', async () => {
    // A 15pt picture lowered 24pt on an exact 12pt line: the line shows none of it.
    const { layout, runs, pictures } = await paintedDocument(
      '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>'
      + `${inlinePictureRun(30, 15, '<w:position w:val="-48"/>')}${letters('Uncommented')}</w:p>`, PICTURE_PART);
    expect(pictures.every((picture) => intersect(picture.box, picture.clip) === null)).toBe(true);
    // The picture keeps its source run: a comment on it needs no fallback.
    const picture = runs.find((run) => run.inlineContent === 'object');
    expect(picture?.sourceRunIndex).toBe(0);
    expect([...textRunSourceIndexForDocument(layout).values()].some((indices) => indices.has(0))).toBe(true);
    expect(rectsOf(runs, 0, 1)).toEqual([]);
  });

  it('keeps the line-box height of a tab that touches only a picture of the range', async () => {
    // A tab before a 5pt-high picture, both commented; the text is not.
    const { runs, pictures } = await paintedDocument(
      '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1440"/></w:tabs></w:pPr>'
      + `<w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:tab/></w:r>${inlinePictureRun(30, 5)}`
      + '<w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>after</w:t></w:r></w:p>', PICTURE_PART);
    const tab = runs.find((run) => run.inlineContent === 'tab')!;
    const picture = intersect(pictures[0]!.box, pictures[0]!.clip)!;
    expect(regionDifference(rectsOf(runs, 0, 2), [
      { left: tab.x, top: tab.y, right: tab.x + tab.w, bottom: tab.y + tab.h }, picture,
    ])).toBeNull();
    expect(tab.h).toBeGreaterThan(5 + 1);
  });

  it('gives a comment on a paragraph that holds only a tab its own geometry, not the next text', async () => {
    const comments = '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:comment w:id="1" w:author="A"><w:p><w:r><w:t>c</w:t></w:r></w:p></w:comment></w:comments>';
    const { layout, model, runs } = await paintedDocument(
      '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="2880"/></w:tabs></w:pPr><w:commentRangeStart w:id="1"/>'
      + `<w:r><w:rPr>${sized}</w:rPr><w:tab/></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>`
      + `<w:p><w:pPr><w:pageBreakBefore/></w:pPr>${letters('Unrelated')}</w:p>`, {
        parts: { 'word/comments.xml': comments },
        relationships: '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
          + 'Target="comments.xml"/>',
        contentTypes: '<Override PartName="/word/comments.xml" '
          + 'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>',
      });
    expect(layout.pages).toHaveLength(2);
    const anchors = collectLayoutSourceCommentRanges(model.comments ?? [], layoutSourceStore(model),
      textRunSourceIndexForDocument(layout));
    expect(anchors).toHaveLength(1);
    expect(anchors[0]!.geometryFallback).toBeUndefined();
    const [thread] = resolveDocxCommentThreads(model.comments ?? [], anchors, runs);
    expect(thread!.anchors[0]!.rects).toHaveLength(1);
    expect(thread!.anchors[0]!.rects[0]!.width).toBeCloseTo(144, 6);
  });

  it('gives each tab of many lines the extent of the text on its own line', async () => {
    // Lines of two sizes in turn: every tab joins its own line into one rect.
    const line = (size: number, index: number) => `<w:r><w:rPr><w:sz w:val="${size}"/></w:rPr><w:t>a</w:t></w:r>`
      + `<w:r><w:rPr><w:sz w:val="${size}"/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:sz w:val="${size}"/></w:rPr><w:t>b</w:t>`
      + `${index < 39 ? '<w:br/>' : ''}</w:r>`;
    const { runs } = await paintedDocument('<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="2880"/></w:tabs></w:pPr>'
      + `${Array.from({ length: 40 }, (_, index) => line(index % 2 === 0 ? 24 : 16, index)).join('')}</w:p>`);
    const rects = rectsOf(runs, 0, Number.MAX_SAFE_INTEGER);
    expect(rects).toHaveLength(40);
    rects.forEach((rect, index) => {
      expect(rect.bottom - rect.top).toBeCloseTo(index % 2 === 0 ? 12 : 8, 6);
      expect(rect.left).toBeCloseTo(72, 6);
    });
  });
});

describe('the work of marking a commented paragraph of many tabbed lines', () => {
  /** `lines` lines of `a<tab>b`, all runs of one comment. */
  const tabbedLines = (lines: number): DocxTextRunInfo[] => Array.from({ length: lines }, (_, line) => {
    const base = { y: line * 14, h: 12, fontSize: 12, font: '12px serif', sourceRunIndex: line };
    return [
      { ...base, text: 'a', x: 0, w: 6 },
      { ...base, text: '', inlineContent: 'tab' as const, x: 6, w: 138 },
      { ...base, text: 'b', x: 144, w: 6 },
    ];
  }).flat();
  /** Rects the tab lookup examines while marking `lines` lines. */
  const examined = (lines: number): number => {
    const runs = tabbedLines(lines);
    const before = overlayRectsExamined();
    expect(overlayBoxes(runs)).toHaveLength(lines);
    return overlayRectsExamined() - before;
  };

  it('grows with the number of lines, not with its square', () => {
    // Four times the lines: a tab that looked at every line's rects would
    // examine sixteen times as many.
    const small = examined(60);
    const large = examined(240);
    expect(small).toBeGreaterThan(0);
    expect(large / small).toBeLessThan(6);
  });
});

describe('a comment on text its clips hide whole', () => {
  it('is no thread of the page', async () => {
    const { runs } = await paintedDocument(exactRowCell(240,
      '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Visible</w:t></w:r>'
      + '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:br/><w:br/><w:t>Hidden</w:t></w:r></w:p>'));
    const hidden = runs.find((run) => run.text === 'Hidden')!;
    const shown = runs.find((run) => run.text === 'Visible')!;
    const threads = (run: DocxTextRunInfo) => resolveDocxCommentThreads([{ id: '1', author: 'A', text: 'c' }], [{
      commentId: '1', source: run.source!, startRunIndex: run.sourceRunIndex!, endRunIndex: run.sourceRunIndex! + 1,
      reference: { source: run.source!, runIndex: run.sourceRunIndex! + 1, affinity: 'preceding' as const },
    }], runs);
    expect(threads(hidden)).toEqual([]);
    expect(threads(shown)).toHaveLength(1);
  });
});

describe('the work of covering overlapping marks once', () => {
  /** `lines` lines of 60 single-character runs, every other one raised by
   * 2px: neighbours touch without overlapping, and no two join. */
  const staggered = (lines: number): DocxTextRunInfo[] => Array.from({ length: lines }, (_, line) => (
    Array.from({ length: 60 }, (_run, index) => ({
      text: 'x', x: index * 6, y: line * 20 - (index % 2) * 2, w: 6, h: 12, fontSize: 12, font: '12px serif',
      sourceRunIndex: line * 60 + index,
    })))).flat();
  const comparisons = (lines: number): number => {
    const runs = staggered(lines);
    const counted = vi.spyOn(Math, 'min');
    try {
      expect(overlayBoxes(runs)).toHaveLength(lines * 60);
      return counted.mock.calls.length;
    } finally {
      counted.mockRestore();
    }
  };

  it('grows with the length of one line, not with its square', () => {
    // One line of touching runs, every other one raised: no two overlap,
    // and four times the runs must not take sixteen times the comparisons.
    const line = (count: number): DocxTextRunInfo[] => Array.from({ length: count }, (_, index) => ({
      text: 'x', x: index * 6, y: 100 - (index % 2) * 2, w: 6, h: 12, fontSize: 12, font: '12px serif', sourceRunIndex: index,
    }));
    const count = (runs: DocxTextRunInfo[]): number => {
      const counted = vi.spyOn(Math, 'min');
      try {
        expect(overlayBoxes(runs)).toHaveLength(runs.length);
        return counted.mock.calls.length;
      } finally {
        counted.mockRestore();
      }
    };
    const small = count(line(300));
    const large = count(line(1200));
    expect(large).toBeLessThan(Math.max(small, 300) * 6);
  });

  it('grows with the number of lines set closer than their glyphs are tall', () => {
    // One 12px glyph per line, lines 8px apart: each overlaps the next, and
    // the whole paragraph is one group.
    const tight = (lines: number): DocxTextRunInfo[] => Array.from({ length: lines }, (_, line) => ({
      text: 'x', x: 72, y: 72 + line * 8, w: 6, h: 12, fontSize: 12, font: '12px serif', sourceRunIndex: line,
    }));
    const count = (runs: DocxTextRunInfo[]): number => {
      const counted = vi.spyOn(Math, 'min');
      try {
        expect(overlayBoxes(runs)).toHaveLength(1);
        return counted.mock.calls.length;
      } finally {
        counted.mockRestore();
      }
    };
    for (const [few, many] of [[20, 80], [200, 800]] as const) {
      const small = count(tight(few));
      const large = count(tight(many));
      expect(small).toBeGreaterThan(0);
      expect(large / small).toBeLessThan(6);
    }
  });

  it('grows with the length of one line of condensed runs that overlap their neighbours', () => {
    // 6px glyphs 4.5px apart, every other one raised by 1px: each overlaps
    // the next, and the whole line is one group.
    const condensed = (count: number): DocxTextRunInfo[] => Array.from({ length: count }, (_, index) => ({
      text: 'x', x: index * 4.5, y: 100 - (index % 2), w: 6, h: 12, fontSize: 12, font: '12px serif', sourceRunIndex: index,
    }));
    const count = (runs: DocxTextRunInfo[]): number => {
      const counted = vi.spyOn(Math, 'min');
      try {
        expect(overlayBoxes(runs).length).toBeGreaterThan(0);
        return counted.mock.calls.length;
      } finally {
        counted.mockRestore();
      }
    };
    for (const [few, many] of [[24, 96], [300, 1200]] as const) {
      const small = count(condensed(few));
      const large = count(condensed(many));
      expect(small).toBeGreaterThan(0);
      expect(large / small).toBeLessThan(6);
    }
  });

  it('grows with the number of lines, not with its square', () => {
    // Four times the lines: comparing every pair of rects would take sixteen
    // times the comparisons.
    const small = comparisons(10);
    const large = comparisons(40);
    expect(large).toBeLessThan(Math.max(small, 600) * 6);
  });
});

describe('a comment on the end of a paragraph that runs over two pages', () => {
  it('is shown on the page that holds the commented run, and not on the page before', async () => {
    const comments = '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:comment w:id="1" w:author="A"><w:p><w:r><w:t>c</w:t></w:r></w:p></w:comment></w:comments>';
    const line = '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Uncommented</w:t><w:br/></w:r>';
    const { layout, model, runs } = await paintedDocument(
      `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>${line.repeat(55)}`
      + '<w:commentRangeStart w:id="1"/><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Commented</w:t></w:r>'
      + '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>', {
        parts: { 'word/comments.xml': comments },
        relationships: '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
          + 'Target="comments.xml"/>',
        contentTypes: '<Override PartName="/word/comments.xml" '
          + 'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>',
      });
    expect(layout.pages).toHaveLength(2);
    expect(runs.some((run) => run.text === 'Commented')).toBe(false);
    const anchors = collectLayoutSourceCommentRanges(model.comments ?? [], layoutSourceStore(model),
      textRunSourceIndexForDocument(layout));
    expect(anchors).toHaveLength(1);
    // Page one holds the paragraph but none of the commented run.
    expect(resolveDocxCommentThreads(model.comments ?? [], anchors, runs)).toEqual([]);
    const second = textRunsForPage(layout, 1, { scale: 1 });
    const [thread] = resolveDocxCommentThreads(model.comments ?? [], anchors, second);
    const commented = second.find((run) => run.text === 'Commented')!;
    expect(thread!.anchors[0]!.rects).toHaveLength(1);
    expect(thread!.anchors[0]!.rects[0]!.x).toBeCloseTo(commented.x, 6);
  });
});

const COMMENT_PART = {
  parts: {
    'word/comments.xml': '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      + '<w:comment w:id="1" w:author="A"><w:p><w:r><w:t>c</w:t></w:r></w:p></w:comment></w:comments>',
  },
  relationships: '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
    + 'Target="comments.xml"/>',
  contentTypes: '<Override PartName="/word/comments.xml" '
    + 'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>',
} as const;
const sized = (value: string) => `<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${value}</w:t></w:r>`;

describe('the card of a comment that starts in clipped text and ends on a later page', () => {
  it('stands on the page that shows the comment, decided once for the document', async () => {
    // The comment starts on text an exact-height row hides (page one) and
    // ends on visible text after a page break (page two).
    const { layout, model, runs } = await paintedDocument(
      exactRowCell(240, '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="800" w:lineRule="exact"/></w:pPr>'
        + `<w:commentRangeStart w:id="1"/>${sized('Hidden')}</w:p>`).replace(/<w:p\/>$/u, '')
      + `<w:p><w:pPr><w:pageBreakBefore/></w:pPr>${sized('Shown')}<w:commentRangeEnd w:id="1"/>`
      + '<w:r><w:commentReference w:id="1"/></w:r></w:p>', COMMENT_PART);
    expect(layout.pages).toHaveLength(2);
    const index = reviewProjectionIndexForDocument(layout);
    const anchors = collectLayoutSourceCommentRangesIfPresent(model.comments, layoutSourceStore(model), index.renderedRunIndex, {
      cardOwners: { sourcePages: index.sourcePages, runsForPage: (page) => textRunsForPage(layout, page, { scale: 1 }) },
    });
    // The owner is the range of the paragraph a page shows, not the first one.
    expect(anchors.filter((anchor) => anchor.cardOwner).map((anchor) => anchor.source.path)).toEqual(
      [textRunsForPage(layout, 1, { scale: 1 }).find((run) => run.text === 'Shown')!.source!.path]);
    installDom();
    const cards = (pageRuns: DocxTextRunInfo[]): number => {
      const tint = makeEl('div');
      const margin = makeEl('div');
      margin.clientHeight = 800;
      buildDocxCommentMargin(tint as never, margin as never, pageRuns, { comments: model.comments ?? [], anchors },
        612, 792, null, () => undefined, 1, 260, false);
      return margin.children.length;
    };
    expect(cards(runs)).toBe(0);
    expect(cards(textRunsForPage(layout, 1, { scale: 1 }))).toBe(1);
  });
});

describe('the card of a comment whose paragraph crosses a page break', () => {
  const crossing = (lines: number) => '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>'
    + '<w:commentRangeStart w:id="1"/>'
    + Array.from({ length: lines }, (_, line) => `${line > 0 ? '<w:r><w:br/></w:r>' : ''}${sized(`L${line}`)}`).join('')
    + '</w:p>';
  const cardsPerPage = (layout: Awaited<ReturnType<typeof paintedDocument>>['layout'],
    model: Awaited<ReturnType<typeof paintedDocument>>['model']) => {
    const index = reviewProjectionIndexForDocument(layout);
    const anchors = collectLayoutSourceCommentRangesIfPresent(model.comments, layoutSourceStore(model), index.renderedRunIndex,
      { cardOwners: { sourcePages: index.sourcePages, runsForPage: (page) => textRunsForPage(layout, page, { scale: 1 }) } });
    installDom();
    return layout.pages.map((_, pageIndex) => {
      const tint = makeEl('div');
      const margin = makeEl('div');
      margin.clientHeight = 800;
      buildDocxCommentMargin(tint as never, margin as never, textRunsForPage(layout, pageIndex, { scale: 1 }),
        { comments: model.comments ?? [], anchors, pageIndex }, 612, 792, null, () => undefined, 1, 260, false);
      return [margin.children.length, tint.children.length > 0];
    });
  };

  it('stands on the first page that shows the owner range only, the range tinted on every page', async () => {
    // The comment starts before the first line of a 70-line paragraph that
    // crosses into page two and ends in a paragraph that starts page three.
    const { layout, model } = await paintedDocument(crossing(70)
      + `<w:p><w:pPr><w:pageBreakBefore/></w:pPr>${sized('Tail')}<w:commentRangeEnd w:id="1"/>`
      + '<w:r><w:commentReference w:id="1"/></w:r></w:p>', COMMENT_PART);
    expect(layout.pages).toHaveLength(3);
    expect(cardsPerPage(layout, model)).toEqual([[1, true], [0, true], [0, true]]);
  });

  it('stands with a reference-only point on the page of the run it follows', async () => {
    // The paragraph crosses into page two; the reference stands right after
    // "Target", a run on page two. Page one shows nothing of the comment.
    const lines = Array.from({ length: 70 }, (_, line) => `${line > 0 ? '<w:r><w:br/></w:r>' : ''}${sized(line === 65 ? 'Target' : `L${line}`)}`
      + (line === 65 ? '<w:r><w:commentReference w:id="1"/></w:r>' : ''));
    const { layout, model } = await paintedDocument('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>'
      + `${lines.join('')}</w:p>`, COMMENT_PART);
    expect(layout.pages).toHaveLength(2);
    expect(cardsPerPage(layout, model)).toEqual([[0, false], [1, true]]);
  });

  it('stands on the first page of a single range that crosses a page break', async () => {
    const { layout, model } = await paintedDocument(crossing(70).replace('</w:p>',
      '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>'), COMMENT_PART);
    expect(layout.pages).toHaveLength(2);
    expect(cardsPerPage(layout, model)).toEqual([[1, true], [0, true]]);
  });
});

describe('a comment that starts at the end of a paragraph', () => {
  it('marks the commented paragraph only, not the run before the range', async () => {
    const { layout, model, runs } = await paintedDocument(
      `<w:p>${sized('Uncommented')}<w:commentRangeStart w:id="1"/></w:p>`
      + `<w:p>${sized('Commented')}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>`, COMMENT_PART);
    const anchors = collectLayoutSourceCommentRanges(model.comments ?? [], layoutSourceStore(model),
      textRunSourceIndexForDocument(layout));
    const [thread] = resolveDocxCommentThreads(model.comments ?? [], anchors, runs);
    const rects = thread!.anchors.flatMap((anchor) => anchor.rects.map(pageBoxOf));
    const commented = runs.find((run) => run.text === 'Commented')!;
    expect(rects).toHaveLength(1);
    expect(rects[0]!.top).toBeCloseTo(commented.y, 6);
    expect(rects[0]!.left).toBeCloseTo(commented.x, 6);
  });
});

describe('a find match over runs whose glyphs overlap', () => {
  it('tints the shared area once', async () => {
    // A condensed 12pt `A` under the start of an 18pt `B`.
    const { runs, glyphs } = await paintedDocument(
      '<w:p><w:r><w:rPr><w:spacing w:val="-80"/><w:sz w:val="24"/></w:rPr><w:t>A</w:t></w:r>'
      + '<w:r><w:rPr><w:sz w:val="36"/></w:rPr><w:t>B</w:t></w:r></w:p>');
    const [a] = paintedBoxes(glyphs, 'A');
    const [b] = paintedBoxes(glyphs, 'B');
    expect(intersect(a!, b!)).not.toBeNull();
    const boxes = withFakeDocument(() => {
      const layer = fakeElement();
      buildDocxHighlightLayer(layer as never, runs, [{
        active: false,
        slices: runs.flatMap((run, runIndex) => (run.text === 'A' || run.text === 'B' ? [{ runIndex, start: 0, end: 1 }] : [])),
      }], 612, 792, () => () => { throw new Error('a projected run is never measured'); });
      return layerBoxes(layer, 612, 792);
    });
    expect(regionDifference(boxes, [a!, b!])).toBeNull();
    for (const [index, box] of boxes.entries()) {
      expect(boxes.slice(index + 1).some((other) => intersect(box, other) !== null)).toBe(false);
    }
  });
});
