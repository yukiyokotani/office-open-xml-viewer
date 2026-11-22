import { beforeAll, describe, expect, it } from 'vitest';
import {
  anchoredRectangle,
  anchoring,
  collisions,
  hasNoPlace,
  layOutPages,
  loadParser,
  P,
  PAGE_H,
  PAGE_W,
  shared,
  T,
  table,
  type Direction,
  type Page,
} from './test-support/vertical-section-upright.test-support.js';

// Generated sequences for `word-vertical-section-upright-block-table`: fixed
// pseudo-random streams (mulberry32) of paragraphs, some anchoring
// square-wrapped rectangles, column breaks, runs of directly following
// tables and tables continued over pages, in one and two columns of both
// directions. Every sequence has to lay out, with no table box sharing an
// area with another table or with a drawing the page has a place for.
//
// Each stream is laid out for `SAMPLES` sequences per direction and column
// count. DOCX_UPRIGHT_GENERATED_SAMPLES sets that number; a larger run lays
// out the same sequences and further ones, since every direction and column
// count has its own stream.

const SAMPLES = Math.max(1, Math.floor(Number(process.env.DOCX_UPRIGHT_GENERATED_SAMPLES ?? 8)));
const TIMEOUT_MS = 60_000 + SAMPLES * 15_000;

const stream = (start: number) => {
  let seed = start;
  return () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let value = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    value = (value + Math.imul(value ^ (value >>> 7), 61 | value)) ^ value;
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
};
const columnBreak = '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
const CONFIGURATIONS = (['tbRl', 'tbLrV'] as const).flatMap((direction) =>
  [1, 2].map((columns) => ({ direction, columns })));

/** Lays `body` out and checks the tables of every page. `avoidsDrawings`:
 * no drawing of the sequence may overlap (ECMA-376 §20.4.2.3). */
const check = (
  direction: Direction, columns: number, body: readonly string[], context: string, avoidsDrawings = false,
) => {
  let pages: Page[] = [];
  expect(() => { pages = layOutPages(direction, body, columns); }, context).not.toThrow();
  expect(collisions(pages, direction, avoidsDrawings), context).toEqual([]);
  return pages;
};
/** Every pair of drawings of different paragraphs of a page sharing an
 * area, but for a later drawing the page has no place for. */
const drawingCollisions = (pages: readonly Page[], direction: Direction) => pages.flatMap((page, pageIndex) =>
  page.drawings.flatMap((first, at) => page.drawings.slice(at + 1).flatMap((second, offset) =>
    page.drawingParagraphs[at] !== page.drawingParagraphs[at + 1 + offset]
      && shared(first.left, first.left + first.width, second.left, second.left + second.width)
      && shared(first.top, first.top + first.height, second.top, second.top + second.height)
      && !hasNoPlace(direction, page, second, true)
      ? [{ pageIndex, first, second }] : [])));
/** Every drawing reaching past an edge of its page. */
const offPage = (pages: readonly Page[]) => pages.flatMap((page, pageIndex) =>
  page.drawings.flatMap((drawing) => drawing.left < -1e-6 || drawing.top < -1e-6
    || drawing.left + drawing.width > PAGE_W + 1e-6 || drawing.top + drawing.height > PAGE_H + 1e-6
    ? [{ pageIndex, drawing }] : []));
/** Runs `sequence` for SAMPLES sequences of every direction and column
 * count, each with its own stream from `start`. */
const forEachSequence = (
  start: number,
  sequence: (random: () => number, direction: Direction, columns: number, context: string) => void,
) => {
  CONFIGURATIONS.forEach(({ direction, columns }, configuration) => {
    const random = stream((start + Math.imul(configuration, 0x9e3779b1)) | 0);
    for (let sample = 0; sample < SAMPLES; sample += 1) {
      sequence(random, direction, columns, `stream ${start} ${direction} ${columns} column(s) sample ${sample}`);
    }
  });
};

beforeAll(loadParser);

describe('generated sequences of upright block tables and square-wrapped drawings', () => {
  it('one-strip paragraphs with rectangles at page or paragraph-relative positions over the band', () => {
    forEachSequence(0x2b12, (random, direction, columns, context) => {
      const pick = <Item>(items: readonly Item[]) => items[Math.floor(random() * items.length)]!;
      const tables = [T, T, table(1), table(6), table(30), table(70), table(1, 400), table(1, 800)];
      const drawing = () => {
        const width = pick([20, 60, 80, 340]);
        const height = pick([1, 10, 20, 40]);
        const paragraphRelative = random() < 0.25;
        return anchoredRectangle(
          40 + Math.floor(random() * (PAGE_W - 80 - width)),
          paragraphRelative ? Math.floor(random() * 40) : 40 + Math.floor(random() * 700),
          width, height, paragraphRelative ? 'paragraph' : 'page',
        );
      };
      const body: string[] = [];
      const length = 3 + Math.floor(random() * 10);
      for (let at = 0; at < length; at += 1) {
        const kind = random();
        if (kind < 0.4) {
          for (let run = random() < 0.25 ? 2 : 1; run > 0; run -= 1) body.push(pick(tables));
        } else if (kind < 0.45 && columns > 1) {
          body.push(columnBreak);
        } else {
          for (let run = 1 + Math.floor(random() * 4); run > 0; run -= 1) {
            body.push(random() < 0.5 ? P.replace('</w:p>', `${drawing()}</w:p>`) : P);
          }
        }
      }
      check(direction, columns, body, context);
    });
  }, TIMEOUT_MS);

  // Paragraphs of one to several strips whose rectangles are positioned
  // from the page, the paragraph, the line or the character along either
  // axis, so that a rectangle follows text that its own exclusion, a table
  // or another rectangle moves; anchored before tables, after them, and by
  // the paragraph directly before a table. With `avoidsDrawings` no
  // rectangle may overlap (ECMA-376 §20.4.2.3): a rectangle then shares no
  // area with a rectangle of an earlier paragraph of its page. (Rectangles
  // of one paragraph are left out: the lower-layer composition of one
  // paragraph keeps authored positions.)
  const followingText = (start: number, avoidsDrawings: boolean) => () => {
    forEachSequence(start, (random, direction, columns, context) => {
      const pick = <Item>(items: readonly Item[]) => items[Math.floor(random() * items.length)]!;
      const tables = [T, T, table(1), table(6), table(30), table(45), table(70), table(1, 400), table(1, 800)];
      const drawing = () => {
        const width = pick([20, 38, 60, 80, 200, 340]);
        const height = pick([1, 10, 20, 50, 150]);
        const horizontalFrom = random() < 0.5 ? 'character' : 'page';
        const verticalFrom = pick(['page', 'paragraph', 'paragraph', 'line'] as const);
        return anchoredRectangle(
          horizontalFrom === 'character'
            ? -60 + Math.floor(random() * 120) : 40 + Math.floor(random() * (PAGE_W - 80 - width)),
          verticalFrom === 'page' ? 40 + Math.floor(random() * 700) : Math.floor(random() * 300),
          width, height, verticalFrom, horizontalFrom, !avoidsDrawings,
        );
      };
      const text = () => {
        const characters = random() < 0.5 ? 5 : pick([20, 35, 66, 90, 140]);
        if (random() < 0.5) return anchoring(characters);
        return anchoring(characters, drawing(), ...(random() < 0.2 ? [drawing()] : []));
      };
      const body: string[] = [];
      const length = 2 + Math.floor(random() * 8);
      for (let at = 0; at < length; at += 1) {
        const kind = random();
        if (kind < 0.3) {
          for (let run = random() < 0.25 ? 2 : 1; run > 0; run -= 1) body.push(pick(tables));
        } else if (kind < 0.45) {
          body.push(anchoring(pick([5, 35, 66, 90]), drawing()), pick(tables));
        } else if (kind < 0.5 && columns > 1) {
          body.push(columnBreak);
        } else {
          for (let run = 1 + Math.floor(random() * 3); run > 0; run -= 1) body.push(text());
        }
      }
      const pages = check(direction, columns, body, context, avoidsDrawings);
      if (avoidsDrawings) expect(drawingCollisions(pages, direction), context).toEqual([]);
    });
  };
  it.each([0x51c3, 0x0a17, 0x7e2d])(
    'rectangles that follow their text, before and after tables (stream %i)',
    (start) => followingText(start, false)(), TIMEOUT_MS,
  );
  it('rectangles that follow their text and may not overlap', followingText(0x3d91, true), TIMEOUT_MS);

  it('tables reaching the page end and rectangles that may not overlap near it', () => {
    // Rectangles positioned on the page, most of them near its bottom edge,
    // so that one moved for a table or for another rectangle often has no
    // room below. Every page-positioned rectangle is anchored within its
    // page, so none may reach past a page edge: no move takes a drawing off
    // its page. (The 70 pt wide ones are positioned down their paragraph,
    // which can start below a table, and can be anchored past the page.) A
    // rectangle lies over a table, or over a rectangle of an earlier
    // paragraph, only where the page has no place for it.
    forEachSequence(0x6a05, (random, direction, columns, context) => {
      const pick = <Item>(items: readonly Item[]) => items[Math.floor(random() * items.length)]!;
      const tables = [T, table(6), table(45), table(1, 650), table(1, 690), table(1, 800)];
      const drawing = () => {
        const width = pick([20, 80, 200, 315, 400]);
        const height = pick([10, 60, 85, 110, 150]);
        const lowest = Math.floor(PAGE_H - height);
        return anchoredRectangle(
          Math.floor(random() * (PAGE_W - width)),
          random() < 0.6 ? lowest - Math.floor(random() * 60) : Math.floor(random() * lowest),
          width, height, 'page', 'page', false,
        );
      };
      const body: string[] = [];
      const length = 2 + Math.floor(random() * 6);
      for (let at = 0; at < length; at += 1) {
        const kind = random();
        if (kind < 0.35) body.push(pick(tables));
        else if (kind < 0.4 && columns > 1) body.push(columnBreak);
        else {
          // Half of the rectangles follow a paragraph whose own rectangle is
          // positioned from its text, which the page's start does not know.
          if (random() < 0.5) body.push(anchoring(5, anchoredRectangle(
            Math.floor(random() * (PAGE_W - 70)), Math.floor(random() * 600), 70, 40, 'paragraph', 'page', false,
          )));
          body.push(random() < 0.75 ? anchoring(5, drawing()) : P);
        }
      }
      const pages = check(direction, columns, body, context, true);
      expect(drawingCollisions(pages, direction), context).toEqual([]);
      expect(offPage(pages).filter(({ drawing }) => drawing.width !== 70), context).toEqual([]);
    });
  }, TIMEOUT_MS);
});
