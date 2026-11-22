import { describe, expect, it } from 'vitest';
import { clusterBoxes } from './visual-cluster-geometry.js';

describe('cluster boxes', () => {
  it('stand the characters upright side by side across the run, each over the cell', () => {
    // '２９' (two 6pt natural advances) in an 8pt cell at x=78, baseline 80:
    // the row is centred on the baseline across the run, first character on
    // the run's far side, as Canvas paints it turned back upright.
    const placement = {
      text: '２９', range: { start: 1, end: 3 }, direction: 'ltr' as const, tateChuYoko: true,
      origin: { xPt: 78, yPt: 80 }, bounds: { xPt: 78, yPt: 72, widthPt: 8, heightPt: 10 },
      clusters: [
        { range: { start: 1, end: 2 }, offset: { xPt: 0, yPt: 0 }, advancePt: 6 },
        { range: { start: 2, end: 3 }, offset: { xPt: 6, yPt: 0 }, advancePt: 6 },
      ],
      paintOps: [{ range: { start: 1, end: 3 }, offset: { xPt: 4, yPt: 0 } }],
    };
    const boxes = clusterBoxes(placement as never, 0, 2, 1, 3, 78);
    expect(boxes.map((box) => [box.xPt, box.widthPt, box.crossPt, box.crossWidthPt])).toEqual([[0, 8, 8, 6], [0, 8, 2, 6]]);
  });

  it('give a floored, condensed glyph the box it is drawn in', () => {
    // 'abc' condensed so that 'b' floors to no advance; its own operation
    // draws it at its 6pt natural advance.
    const placement = {
      text: 'abc', range: { start: 0, end: 3 }, direction: 'ltr' as const,
      origin: { xPt: 0, yPt: 10 }, bounds: { xPt: 0, yPt: 0, widthPt: 8, heightPt: 12 },
      clusters: [
        { range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 }, advancePt: 4 },
        { range: { start: 1, end: 2 }, offset: { xPt: 4, yPt: 0 }, advancePt: 0 },
        { range: { start: 2, end: 3 }, offset: { xPt: 4, yPt: 0 }, advancePt: 4 },
      ],
      paintOps: [
        { range: { start: 0, end: 1 }, offset: { xPt: 0, yPt: 0 } },
        { range: { start: 1, end: 2 }, offset: { xPt: 4, yPt: 0 }, naturalAdvancePt: 6 },
        { range: { start: 2, end: 3 }, offset: { xPt: 4, yPt: 0 } },
      ],
    };
    const boxes = clusterBoxes(placement as never, 0, 3, 0, 3, 0);
    expect(boxes.map((box) => [box.xPt, box.widthPt])).toEqual([[0, 4], [4, 6], [4, 4]]);
  });

  it('end a compressed right-to-left fitted run at its box\'s right edge', () => {
    // אבגד fitted into 12pt with 6pt advances: pitch -4, clusters of 2pt.
    const placement = {
      text: 'אבגד', range: { start: 0, end: 4 }, direction: 'rtl' as const,
      origin: { xPt: 168, yPt: 10 }, bounds: { xPt: 168, yPt: 0, widthPt: 12, heightPt: 12 },
      clusters: [0, 1, 2, 3].map((index) => ({ range: { start: index, end: index + 1 }, offset: { xPt: 2 * index, yPt: 0 }, advancePt: 2 })),
    };
    const boxes = clusterBoxes(placement as never, 0, 4, 0, 4, 168);
    // א (logical first) is painted rightmost, ending at the box edge.
    expect(boxes[0]!.xPt + boxes[0]!.widthPt).toBeCloseTo(12, 9);
    expect(Math.min(...boxes.map((box) => box.xPt))).toBeCloseTo(0, 9);
  });

  it('extend the final cluster of a compressed fitted run to the run\'s end', () => {
    // fitText 12pt over four 6pt characters: pitch -4, so each cluster
    // advances 2pt while the last glyph paints to the box edge.
    const placement = {
      text: 'abcd', range: { start: 0, end: 4 }, direction: 'ltr' as const,
      origin: { xPt: 0, yPt: 10 }, bounds: { xPt: 0, yPt: 0, widthPt: 12, heightPt: 12 },
      clusters: [0, 1, 2, 3].map((index) => ({ range: { start: index, end: index + 1 }, offset: { xPt: 2 * index, yPt: 0 }, advancePt: 2 })),
    };
    const boxes = clusterBoxes(placement as never, 0, 4, 0, 4, 0);
    expect(boxes.at(-1)).toMatchObject({ xPt: 6, widthPt: 6, trackingExcluded: true });
  });

  it('read a placement\'s paint operations once for all of its source owners', () => {
    // n one-letter owners of one condensed placement, one of them floored:
    // each owner's boxes look up the floored glyphs; reads stay linear in n.
    const reads = (count: number): number => {
      let read = 0;
      const operations = Array.from({ length: count }, (_, index) => ({
        range: { start: index, end: index + 1 }, offset: { xPt: 4 * index, yPt: 0 },
        ...(index === 1 ? { naturalAdvancePt: 6 } : {}),
      }));
      const paintOps = new Proxy(operations, {
        get(target, key, receiver) {
          if (typeof key === 'string' && /^\d+$/u.test(key)) read += 1;
          return Reflect.get(target, key, receiver);
        },
      });
      const placement = {
        text: 'a'.repeat(count), range: { start: 0, end: count }, direction: 'ltr' as const,
        origin: { xPt: 0, yPt: 10 }, bounds: { xPt: 0, yPt: 0, widthPt: 4 * count, heightPt: 12 },
        clusters: Array.from({ length: count }, (_, index) => ({
          range: { start: index, end: index + 1 }, offset: { xPt: 4 * index, yPt: 0 }, advancePt: 4 })),
        paintOps,
      };
      for (let index = 0; index < count; index += 1) {
        expect(clusterBoxes(placement as never, index, index + 1, index, index + 1, 0)).toHaveLength(1);
      }
      return read;
    };
    expect(reads(32)).toBeLessThanOrEqual(32);
    expect(reads(128)).toBeLessThanOrEqual(128);
  });
});
