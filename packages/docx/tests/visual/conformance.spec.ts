import { expect, test } from '@playwright/test';
import { CONFORMANCE_CASES } from '../../src/conformance/cases.js';
import { readFileSync } from 'node:fs';
const mainSpacingStreams = JSON.parse(readFileSync(
  new URL('../../src/spacing-arithmetic-main.json', import.meta.url), 'utf8',
));

interface BrowserConformanceReport {
  readonly caseId: string;
  readonly expected: (typeof CONFORMANCE_CASES)[number]['expected'];
  readonly pageCount: number;
  readonly mainSizes: readonly { widthPt: number; heightPt: number }[];
  readonly workerSizes: readonly { widthPt: number; heightPt: number }[];
  readonly runsFinite: boolean;
  readonly targetTextPresent: boolean;
  readonly cloneSafe: boolean;
  readonly layoutFingerprint: string;
  readonly mainWorkerParity: {
    readonly pageCount: boolean;
    readonly sizes: boolean;
    readonly runs: boolean;
  };
  readonly authoredGeometry: {
    readonly pageBoxes: readonly {
      readonly widthPt: number;
      readonly heightPt: number;
      readonly contentTopPt: number;
      readonly contentBottomPt: number;
    }[];
    readonly tables: readonly {
      readonly id: string;
      readonly columnWidthsPt: readonly number[];
      readonly borderEndpoints: readonly (readonly number[])[];
    }[];
    readonly drawings: readonly {
      readonly id: string;
      readonly widthPt: number;
      readonly heightPt: number;
    }[];
  };
}

test('fractional unchanged-text paint and advances equal main bit for bit', async ({ page }) => {
  await page.goto('/tests/visual/conformance-fixture.html');
  const streams = await page.evaluate(async () => {
    const { spacingFixtures, recordSpacingFixture } = await import('/src/test-support/spacing-arithmetic.test-support.ts');
    return Object.fromEntries(await Promise.all(spacingFixtures.map(async (fixture: { id: string }) =>
      [fixture.id, await recordSpacingFixture(fixture)])));
  });
  // Source-owned Arabic spaces intentionally change first-line paint packets;
  // retain exact arithmetic for every unaffected fixture and the Latin tail.
  for (const [id, stream] of Object.entries(streams)) {
    if (id !== 'kashida') expect(stream).toEqual(mainSpacingStreams[id]);
  }
  const actual = streams.kashida;
  const previous = mainSpacingStreams.kashida;
  expect(actual.advances.slice(1)).toEqual(previous.advances.slice(1));
  expect(actual.fills.filter((fill: readonly unknown[]) => String(fill[0]).startsWith('W')))
    .toEqual(previous.fills.filter((fill: readonly unknown[]) => String(fill[0]).startsWith('W')));
  expect(actual.advances[0].map((piece: readonly unknown[]) => piece[0]).join('')).toBe('بببب بببب ');
  const arabicPaint = actual.fills.map((fill: readonly unknown[]) => String(fill[0]))
    .filter((text: string) => /[\u0600-\u06ff]/u.test(text)).join('');
  expect(arabicPaint).toContain('ـ');
  expect(arabicPaint.replace(/[ـ ]/gu, '')).toBe('بببببببب');
});

for (let index = 0; index < CONFORMANCE_CASES.length; index += 1) {
  const testCase = CONFORMANCE_CASES[index]!;
  test(testCase.id, async ({ page }) => {
    test.setTimeout(60_000);
    await page.goto('/tests/visual/conformance-fixture.html');
    await page.waitForFunction(() => document.body.dataset.status === 'ready');
    await expect.poll(() => page.evaluate(() =>
      Number(document.body.dataset.caseCount))).toBe(CONFORMANCE_CASES.length);

    const report = await page.evaluate(async (caseIndex) =>
      (window as unknown as {
        runConformanceCase: (index: number) => Promise<BrowserConformanceReport>;
      }).runConformanceCase(caseIndex), index);

    expect(report.caseId).toBe(testCase.id);
    expect(report.expected).toEqual(testCase.expected);
    expect(report.pageCount).toBe(testCase.expected.pageCount);
    expect(report.mainSizes).toEqual([{
      widthPt: testCase.expected.pageWidthPt,
      heightPt: testCase.expected.pageHeightPt,
    }]);
    expect(report.workerSizes).toEqual(report.mainSizes);
    expect(report.runsFinite).toBe(true);
    expect(report.targetTextPresent).toBe(true);
    expect(report.cloneSafe).toBe(true);
    expect(report.layoutFingerprint.length).toBeGreaterThan(0);
    expect(report.mainWorkerParity).toEqual({
      pageCount: true,
      sizes: true,
      runs: true,
    });

    // These values come entirely from authored twips/EMU and page margins.
    // Native Canvas shaping is deliberately absent from this exact
    // cross-browser comparison.
    expect(report.authoredGeometry.pageBoxes).toMatchObject([{
      widthPt: 612,
      heightPt: 792,
    }]);
    const pageBox = report.authoredGeometry.pageBoxes[0]!;
    expect(pageBox.contentTopPt).toBeGreaterThanOrEqual(0);
    expect(pageBox.contentBottomPt).toBeLessThanOrEqual(pageBox.heightPt);
    expect(pageBox.contentBottomPt).toBeGreaterThan(pageBox.contentTopPt);
    expect(report.authoredGeometry.tables).toHaveLength(testCase.expected.tableDepth);
    const tableWidths = report.authoredGeometry.tables
      .flatMap(({ columnWidthsPt }) => columnWidthsPt)
      .sort((left, right) => right - left);
    expect(tableWidths).toEqual(
      testCase.expected.tableDepth === 2
        ? [360, 180]
        : testCase.expected.tableDepth === 1
          ? [360]
          : [],
    );
    for (const table of report.authoredGeometry.tables) {
      expect(table.borderEndpoints.flat().every(Number.isFinite)).toBe(true);
      const horizontalSpans = table.borderEndpoints
        .filter(([, fromY, , toY]) => fromY === toY)
        .map(([fromX, , toX]) => Math.abs(toX! - fromX!));
      expect(Math.max(...horizontalSpans)).toBe(
        table.columnWidthsPt.reduce((sum, width) => sum + width, 0),
      );
    }
    expect(report.authoredGeometry.drawings).toHaveLength(
      testCase.expected.drawingCount,
    );
    for (const drawing of report.authoredGeometry.drawings) {
      expect(drawing).toMatchObject({ widthPt: 36, heightPt: 21.6 });
    }
  });
}


test('unselected literal soft hyphens preserve contextual cs paint in every browser', async ({ page }) => {
  await page.goto('/tests/visual/conformance-fixture.html');
  await page.waitForFunction(() => document.body.dataset.status === 'ready');
  const results = await page.evaluate(async () => {
    const { renderDocumentToCanvas } = await import('/src/renderer.ts');
    const render = async (text: string) => {
      const model = {
        section: { pageWidth: 300, pageHeight: 100, marginTop: 10, marginRight: 10,
          marginBottom: 10, marginLeft: 10, headerDistance: 0, footerDistance: 0 },
        body: [{ type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0,
          indentFirst: 0, spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
          numbering: null, tabStops: [], runs: [{ type: 'text', text, cs: true,
            fontFamily: 'Arial', fontFamilyCs: 'Arial', fontSize: 20, charSpacing: 2,
            bold: false, italic: false, underline: false, strikethrough: false,
            color: null, isLink: false, background: null, vertAlign: null, hyperlink: null }],
        }], headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
      };
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
      const fills: string[] = [];
      const fillText = ctx.fillText.bind(ctx);
      ctx.fillText = (value, x, y) => { fills.push(value); fillText(value, x, y); };
      await renderDocumentToCanvas(model, canvas, 0, { width: 300, dpr: 1 });
      return { bitmap: canvas.toDataURL(), fills };
    };
    return { optional: await render('a\u00ADb'), plain: await render('ab') };
  });
  expect(results.optional.fills).toEqual(['ab']);
  expect(results.optional).toEqual(results.plain);
});

// The recording-canvas test pins the exact main font/coordinates. Here a native
// whole-run Canvas draw is the bitmap expectation, so real kerning and Arabic
// joining are exercised without OS-specific stored screenshots. This contract
// also passes on origin/main 4bcb91c487e38393aadb3f3b70a732adf2c7b708.
test('kerning and Arabic joining preserve whole-run paint with character spacing', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('/tests/visual/conformance-fixture.html');
  await page.waitForFunction(() => document.body.dataset.status === 'ready');
  const results = await page.evaluate(async () => {
    const { renderDocumentToCanvas } = await import('/src/renderer.ts');
    const cases = [
      ...['AV', 'To'].flatMap(text => [undefined, 21, 20].flatMap(kerning =>
        [0, 2].map(charSpacing => ({ text, kerning, charSpacing, cs: false })))),
      ...[0, 2].map(charSpacing => ({ text: 'سلام', kerning: undefined, charSpacing, cs: true })),
    ];
    const results = [];
    for (const { text, kerning, charSpacing, cs } of cases) {
      const run = { type: 'text', text, kerning, charSpacing, cs,
        fontFamily: 'Arial', fontFamilyCs: 'Arial', fontSize: 20,
        bold: false, italic: false, underline: false, strikethrough: false,
        color: null, isLink: false, background: null, vertAlign: null, hyperlink: null };
      const model = {
        section: { pageWidth: 300, pageHeight: 100, marginTop: 0, marginRight: 0,
          marginBottom: 0, marginLeft: 0, headerDistance: 0, footerDistance: 0 },
        body: [{ type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0,
          indentFirst: 0, spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
          numbering: null, tabStops: [], runs: [run, { ...run, text: '|',
            cs: false, charSpacing: 0, kerning: undefined, bold: true }],
        }], headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
      };
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
      const fills: { text: string; x: number; y: number; font: string;
        letterSpacing: string; fontKerning: CanvasFontKerning }[] = [];
      const fillText = ctx.fillText.bind(ctx);
      ctx.fillText = (value, x, y) => {
        fills.push({ text: value, x, y, font: ctx.font,
          letterSpacing: ctx.letterSpacing, fontKerning: ctx.fontKerning });
        fillText(value, x, y);
      };
      await renderDocumentToCanvas(model, canvas, 0, { width: 300, dpr: 1 });
      if (fills.length !== 2) throw new Error(`Expected two contextual draws for ${text}`);

      const reference = document.createElement('canvas');
      reference.width = canvas.width;
      reference.height = canvas.height;
      const expected = reference.getContext('2d') as CanvasRenderingContext2D;
      expected.fillStyle = '#ffffff';
      expected.fillRect(0, 0, reference.width, reference.height);
      expected.fillStyle = '#000000';
      expected.textAlign = 'left';
      expected.textBaseline = 'alphabetic';
      expected.direction = 'ltr';
      expected.font = fills[0].font;
      // WORD_ARABIC_PAIR_KERNING: Arabic text is pair-kerned whatever w:kern says.
      expected.fontKerning = kerning === 20 || cs ? 'normal' : 'none';
      // Measure the whole string with zero tracking, then add authored pitch.
      // Measuring individual glyphs would lose the pair/joining advance.
      const advance = expected.measureText(text).width + [...text].length * charSpacing;
      expected.letterSpacing = `${charSpacing}px`;
      expected.fillText(text, 0, fills[0].y);
      expected.font = fills[1].font;
      expected.fontKerning = 'none';
      expected.letterSpacing = '0px';
      expected.fillText('|', advance, fills[0].y);
      results.push({ text, kerning, charSpacing, cs, fills, advance,
        bitmapMatches: canvas.toDataURL() === reference.toDataURL() });
    }
    return results;
  });
  expect(results).toHaveLength(14);
  for (const result of results) {
    expect(result.fills.map(({ text, x, y, letterSpacing, fontKerning }) =>
      // Chrome can serialize zero tracking as "" after context save/restore.
      // Unit tests pin the exact assigned strings; compare native zero by value.
      ({ text, x, y, letterSpacing: letterSpacing || '0px', fontKerning }))).toEqual([
      { text: result.text, x: 0, y: result.fills[0].y,
        letterSpacing: `${result.charSpacing}px`,
        fontKerning: result.kerning === 20 || result.cs ? 'normal' : 'none' },
      { text: '|', x: result.advance, y: result.fills[0].y,
        letterSpacing: '0px', fontKerning: 'none' },
    ]);
    expect(result.bitmapMatches, JSON.stringify({ text: result.text,
      kerning: result.kerning, tracking: result.charSpacing })).toBe(true);
  }
});
