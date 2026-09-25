import { test, expect } from '@playwright/test';
import { existsSync } from 'node:fs';

// ECMA-376 §17.6.12 on REAL private samples, against Word PDF ground truth (measured
// with pdftotext). Two shapes are covered:
//
// (1) NON-REGRESSION — sample-12/13 both carry `<w:pgNumType>` (sample-12: continuous
//     start=1; sample-13: nextPage start=1 + a continuous start=2) plus a footer
//     PAGE field, and Word prints SEQUENTIAL footers. No restart is VISIBLE:
//     start=1 on the first section is the identity, and sample-13's continuous
//     start=2 section begins exactly AT a page boundary (probed: its content first
//     appears on physical page 2, the SAME page whose top it owns — see the module
//     header of page-numbering.ts), so its restart fires with anchor offset 0 and
//     shows start=2, which equals the natural continuation (1+1). The DISPLAYED
//     footer number therefore equals the physical page number. (sample-14 also
//     carries pgNumType but its footer shares the bottom band with other numeric
//     content, so the position heuristic can't isolate it; its structure — nextPage
//     start=1 on the first section — is covered deterministically by
//     page-numbering.test.ts.)
//
// (2) RESTART — sample-27 (synthesized for issue #804): section 1 → continuous break
//     + `w:pgNumType w:start="50"` → section 2 sharing p.1 and spilling to p.2/p.3,
//     footer PAGE. Word prints [Page 1, Page 51, Page 52]: the continuous section's
//     series counts the SHARED page (its first appearance) as page 50, so its owned
//     pages show 51 and 52. This is the case the pre-#804 code got wrong ([1, 50, 51]).
//
// Skips gracefully when the (gitignored) sample is absent.
const CASES: { file: string; pageCount: number; width: number; expected: string[] }[] = [
  { file: 'private/sample-12', pageCount: 4, width: 595, expected: ['1', '2', '3', '4'] },
  { file: 'private/sample-13', pageCount: 6, width: 595, expected: ['1', '2', '3', '4', '5'] },
  // §17.6.12 continuous restart (#804): continuous start=50 shares p.1 and spills.
  // The DISTINGUISHING signal is that the section's FIRST OWNED page shows 51 (not
  // 50) — the shared page it does not own counts as the section's page 50. Word's
  // PDF (public/private/docx/sample-27.pdf) is 3 pages [Page 1, Page 51, Page 52]; this
  // renderer packs ~50 body paragraphs per page vs Word's ~46, so it fits section 2
  // in ONE fewer page and shows [1, 51] over 2 pages. That page-DENSITY gap is a
  // separate pagination-fidelity concern, out of scope for #804 — the RESTART
  // semantics (51, not 50) are what this asserts. The full [1, 51, 52] series is
  // pinned deterministically in page-number-field-render.test.ts.
  { file: 'private/sample-27', pageCount: 2, width: 595, expected: ['1', '51'] },
];

test.describe('page-number restart non-regression (§17.6.12)', () => {
  for (const { file, pageCount, width, expected } of CASES) {
    test(`${file}: footer PAGE numbers stay sequential`, async ({ page }) => {
      if (!existsSync(`${process.cwd()}/public/${file}.docx`)) {
        test.skip(true, `gitignored sample ${file}.docx not present`);
      }
      // Load a page in the Vite module graph first so a dynamic `import('/src/...')`
      // inside page.evaluate resolves against the dev server (the fixture imports it).
      await page.goto(`/tests/visual/fixture.html?file=demo/sample-1.docx&page=0&width=595`);
      await page.waitForFunction(
        () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
        { timeout: 30_000 },
      );
      // Render each page and collect the bottom-most SHORT text run (the footer
      // page number). A footer number is a 1–3 char numeric/roman/letter token near
      // the page bottom, so we take the lowest run whose text is <= 4 chars.
      const numbers = await page.evaluate(
        async ({ file, pageCount, width }) => {
          const { DocxDocument } = await import('/src/index.ts');
          const doc = await DocxDocument.load('/' + file + '.docx');
          const out: (string | null)[] = [];
          for (let i = 0; i < pageCount; i++) {
            const canvas = document.createElement('canvas');
            const runs: { text: string; y: number }[] = [];
            await doc.renderPage(canvas, i, {
              width, dpr: 1,
              onTextRun: (r: { text: string; y: number }) => runs.push({ text: r.text, y: r.y }),
            });
            // Footer page number = the lowest-on-page run whose trimmed text is a
            // pure page-number token (Arabic digits, or roman/letter glyphs). This
            // excludes footnote/dagger marks (†) and body text. Prefer the bottom-most.
            const isPageToken = (t: string) => /^[0-9]{1,4}$/.test(t) || /^[ivxlcdmIVXLCDM]{1,4}$/.test(t) || /^[a-zA-Z]{1,3}$/.test(t);
            const tokens = runs
              .map((r) => ({ text: r.text.trim(), y: r.y }))
              .filter((r) => isPageToken(r.text))
              .sort((a, b) => b.y - a.y);
            out.push(tokens.length ? tokens[0].text : null);
          }
          return out;
        },
        { file, pageCount, width },
      );
      // The first `expected.length` pages carry a footer; assert those.
      for (let i = 0; i < expected.length; i++) {
        expect(numbers[i], `${file} page ${i + 1} footer number`).toBe(expected[i]);
      }
    });
  }
});
