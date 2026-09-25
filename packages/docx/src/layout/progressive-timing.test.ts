import { writeFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import {
  installStubCanvas,
  syntheticDocxModel,
  type SyntheticDocumentShape,
} from '../testing/synthetic-document.js';
import { paginateBody } from './body-paginator.js';
import { layoutOptionsForRender } from './options.js';
import { layoutDocumentProgressively } from './progressive.js';
import { setDocumentLayoutValidation } from './validation-policy.js';

/**
 * Progressive-layout latency harness. Run with `pnpm bench:progressive`; it is
 * skipped in an ordinary `pnpm test` because a full matrix takes minutes.
 *
 * ## What it measures
 *
 * Per document shape: how long until a paintable prefix exists, how long until
 * the authoritative layout exists, and what one straight-through blocking pass
 * costs. From those, the two ratios that actually matter — the time-to-first-
 * page win, and the price checkpoint composition pays for it.
 *
 * The last two columns split what would otherwise be one indistinguishable
 * overhead number. `checkpoint overhead` measures composing immutable,
 * paintable snapshots from the SAME resumable pagination session with yielding
 * disabled. `slicing` is what spreading that session over event-loop turns adds
 * on top; it buys responsiveness, and in worker mode it also keeps the worker
 * able to answer render requests mid-pagination.
 *
 * ## What it does NOT measure — read before quoting any number
 *
 * These are LAYOUT-COST numbers, taken in-process against a stub canvas. They
 * exclude, entirely:
 *
 *   - worker spin-up, the render-worker module fetch/eval and the init handshake
 *   - WASM instantiation, ZIP/XML parse and model materialization
 *   - font preload and resource metrics (Google, embedded, native) — all of which precede
 *     pagination and all of which are on the real first-paint critical path
 *   - math conversion and rasterization
 *   - the wire: structured clone per publication, and ImageBitmap transfer
 *   - painting: `renderLayoutSourceToCanvas` is never called
 *   - main-thread frame contention, which is the entire point of worker mode
 *     and the one thing an in-process harness structurally cannot show
 *
 * ## main mode vs worker mode
 *
 * This harness deliberately reports ONE set of layout numbers, because there is
 * only one engine: `mode: 'main'` and `mode: 'worker'` run the identical
 * `layoutDocumentProgressively` over identical services, so their layout cost is
 * the same by construction and measuring both would just be measuring twice.
 *
 * The mode decides whose thread pays that cost, which is what `pre-paint
 * block`, `longest block` and `yields` are for. With the opening preview
 * drained through the same scheduler as the rest, `pre-paint block` should sit
 * near the slice budget: main mode stays interactive throughout, paying the
 * totals above in slice-sized installments of UI-thread time. `longest block`
 * is bounded below by the slowest single body entry — the scheduler can only
 * release the thread BETWEEN entries. Worker mode moves all of it off the UI
 * thread entirely; what remains unpriceable in-process is its spin-up, parse
 * and per-page bitmap transfer. In main mode the totals above are time the UI thread is
 * occupied — spread over `yields` slices, none longer than `longest block`
 * (and, for a non-progressive load, one single block as long as the whole
 * `blocking` column). In worker mode the UI thread pays none of it. That
 * difference is architectural rather than empirical; what an in-process harness
 * genuinely cannot price is worker spin-up, a second WASM instantiation, font
 * preload in the worker, and per-page `ImageBitmap` transfer. For those, the
 * browser is the only honest instrument.
 *
 * The stub canvas also substitutes arithmetic glyph widths for real shaping.
 * Since measurement dominates the block loop, the absolute milliseconds are
 * optimistic and the progressive:blocking RATIO is the portable finding. For an
 * end-to-end number including everything above, read the Storybook
 * "progressive first paint" story's `First paint … → full layout …` line.
 */

const ENABLED = !!process.env.OOXML_DOCX_PROGRESSIVE_BENCH;

const CASES: readonly (readonly [SyntheticDocumentShape, number])[] = [
  ['plain', 400],
  ['header-footer', 200],
  ['fields', 200],
  ['tables', 60],
  ['tracked-fields', 200],
];

interface Row {
  shape: string;
  paragraphs: number;
  firstPreviewMs: number | null;
  previewPages: number;
  progressiveMs: number;
  unslicedMs: number;
  blockingMs: number;
  longestBlockMs: number;
  prePaintBlockMs: number;
  yields: number;
  pages: number;
  measureCalls: number;
}

function fixture(shape: SyntheticDocumentShape, paragraphs: number) {
  // Rebuilt per timing so neither run inherits the other's warmed caches; the
  // model itself is LCG-derived and therefore identical every time.
  const source = layoutSourceStore(syntheticDocxModel(shape, { paragraphs }));
  return { source, services: createLayoutServices(source) };
}

function fixed(value: number, places = 1): string {
  return value.toFixed(places);
}

describe.skipIf(!ENABLED)('progressive layout latency', () => {
  let measureTextCalls: () => number;

  beforeAll(() => {
    ({ measureTextCalls } = installStubCanvas());
    // `validation-policy.ts` turns path-precise retained-layout validation ON
    // whenever VITEST is set. Leaving it on would measure what CI pays, not
    // what a shipped viewer pays.
    setDocumentLayoutValidation(false);
  });

  it('reports time-to-first-page against a blocking layout', async () => {
    const options = layoutOptionsForRender({ defaultCurrentDateMs: 1_700_000_000_000 });
    const rows: Row[] = [];

    for (const [shape, paragraphs] of CASES) {
      // 1. Progressive: timestamp the first publication, then run to completion.
      const progressive = fixture(shape, paragraphs);
      const callsBefore = measureTextCalls();
      // Instrumented yield: the gap between releases IS the uninterrupted block
      // a main-mode UI would be frozen for. This is the only column that
      // distinguishes the two render modes at all — see the header.
      let longestBlockMs = 0;
      // The block up to the FIRST release — the longest a main-mode UI can be
      // frozen before layout first lets go of the thread. Since the opening
      // preview drains through the same scheduler as everything else, this
      // should sit near the slice budget; a large value here means slicing
      // regressed somewhere on the way to first paint.
      let prePaintBlockMs = 0;
      let yields = 0;
      let lastYield = performance.now();
      const yieldToHost = async (): Promise<void> => {
        const block = performance.now() - lastYield;
        longestBlockMs = Math.max(longestBlockMs, block);
        if (yields === 0) prePaintBlockMs = block;
        yields += 1;
        await new Promise<void>((resolve) => {
          const channel = new MessageChannel();
          channel.port1.onmessage = () => { channel.port1.close(); resolve(); };
          channel.port2.postMessage(null);
        });
        lastYield = performance.now();
      };
      const started = performance.now();
      lastYield = started;
      let firstPreviewMs: number | null = null;
      let previewPages = 0;
      const full = await layoutDocumentProgressively(
        progressive.source.bodyLayoutInput,
        progressive.services,
        options,
        {
          scheduler: { yieldToHost },
          onPreview: (preview) => {
            firstPreviewMs ??= performance.now() - started;
            if (previewPages === 0) previewPages = preview.layout.pages.length;
          },
        },
      );
      const progressiveMs = performance.now() - started;
      const measureCalls = measureTextCalls() - callsBefore;

      // 2. Progressive again, but never yielding. Separates checkpoint
      //    composition from the cost of spreading one resumable session over
      //    event-loop turns.
      const unslicedFixture = fixture(shape, paragraphs);
      const unslicedStart = performance.now();
      await layoutDocumentProgressively(
        unslicedFixture.source.bodyLayoutInput,
        unslicedFixture.services,
        options,
        {
          scheduler: { sliceMs: Number.POSITIVE_INFINITY },
          onPreview: () => {},
        },
      );
      const unslicedMs = performance.now() - unslicedStart;

      // 3. Baseline: one straight-through pass, on a cold fixture.
      const blockingFixture = fixture(shape, paragraphs);
      const blockingStart = performance.now();
      const blocking = paginateBody(
        blockingFixture.source.bodyLayoutInput,
        blockingFixture.services,
        options,
      );
      const blockingMs = performance.now() - blockingStart;

      // The guarantee the whole feature rests on: progressive changes WHEN
      // pages appear, never WHICH pages appear.
      expect(full.pages.length).toBe(blocking.pages.length);

      rows.push({
        shape,
        paragraphs,
        firstPreviewMs,
        previewPages,
        progressiveMs,
        unslicedMs,
        blockingMs,
        pages: full.pages.length,
        longestBlockMs,
        prePaintBlockMs,
        yields,
        measureCalls,
      });
    }

    const lines = [
      '',
      '| shape | body | pages | first preview | progressive total | unsliced | blocking | speedup to 1st page | checkpoint overhead | slicing | pre-paint block | longest block | yields |',
      '| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |',
      ...rows.map((row) => {
        const first = row.firstPreviewMs === null ? 'none' : `${fixed(row.firstPreviewMs)}ms`;
        const speedup = row.firstPreviewMs === null
          ? '—'
          : `${fixed(row.blockingMs / row.firstPreviewMs, 2)}×`;
        return `| ${row.shape} | ${row.paragraphs} | ${row.pages} | ${first} | `
          + `${fixed(row.progressiveMs)}ms | ${fixed(row.unslicedMs)}ms | `
          + `${fixed(row.blockingMs)}ms | ${speedup} | `
          + `${fixed(row.unslicedMs / row.blockingMs, 2)}× | `
          + `${fixed(row.progressiveMs / row.unslicedMs, 2)}× | `
          + `${fixed(row.prePaintBlockMs)}ms | ${fixed(row.longestBlockMs)}ms | `
          + `${row.yields} |`;
      }),
      '',
      'Layout cost only — excludes worker spin-up, WASM parse, font preload,',
      'wire transfer and paint. See this file’s header before quoting these.',
      '',
      'main vs worker: every column except the last three is IDENTICAL in both',
      'render modes — it is one engine, and these run it in-process. What the',
      'mode changes is WHOSE thread pays. In `main` the numbers above are time',
      'the UI thread is occupied, in slices of at most `longest block`; a',
      'blocking (non-progressive) main-mode load makes that one block as long',
      'as the whole `blocking` column. In `worker` the UI thread pays none of',
      'it and stays free for the whole duration, at the cost of spin-up, parse',
      'and per-page bitmap transfer that this harness does not measure.',
      '',
      '`pre-paint block` is the longest the UI can freeze before layout first',
      'releases the thread; with the preview sliced it should sit near the',
      '16ms budget. `longest block` is bounded by the slowest single body',
      'entry — the scheduler only yields between entries.',
      '',
    ];
    const report = lines.join('\n');
    // Written as well as printed: vitest intercepts console output, and the
    // table is the point of the run.
    process.stdout.write(`${report}\n`);
    const out = process.env.OOXML_DOCX_PROGRESSIVE_BENCH_OUTPUT;
    if (out) writeFileSync(out, `${report}\n`);

    expect(rows).toHaveLength(CASES.length);
  }, 1_800_000);
});
