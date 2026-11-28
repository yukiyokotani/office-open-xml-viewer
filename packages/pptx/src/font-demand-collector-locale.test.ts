import { afterEach, expect, it, vi } from 'vitest';

afterEach(() => {
  vi.resetModules();
  vi.doUnmock('@silurus/ooxml-core/internal/worker-sidecar');
});

// A parse worker adopts the main thread's locale in ITS number-symbols
// module, but the sidecar is a separate module graph (own state, navigator
// en-US here). The locale must therefore travel as an argument.
it('demands the adopted locale\'s group separator from an isolated sidecar graph', async () => {
  vi.resetModules();
  // The sidecar's own, fresh graph: it never saw adoptHostLocale('fr-FR').
  const sidecar = await import('./internal/font-demand-sidecar.js');
  vi.resetModules();
  vi.doMock('@silurus/ooxml-core/internal/worker-sidecar', () => ({ importWorkerSidecar: async () => sidecar }));
  const { importFontDemandCollector } = await import('./font-demand-collector.js');
  const core = await import('@silurus/ooxml-core');
  core.adoptHostLocale('fr-FR');
  const slide = { index: 0, slideNumber: 1, background: null, elements: [{ type: 'chart', chart: { chartType: 'bar', series: [] } }] } as never;

  const collector = await importFontDemandCollector('https://example.test/font-demand.mjs', core.hostLocale());
  const demand = collector(slide, new Set(), 100000);
  expect(demand).not.toBe('all');
  expect(demand).toContain(0x202f);

  // Without the explicit locale the sidecar falls back to navigator (en-US).
  const noLocale = (await importFontDemandCollector('https://example.test/font-demand.mjs'))(slide, new Set(), 100000);
  expect(noLocale).not.toContain(0x202f);
});
