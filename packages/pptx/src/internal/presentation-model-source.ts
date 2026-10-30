import { resolveBundledAssetUrl } from '@silurus/ooxml-core/internal/bundled-asset-url';
import {
  resolveCjkFallback,
  type FontPreloadDemand,
  workerRendererDescriptors,
  type AdmittedModelSourceLoad,
  type CjkLang,
} from '@silurus/ooxml-core';
import {
  disposeRejectedLoad,
  normalizeLoadResourceOptions,
  OoxmlResourceMetricsSession,
  type NormalizedOoxmlResourcePolicy,
  type WorkerRendererDescriptors,
} from '@silurus/ooxml-core/worker';
import { selectModelSource, beginModelSourceLoad, effectiveModelSources, modelSourceWorker } from '@silurus/ooxml-core/internal/model-source';
import type { FontSource } from '@silurus/ooxml-core/internal/font-sources';
import type { PptxFontDemandFacts } from '../font-requests.js';
import { PptxPresentation, type LoadOptions } from '../presentation.js';
import { deferred } from '@silurus/ooxml-core/internal/deferred';

type MutablePresentation = {
  _metrics: OoxmlResourceMetricsSession;
  _cjkFallback: CjkLang;
  _math: LoadOptions['math'];
  _threeD: LoadOptions['threeD'];
  _regionMap: LoadOptions['regionMap'];
  _chartEx: LoadOptions['chartEx'];
  _tiff: LoadOptions['tiff'];
  _fonts: readonly FontSource[] | undefined;
  _preflight: {
    fontPreloadNames: readonly (string | null)[];
    fontPreloadFacts?: PptxFontDemandFacts;
    fontPreloadDemand?: FontPreloadDemand;
  } | null;
  _ensureFontSources: (
    names: readonly (string | null)[],
    facts: PptxFontDemandFacts | undefined,
    scalars: FontPreloadDemand | undefined,
  ) => Promise<void>;
  _parse: (
    buffer: ArrayBuffer,
    resourcePolicy: NormalizedOoxmlResourcePolicy,
    fonts: readonly FontSource[] | undefined,
    timeoutMs: number | undefined,
    onUsage: (usage: import('@silurus/ooxml-core').OoxmlResourceUsageSnapshot) => void,
    renderers: WorkerRendererDescriptors | undefined,
    progressive?: unknown,
  ) => Promise<void>;
};

/** All source admission, wire decoration and cleanup live behind opt-in import.
 * `opts` arrives normalised by {@link PptxPresentation.load} (`fonts` only,
 * never `useGoogleFonts`). */
export async function loadPptxModelSource(
  input: string | ArrayBuffer,
  opts: LoadOptions,
): Promise<PptxPresentation> {
  const mode = opts.mode ?? 'main';
  const resourceOptions = normalizeLoadResourceOptions(opts);
  const metrics = new OoxmlResourceMetricsSession({
    enabled: true, format: 'pptx', mode, policy: resourceOptions.policy,
    onMetrics: resourceOptions.onResourceMetrics, emitToConsole: resourceOptions.debug,
  });
  try {
    if (mode === 'worker' && (typeof Worker === 'undefined' || typeof OffscreenCanvas === 'undefined')) {
      throw new Error("mode: 'worker' requires Worker and OffscreenCanvas support");
    }
    let buffer: ArrayBuffer;
    if (typeof input === 'string') {
      const response = await fetch(input);
      if (!response.ok) throw new Error(`Failed to fetch: ${response.status} ${response.statusText}`);
      buffer = await response.arrayBuffer();
    } else {
      buffer = input;
    }
    // Selection happens once: an unclaimed input continues on the OOXML path
    // with every source option cleared, so it is never offered again.
    const selected = selectModelSource(effectiveModelSources(opts.legacyPpt, opts.odp), 'pptx', new Uint8Array(buffer));
    if (!selected) return PptxPresentation.load(buffer, { ...opts, legacyPpt: undefined, odp: undefined });
    const load = beginModelSourceLoad(selected, 'pptx');
    try {
      metrics.setSourceBytes(buffer.byteLength);
      metrics.checkpoint('container ready');
      const worker = mode === 'worker'
        ? (await import('../render-worker-host.js')).createRenderWorker()
        : new (await import('../worker-source.ts?worker&inline')).default();
      const wiredWorker = sourceWorker(worker, load);
      const renderers = mode === 'worker' ? workerRendererDescriptors(opts) : undefined;
      let presentation: PptxPresentation | undefined;
      try {
        // TypeScript's private constructor is compile-time only. This friend
        // module owns source setup; the returned object remains the public class.
        const Constructor = PptxPresentation as unknown as new (
          worker: Worker, mode: 'main' | 'worker', wasmUrl?: string | URL,
        ) => PptxPresentation;
        presentation = new Constructor(wiredWorker, mode, opts.wasmUrl);
        const state = presentation as unknown as MutablePresentation;
        state._metrics = metrics;
        state._cjkFallback = resolveCjkFallback(opts.cjkFallback);
        state._math = mode === 'worker' ? undefined : opts.math;
        state._threeD = mode === 'worker' ? undefined : opts.threeD;
        state._regionMap = mode === 'worker' ? undefined : opts.regionMap;
        state._chartEx = mode === 'worker' ? undefined : opts.chartEx;
        state._tiff = mode === 'worker' ? undefined : opts.tiff;
        state._fonts = opts.fonts;
        warnUnserializableRenderers(opts, mode, renderers);
        const progressive = opts.progressiveLayout ? {
          onProgress: opts.onLayoutProgress,
          onPartial: opts.onLayoutPartial,
          onComplete: opts.onLayoutComplete,
          firstPublication: deferred<void>(),
          published: false, deferred: false, settled: false,
        } : undefined;
        await state._parse(
          buffer, resourceOptions.policy, opts.fonts,
          opts.workerTimeoutMs, (usage) => metrics.observeUsage(usage),
          renderers, progressive,
        );
        metrics.checkpoint('presentation preflight ready');
        if (mode === 'main' && opts.fonts && state._preflight && !progressive) {
          await state._ensureFontSources(
            state._preflight.fontPreloadNames,
            state._preflight.fontPreloadFacts,
            state._preflight.fontPreloadDemand,
          );
        }
        metrics.succeed({ slides: presentation.slideCount });
        load.release();
        return presentation;
      } catch (error) {
        const rejected = presentation;
        disposeRejectedLoad(worker, rejected ? () => rejected.destroy() : undefined);
        throw error;
      }
    } finally {
      load.release();
    }
  } catch (error) {
    metrics.fail(error);
    throw error;
  }
}

function sourceWorker(worker: Worker, load: AdmittedModelSourceLoad): Worker {
  const sourceOwnerUrl = resolveBundledAssetUrl(new URL(
    import.meta.env.DEV ? './worker-presentation-source.ts' : './pptx-source-worker.mjs',
    import.meta.url,
  ).href, import.meta.url);
  return modelSourceWorker(worker, {
    discriminator: 'kind',
    load,
    parseMessage: (wire) => ({ ...wire, source: load.module, sourceOwnerUrl,
      ...(load.transfer.length ? { sourceTransfer: load.transfer } : {}) }),
  });
}

function warnUnserializableRenderers(
  opts: LoadOptions,
  mode: 'main' | 'worker',
  renderers: WorkerRendererDescriptors | undefined,
): void {
  if (mode !== 'worker') return;
  if (opts.math && !renderers?.math) console.warn(
    "[ooxml] a custom math renderer cannot cross the worker boundary; equations will be skipped in mode: 'worker'. Use the math renderer from @silurus/ooxml/math.",
  );
  if (opts.threeD && !renderers?.threeD) console.warn(
    "[ooxml] a custom 3-D chart renderer cannot cross the worker boundary; charts use their 2-D family fallback in mode: 'worker'. Use the renderer from @silurus/ooxml/three-d.",
  );
  if (opts.regionMap && !renderers?.regionMap) console.warn(
    "[ooxml] a custom Region Map renderer cannot cross the worker boundary; geospatial charts use the unsupported-chart placeholder in mode: 'worker'. Use the renderer from @silurus/ooxml/region-map.",
  );
  if (opts.chartEx && !renderers?.chartEx) console.warn(
    "[ooxml] a custom ChartEx renderer cannot cross the worker boundary; ChartEx charts use the unsupported-chart placeholder in mode: 'worker'. Use the renderer from @silurus/ooxml/chart-ex.",
  );
  if (opts.tiff && !renderers?.tiff) console.warn(
    "[ooxml] a custom TIFF codec cannot cross the worker boundary; recognized TIFF images will use an unavailable-image placeholder in mode: 'worker'. Use the codec from @silurus/ooxml/tiff to display them.",
  );
}
