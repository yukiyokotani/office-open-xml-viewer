import {
  resolveCjkFallback,
  type AdmittedModelSourceLoad,
  type OoxmlResourceUsageSnapshot,
} from '@silurus/ooxml-core';
import {
  disposeRejectedLoad,
  normalizeLoadResourceOptions,
  OoxmlResourceMetricsSession,
  normalizeXlsxWorksheetPolicy,
  type NormalizedOoxmlResourcePolicy,
} from '@silurus/ooxml-core/worker';
import { beginModelSourceLoad, selectModelSource } from '@silurus/ooxml-core/internal/model-source';
import { computeMdw, pinXlsxGridGeometry } from '../renderer.js';
import { respondToHostLayoutRequest } from './host-layout.js';
import type { XlsxWorkbook, LoadOptions } from '../workbook.js';
import type { ParsedWorkbook, Worksheet, ViewportRange, RenderViewportOptions } from '../types.js';
import { extractViewerRenderContext, withViewerRenderContext } from '../worker-protocol.js';

type WorkbookConstructor = new (worker: Worker, mode: 'main' | 'worker', wasmUrl?: string | URL) => XlsxWorkbook;

type MutableWorkbook = {
  metrics: OoxmlResourceMetricsSession;
  parsedWorkbook: ParsedWorkbook | null;
  _load(
    data: ArrayBuffer,
    opts: LoadOptions,
    policy: NormalizedOoxmlResourcePolicy,
    onUsage: (usage: OoxmlResourceUsageSnapshot) => void,
    preserveCallerBuffer: boolean,
  ): Promise<void>;
};

/** Admit a source before container resolution, then run the ordinary parser client on its own worker. */
export async function loadXlsxModelSource(
  input: string | ArrayBuffer,
  opts: LoadOptions,
  workbookType: typeof import('../workbook.js')['XlsxWorkbook'],
): Promise<XlsxWorkbook> {
  const worksheetPolicy = normalizeXlsxWorksheetPolicy(opts);
  opts = { ...opts, xlsxWorksheetLimits: worksheetPolicy.worksheet };
  opts = { ...opts, cjkFallback: resolveCjkFallback(opts.cjkFallback) };
  const resourceOptions = normalizeLoadResourceOptions(opts);
  const mode = opts.mode ?? 'main';
  const metrics = new OoxmlResourceMetricsSession({
    enabled: true, format: 'xlsx', mode, policy: resourceOptions.policy,
    xlsxWorksheetPolicy: worksheetPolicy,
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
    const selected = selectModelSource(opts.modelSources, 'xlsx', new Uint8Array(buffer));
    if (!selected) return workbookType.load(buffer, { ...opts, modelSources: undefined });
    // The caller already owns this public constructor. Passing it avoids a
    // static source-only back-edge that splits the ordinary workbook graph;
    // source overrides remain confined to this admitted source's subclass.
    const BaseWorkbook = workbookType as unknown as WorkbookConstructor;
    /** Source-only behavior lives in this subclass, outside the ordinary entry. */
    class SourceWorkbook extends BaseWorkbook {
      sourceMdw: number | undefined;

      override async getWorksheet(sheetIndex: number): Promise<Worksheet> {
        const worksheet = await super.getWorksheet(sheetIndex);
        if (this.sourceMdw !== undefined) pinXlsxGridGeometry(worksheet, this.sourceMdw);
        return worksheet;
      }

      override async renderViewport(
        target: HTMLCanvasElement | OffscreenCanvas,
        sheetIndex: number,
        viewport: ViewportRange,
        options: RenderViewportOptions = {},
      ): Promise<void> {
        const resolved = this.sourceMdw !== undefined
          && !extractViewerRenderContext(options).layoutMetrics
          ? withViewerRenderContext(options, this.sourceMdw)
          : options;
        return super.renderViewport(target, sheetIndex, viewport, resolved);
      }
    }
    const load = beginModelSourceLoad(selected, 'xlsx');
    try {
      metrics.setSourceBytes(buffer.byteLength);
      metrics.checkpoint('container ready');
      const worker = mode === 'worker'
        ? (await import('../render-worker-source-host.js')).createRenderWorker()
        : new (await import('../worker-source.ts?worker&inline')).default();
      let workbook: SourceWorkbook | undefined;
      const wiredWorker = sourceWorker(worker, load, opts, (mdw) => { if (workbook) workbook.sourceMdw = mdw; });
      try {
        workbook = new SourceWorkbook(wiredWorker, mode, opts.wasmUrl);
        const state = workbook as unknown as MutableWorkbook;
        state.metrics = metrics;
        await state._load(buffer, opts, resourceOptions.policy, (usage) => metrics.observeUsage(usage), true);
        if (workbook.sourceMdw !== undefined && state.parsedWorkbook) {
          state.parsedWorkbook.layoutMetrics = { maximumDigitWidth: workbook.sourceMdw };
        }
        metrics.checkpoint('workbook index ready');
        metrics.succeed({ sheets: workbook.sheetCount });
        load.release();
        return workbook;
      } catch (error) {
        const rejected = workbook;
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

function sourceWorker(
  worker: Worker,
  load: AdmittedModelSourceLoad,
  opts: LoadOptions,
  onMdw: (value: number) => void,
): Worker {
  const sourceOwnerUrl = new URL(
    import.meta.env.DEV ? './worker-worksheet-source.ts' : './xlsx-source-worker.mjs',
    import.meta.url,
  ).href;
  return new Proxy(worker, {
    get(target, key) {
      if (key === 'postMessage') return (message: unknown, transfer?: Transferable[]) => {
        if (typeof message === 'object' && message !== null) {
          const wire = message as { type?: string };
          if (wire.type === 'init') return;
          if (wire.type === 'parse') {
            target.postMessage({ ...wire, source: load.module, sourceOwnerUrl,
              ...(load.transfer.length ? { sourceTransfer: load.transfer } : {}) },
            [...(transfer ?? []), ...load.transfer]);
            return;
          }
        }
        target.postMessage(message, transfer ?? []);
      };
      if (key === 'addEventListener') return (type: string, listener: EventListener) => {
        if (type !== 'message') return target.addEventListener(type, listener);
        return target.addEventListener('message', ((event: MessageEvent) => {
          const response = event.data as { type?: string; layoutMetrics?: { maximumDigitWidth: number } };
          if (response.type === 'parsed' && response.layoutMetrics) {
            onMdw(response.layoutMetrics.maximumDigitWidth);
          }
          if (type === 'message') respondToHostLayoutRequest(
            (reply) => target.postMessage(reply), event.data,
            (font) => computeMdw(font.family, font.sizePt, undefined, !!opts.useGoogleFonts,
              font.bold ? 700 : 400, font.italic ? 'italic' : 'normal'),
          );
          listener(event);
        }) as EventListener);
      };
      const value = Reflect.get(target, key, target);
      return typeof value === 'function' ? value.bind(target) : value;
    },
  });
}
