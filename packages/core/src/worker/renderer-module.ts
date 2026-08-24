import type { ChartRegionMapRenderer } from '../chart/region-map-contract.js';
import type { ChartThreeDRenderer } from '../chart/three-d-contract.js';
import type { ChartExRenderer } from '../chart/chart-ex-contract.js';
import type { MathRenderer } from '../math/mathjax.js';
import {
  assertWorkerRendererDescriptor,
  type WorkerRendererDescriptor,
  type WorkerRendererDescriptors,
} from './renderer-module-contract.js';

export interface LoadedWorkerRenderers {
  readonly math?: MathRenderer;
  readonly threeD?: ChartThreeDRenderer;
  readonly regionMap?: ChartRegionMapRenderer;
  readonly chartEx?: ChartExRenderer;
}

/** Import a renderer's named export in the calling realm (normally a render
 * worker). The descriptor comes only from an explicit load option; document
 * content never selects executable module URLs. */
export async function loadWorkerRenderer<T>(descriptor: WorkerRendererDescriptor): Promise<T> {
  assertWorkerRendererDescriptor(descriptor);
  return loadBuiltinRenderer(descriptor) as Promise<T>;
}

async function loadBuiltinRenderer(descriptor: WorkerRendererDescriptor): Promise<object> {
  switch (descriptor.builtin) {
    case 'math': {
      const engine = await import('../math/engine.js');
      return Object.freeze({
        loadMathJax: () => engine.loadMathJaxFromAsset(descriptor.engineAssetUrl),
        mathMLToSvg: (mathml: string) => engine.mathMLToSvgFromAsset(
          mathml,
          descriptor.engineAssetUrl,
        ),
      });
    }
    case 'threeD': {
      const renderer = await import('../chart/three-d-renderer.js');
      return Object.freeze({ render: renderer.renderSimpleThreeDChart });
    }
    case 'regionMap': {
      const renderer = await import('../chart/region-map-renderer.js');
      return Object.freeze({ render: renderer.renderRegionMapChart });
    }
    case 'chartEx': {
      const renderer = await import('../chart/chart-ex-renderer.js');
      return Object.freeze({ render: renderer.renderChartExChart });
    }
  }
}

function requireRendererMethods<T extends object>(
  rendererName: string,
  value: unknown,
  methods: readonly string[],
): T {
  if (typeof value !== 'object' || value === null) {
    throw new TypeError(`Worker ${rendererName} renderer export must be an object`);
  }
  const record = value as Record<string, unknown>;
  for (const method of methods) {
    if (typeof record[method] !== 'function') {
      throw new TypeError(`Worker ${rendererName} renderer must implement ${method}()`);
    }
  }
  return value as T;
}

/** Recreate all explicitly supplied renderers in the current worker realm. */
export async function loadWorkerRenderers(
  descriptors: WorkerRendererDescriptors | undefined,
): Promise<LoadedWorkerRenderers> {
  const [math, threeD, regionMap, chartEx] = await Promise.all([
    descriptors?.math ? loadWorkerRenderer(descriptors.math) : undefined,
    descriptors?.threeD ? loadWorkerRenderer(descriptors.threeD) : undefined,
    descriptors?.regionMap ? loadWorkerRenderer(descriptors.regionMap) : undefined,
    descriptors?.chartEx ? loadWorkerRenderer(descriptors.chartEx) : undefined,
  ]);
  return Object.freeze({
    ...(math ? {
      math: requireRendererMethods<MathRenderer>(
        'math',
        math,
        ['loadMathJax', 'mathMLToSvg'],
      ),
    } : {}),
    ...(threeD ? {
      threeD: requireRendererMethods<ChartThreeDRenderer>('threeD', threeD, ['render']),
    } : {}),
    ...(regionMap ? {
      regionMap: requireRendererMethods<ChartRegionMapRenderer>(
        'regionMap',
        regionMap,
        ['render'],
      ),
    } : {}),
    ...(chartEx ? {
      chartEx: requireRendererMethods<ChartExRenderer>('chartEx', chartEx, ['render']),
    } : {}),
  });
}
