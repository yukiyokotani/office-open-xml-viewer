import { renderChartExChart } from '@silurus/ooxml-core/internal/chart-ex-renderer';
import { renderRegionMapChart } from '@silurus/ooxml-core/internal/region-map-renderer';
import { renderSimpleThreeDChart } from '@silurus/ooxml-core/internal/three-d-renderer';
import { renderTiffToBitmap } from '@silurus/ooxml-core/internal/tiff-renderer';

/**
 * The extension is a ready-to-use product rather than a consumer-controlled
 * application bundle, so it enables every first-party chart renderer.
 */
export const advancedChartRenderers = Object.freeze({
  chartEx: Object.freeze({ render: renderChartExChart }),
  regionMap: Object.freeze({ render: renderRegionMapChart }),
  threeD: Object.freeze({ render: renderSimpleThreeDChart }),
});

/** Every optional first-party renderer/codec enabled by the ready-to-use extension. */
export const fullRenderers = Object.freeze({
  ...advancedChartRenderers,
  tiff: Object.freeze({ render: renderTiffToBitmap }),
});
