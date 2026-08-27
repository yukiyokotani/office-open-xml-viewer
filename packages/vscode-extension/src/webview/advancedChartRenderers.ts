import { renderChartExChart } from '@silurus/ooxml-core/internal/chart-ex-renderer';
import { renderRegionMapChart } from '@silurus/ooxml-core/internal/region-map-renderer';
import { renderSimpleThreeDChart } from '@silurus/ooxml-core/internal/three-d-renderer';

/**
 * The extension is a ready-to-use product rather than a consumer-controlled
 * application bundle, so it enables every first-party chart renderer.
 */
export const advancedChartRenderers = Object.freeze({
  chartEx: Object.freeze({ render: renderChartExChart }),
  regionMap: Object.freeze({ render: renderRegionMapChart }),
  threeD: Object.freeze({ render: renderSimpleThreeDChart }),
});
