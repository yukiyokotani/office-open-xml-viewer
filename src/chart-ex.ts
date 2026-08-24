// Opt-in Microsoft ChartEx renderer entry point.

import { renderChartExChart } from '../packages/core/src/chart/chart-ex-renderer.js';
import type { ChartExRenderer } from '../packages/core/src/chart/chart-ex-contract.js';
import { registerBuiltinWorkerRenderer } from '../packages/core/src/worker/renderer-module-contract.js';

/** Optional renderer for the newer Microsoft ChartEx (`cx:*`) families. */
export const chartEx: ChartExRenderer = registerBuiltinWorkerRenderer({
  render: renderChartExChart,
}, 'chartEx');

export type { ChartExRenderer } from '../packages/core/src/chart/chart-ex-contract.js';
