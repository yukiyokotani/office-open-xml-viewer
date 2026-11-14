// Internal shared demand collector: intentionally excluded from the package
// root API so the inert-scalar table stays in lazily imported chunks.
export {
  FontDemandScalarCollector,
  collectFontDemandScalars,
  type FontDemandScalarOptions,
  type FontDemandScalars,
} from '../fonts/scalar-demand.js';
export {
  CHART_SYNTHESIZED_TEXT,
  VERTICAL_PRESENTATION_FORMS,
  chartPaintTexts,
  paintsTextFreeImage,
} from '../fonts/chart-paint-texts.js';
