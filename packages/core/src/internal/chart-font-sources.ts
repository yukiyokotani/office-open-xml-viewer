// Internal chart font-source helpers (chart face demand and the loaded
// substitution map charts paint), a light entry the format renderers import
// statically without the font-source session barrel.
export {
  chartFontFamilies,
  chartFontSubstitutions,
  withChartFontSubstitutions,
  type LoadedChartFontSources,
} from '../fonts/chart-font-sources.js';
