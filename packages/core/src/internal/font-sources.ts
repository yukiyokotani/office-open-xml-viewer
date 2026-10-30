// Internal font-source descriptors, factories, substitution data, session and
// the `useGoogleFonts` migration. The public factories are re-exported by the
// format entries that consume the `fonts` load option.
export {
  FONT_SOURCE_PROTOCOL,
  assertFontSourceDescriptors,
  fontCssApi,
  fontCssApiRequests,
  fontCss,
  fontFiles,
  fontResolver,
  fontResolverFunction,
  validateFontFaceDeclaration,
  type FontCssApiOptions,
  type FontCssApiRequest,
  type FontCssApiSource,
  type FontDemand,
  type FontCssOptions,
  type FontCssSource,
  type FontFaceDeclaration,
  type FontFilesOptions,
  type FontFilesSource,
  type FontResolveContext,
  type FontResolveFn,
  type FontResolverOptions,
  type FontResolverSource,
  type FontSource,
  type FontSourceCredentials,
} from '../fonts/source.js';
export {
  OFFICE_FONT_SUBSTITUTIONS,
  fontFamilyKey,
  loadedRegularAliases,
  matchFontDeclaration,
  substitutionsFor,
  type FontDeclarationMatch,
  type FontSubstitution,
  type LoadedFaceFacts,
} from '../fonts/substitutions.js';
export {
  GOOGLE_FONTS_CATALOG,
  GOOGLE_FONTS_DEFAULT_CSS_ORIGIN,
  GOOGLE_FONTS_SUBSTITUTIONS,
  GOOGLE_NOTO_CJK_NAMING,
  googleFonts,
  type GoogleFontsOptions,
} from '../fonts/presets/google-fonts.js';
export { fontSourceLimitError } from '../errors/font-source-limit.js';
export {
  FontSourceSession,
  normalizeFontDemand,
  type FontResolverFailures,
  type FontSessionDemand,
  type FontResolverRunner,
  type FontSourceSessionOptions,
  type LoadedFontSources,
  type LoadedSourceFace,
} from '../fonts/source-session.js';
export { fetchFontStylesheet, parseFontFaceCss, parseUnicodeRange } from '../fonts/css-source.js';
export {
  WorkerFontResolverClient,
  isMainFontResolveReply,
  isWorkerFontResolveMessage,
  type FontResolveOutcome,
  type MainFontResolveReply,
  type WorkerFontResolveCancel,
  type WorkerFontResolveMessage,
  type WorkerFontResolveRequest,
} from '../fonts/resolver-bridge.js';
export { FontResolverBridgeHost, type FontResolverBridgeHostOptions } from '../fonts/resolver-bridge-host.js';
export {
  USE_GOOGLE_FONTS_DEPRECATION_MESSAGE,
  normalizeFontOptions,
  resetFontDeprecationWarningsForTests,
  type FontLoadOptions,
} from '../fonts/deprecations.js';
export { FONT_SUBSTITUTE_SCRIPTS } from '../fonts/substitute-script.js';
export { probeFontFamilyPresence, type FontFamilyPresence } from '../fonts/office-fallback.js';
export {
  InstalledFamilyExclusion,
  type InstalledFamilyProbe,
} from '../fonts/installed-families.js';
export {
  chartFontFamilies,
  chartFontSubstitutions,
  withChartFontSubstitutions,
  type LoadedChartFontSubstitutions,
} from '../fonts/chart-font-sources.js';
