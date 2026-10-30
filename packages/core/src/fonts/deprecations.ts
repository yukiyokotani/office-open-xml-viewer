/**
 * The `useGoogleFonts` → `fonts` migration (library API policy, not OOXML
 * behaviour).
 *
 * `useGoogleFonts` stays accepted for exactly one minor release (0.90) as an
 * alias of `fonts: [googleFonts()]`, and is removed in 0.91. Each load entry
 * point calls {@link normalizeFontOptions} before anything else reads its
 * options, so the rest of the load sees only `fonts`. Viewers forward their
 * options to those entry points, and a render worker receives the normalised
 * descriptors, never the key, so the warning is emitted by the realm that
 * received the option and nowhere else.
 */
import { googleFonts } from './presets/google-fonts.js';
import { assertFontSourceDescriptors, type FontSource } from './source.js';

export const USE_GOOGLE_FONTS_DEPRECATION_MESSAGE =
  '[ooxml] useGoogleFonts is deprecated and will be removed in 0.91.0. Replace it with fonts: [googleFonts()] — import { googleFonts } from \'@silurus/ooxml\' (also exported from the docx/xlsx/pptx entries). A custom CSS origin is googleFonts({ cssOrigin }); other sources are fontCssApi(), fontCss(), fontFiles() and fontResolver(). See docs/migration-0.90.md.';

// One realm's module instance (main window, each popup, each worker) warns
// once, however many documents or viewers pass the deprecated key.
let warned = false;

/** @internal Test hook: forget that this realm already warned. */
export function resetFontDeprecationWarningsForTests(): void {
  warned = false;
}

/** The font fields a load entry point accepts. */
export interface FontLoadOptions {
  readonly fonts?: readonly FontSource[];
  readonly useGoogleFonts?: boolean;
}

/**
 * Normalise the font fields of one load request: the returned options carry a
 * validated `fonts` (omitted when there are no sources) and never
 * `useGoogleFonts`.
 *
 * - `useGoogleFonts: true` becomes `fonts: [googleFonts()]`;
 * - `useGoogleFonts: false` is a no-op;
 * - either value warns once per realm (the key was used), and `true` together
 *   with `fonts` throws a `TypeError` (the caller must choose one);
 * - an undefined value (a viewer forwarding an unset option) is ignored;
 * - `fonts` is validated as the factories validate it (`TypeError` for a
 *   malformed descriptor, `OoxmlError('font-source-limit')` beyond a ceiling).
 */
export function normalizeFontOptions<T extends FontLoadOptions>(
  options: T,
): Omit<T, 'useGoogleFonts' | 'fonts'> & { readonly fonts?: readonly FontSource[] } {
  const { useGoogleFonts, fonts, ...rest } = options;
  if (useGoogleFonts !== undefined) {
    if (useGoogleFonts === true && fonts !== undefined) {
      throw new TypeError('useGoogleFonts cannot be combined with fonts; add googleFonts() to fonts instead');
    }
    if (!warned) {
      warned = true;
      console.warn(USE_GOOGLE_FONTS_DEPRECATION_MESSAGE);
    }
  }
  const sources = useGoogleFonts === true
    ? [googleFonts()]
    : fonts === undefined ? undefined : assertFontSourceDescriptors(fonts);
  return sources !== undefined && sources.length > 0 ? { ...rest, fonts: Object.freeze([...sources]) } : rest;
}
