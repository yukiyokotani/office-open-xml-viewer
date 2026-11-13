/**
 * The `useGoogleFonts` → `fonts` migration (library API policy, not OOXML
 * behaviour).
 *
 * `useGoogleFonts` stays accepted for exactly one minor release (0.90) as an
 * alias of `fonts: [googleFonts()]`. From 0.91 the key is ignored with the
 * same once-per-realm warning (never a `TypeError`, so an application that
 * upgrades without migrating keeps working, offline). Each load entry point
 * calls {@link normalizeFontOptions} before anything else reads its options,
 * so the rest of the load sees only `fonts`. Viewers forward their options to
 * those entry points, and a render worker receives the normalised
 * descriptors, never the key, so the warning is emitted by the realm that
 * received the option and nowhere else.
 */
import { googleFonts } from './presets/google-fonts.js';
import { fontSourceDescriptors, type FontSource } from './source.js';

export const USE_GOOGLE_FONTS_DEPRECATION_MESSAGE =
  '[ooxml] useGoogleFonts is deprecated: from 0.91.0 it is ignored (no web fonts are loaded). Replace it with fonts: [googleFonts()] — import { googleFonts } from \'@silurus/ooxml\' (also exported from the docx/xlsx/pptx entries). A custom CSS origin is googleFonts({ cssOrigin }); other sources are fontCssApi(), fontCss(), fontFiles() and fontResolver(). See docs/migration-0.90.md.';

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
 * Normalise the font fields of one load request: the returned options carry
 * the checked `fonts` (omitted when there are no sources; read their
 * descriptors with `optionFontSources`) and never `useGoogleFonts`.
 *
 * - `useGoogleFonts: true` becomes `fonts: [googleFonts()]`;
 * - `useGoogleFonts: false` is a no-op;
 * - either value warns once per realm (the key was used), and `true` together
 *   with `fonts` throws a `TypeError` (the caller must choose one);
 * - an undefined value (a viewer forwarding an unset option) is ignored;
 * - `fonts` must hold only values the factories returned (`TypeError`
 *   otherwise; see {@link fontSourceDescriptors}).
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
  const sources = useGoogleFonts === true ? [googleFonts()] : fonts;
  // Checked here, before anything else reads the options; the values stay
  // the factories' own (opaque) objects.
  if (sources !== undefined) fontSourceDescriptors(sources);
  return sources !== undefined && sources.length > 0 ? { ...rest, fonts: Object.freeze([...sources]) } : rest;
}
