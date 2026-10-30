/**
 * Google Fonts as an ordinary font source: a {@link fontCssApi} descriptor filled
 * with Google-owned data. Nothing outside this module knows about Google.
 *
 * What is Google's here, and only here:
 * - the default CSS origin and the CSS2 API URL grammar
 *   (`/css2?family=<Family+Name>:<axes>&display=swap`);
 * - the families the library asks Google for, each with the exact axis query
 *   the loaders have always requested (so requests are byte-identical to the
 *   former per-format tables, which the oracle test pins);
 * - Google's distribution names for the native Noto CJK packages
 *   ("Noto Sans CJK SC" is served as "Noto Sans SC"). Google serves no
 *   Noto Serif HK, so the native serif HK name has no substitution.
 *
 * The Office substitutions (Calibri → Carlito, …) are the vendor-neutral
 * {@link OFFICE_FONT_SUBSTITUTIONS}, adopted here like any source may adopt
 * them. Together with the Noto CJK naming they form
 * {@link GOOGLE_FONTS_SUBSTITUTIONS}, the default `substitutes`; a caller's
 * `substitutes` replaces that list, exactly as it is every other source's
 * whole list.
 */
import { fontCssApi, type FontCssApiSource } from '../source.js';
import { OFFICE_FONT_SUBSTITUTIONS, type FontSubstitution } from '../substitutions.js';

export const GOOGLE_FONTS_DEFAULT_CSS_ORIGIN = 'https://fonts.googleapis.com';

/** Regular, bold and their italics. */
const TEXT_FACES = 'ital,wght@0,400;0,700;1,400;1,700';
/** Regular and bold, upright only: the Noto script families. */
const SCRIPT_FACES = 'wght@400;700';

/** Served family (Google's spelling) → CSS2 axis query. */
export const GOOGLE_FONTS_CATALOG: Readonly<Record<string, string>> = Object.freeze({
  // Office substitutes (see OFFICE_FONT_SUBSTITUTIONS). Libre Franklin also
  // carries 600 so Franklin Gothic Medium has its own weight available.
  'Carlito': TEXT_FACES,
  'Caladea': TEXT_FACES,
  'Libre Franklin': 'ital,wght@0,400;0,600;0,700;1,400;1,600;1,700',
  // Popular free web fonts Office templates name directly.
  'Nunito Sans': TEXT_FACES,
  'Nunito': TEXT_FACES,
  'Open Sans': TEXT_FACES,
  'Roboto': TEXT_FACES,
  'Lato': TEXT_FACES,
  'Montserrat': TEXT_FACES,
  'Poppins': TEXT_FACES,
  'Raleway': TEXT_FACES,
  'Playfair Display': TEXT_FACES,
  'Ubuntu': TEXT_FACES,
  // Arabic substitutes and library Arabic fallbacks.
  'Noto Naskh Arabic': SCRIPT_FACES,
  'Noto Sans Arabic': SCRIPT_FACES,
  // CJK: sans and serif per language, except HK (sans only on Google Fonts).
  'Noto Sans KR': SCRIPT_FACES,
  'Noto Sans SC': SCRIPT_FACES,
  'Noto Sans TC': SCRIPT_FACES,
  'Noto Sans HK': SCRIPT_FACES,
  'Noto Sans JP': SCRIPT_FACES,
  'Noto Serif KR': SCRIPT_FACES,
  'Noto Serif SC': SCRIPT_FACES,
  'Noto Serif TC': SCRIPT_FACES,
  'Noto Serif JP': SCRIPT_FACES,
  // Latin/Cyrillic/Greek, Indic / South-East Asian and Hebrew fallbacks.
  'Noto Sans': SCRIPT_FACES,
  'Noto Serif': SCRIPT_FACES,
  'Noto Sans Devanagari': SCRIPT_FACES,
  'Noto Sans Thai': SCRIPT_FACES,
  'Noto Sans Hebrew': SCRIPT_FACES,
  'Noto Serif Hebrew': SCRIPT_FACES,
});

/** Native Noto CJK package names → Google's distribution names. */
export const GOOGLE_NOTO_CJK_NAMING: readonly FontSubstitution[] = Object.freeze([
  ['Noto Sans CJK SC', 'Noto Sans SC'],
  ['Noto Serif CJK SC', 'Noto Serif SC'],
  ['Noto Sans CJK TC', 'Noto Sans TC'],
  ['Noto Serif CJK TC', 'Noto Serif TC'],
  ['Noto Sans CJK JP', 'Noto Sans JP'],
  ['Noto Serif CJK JP', 'Noto Serif JP'],
  ['Noto Sans CJK KR', 'Noto Sans KR'],
  ['Noto Serif CJK KR', 'Noto Serif KR'],
  ['Noto Sans CJK HK', 'Noto Sans HK'],
].map(([authored, served]) => Object.freeze({ for: authored, use: served })));

/** The default `googleFonts()` substitutes: the Office stand-ins, then
 * Google's names for the Noto CJK packages. Spread it to extend it. */
export const GOOGLE_FONTS_SUBSTITUTIONS: readonly FontSubstitution[] = Object.freeze([
  ...OFFICE_FONT_SUBSTITUTIONS,
  ...GOOGLE_NOTO_CJK_NAMING,
]);

export interface GoogleFontsOptions {
  /**
   * Google-compatible CSS origin (scheme, host and optional port only), for
   * example a regional or internal mirror. Default `https://fonts.googleapis.com`.
   * Font file URLs come from the returned CSS, resolved against the
   * stylesheet's final URL.
   */
  readonly cssOrigin?: string;
  /**
   * Family substitutions, as for every other source. When given, it replaces
   * the default {@link GOOGLE_FONTS_SUBSTITUTIONS} (Office stand-ins and the
   * Noto CJK naming) entirely; spread that list to extend it. Only families
   * in the Google Fonts catalog the preset requests can be served.
   */
  readonly substitutes?: readonly FontSubstitution[];
}

function validatedOrigin(value: unknown): string {
  if (typeof value !== 'string') throw new TypeError('googleFonts() cssOrigin must be a string');
  const origin = value.endsWith('/') ? value.slice(0, -1) : value;
  let url: URL;
  try {
    url = new URL(origin);
  } catch {
    throw new TypeError('googleFonts() cssOrigin must be an absolute http(s) origin');
  }
  if ((url.protocol !== 'https:' && url.protocol !== 'http:') || url.origin !== origin) {
    throw new TypeError('googleFonts() cssOrigin must be an http(s) origin without path, query or credentials');
  }
  return origin;
}

/** Google Fonts (or a Google-compatible mirror) as a font source. */
export function googleFonts(options: GoogleFontsOptions = {}): FontCssApiSource {
  if (!options || typeof options !== 'object' || Array.isArray(options)) {
    throw new TypeError('googleFonts() options must be an object');
  }
  for (const key of Object.keys(options)) {
    if (key !== 'cssOrigin' && key !== 'substitutes') throw new TypeError(`googleFonts() options has unknown property "${key}"`);
  }
  const origin = options.cssOrigin === undefined
    ? GOOGLE_FONTS_DEFAULT_CSS_ORIGIN
    : validatedOrigin(options.cssOrigin);
  return fontCssApi({
    urlTemplate: `${origin}/css2?family={family}:{query}&display=swap`,
    families: GOOGLE_FONTS_CATALOG,
    substitutes: options.substitutes === undefined ? GOOGLE_FONTS_SUBSTITUTIONS : options.substitutes,
  });
}
