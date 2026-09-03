/**
 * Office-font → Google-Fonts substitute registry used by
 * {@link GoogleFontsProvider}.
 *
 * These are the well-known free webfont alternatives Microsoft Office templates
 * pull from, plus the metric-compatible pairings Microsoft and Google both
 * publish (Calibri → Carlito, Cambria → Caladea: same advance widths and
 * ascender / descender). Loading the substitute on a system that lacks the
 * Office face keeps text-width measurements close to Word / PowerPoint / Excel.
 * Entries whose substitute family name differs from the requested face carry a
 * `loadFamily` so the FontFaceSet load is driven against the substitute; the
 * rest omit it because Google Fonts serves the same family name we request.
 *
 * The library ships NO font binaries — every family here is fetched on demand
 * from the Google Fonts CDN only when the caller opts in with
 * `useGoogleFonts: true`, and only when a document actually references the key.
 * A key that never appears in a document is inert (no network request).
 *
 * ## Loading is async — measure timing caveat
 *
 * Fonts load through FontFaceSet asynchronously. Text measured BEFORE a
 * substitute finishes loading uses the system fallback; once loaded, the same
 * text measures against the substitute — so a first paint racing a slow font
 * fetch can differ from a repaint (the preload paths await the load before
 * rendering, which is why this rarely shows in practice). Adding a name to
 * this registry therefore CHANGES how documents using that name measure: from
 * "whatever the OS falls back to" to the (better, deterministic) substitute.
 * Keep the list to faces with a published metric-compatible or well-known
 * substitute; do not add speculative names.
 *
 * ## Why these live in ONE table
 *
 * A DOCX template requesting Roboto, a PPTX theme requesting Calibri Light and
 * an XLSX cell styled Cambria Math all describe the SAME concept — an Office
 * face the host may not ship — and all want the SAME metric-compatible
 * substitute. None of these substitutions is specific to one file format, so
 * the built-in provider owns one shared registry. The format packages only
 * discover names from their parsed models; they do not own Google URLs or
 * substitution policy. Script-fallback Noto faces live separately in
 * {@link SCRIPT_GOOGLE_FONTS}, because CJK ordering is language-dependent.
 *
 * All keys are lower-cased family names (callers lower-case the requested face
 * before lookup).
 */
import type { FontPreloadEntry } from './google-css.js';

const NOTO_NASKH_ARABIC_URL =
  'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap';
const NOTO_SANS_ARABIC_URL =
  'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap';
const LIBRE_FRANKLIN_URL =
  'https://fonts.googleapis.com/css2?family=Libre+Franklin:ital,wght@0,400;0,600;0,700;1,400;1,600;1,700&display=swap';

export const GOOGLE_FONT_SUBSTITUTES: Record<string, FontPreloadEntry> = {
  // Metric-compatible Office substitutes (same advance widths / vertical
  // metrics). "Calibri Light" (Office theme heading default) and "Cambria Math"
  // (OMML equation font) reduce to the same substitute as their base family.
  'calibri':           { url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Carlito' },
  'calibri light':     { url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Carlito' },
  'cambria':           { url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Caladea' },
  'cambria math':      { url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap', loadFamily: 'Caladea' },
  // Libre Franklin is the open Franklin-family substitute. Office templates
  // distinguish Book (regular) and Medium in the family name rather than with
  // rPr@b; exposing both keys preserves that authored face choice on hosts that
  // do not ship Franklin Gothic (notably macOS/Linux/browser workers).
  'franklin gothic book':   { url: LIBRE_FRANKLIN_URL, loadFamily: 'Libre Franklin' },
  'franklin gothic medium': { url: LIBRE_FRANKLIN_URL, loadFamily: 'Libre Franklin' },
  // Popular free Google web fonts frequently used as Office template body /
  // heading faces. Google serves each under the requested family name, so no
  // `loadFamily` redirect is needed.
  'nunito sans':       { url: 'https://fonts.googleapis.com/css2?family=Nunito+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'nunito':            { url: 'https://fonts.googleapis.com/css2?family=Nunito:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'open sans':         { url: 'https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'roboto':            { url: 'https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'lato':              { url: 'https://fonts.googleapis.com/css2?family=Lato:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'montserrat':        { url: 'https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'poppins':           { url: 'https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'raleway':           { url: 'https://fonts.googleapis.com/css2?family=Raleway:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'playfair display':  { url: 'https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  // Ubuntu — the minorFont in some templates. Without this the renderer falls
  // back to a system sans whose horizontal metrics are narrower than Ubuntu's,
  // so content sized against the Ubuntu width (e.g. a table cell measured to
  // wrap into two lines) wraps differently from the authoring app.
  'ubuntu':            { url: 'https://fonts.googleapis.com/css2?family=Ubuntu:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  // Common Arabic-script faces that hosts rarely ship. Map them to Noto
  // substitutes so RTL documents (which request e.g. Sakkal Majalla / Univers
  // Next Arabic) render with a real web font instead of an oversized OS
  // fallback. "Naskh" covers traditional serif-like Arabic faces; "Sans" covers
  // the modern geometric ones.
  'sakkal majalla':      { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'traditional arabic':  { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'simplified arabic':   { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'arabic typesetting':  { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'univers next arabic': { url: NOTO_SANS_ARABIC_URL, loadFamily: 'Noto Sans Arabic' },
  // Self-referencing entries so the generic Arabic fallback fonts (appended to
  // the renderer's font chain) are themselves loaded whenever useGoogleFonts is
  // enabled — the loaders always queue these names.
  'noto naskh arabic':   { url: NOTO_NASKH_ARABIC_URL, loadFamily: 'Noto Naskh Arabic' },
  'noto sans arabic':    { url: NOTO_SANS_ARABIC_URL, loadFamily: 'Noto Sans Arabic' },
};
