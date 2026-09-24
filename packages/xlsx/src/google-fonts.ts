import {
  classifyCjkFont,
  cjkFallbackForText,
  scriptPreloadNamesForText,
  GOOGLE_FONT_SUBSTITUTES,
  SCRIPT_GOOGLE_FONTS,
  findReferenceFontMetrics,
  type CjkLang,
  type FontPreloadEntry,
  type OfficeFontFallbackRequest,
} from '@silurus/ooxml-core';
import type { ParsedWorkbook, Worksheet } from './types.js';
import { officeRequestKey, singleNaturalShapeRun } from './shape-office-line.js';

/** Office font name → Google Fonts substitute for XLSX cells.
 *
 *  {@link GOOGLE_FONT_SUBSTITUTES} supplies the Office substitutes (Calibri →
 *  Carlito, Cambria → Caladea — advance-width alternatives for base text faces,
 *  without an exact Excel layout guarantee), the popular free web fonts
 *  and the Arabic Noto fallbacks — shared with docx/pptx. {@link
 *  SCRIPT_GOOGLE_FONTS} adds the CJK / Cyrillic / Thai / Devanagari / Hebrew
 *  Noto faces (the renderer chooses the CJK Noto per cell from the cell's font
 *  name; non-CJK scripts append to the default chain). Both load only when
 *  `useGoogleFonts` is on — no binaries ship in the bundle. XLSX currently has
 *  no format-specific additions. */
export const XLSX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  ...GOOGLE_FONT_SUBSTITUTES,
  ...SCRIPT_GOOGLE_FONTS,
};

/** Yield every textual cell value carried by the parsed workbook: the shared
 *  string table (`text` plus rich-text `runs[].text`). This is the bulk of a
 *  workbook's painted text and is present in BOTH main and worker at parse time
 *  (sheets parse lazily, but the shared string table is workbook-level).
 *  Numbers / dates carry no script-specific glyphs, so they are irrelevant. */
function* xlsxTextRuns(wb: ParsedWorkbook | undefined): Generator<string> {
  for (const s of wb?.sharedStrings ?? []) {
    if (s.runs && s.runs.length > 0) {
      for (const r of s.runs) yield r.text;
    } else {
      yield s.text;
    }
  }
}

/**
 * The font-family names to preload for a workbook: every styled cell font, plus
 * only the script-fallback Noto faces whose script the workbook's TEXT actually
 * contains ({@link scriptPreloadNamesForText}). Office faces map to
 * advance-width substitutes (Calibri → Carlito, Cambria → Caladea); the
 * renderer's default chain still ends with the full Noto set, but eagerly
 * fetching the multi-MB CJK families for a workbook that has no CJK glyphs would
 * block first paint for nothing; an un-preloaded face loads lazily if it ever
 * proves needed. A workbook using only system fonts (no map entries) still
 * produces zero network requests.
 *
 * Single source of truth shared by the main-thread `_load()` and the render
 * worker. Both derive the set from the SAME parsed {@link ParsedWorkbook}, so
 * both modes preload an identical set — worker/main rendering must stay
 * pixel-equivalent.
 */
export function xlsxFontPreloadNames(wb: ParsedWorkbook | undefined, fallback?: CjkLang): Set<string> {
  const names = new Set<string>();
  let cjkLang: CjkLang | null = null;
  for (const f of wb?.styles?.fonts ?? []) {
    if (f.name) {
      names.add(f.name);
      cjkLang ??= classifyCjkFont(f.name);
    }
  }
  for (const n of scriptPreloadNamesForText(xlsxTextRuns(wb), cjkLang ?? fallback ?? null)) {
    names.add(n);
  }
  return names;
}

/** The same workbook hint is used for preloading and ambiguous cell fallback. */
export function xlsxCjkFallback(wb: ParsedWorkbook | undefined, fallback: CjkLang): CjkLang {
  for (const font of wb?.styles?.fonts ?? []) {
    const region = classifyCjkFont(font.name);
    if (region) return cjkFallbackForText(xlsxTextRuns(wb), region);
  }
  return cjkFallbackForText(xlsxTextRuns(wb), fallback);
}

/** Exact Calibri style slots in the workbook style and shared-string tables.
 * An omitted font name uses the workbook default Calibri chain; unrelated
 * authored families never borrow Carlito merely through a CSS fallback tail.
 * Styles are prepared before a sheet is pulled so Normal-font MDW and viewer
 * geometry cannot be captured against a different fallback resource. This can
 * prepare an unused style slot, bounded by the four supported tuples. */
export function xlsxOfficeFontRequests(wb: ParsedWorkbook | undefined): OfficeFontFallbackRequest[] {
  const found = new Map<string, OfficeFontFallbackRequest>();
  const add = (name: string | null | undefined, bold: boolean, italic: boolean) => {
    if ((name?.trim().toLowerCase() || 'calibri') !== 'calibri') return;
    const weight = bold ? 700 : 400;
    const style = italic ? 'italic' : 'normal';
    found.set(`${weight}:${style}`, { family: 'Calibri', weight, style });
  };
  for (const font of wb?.styles?.fonts ?? []) add(font.name, font.bold, font.italic);
  for (const shared of wb?.sharedStrings ?? []) {
    for (const run of shared.runs ?? []) {
      if (run.font) add(run.font.name, run.font.bold, run.font.italic);
    }
  }
  return [...found.values()];
}

/** Inline strings and DrawingML shapes are worksheet-local and absent from the
 * bootstrap shared-string table. Shape preflight is limited to one natural
 * text run with a catalogued exact style; cell requests remain Calibri-only. */
export function xlsxWorksheetOfficeFontRequests(ws: Worksheet): OfficeFontFallbackRequest[] {
  const found = new Map<string, OfficeFontFallbackRequest>();
  // The parser resolves Normal through cellStyleXfs[0].fontId. Preflight that
  // authored face before measuring column MDW, including families that occur
  // nowhere in a cell's text. Exact local bytes win over catalog references.
  const normalFamily = ws.defaultFontFamily?.trim();
  if (normalFamily && findReferenceFontMetrics(normalFamily, { weight: 400, style: 'normal' }).length) {
    const request = { family: normalFamily, weight: 400, style: 'normal' } as const;
    found.set(officeRequestKey(request), request);
  }
  for (const row of ws.rows) for (const cell of row.cells) {
    if (cell.value.type !== 'text') continue;
    for (const run of cell.value.runs ?? []) {
      const font = run.font;
      if (!font || (font.name?.trim().toLowerCase() || 'calibri') !== 'calibri') continue;
      const weight = font.bold ? 700 : 400;
      const style = font.italic ? 'italic' : 'normal';
      const request = { family: 'Calibri', weight, style } as const;
      found.set(officeRequestKey(request), request);
    }
  }
  for (const anchor of ws.shapeGroups ?? []) for (const shape of anchor.shapes) {
    if (!shape.text) continue;
    const run = singleNaturalShapeRun(shape.text);
    if (!run) continue;
    const weight = run.bold ? 700 : 400;
    const style = run.italic ? 'italic' : 'normal';
    if (findReferenceFontMetrics(run.fontFace!, { weight, style }).length === 0) continue;
    const request = { family: run.fontFace!.trim(), weight, style } as const;
    found.set(officeRequestKey(request), request);
  }
  return [...found.values()];
}
