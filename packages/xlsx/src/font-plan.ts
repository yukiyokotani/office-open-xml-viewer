import {
  chartFontFamilies,
  classifyCjkFont,
  scriptPreloadNamesForText,
  type CjkLang,
} from '@silurus/ooxml-core';
import type { ParsedWorkbook, Worksheet } from './types.js';

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
 * metric-compatible substitutes (Calibri → Carlito, Cambria → Caladea); the
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
export function xlsxFontPreloadNames(wb: ParsedWorkbook | undefined): Set<string> {
  const names = new Set<string>();
  let cjkLang: CjkLang | null = null;
  for (const f of wb?.styles?.fonts ?? []) {
    if (f.name) {
      names.add(f.name);
      cjkLang ??= classifyCjkFont(f.name);
    }
  }
  for (const n of scriptPreloadNamesForText(xlsxTextRuns(wb), cjkLang)) {
    names.add(n);
  }
  return names;
}

/** Authored families resolved by an application-provided font provider. */
export function xlsxFontProviderNames(workbook: ParsedWorkbook | undefined): string[] {
  const names = new Set<string>();
  for (const font of workbook?.styles.fonts ?? []) {
    if (font.name?.trim()) names.add(font.name.trim());
  }
  for (const item of workbook?.sharedStrings ?? []) {
    for (const run of item.runs ?? []) {
      if (run.font?.name?.trim()) names.add(run.font.name.trim());
    }
  }
  return [...names];
}

/** Families discovered only after a worksheet is parsed lazily. */
export function xlsxWorksheetFontProviderNames(worksheet: Worksheet): string[] {
  const names = new Set<string>();
  for (const row of worksheet.rows) {
    for (const cell of row.cells) {
      if (cell.value.type !== 'text') continue;
      for (const run of cell.value.runs ?? []) {
        if (run.font?.name?.trim()) names.add(run.font.name.trim());
      }
    }
  }
  for (const anchor of worksheet.shapeGroups ?? []) {
    for (const shape of anchor.shapes) {
      for (const paragraph of shape.text?.paragraphs ?? []) {
        for (const run of paragraph.runs) {
          if (run.type !== 'text') continue;
          if (run.fontFace?.trim()) names.add(run.fontFace.trim());
          if (run.fontFaceEa?.trim()) names.add(run.fontFaceEa.trim());
          if (run.fontFaceCs?.trim()) names.add(run.fontFaceCs.trim());
        }
      }
    }
  }
  for (const anchor of worksheet.charts) {
    for (const family of chartFontFamilies(anchor.chart)) names.add(family);
  }
  return [...names];
}
