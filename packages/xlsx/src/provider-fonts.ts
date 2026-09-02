import { chartFontFamilies, providerFontFamily, type FontFamilyRoutes } from '@silurus/ooxml-core';
import type { CellFont, ParsedWorkbook, RunFont, Worksheet } from './types.js';

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

function applyFace(font: CellFont | RunFont | null | undefined, routes: FontFamilyRoutes): void {
  if (!font?.name) return;
  font.providerFamily = providerFontFamily(routes, font.name);
}

export function applyWorkbookFontRoutes(workbook: ParsedWorkbook, routes: FontFamilyRoutes): void {
  workbook.styles.providerFontRoutes = routes;
  for (const font of workbook.styles.fonts) applyFace(font, routes);
  for (const dxf of workbook.styles.dxfs) applyFace(dxf.font, routes);
  for (const item of workbook.sharedStrings ?? []) {
    for (const run of item.runs ?? []) applyFace(run.font, routes);
  }
}

export function applyWorksheetFontRoutes(worksheet: Worksheet, routes: FontFamilyRoutes): void {
  worksheet.providerFontRoutes = routes;
  for (const row of worksheet.rows) {
    for (const cell of row.cells) {
      if (cell.value.type !== 'text') continue;
      for (const run of cell.value.runs ?? []) applyFace(run.font, routes);
    }
  }
}
