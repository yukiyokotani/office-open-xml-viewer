import { providerFontFamily, type FontFamilyRoutes } from '@silurus/ooxml-core';
import type { CellFont, ParsedWorkbook, RunFont, Worksheet } from './types.js';

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
