import { xlsxFontPreloadNames, xlsxCjkFallback, xlsxWorksheetOfficeFontRequests } from './google-fonts.js';
import type { ParsedWorkbook, Worksheet } from './types.js';

it('preflights the Normal font even when no text cell uses it', () => {
  const worksheet = {
    defaultFontFamily: 'Arial', defaultFontBold: true,
    defaultFontItalic: true, rows: [], shapeGroups: [],
  } as unknown as Worksheet;
  expect(xlsxWorksheetOfficeFontRequests(worksheet)).toEqual([
    { family: 'Arial', weight: 700, style: 'italic' },
  ]);
});
import { describe, expect, it } from 'vitest';
import type { FontPreloadEntry } from '@silurus/ooxml-core';
import { XLSX_GOOGLE_FONTS } from './google-fonts.js';

// Verbatim snapshot of the XLSX Office-font substitute map BEFORE the shared
// registry consolidation (Phase 3 C7), excluding the SCRIPT_GOOGLE_FONTS spread
// (unchanged, shared already). Frozen as the oracle so the consolidated map's
// effective entries can only ADD keys, never drop or alter one. This was the
// smallest of the three maps (Calibri/Cambria + Arabic only).
const XLSX_GOOGLE_FONTS_OLD: Record<string, FontPreloadEntry> = {
  'calibri': {
    url: 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap',
    loadFamily: 'Carlito',
  },
  'cambria': {
    url: 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap',
    loadFamily: 'Caladea',
  },
  'sakkal majalla': { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'traditional arabic': { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'simplified arabic': { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'arabic typesetting': { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'univers next arabic': { url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Sans Arabic' },
  'noto naskh arabic': { url: 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Naskh Arabic' },
  'noto sans arabic': { url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&display=swap', loadFamily: 'Noto Sans Arabic' },
};

// Generic web fonts + Office face names the shared registry now contributes to
// XLSX (previously only in docx/pptx). Each is either a plain Google web font
// served under its own family name, or an Office face reducing to a metric
// substitute already present. Calibri Light and Cambria Math have distinct
// capabilities, so neither inherits the base text-face substitution.
const EXPECTED_ADDED = new Set([
  'franklin gothic book',
  'franklin gothic medium',
  'nunito sans',
  'nunito',
  'open sans',
  'roboto',
  'lato',
  'montserrat',
  'poppins',
  'raleway',
  'playfair display',
  'ubuntu',
]);

describe('XLSX_GOOGLE_FONTS — shared registry consolidation (oracle)', () => {
  it('preserves every pre-consolidation entry byte-for-byte', () => {
    for (const [key, entry] of Object.entries(XLSX_GOOGLE_FONTS_OLD)) {
      expect(XLSX_GOOGLE_FONTS[key], `entry "${key}"`).toEqual(entry);
    }
  });

  it('adds only the safe, documented web-font / Office-face keys', () => {
    const oldKeys = new Set(Object.keys(XLSX_GOOGLE_FONTS_OLD));
    const added = Object.keys(XLSX_GOOGLE_FONTS).filter(
      (k) => !oldKeys.has(k) && !k.startsWith('noto '),
    );
    expect(new Set(added)).toEqual(EXPECTED_ADDED);
    expect(XLSX_GOOGLE_FONTS['calibri light']).toBeUndefined();
    expect(XLSX_GOOGLE_FONTS['cambria math']).toBeUndefined();
    expect(XLSX_GOOGLE_FONTS['franklin gothic medium']).toMatchObject({
      loadFamily: 'Libre Franklin',
    });
  });
});


it('matches workbook preload and render preferences without fetching fonts for Latin-only text', () => {
  const wb = { styles: { fonts: [{ name: 'Calibri' }] }, sharedStrings: [{ text: '漢字' }] } as ParsedWorkbook;
  expect(xlsxCjkFallback(wb, 'sc')).toBe('sc');
  expect(xlsxFontPreloadNames(wb, 'sc')).toContain('Noto Sans SC');
  expect(xlsxFontPreloadNames(wb, 'sc')).not.toContain('Noto Sans JP');
  wb.styles.fonts[0].name = 'Meiryo';
  expect(xlsxCjkFallback(wb, 'sc')).toBe('jp');
  expect(xlsxFontPreloadNames(wb, 'sc')).toContain('Noto Sans JP');
  wb.sharedStrings = [{ text: 'Hello' }];
  expect([...xlsxFontPreloadNames(wb, 'sc')]).toEqual(['Meiryo']);
});


it('uses the same strong script evidence for workbook rendering and preload', () => {
  const wb = { styles: { fonts: [{ name: 'Calibri' }] }, sharedStrings: [{ text: '漢字かな' }] } as ParsedWorkbook;
  expect(xlsxCjkFallback(wb, 'sc')).toBe('jp');
  expect(xlsxFontPreloadNames(wb, 'sc')).toContain('Noto Sans JP');
  expect(xlsxFontPreloadNames(wb, 'sc')).not.toContain('Noto Sans SC');
});
