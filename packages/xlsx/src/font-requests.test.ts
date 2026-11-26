import { expect, it } from 'vitest';
import { xlsxCjkFallback, xlsxWorksheetOfficeFontRequests } from './font-requests.js';
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

it('resolves the workbook CJK fallback from style fonts and strong script evidence', () => {
  const wb = { styles: { fonts: [{ name: 'Calibri' }] }, sharedStrings: [{ text: '漢字' }] } as ParsedWorkbook;
  expect(xlsxCjkFallback(wb, 'sc')).toBe('sc');
  wb.styles.fonts[0].name = 'Meiryo';
  expect(xlsxCjkFallback(wb, 'sc')).toBe('jp');
  const kana = { styles: { fonts: [{ name: 'Calibri' }] }, sharedStrings: [{ text: '漢字かな' }] } as ParsedWorkbook;
  expect(xlsxCjkFallback(kana, 'sc')).toBe('jp');
});

it('does not request a DOCX-only catalog face for the workbook default', () => {
  const worksheet = {
    defaultFontFamily: 'Noto Sans JP', defaultFontBold: false,
    defaultFontItalic: false, rows: [], shapeGroups: [],
  } as unknown as Worksheet;
  expect(xlsxWorksheetOfficeFontRequests(worksheet)).toEqual([]);
});
