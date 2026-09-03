import { describe, expect, it } from 'vitest';
import {
  xlsxFontPreloadNames,
  xlsxFontProviderNames,
  xlsxWorksheetFontProviderNames,
} from './font-plan.js';
import type { ParsedWorkbook, Worksheet } from './types.js';

describe('XLSX font plan', () => {
  it('adds script fallbacks only to the Google preload plan', () => {
    const workbook = {
      styles: { fonts: [{ name: 'Calibri' }] },
      sharedStrings: [{ text: '日本語' }],
    } as ParsedWorkbook;

    expect(xlsxFontPreloadNames(workbook)).toEqual(new Set([
      'Calibri', 'Noto Sans JP', 'Noto Serif JP',
    ]));
    expect(xlsxFontProviderNames(workbook)).toEqual(['Calibri']);
  });

  it('discovers lazy worksheet text, shape, and chart families', () => {
    const worksheet = {
      rows: [{ cells: [{ value: {
        type: 'text', text: 'Cell', runs: [{ text: 'Cell', font: { name: 'Cell Face' } }],
      } }] }],
      shapeGroups: [{ shapes: [{ text: { paragraphs: [{ runs: [{
        type: 'text', text: 'Shape', fontFace: 'Shape Face',
      }] }] } }] }],
      charts: [{ chart: { titleFontFace: 'Chart Face' } }],
    } as unknown as Worksheet;

    expect(xlsxWorksheetFontProviderNames(worksheet)).toEqual(expect.arrayContaining([
      'Cell Face', 'Shape Face', 'Chart Face',
    ]));
  });
});
