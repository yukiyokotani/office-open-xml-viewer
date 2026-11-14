import { describe, expect, it } from 'vitest';
import { wordPerColumnFootnoteBandsApply } from './body-pagination-compatibility.js';
import { buildDocxCompatProfile } from './compat-profile.js';

const bands = (mode: number | undefined, columns: number, producer?: boolean) =>
  wordPerColumnFootnoteBandsApply(
    buildDocxCompatProfile({ compatibilityMode: mode }, { footnoteColumnBands: producer }),
    columns,
  );

describe('producer footnote column bands', () => {
  it('gives every column its band for a producer that lays notes out per column', () => {
    expect(bands(undefined, 2, true)).toBe(true);
    expect(bands(15, 3, true)).toBe(true);
    expect(bands(undefined, 1, true)).toBe(false);
  });

  it('keeps the Word compatibility-mode gate for DOCX documents', () => {
    expect(bands(undefined, 2)).toBe(false);
    expect(bands(14, 2)).toBe(true);
    expect(bands(15, 2)).toBe(false);
  });
});
