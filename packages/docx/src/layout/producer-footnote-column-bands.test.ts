import { describe, expect, it } from 'vitest';
import { wordPerColumnFootnoteBandsApply } from './body-pagination-compatibility.js';

describe('producer footnote column bands', () => {
  it('gives every column its band for a producer that lays notes out per column', () => {
    expect(wordPerColumnFootnoteBandsApply(undefined, 2, true)).toBe(true);
    expect(wordPerColumnFootnoteBandsApply(15, 3, true)).toBe(true);
    expect(wordPerColumnFootnoteBandsApply(undefined, 1, true)).toBe(false);
  });

  it('keeps the Word compatibility-mode gate for DOCX documents', () => {
    expect(wordPerColumnFootnoteBandsApply(undefined, 2)).toBe(false);
    expect(wordPerColumnFootnoteBandsApply(14, 2)).toBe(true);
    expect(wordPerColumnFootnoteBandsApply(15, 2)).toBe(false);
  });
});
