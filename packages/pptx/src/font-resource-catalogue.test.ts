import { describe, expect, it } from 'vitest';
import { findReferenceFontMetrics } from '@silurus/ooxml-core';
import { referenceFontCoversSymbol, referenceFontCoversCjk } from './font-resource-catalogue.js';

describe('per-cut resource catalogue', () => {
  it('distinguishes cmap presence, absence and unrecorded symbols in a concrete cut', () => {
    const tahoma = findReferenceFontMetrics('Tahoma', { source: 'office-mac', weight: 400, style: 'normal' })[0];
    const extB = findReferenceFontMetrics('SimSun-ExtB', { source: 'office-mac', weight: 400, style: 'normal' })[0];
    expect(referenceFontCoversSymbol(tahoma, 0x25A0)).toBe(true);
    expect(referenceFontCoversSymbol(tahoma, 0x25C6)).toBe(false);
    expect(referenceFontCoversSymbol(extB, 0x00A7)).toBe(false);
    expect(referenceFontCoversSymbol(tahoma, 0x4E00)).toBeUndefined();
    expect(referenceFontCoversSymbol(tahoma, 0x25A0 + 0.5)).toBeUndefined();
    const open = findReferenceFontMetrics('BIZ UDMincho', { weight: 400, style: 'normal' })[0];
    expect(referenceFontCoversSymbol(open, 0x25A0)).toBeUndefined();
  });

  it('keeps selected-cut CJK presence separate from family-wide Han coverage', () => {
    const deng = findReferenceFontMetrics('DengXian', { source: 'office-mac', weight: 400, style: 'normal' })[0];
    expect(referenceFontCoversCjk(deng, 0x6f22)).toBe(true);
    expect(referenceFontCoversCjk(deng, 0xd55c)).toBe(false);
    expect(referenceFontCoversCjk(deng, 0x41)).toBeUndefined();
    const open = findReferenceFontMetrics('BIZ UDMincho', { weight: 400, style: 'normal' })[0];
    expect(referenceFontCoversCjk(open, 0x6f22)).toBeUndefined();
  });

});

it('never lends a generated cut certificate to an equal-looking foreign profile', () => {
  const tahoma = findReferenceFontMetrics('Tahoma', {source: 'office-mac', weight: 400, style: 'normal'})[0];
  expect(referenceFontCoversSymbol(tahoma, 0x25a0)).toBe(true);
  expect(referenceFontCoversSymbol({...tahoma}, 0x25a0)).toBeUndefined();
});
