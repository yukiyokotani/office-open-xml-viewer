import { describe, expect, it } from 'vitest';
import { createSymbolFaceOracle } from '@silurus/ooxml-core';
import { loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
const { Canvas } = (skia ?? {}) as typeof import('skia-canvas');

// A Node canvas has no FontFaceSet, so no route sentinel can be registered:
// the symbol oracle (core symbol-face.ts) judges host fonts in its plain mode.
// Host fonts differ per machine, so only a missing family and the structural
// invariant of whatever Symbol/Wingdings the host has are asserted.
describe.skipIf(!skia)('symbol face oracle on a Node canvas (#2177)', () => {
  const context = () => new Canvas(16, 16).getContext('2d') as unknown as Parameters<typeof createSymbolFaceOracle>[0];
  const absent = '32px "Ooxml Absent Symbol Face 7", sans-serif';

  it('converts every code of a missing symbol font, bare and F0-shifted', () => {
    const oracle = createSymbolFaceOracle(context());
    expect([oracle.draws('a', absent), oracle.draws('', absent)]).toEqual([false, false]);
    expect(oracle.segments('a', 'Symbol', absent)).toEqual([{ text: 'αβα', mapped: true }]);
    expect(oracle.segments('', 'Wingdings 3', absent).map((part) => part.mapped)).toEqual([true]);
  });

  it.each(['Symbol', 'Wingdings'])('keeps a code of the host %s only when it proves both forms draw one glyph', (family) => {
    const oracle = createSymbolFaceOracle(context());
    const font = `32px "${family}", sans-serif`;
    let pua = 0;
    let bare = 0;
    for (let code = 0x21; code <= 0xfe; code++) {
      if (code === 0x7f || code === 0xad) continue;
      const verdicts = [oracle.draws(String.fromCharCode(0xf000 + code), font), oracle.draws(String.fromCharCode(code), font)];
      expect(verdicts.includes(undefined)).toBe(false);
      // The two forms stand or fall together here.
      expect(verdicts[0]).toBe(verdicts[1]);
      if (verdicts[0]) pua++;
      if (verdicts[1]) bare++;
    }
    // Missing or Unicode-encoded (nothing native), or a symbol-encoded face.
    expect(pua === 0 || pua > 180).toBe(true);
    expect(bare).toBe(pua);
    // A bare default-ignorable code is never native.
    expect(oracle.draws('­', font)).toBe(false);
  });
});
