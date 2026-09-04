import { describe, it, expect } from 'vitest';
import { cssTailFor, fontStackFor } from './renderer.js';

describe('fontStackFor — default Latin chain (regression)', () => {
  it('returns the Calibri/Carlito default chain for an unnamed cell', () => {
    const stack = fontStackFor(null);
    expect(stack.startsWith('"Calibri", "Carlito"')).toBe(true);
    // Arabic + non-CJK script fallbacks retained.
    expect(stack).toContain('"Noto Naskh Arabic"');
    expect(stack).toContain('"Noto Sans Hebrew"');
    expect(stack).toContain('"Noto Sans Thai"');
    expect(stack).toContain('"Noto Sans Devanagari"');
    expect(stack.endsWith('sans-serif')).toBe(true);
  });

  it('leads with the named Latin face, then the default chain', () => {
    const stack = fontStackFor('Arial');
    expect(stack.startsWith('"Arial", "Calibri", "Carlito"')).toBe(true);
  });
});

describe('fontStackFor — CJK language-specific Noto ordering', () => {
  it('routes a native Noto CJK name through the loaded Google family', () => {
    const sans = fontStackFor(' Noto Sans CJK JP ');
    expect(sans.startsWith('"Noto Sans CJK JP", "Noto Sans JP", ')).toBe(true);
    expect(sans.match(/"Noto Sans JP"/g)).toHaveLength(1);

    const serif = fontStackFor('Noto Serif CJK KR');
    expect(serif.startsWith('"Noto Serif CJK KR", "Noto Serif KR", ')).toBe(true);
    expect(serif.endsWith('serif')).toBe(true);
  });

  it('routes HK sans through Google Fonts without inventing an HK serif alias', () => {
    expect(fontStackFor('Noto Sans CJK HK').startsWith(
      '"Noto Sans CJK HK", "Noto Sans HK", ',
    )).toBe(true);

    const serif = fontStackFor('Noto Serif CJK HK');
    expect(serif).not.toContain('"Noto Serif HK"');
    expect(serif).not.toContain('"Noto Serif TC"');
    expect(serif).not.toMatch(/,\s*,/);
  });

  it('Korean sans (Malgun Gothic) → Noto Sans KR leads the tail', () => {
    const tail = cssTailFor('Malgun Gothic');
    expect(tail.startsWith('"Noto Sans KR"')).toBe(true);
    expect(tail.indexOf('Noto Sans KR')).toBeLessThan(tail.indexOf('Noto Sans JP'));
    expect(tail.endsWith('sans-serif')).toBe(true);
  });

  it('Simplified Chinese serif (SimSun) → Noto Serif SC leads', () => {
    const tail = cssTailFor('SimSun');
    expect(tail.startsWith('"Noto Serif SC"')).toBe(true);
    expect(tail.endsWith('serif')).toBe(true);
  });

  it('Simplified Chinese sans (Microsoft YaHei) → Noto Sans SC leads', () => {
    expect(cssTailFor('Microsoft YaHei').startsWith('"Noto Sans SC"')).toBe(true);
  });

  it('Traditional Chinese (PMingLiU serif, JhengHei sans)', () => {
    expect(cssTailFor('PMingLiU').startsWith('"Noto Serif TC"')).toBe(true);
    expect(cssTailFor('Microsoft JhengHei').startsWith('"Noto Sans TC"')).toBe(true);
  });

  it('Japanese faces lead with Noto Sans JP (xlsx previously had no CJK fallback)', () => {
    const tail = cssTailFor('Meiryo');
    expect(tail.startsWith('"Noto Sans JP"')).toBe(true);
    // Still keeps the Latin metric substitutes after the CJK face.
    expect(tail).toContain('"Calibri"');
  });

  it('non-CJK SANS named face falls back to the default sans chain', () => {
    expect(cssTailFor('Arial')).toBe(fontStackFor(null));
  });
});

describe('cssTailFor / fontStackFor — Latin serif & mono (bug fix)', () => {
  it('a Latin serif the host lacks (Century) degrades to a serif, not sans', () => {
    const tail = cssTailFor('Century');
    expect(tail.endsWith('serif')).toBe(true);
    expect(tail.endsWith('sans-serif')).toBe(false);
    // Office serif (Cambria) and its metric clone (Caladea) lead the serif default.
    expect(tail).toContain('"Cambria"');
    expect(tail).toContain('"Caladea"');
    // Pure Latin serif — no CJK Noto face should lead the chain.
    expect(tail.startsWith('"Noto Serif')).toBe(false);
    expect(tail.startsWith('"Noto Sans')).toBe(false);
  });

  it('fontStackFor leads with the named serif face then the serif default', () => {
    expect(fontStackFor('Century')).toBe(`"Century", ${cssTailFor('Century')}`);
    expect(fontStackFor('Century').startsWith('"Century", "Cambria"')).toBe(true);
  });

  it('other Latin serifs (Garamond, Times New Roman) also end in serif', () => {
    expect(cssTailFor('Garamond').endsWith('serif')).toBe(true);
    expect(cssTailFor('Times New Roman').endsWith('serif')).toBe(true);
  });

  it('a monospaced face the host lacks (Consolas) degrades to monospace', () => {
    expect(cssTailFor('Consolas').endsWith('monospace')).toBe(true);
  });

  it('regression: a Latin sans face / unnamed cell still ends in sans-serif', () => {
    expect(cssTailFor('Arial').endsWith('sans-serif')).toBe(true);
    expect(fontStackFor(null).endsWith('sans-serif')).toBe(true);
    expect(fontStackFor('Arial').startsWith('"Arial", "Calibri", "Carlito"')).toBe(true);
    // "Century Gothic" is SANS despite the "century" token — must not regress to
    // the serif default just because the serif Century family does.
    expect(cssTailFor('Century Gothic').endsWith('sans-serif')).toBe(true);
    expect(cssTailFor('Century Gothic')).toBe(fontStackFor(null));
  });

  it('regression: CJK serif/sans ordering unchanged', () => {
    expect(cssTailFor('SimSun').startsWith('"Noto Serif SC"')).toBe(true);
    expect(cssTailFor('SimSun').endsWith('serif')).toBe(true);
    expect(cssTailFor('Microsoft YaHei').startsWith('"Noto Sans SC"')).toBe(true);
    expect(cssTailFor('Microsoft YaHei').endsWith('sans-serif')).toBe(true);
  });
});
