import { describe, it, expect } from 'vitest';
import { cssTailFor, fontStackFor } from './renderer.js';
import { xlsxOfficeFontRequests, xlsxWorksheetOfficeFontRequests } from './google-fonts.js';
import type { ParsedWorkbook, Worksheet } from './types.js';

describe('XLSX exact Office face requests', () => {
  it('retains only Calibri tuples used in styled fonts and rich text', () => {
    const workbook = { styles: { fonts: [
      { name: null, bold: false, italic: false },
      { name: 'Arial', bold: true, italic: false },
      { name: 'Calibri', bold: true, italic: true },
    ] }, sharedStrings: [{ runs: [
      { text: 'x', font: { name: 'Calibri', bold: false, italic: true } },
    ] }] } as unknown as ParsedWorkbook;
    expect(xlsxOfficeFontRequests(workbook)).toEqual([
      { family: 'Calibri', weight: 400, style: 'normal' },
      { family: 'Calibri', weight: 700, style: 'italic' },
      { family: 'Calibri', weight: 400, style: 'italic' },
    ]);
  });

  it('finds a styled inline run only after its worksheet is pulled', () => {
    const worksheet = { rows: [{ cells: [
      { value: { type: 'text', text: 'x', runs: [
        { text: 'x', font: { name: 'Calibri', bold: true, italic: false } },
        { text: 'y', font: { name: 'Arial', bold: false, italic: true } },
      ] } },
    ] }], shapeGroups: [{ shapes: [{ text: { paragraphs: [{ runs: [
      { type: 'text', text: 'z', fontFace: 'Calibri', bold: false, italic: true },
    ] }] } }] }] } as unknown as Worksheet;
    expect(xlsxWorksheetOfficeFontRequests(worksheet)).toEqual([
      { family: 'Calibri', weight: 700, style: 'normal' },
      { family: 'Calibri', weight: 400, style: 'italic' },
    ]);
  });
});

describe('fontStackFor — default Latin chain (regression)', () => {
  it('uses a retained tuple alias only for the requested Calibri face', () => {
    const route = {
      requestedFamily: 'Calibri' as const, family: '__pinned_regular',
      source: 'local' as const, resourceIdentity: 'office-local:calibri:test',
      weight: 400 as const, style: 'normal' as const,
      metric: { family: '__pinned_regular' },
    };
    expect(fontStackFor('Calibri', undefined, 'text', route).startsWith('"__pinned_regular"')).toBe(true);
    expect(fontStackFor(null, undefined, 'text', route).startsWith('"__pinned_regular"')).toBe(true);
    expect(fontStackFor('Arial', undefined, 'text', route)).not.toContain('__pinned_regular');
  });
  it('does not substitute Carlito for Calibri by default', () => {
    const stack = fontStackFor(null);
    expect(stack.startsWith('"Calibri", Arial')).toBe(true);
    expect(stack).not.toContain('"Carlito"');
    expect(stack).not.toContain('"Caladea"');
    // Only script-appropriate fallbacks precede the generic.
    expect(stack).not.toContain('"Noto Naskh Arabic"');
    expect(stack).not.toContain('"Noto Sans Arabic"');
    expect(stack).toContain('"Noto Sans Hebrew"');
    expect(stack).toContain('"Noto Sans Thai"');
    expect(stack).toContain('"Noto Sans Devanagari"');
    expect(stack.endsWith('sans-serif')).toBe(true);
  });

  it('leads with a named Latin face, then its generic fallback', () => {
    const stack = fontStackFor('Arial');
    expect(stack.startsWith('"Arial", Arial, Helvetica')).toBe(true);
    expect(stack).not.toContain('"Carlito"');
    expect(stack).not.toContain('"Caladea"');
  });

  it('keeps same-class generics by default and opts into published web aliases', () => {
    expect(fontStackFor('Calibri').startsWith('"Calibri", Arial')).toBe(true);
    expect(fontStackFor('Calibri')).not.toContain('"Carlito"');
    expect(fontStackFor('Cambria').startsWith('"Cambria", "Times New Roman"')).toBe(true);
    expect(fontStackFor('Cambria')).not.toContain('"Caladea"');
    expect(fontStackFor('Calibri', undefined, 'text', undefined, true))
      .toContain('"Carlito"');
    expect(fontStackFor('Cambria', undefined, 'text', undefined, true))
      .toContain('"Caladea"');
    expect(fontStackFor('Calibri Light')).not.toContain('"Carlito"');
    expect(fontStackFor('Cambria Math')).not.toContain('"Caladea"');
  });

  it('adds Arabic support only for Arabic text without importing a serif face into a sans tail', () => {
    expect(fontStackFor('Arial', undefined, 'Latin')).not.toContain('Noto Sans Arabic');
    const sans = fontStackFor('Arial', undefined, 'العربية');
    expect(sans).toContain('"Noto Sans Arabic"');
    expect(sans).not.toContain('"Noto Naskh Arabic"');
    const serif = fontStackFor('Cambria', undefined, 'العربية');
    expect(serif).toContain('"Noto Naskh Arabic"');
    expect(serif.endsWith('serif')).toBe(true);
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
    // Latin glyphs fall back by generic class, not to an unrelated Office face.
    expect(tail).toContain('Arial');
    expect(tail).not.toContain('"Carlito"');
    expect(tail).not.toContain('"Caladea"');
  });

  it('non-CJK sans named face uses a generic sans chain', () => {
    expect(cssTailFor('Arial').endsWith('sans-serif')).toBe(true);
    expect(cssTailFor('Arial')).not.toContain('"Carlito"');
  });
});

describe('cssTailFor / fontStackFor — Latin serif & mono (bug fix)', () => {
  it('a Latin serif the host lacks (Century) degrades to a serif, not sans', () => {
    const tail = cssTailFor('Century');
    expect(tail.endsWith('serif')).toBe(true);
    expect(tail.endsWith('sans-serif')).toBe(false);
    expect(tail).toContain('"Times New Roman"');
    expect(tail).not.toContain('"Caladea"');
    // Pure Latin serif — no CJK Noto face should lead the chain.
    expect(tail.startsWith('"Noto Serif')).toBe(false);
    expect(tail.startsWith('"Noto Sans')).toBe(false);
  });

  it('fontStackFor leads with the named serif face then a generic serif tail', () => {
    expect(fontStackFor('Century')).toBe(`"Century", ${cssTailFor('Century')}`);
    expect(fontStackFor('Century').startsWith('"Century", "Times New Roman"')).toBe(true);
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
    expect(fontStackFor('Arial').startsWith('"Arial", Arial, Helvetica')).toBe(true);
    // "Century Gothic" is SANS despite the "century" token — must not regress to
    // the serif default just because the serif Century family does.
    expect(cssTailFor('Century Gothic').endsWith('sans-serif')).toBe(true);
    expect(cssTailFor('Century Gothic')).toBe(cssTailFor('Arial'));
  });

  it('regression: CJK serif/sans ordering unchanged', () => {
    expect(cssTailFor('SimSun').startsWith('"Noto Serif SC"')).toBe(true);
    expect(cssTailFor('SimSun').endsWith('serif')).toBe(true);
    expect(cssTailFor('Microsoft YaHei').startsWith('"Noto Sans SC"')).toBe(true);
    expect(cssTailFor('Microsoft YaHei').endsWith('sans-serif')).toBe(true);
  });
});


it('uses per-workbook fallback only for Han while preserving authored regions', () => {
  for (const family of [null, 'Calibri', 'Arial']) {
    const sc = fontStackFor(family, 'sc', '漢');
    expect(sc).not.toContain('"Carlito"');
    if (family !== 'Arial') {
      const optedIn = fontStackFor(family, 'sc', '漢', undefined, true);
      expect(optedIn.indexOf('"Carlito"')).toBeLessThan(optedIn.indexOf('"Noto Sans SC"'));
    }
    expect(sc.indexOf('"Noto Sans SC"')).toBeLessThan(sc.indexOf('"Noto Sans JP"'));
    const tc = fontStackFor(family, 'tc', '漢');
    expect(tc.indexOf('"Noto Sans TC"')).toBeLessThan(tc.indexOf('"Noto Sans SC"'));
    expect(fontStackFor(family, 'sc', '→')).not.toContain('Noto Sans SC');
  }
  const jp = fontStackFor('Meiryo', 'sc', '→');
  expect(jp.indexOf('"Noto Sans JP"')).toBeLessThan(jp.indexOf('"Noto Sans SC"'));
});
