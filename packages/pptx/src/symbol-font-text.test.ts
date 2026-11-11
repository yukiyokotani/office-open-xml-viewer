import { describe, expect, it } from 'vitest';
import { createSymbolFaceOracle, symbolCodeSet, type SymbolFaceEncoding, type TextRunData } from '@silurus/ooxml-core';
import { renderTextBody, type RenderContext } from './renderer.js';
import type { Paragraph, TextBody } from './types';

// Issue #2177: text in a symbol font is stored as the font's own codes,
// F0-shifted (U+F061) or bare ("a") in a run whose Latin face is the symbol
// font. Each code is judged against the face the render selects: one the
// face draws stays in it; any other code converts to Unicode (α) and is drawn
// in a Unicode face. a:sym (§21.1.2.3.10) owns only the F0-shifted codes.

interface Fill { text: string; font: string }

/** A painting context whose "Symbol" face, when installed, is symbol-encoded
 *  (codes 0x21..0xFE, bare and F0-shifted, one glyph each). */
function context(installed: boolean, fills: Fill[]): CanvasRenderingContext2D {
  let font = '20px serif';
  return {
    get font() { return font; }, set font(v: string) { font = v; },
    fillStyle: '', direction: 'ltr', letterSpacing: '0px',
    measureText: (s: string) => {
      const cp = s.codePointAt(0) ?? 0;
      const code = cp >= 0xf000 ? cp - 0xf000 : cp;
      const owned = installed && font.includes('"Symbol"') && code >= 0x21 && code <= 0xfe;
      return {
        width: s.length === 1 ? (owned ? 100 + code : cp >= 0xf000 ? 7 : 5 + (cp % 5)) : [...s].length * 20,
        actualBoundingBoxAscent: 16, actualBoundingBoxDescent: 4,
        fontBoundingBoxAscent: 16, fontBoundingBoxDescent: 4,
      } as TextMetrics;
    },
    fillText: (t: string) => fills.push({ text: t, font }),
    strokeText: () => {}, fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    setLineDash: () => {}, closePath: () => {}, arc: () => {},
    strokeStyle: '#000', lineWidth: 1, lineJoin: 'miter',
  } as unknown as CanvasRenderingContext2D;
}

function render(runs: TextRunData[], options: Readonly<{
  installed?: boolean; encodings?: ReadonlyMap<string, SymbolFaceEncoding>; bullet?: Paragraph['bullet'];
}> = {}): Fill[] {
  const fills: Fill[] = [];
  const ctx = context(options.installed ?? false, fills);
  const para = {
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: options.bullet ?? { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops: [], rtl: false, runs,
  } as unknown as Paragraph;
  const body = {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: 20,
    defaultBold: null, defaultItalic: null, lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert: 'horz', autoFit: 'none',
  } as unknown as TextBody;
  const rc: RenderContext = {
    themeMajorFont: null, themeMinorFont: null, dpr: 1,
    symbolFaces: createSymbolFaceOracle(ctx),
    ...(options.encodings ? { symbolFaceEncodings: options.encodings } : {}),
  };
  renderTextBody(ctx, body, 0, 0, 2000, 400, 1 / 12700, null, 0, false, false, '#000000', undefined, rc, () => {});
  return fills;
}

function run(text: string, fontFamily: string, fontFamilySym?: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', fontFamily,
    ...(fontFamilySym ? { fontFamilySym } : {}),
  } as TextRunData;
}

const painted = (fills: Fill[]) => fills.map((fill) => fill.text).join('');

describe('symbol-encoded PPTX runs (#2177)', () => {
  it('converts F0-shifted and bare codes the selected face lacks, in a Unicode face', () => {
    const fills = render([run('\uF061\uF062\uF063\uF070\uF0E5\uF0AE', 'Symbol'), run('abcp 1', 'Symbol')]);
    expect(painted(fills)).toBe('αβχπ∑→αβχπ 1');
    expect(fills.filter((fill) => fill.text.includes('α')).every((fill) => !fill.font.includes('Symbol'))).toBe(true);
  });

  it('keeps the codes a symbol-encoded host face draws, in that face', () => {
    const fills = render([run('a\uF062', 'Symbol')], { installed: true });
    expect(painted(fills)).toBe('a\uF062');
    expect(fills.every((fill) => fill.font.includes('Symbol'))).toBe(true);
  });

  it('prefers the cmap of held bytes over measurement', () => {
    // The measured host face owns everything, but the held bytes of the face
    // that paints map only 0x61: 0x62 converts.
    const held = new Map([['symbol|400|normal', { kind: 'known', pua: symbolCodeSet([0x61]), bare: symbolCodeSet([0x61]) } as const]]);
    expect(painted(render([run('\uF061\uF062', 'Symbol')], { installed: true, encodings: held }))).toBe('\uF061β');
  });

  it('leaves ASCII alone when only a:sym names Symbol', () => {
    expect(painted(render([run('ab\uF061', 'Arial', 'Symbol')]))).toBe('abα');
  });

  it('converts a symbol bullet the face lacks and keeps one it draws (a:buFont)', () => {
    const bullet = { type: 'char', char: '\uF0B7', fontFamily: 'Symbol' } as unknown as Paragraph['bullet'];
    const missing = render([run('x', 'Arial')], { bullet });
    expect(missing.some((fill) => fill.text === '•' && !fill.font.includes('Symbol'))).toBe(true);
    const installed = render([run('x', 'Arial')], { bullet, installed: true });
    expect(installed.some((fill) => fill.text === '\uF0B7' && fill.font.includes('Symbol'))).toBe(true);
  });
});
