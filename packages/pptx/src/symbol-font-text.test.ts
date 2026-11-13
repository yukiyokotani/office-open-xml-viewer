import { describe, expect, it } from 'vitest';
import { createSymbolFaceOracle, type TextRunData } from '@silurus/ooxml-core';
import { renderTextBody, type RenderContext } from './renderer.js';
import type { Paragraph, TextBody } from './types';

// Issue #2177: text in a symbol font is stored as the font's own codes,
// F0-shifted (U+F061) or bare ("a") in a run whose Latin face is the symbol
// font. Each code is judged against the font the render paints with: one it
// draws stays in it; any other code converts to Unicode (α) and is drawn
// in a Unicode face. a:sym (§21.1.2.3.10) owns only the F0-shifted codes.

interface Fill { text: string; font: string }

/** A painting context whose "Symbol" face, when installed, is symbol-encoded
 *  (codes 0x21..0xFE, bare and F0-shifted, one glyph each). */
function context(installed: boolean | string, fills: Fill[]): CanvasRenderingContext2D {
  const face = `"${typeof installed === 'string' ? installed : 'Symbol'}"`;
  let font = '20px serif';
  return {
    get font() { return font; }, set font(v: string) { font = v; },
    fillStyle: '', direction: 'ltr', letterSpacing: '0px',
    measureText: (s: string) => {
      const cp = s.codePointAt(0) ?? 0;
      const code = cp >= 0xf000 ? cp - 0xf000 : cp;
      const owned = installed !== false && font.includes(face) && code >= 0x21 && code <= 0xfe;
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
  installed?: boolean | string; bullet?: Paragraph['bullet'];
  substitute?: string;
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
    ...(options.substitute ? { fontSources: { substitutions: new Map([['symbol', { family: options.substitute }]]), families: new Map() } } : {}),
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
  it('converts F0-shifted and bare codes the painted font lacks, in a Unicode face', () => {
    const fills = render([run('\uF061\uF062\uF063\uF070\uF0E5\uF0AE', 'Symbol'), run('abcp 1', 'Symbol')]);
    expect(painted(fills)).toBe('αβχπ∑→αβχπ 1');
    expect(fills.filter((fill) => fill.text.includes('α')).every((fill) => !fill.font.includes('Symbol'))).toBe(true);
  });

  it('keeps the codes a symbol-encoded host face draws, in that face', () => {
    const fills = render([run('a\uF062', 'Symbol')], { installed: true });
    expect(painted(fills)).toBe('a\uF062');
    expect(fills.every((fill) => fill.font.includes('Symbol'))).toBe(true);
  });

  it('judges the face that paints: a loaded substitute of the authored family', () => {
    // Symbol is missing; the served substitute "Fixture Sub" is symbol-encoded.
    const fills = render([run('\uF061', 'Symbol')], { installed: 'Fixture Sub', substitute: 'Fixture Sub' });
    expect(painted(fills)).toBe('\uF061');
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

  it('paints a kept bare code that bidi mirroring would replace in its F0-shifted form', () => {
    const bullet = { type: 'char', char: '\u00AB', fontFamily: 'Symbol' } as unknown as Paragraph['bullet'];
    const fills = render([run('x\u00AB', 'Symbol')], { bullet, installed: true });
    expect(fills.filter((fill) => fill.text.includes('\uF0AB')).length).toBeGreaterThanOrEqual(2);
    expect(fills.some((fill) => fill.text.includes('\u00AB'))).toBe(false);
  });

  it('keeps a bullet as authored when its font draws only some of its codes', () => {
    const bullet = { type: 'char', char: '\uF0B7\uF0FF', fontFamily: 'Symbol' } as unknown as Paragraph['bullet'];
    const fills = render([run('x', 'Arial')], { bullet, installed: true });
    expect(fills.some((fill) => fill.text === '\uF0B7\uF0FF' && fill.font.includes('Symbol'))).toBe(true);
  });

  it('keeps the mirror-safe form of the native codes of a partly native bullet', () => {
    const bullet = { type: 'char', char: '\u00AB\uF0FF', fontFamily: 'Symbol' } as unknown as Paragraph['bullet'];
    const fills = render([run('x', 'Arial')], { bullet, installed: true });
    expect(fills.some((fill) => fill.text === '\uF0AB\uF0FF' && fill.font.includes('Symbol'))).toBe(true);
  });
});
