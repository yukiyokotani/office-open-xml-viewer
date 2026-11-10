import { describe, expect, it } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { Paragraph, TextBody } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

// Issue #2177: text in the Symbol font is stored in Symbol's own encoding,
// either F0-shifted (U+F061) or as bare ASCII ("a") in a run whose Latin face
// is Symbol. Both must paint the Unicode character (α) in a Unicode face. The
// a:sym face (§21.1.2.3.10) owns only the F0-shifted codes, so ASCII in a run
// whose Latin face is ordinary stays Latin.

interface Fill { text: string; font: string }

function render(runs: TextRunData[]): Fill[] {
  let font = '20px serif';
  const fills: Fill[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    fillStyle: '', direction: 'ltr', letterSpacing: '0px',
    measureText: (s: string) => ({
      width: [...s].length * 20,
      actualBoundingBoxAscent: 16, actualBoundingBoxDescent: 4,
      fontBoundingBoxAscent: 16, fontBoundingBoxDescent: 4,
    }) as TextMetrics,
    fillText: (t: string) => fills.push({ text: t, font }),
    strokeText: () => {}, fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    setLineDash: () => {}, closePath: () => {}, arc: () => {},
    strokeStyle: '#000', lineWidth: 1, lineJoin: 'miter',
  } as unknown as CanvasRenderingContext2D;
  const para = {
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops: [], rtl: false, runs,
  } as unknown as Paragraph;
  const body = {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: 20,
    defaultBold: null, defaultItalic: null, lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert: 'horz', autoFit: 'none',
  } as unknown as TextBody;
  renderTextBody(
    ctx, body, 0, 0, 2000, 400, 1 / 12700,
    null, 0, false, false, '#000000', undefined,
    { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
    () => {},
  );
  return fills;
}

function run(text: string, fontFamily: string, fontFamilySym?: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', fontFamily,
    ...(fontFamilySym ? { fontFamilySym } : {}),
  } as TextRunData;
}

describe('Symbol- and dingbat-encoded PPTX runs (#2177)', () => {
  it('maps F0-shifted Greek, operators and arrows in a Symbol run', () => {
    const fills = render([run('\uF061\uF062\uF063\uF070\uF0E5\uF0AE', 'Symbol')]);
    expect(fills.map((f) => f.text).join('')).toBe('αβχπ∑→');
    expect(fills.every((f) => !f.font.includes('Symbol'))).toBe(true);
  });

  it('maps bare ASCII in a Symbol run and keeps identity codes in the run face', () => {
    const fills = render([run('abcp 1', 'Symbol')]);
    expect(fills.map((f) => f.text).join('')).toBe('αβχπ 1');
    const greek = fills.find((f) => f.text.includes('α'));
    expect(greek?.font).not.toContain('Symbol');
  });

  it('maps bare and F0-shifted codes in Wingdings and Webdings runs', () => {
    expect(render([run('J\uF0FC', 'Wingdings')]).map((f) => f.text).join('')).toBe('☺✓');
    expect(render([run('\uF021', 'Webdings')]).map((f) => f.text).join('')).toBe('\u{1F577}');
  });

  it('leaves ASCII alone when only a:sym names Symbol', () => {
    const fills = render([run('ab\uF061', 'Arial', 'Symbol')]);
    expect(fills.map((f) => f.text).join('')).toBe('abα');
  });
});
