import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { Paragraph, TextBody } from './types.js';
import type { TextRunData } from '@silurus/ooxml-core';

const { buildWarpEnvelope } = vi.hoisted(() => ({ buildWarpEnvelope: vi.fn() }));

vi.mock('@silurus/ooxml-core', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@silurus/ooxml-core')>();
  buildWarpEnvelope.mockImplementation((_preset, _adj, width = 620) => ({
    top: [{ x: 0, y: 0 }, { x: width, y: 0 }],
    bottom: [{ x: 0, y: 0 }, { x: width, y: 0 }],
    topLen: [0, width], bottomLen: [0, width], singleEdge: true,
  }));
  return { ...actual, buildWarpEnvelope };
});

import { renderTextBody } from './renderer.js';

function run(text: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 40, color: '000000', fontFamily: 'Arial',
  };
}

function bodyWithEmptyLines(emptyLineCount: number): TextBody {
  const paragraph: Paragraph = {
    alignment: 'ctr', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null,
    defBold: null, defItalic: null, defFontFamily: null, tabStops: [],
    eaLnBrk: true,
    runs: [run('A'), ...Array.from({ length: emptyLineCount }, () => ({ type: 'break' as const }))],
  } as Paragraph;
  return {
    verticalAnchor: 'ctr', paragraphs: [paragraph], defaultFontSize: 40,
    defaultBold: null, defaultItalic: null, lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert: 'horz', autoFit: 'none',
    textWarp: { preset: 'textArchUp', adj: [] },
  } as TextBody;
}

function mockCtx(): CanvasRenderingContext2D {
  return {
    font: '', fillStyle: '', direction: 'ltr', textAlign: 'left', textBaseline: 'alphabetic',
    measureText: (text: string) => ({
      width: text.length * 10, actualBoundingBoxAscent: 7, actualBoundingBoxDescent: 2,
    }) as TextMetrics,
    save() {}, restore() {}, translate() {}, rotate() {}, scale() {}, fillText() {},
  } as unknown as CanvasRenderingContext2D;
}

describe('WordArt rendering resource bounds', () => {
  beforeEach(() => buildWarpEnvelope.mockClear());

  it('skips per-line envelope construction for empty manual lines', () => {
    renderTextBody(mockCtx(), bodyWithEmptyLines(100), 0, 0, 620, 150, 1 / 12700);

    // One base envelope and, at most, one expanded envelope for the painted line.
    expect(buildWarpEnvelope).toHaveBeenCalledTimes(2);
  });
});
