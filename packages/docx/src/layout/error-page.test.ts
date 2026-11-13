import { describe, expect, it } from 'vitest';
import { layoutParseErrorPage } from './error-page.js';
import { createFontResolver } from './font-service.js';
import { createTextLayoutService } from './text.js';
import { paintLayoutPage } from '../paint/canvas-page.js';
import { canvasFontString } from '@silurus/ooxml-core';

describe('parse-error page layout', () => {
  it('wraps detail text during layout and never measures text during paint', async () => {
    const measuredFonts: string[] = [];
    const text = createTextLayoutService({
      fonts: createFontResolver([]),
      measurer: {
        fingerprint: 'fixed-width-v1',
        measure: (request) => {
          measuredFonts.push(canvasFontString(request.fontRoute, request.fontSizePt, request.weight, request.style));
          return {
            advancePt: request.text.length * request.fontSizePt * 0.5,
            ascentPt: request.fontSizePt * 0.8,
            descentPt: request.fontSizePt * 0.2,
          };
        },
      },
    });
    const layout = layoutParseErrorPage(
      'word/document.xml: a deliberately long parse failure that must wrap before painting',
      { widthPt: 300, heightPt: 400 },
      text,
    );
    const textCommands = layout.pages[0]?.layers.body.flatMap((node) =>
      node.kind === 'drawing' ? node.commands.filter((command) => command.kind === 'text') : [],
    ) ?? [];
    const layers = layout.pages[0]!.layers;

    expect(Object.isFrozen(layers)).toBe(true);
    expect(Object.isFrozen(layers.roots)).toBe(true);
    expect(Object.isFrozen(layers.roots[0])).toBe(true);
    expect(layers.roots[0]!.node).toBe(layers.body[0]);
    expect(textCommands.length).toBeGreaterThan(3);
    expect(textCommands.map((command) => command.kind === 'text' ? command.text : '').join(' '))
      .toContain('word/document.xml:');

    const calls: string[] = [];
    const paintedFonts: string[] = [];
    let paintedFont = '';
    const ctx = {
      save: () => calls.push('save'),
      restore: () => calls.push('restore'),
      setTransform: () => calls.push('setTransform'),
      clearRect: () => calls.push('clearRect'),
      fillRect: () => calls.push('fillRect'),
      strokeRect: () => calls.push('strokeRect'),
      setLineDash: () => calls.push('setLineDash'),
      fillText: () => calls.push('fillText'),
      measureText: () => { throw new Error('paint measured text'); },
      fillStyle: '',
      strokeStyle: '',
      lineWidth: 1,
      get font() { return paintedFont; },
      set font(value: string) { paintedFont = value; paintedFonts.push(value); },
      textAlign: 'start',
      textBaseline: 'alphabetic',
    } as unknown as CanvasRenderingContext2D;
    const canvas = {
      width: 0,
      height: 0,
      getContext: () => ctx,
    } as unknown as HTMLCanvasElement;

    await expect(paintLayoutPage(layout, 0, canvas, { scale: 1, dpr: 1 })).resolves.toBeUndefined();
    expect(calls).toContain('fillText');
    expect(measuredFonts.length).toBeGreaterThan(0);
    for (const measured of new Set(measuredFonts)) expect(paintedFonts).toContain(measured);
  });

  it('emergency-splits an overlong token only at grapheme boundaries', () => {
    const text = createTextLayoutService({
      fonts: createFontResolver([]),
      measurer: {
        fingerprint: 'fixed-width-v1',
        measure: (request) => ({
          advancePt: [...request.text].length * 10,
          ascentPt: 8,
          descentPt: 2,
        }),
      },
    });
    const token = 'word/' + 'a\u0301'.repeat(12) + '.xml';
    const layout = layoutParseErrorPage(token, { widthPt: 180, heightPt: 300 }, text);
    const lines = layout.pages[0]?.layers.body.flatMap((node) =>
      node.kind === 'drawing'
        ? node.commands.filter((command) => command.kind === 'text').slice(2)
        : [],
    ) ?? [];

    expect(lines.length).toBeGreaterThan(1);
    expect(lines.map((line) => line.kind === 'text' ? line.text : '').join('')).toBe(token);
    expect(lines.every((line) => line.kind !== 'text' || !line.text.startsWith('\u0301'))).toBe(true);
  });

  it('caps five-plus emergency chunks at maxLines with a grapheme-safe ellipsis', () => {
    const text = createTextLayoutService({
      fonts: createFontResolver([]),
      measurer: {
        fingerprint: 'fixed-graphemes-v1',
        measure: (request) => ({
          advancePt: [...request.text].length * 10,
          ascentPt: 8,
          descentPt: 2,
        }),
      },
    });
    const token = 'a\u0301'.repeat(80);
    const layout = layoutParseErrorPage(token, { widthPt: 180, heightPt: 300 }, text);
    const lines = layout.pages[0]?.layers.body.flatMap((node) =>
      node.kind === 'drawing'
        ? node.commands.filter((command) => command.kind === 'text').slice(2)
        : [],
    ) ?? [];
    const values = lines.map((line) => line.kind === 'text' ? line.text : '');

    expect(values).toHaveLength(4);
    expect(values.at(-1)).toMatch(/…$/u);
    expect(values.every((line) => !line.startsWith('\u0301'))).toBe(true);
  });

  it('ellipsizes a long caller title to one line inside the frame', () => {
    const text = createTextLayoutService({
      fonts: createFontResolver([]),
      measurer: {
        fingerprint: 'fixed-title-v1',
        measure: (request) => ({ advancePt: [...request.text].length * 10, ascentPt: 8, descentPt: 2 }),
      },
    });
    const layout = layoutParseErrorPage('detail', { widthPt: 180, heightPt: 300 }, text, {
      title: 'Title '.repeat(2_000),
      detail: undefined,
    });
    const title = layout.pages[0]?.layers.body.flatMap((node) =>
      node.kind === 'drawing' ? node.commands.filter((command) => command.kind === 'text').slice(1, 2) : [],
    )[0];
    // Frame text width: 180 - 4 * 18 pt.
    expect(title?.kind === 'text' && title.text).toBe('Title…');
  });
});
