/// <reference types="node" />
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { excelDateTimeSectionNameTexts } from '@silurus/ooxml-core';
import { CHART_SYNTHESIZED_TEXT, chartPaintTexts } from './chart-font-demand.js';

/** Every name a date section can paint: all of the formatter's tables. */
const DATE_NAMES = excelDateTimeSectionNameTexts('[$-411]ggg aaaa mmm ddd AM/PM [QQ]');
import { collectSlideFontDemand } from './font-preload-demand.js';
import type { Slide } from './types';

const chartSlide = (chart: Record<string, unknown>): Slide => ({
  index: 0, slideNumber: 1, background: null,
  elements: [{ type: 'chart', x: 0, y: 0, width: 1, height: 1, rotation: 0, flipH: false, flipV: false, chart }],
} as unknown as Slide);

const scalars = (text: string) => [...text].map((ch) => ch.codePointAt(0)!);

afterEach(() => { vi.unstubAllGlobals(); });

describe('chart paint texts for font demand', () => {
  it('demands authored chart strings and generated ASCII, but not face names', () => {
    const demand = collectSlideFontDemand(chartSlide({
      chartType: 'bar',
      title: '売上',
      titleFontFace: '游ゴシック',
      categories: ['東京', 'Øsaka'],
      series: [{ name: 'Δ series', values: [1, 2], valFormatCode: '#,##0"円"' }],
    }), new Set(), Number.MAX_SAFE_INTEGER);
    expect(demand).not.toBe('all');
    expect(demand).toEqual(expect.arrayContaining(scalars('売上東京ØΔ円' + CHART_SYNTHESIZED_TEXT)));
    expect(demand).not.toContain('游'.codePointAt(0));
    // Elision measures lone-surrogate prefixes as U+FFFD.
    expect(demand).toContain(0xfffd);
  });

  it('adds the formatter name tables for every date section, as the formatter detects it', () => {
    const texts = (code: string) => chartPaintTexts({ chartType: 'line', valAxisFormatCode: code, series: [] }) as string[];
    for (const code of ['General', '0.00E+00', '#,##0;[Red]-#,##0', '"Items" 0', '0%', '\\d 0']) {
      expect(texts(code), code).not.toContain('令和');
      expect(texts(code), code).not.toContain('January');
    }
    expect(texts('[$-411]rr')).toEqual(expect.arrayContaining(['令和']));
    expect(texts('[$-411]ggge"年"')).toEqual(expect.arrayContaining(['令和', '平']));
    expect(texts('[$-411][QQ]')).toEqual(expect.arrayContaining(['第１四半期']));
    expect(texts('aaa')).toEqual(expect.arrayContaining(['水']));
    expect(texts('0;[$-411]mmmm')).toEqual(expect.arrayContaining(['January']));
    // A format code under any key, including per-point category arrays.
    expect(chartPaintTexts({ chartType: 'line', series: [{ catFormatCodes: [null, '[$-411]gge'] }] }))
      .toEqual(expect.arrayContaining(['令和']));
  });

  it('declines region maps, metafile or unknown image fills and incomplete locale short dates', () => {
    expect(chartPaintTexts({ chartType: 'regionMap', series: [] })).toBe('all');
    expect(chartPaintTexts({ chartType: 'bar', series: [{ fill: { fillType: 'image', imagePath: 'ppt/media/i.emf', mimeType: 'image/emf' } }] })).toBe('all');
    expect(chartPaintTexts({ chartType: 'bar', series: [{ fill: { fillType: 'image', imagePath: 'ppt/media/i.bin', mimeType: 'application/octet-stream' } }] })).toBe('all');
    expect(chartPaintTexts({ chartType: 'bar', series: [{ fill: { fillType: 'image', imagePath: 'ppt/media/i.png', mimeType: 'image/png' } }] })).not.toBe('all');
    const builtin14 = { chartType: 'bar', dataTable: {}, series: [{ catFormatBuiltinId: 14 }] };
    vi.stubGlobal('navigator', { language: 'en-US' });
    expect(chartPaintTexts(builtin14)).not.toBe('all');
    vi.stubGlobal('navigator', { language: 'ja-JP-u-ca-japanese' });
    expect(chartPaintTexts(builtin14)).toBe('all');
    expect(collectSlideFontDemand(chartSlide(builtin14), new Set(), Number.MAX_SAFE_INTEGER)).toBe('all');
  });

  it('lists every non-ASCII literal the chart painter and its number formatter can write', () => {
    // Guards the synthesized sets against new painter literals: every
    // non-ASCII character in non-comment code of core's chart painter and
    // §18.8.30 formatter must be demanded. The region-map renderer and its
    // geodata are excluded: a region map declines to 'all'.
    const core = fileURLToPath(new URL('../../core/src/', import.meta.url));
    const sources: string[] = ['excel-number-format.ts', 'excel-date.ts'];
    const walk = (dir: string) => {
      for (const entry of readdirSync(join(core, dir), { withFileTypes: true })) {
        const path = join(dir, entry.name);
        if (entry.isDirectory()) walk(path);
        else if (/\.ts$/.test(entry.name) && !/\.(test|stories)\.ts$|^region-map/.test(entry.name)) sources.push(path);
      }
    };
    walk('chart');
    const allowed = new Set(scalars(CHART_SYNTHESIZED_TEXT + DATE_NAMES.join('')));
    const missing = new Set<string>();
    for (const source of sources) {
      const code = readFileSync(join(core, source), 'utf8')
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .split('\n').map((line) => line.replace(/(^|[^:'"`\\])\/\/.*$/, '$1')).join('\n');
      const literals = code.replace(/\\u\{?([0-9a-fA-F]{4,6})\}?/g, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)));
      for (const ch of literals) {
        if (ch.codePointAt(0)! > 0x7e && !allowed.has(ch.codePointAt(0)!)) missing.add(`${source}: ${ch}`);
      }
    }
    expect([...missing]).toEqual([]);
  });
});
