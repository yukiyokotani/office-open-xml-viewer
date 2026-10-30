import { afterEach, describe, expect, it, vi } from 'vitest';
import { formatLocalizedExcelShortDate } from '@silurus/ooxml-core';
import { OFFICE_FONT_SUBSTITUTIONS } from '@silurus/ooxml-core/internal/font-sources';
import { XlsxFontDemand } from './font-source-load.js';
import type { CellFont, ParsedWorkbook, Row, Worksheet } from './types.js';

const font = (name: string | null, extra: Partial<CellFont> = {}): CellFont =>
  ({ name, bold: false, italic: false, underline: false, strike: false, size: 11, color: null, ...extra });

function workbook(overrides: Partial<{ fonts: CellFont[]; sharedStrings: ParsedWorkbook['sharedStrings']; numFmts: { numFmtId: number; formatCode: string }[]; cellXfs: object[] }> = {}): ParsedWorkbook {
  return {
    styles: {
      fonts: overrides.fonts ?? [font('Calibri')],
      fills: [], borders: [],
      cellXfs: overrides.cellXfs ?? [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }],
      numFmts: overrides.numFmts ?? [],
      dxfs: [],
    },
    sharedStrings: overrides.sharedStrings ?? [],
    workbook: { sheets: [] },
  } as unknown as ParsedWorkbook;
}

const textRow = (text: string, styleIndex = 0): Row =>
  ({ index: 1, height: null, cells: [{ col: 1, row: 1, styleIndex, value: { type: 'text', text } }] }) as unknown as Row;

const sheet = (rows: Row[], extra: Partial<Worksheet> = {}) =>
  ({ rows, charts: [], images: [], shapeGroups: [], ...extra }) as unknown as Worksheet;

const scalarsOf = (text: string) => [...text].map((ch) => ch.codePointAt(0) as number);

describe('XLSX font-source demand', () => {
  it('demands every style and rich-run font, the base of a regular full name, and Calibri for an unnamed font', () => {
    const demand = new XlsxFontDemand(workbook({
      fonts: [font('Lato Regular'), font(null, { bold: true }), font('Cambria', { italic: true })],
      sharedStrings: [{ text: 'ab', runs: [{ text: 'ab', font: { name: 'Roboto', bold: true, italic: false } }] }] as never,
    }), 'jp', OFFICE_FONT_SUBSTITUTIONS).take();
    expect(demand?.families).toEqual(['Lato Regular', 'Lato', 'Calibri', 'Cambria', 'Roboto']);
    expect(demand?.fallbackFamilies).toEqual([]);
  });

  it('demands printable ASCII, workbook text and number-format literals, and the formatter name tables', () => {
    const demand = new XlsxFontDemand(workbook({
      numFmts: [{ numFmtId: 164, formatCode: '"€"#,##0.00' }, { numFmtId: 165, formatCode: 'dddd mmmm' }],
      cellXfs: [{ numFmtId: 164, fontId: 0 }, { numFmtId: 165, fontId: 0 }, { numFmtId: 31, fontId: 0 }],
      sharedStrings: [{ text: 'Straße' }] as never,
    }), 'jp', []).take();
    const scalars = new Set(demand?.scalars as number[]);
    for (const cp of [...scalarsOf(' ~09AZaz…ß€年月日'), ...scalarsOf('Wednesday September')]) {
      expect(scalars.has(cp), `U+${cp.toString(16)}`).toBe(true);
    }
    // Built-in 31 (yyyy"年"m"月"d"日") is in use; Japanese weekdays are not.
    expect(scalars.has('曜'.codePointAt(0) as number)).toBe(false);
  });

  it('adds script fallbacks for the text the workbook paints', () => {
    const demand = new XlsxFontDemand(workbook({ sharedStrings: [{ text: '漢字' }] as never }), 'sc', []).take();
    expect(demand?.fallbackFamilies).toEqual(['Noto Sans SC', 'Noto Serif SC']);
  });

  it('emits per-sheet deltas with only new families and scalars, and none when nothing is new', () => {
    const demand = new XlsxFontDemand(workbook(), 'jp', []);
    demand.take();
    expect(demand.take()).toBeNull();
    demand.rows([textRow('ABC')]);
    expect(demand.take()).toBeNull();
    demand.rows([textRow('Ωmega')]);
    const delta = demand.take();
    expect(delta?.families).toEqual([]);
    expect(delta?.scalars).toEqual(scalarsOf('Ω'));
    expect(delta?.fallbackFamilies).toEqual(['Noto Sans', 'Noto Serif']);
    const ws = sheet([], { shapeGroups: [{ shapes: [{ geom: { type: 'preset', name: 'rect' }, text: {
      paragraphs: [{ runs: [{ type: 'text', text: 'Шрифт', fontFace: 'Nunito', bold: true, italic: false, size: 11 }] }],
    } }] }] as never });
    demand.worksheetParts(ws);
    const shapeDelta = demand.take();
    expect(shapeDelta?.families).toEqual(['Nunito']);
    expect(shapeDelta?.fallbackFamilies).toEqual([]);
    expect(shapeDelta?.scalars).toEqual(scalarsOf('Шрифт').filter((cp, i, all) => all.indexOf(cp) === i));
    demand.worksheetParts(ws);
    expect(demand.take()).toBeNull();
  });

  it('scans only rows a growing worksheet added since the last delta', () => {
    const demand = new XlsxFontDemand(workbook(), 'jp', []);
    demand.take();
    const ws = sheet([textRow('α')]);
    demand.newRowsOf(ws);
    expect(demand.take()?.scalars).toEqual(scalarsOf('α'));
    ws.rows.push(textRow('β'));
    demand.newRowsOf(ws);
    expect(demand.take()?.scalars).toEqual(scalarsOf('β'));
  });

  it('proves a script-scoped substitute only from text painted in that family', () => {
    const demand = new XlsxFontDemand(workbook({
      fonts: [font('Calibri'), font('Sakkal Majalla')],
      cellXfs: [{ numFmtId: 0, fontId: 0 }, { numFmtId: 0, fontId: 1 }],
    }), 'jp', OFFICE_FONT_SUBSTITUTIONS);
    expect(demand.take()?.substituteScripts?.['sakkal majalla']).toEqual([]);
    // Arabic text in a Calibri cell proves nothing for Sakkal Majalla.
    demand.rows([textRow('مرحبا', 0)]);
    expect(demand.take()?.substituteScripts?.['sakkal majalla']).toBeUndefined();
    demand.rows([textRow('مرحبا', 1)]);
    const widened = demand.take();
    expect(widened?.families).toEqual(['Sakkal Majalla']);
    expect(widened?.substituteScripts?.['sakkal majalla']).toEqual(['arabic']);
  });

  it("declines to 'all' for paint the model text does not spell", () => {
    const cases: Array<[string, (demand: XlsxFontDemand) => void, ParsedWorkbook?]> = [
      ['chart', (d) => d.worksheetParts(sheet([], { charts: [{ chart: { themeMinorFontLatin: 'Aptos' } }] } as never))],
      ['equation', (d) => d.worksheetParts(sheet([], { shapeGroups: [{ shapes: [{ geom: { type: 'preset', name: 'rect' },
        text: { paragraphs: [{ runs: [{ type: 'math', nodes: [], display: false }] }] } }] }] } as never))],
      ['vertical shape text', (d) => d.worksheetParts(sheet([], { shapeGroups: [{ shapes: [{ geom: { type: 'preset', name: 'rect' },
        text: { vert: 'eaVert', paragraphs: [] } }] }] } as never))],
      ['metafile picture', (d) => d.worksheetParts(sheet([], { images: [{ imagePath: 'xl/media/image1.emf', mimeType: 'image/x-emf' }] } as never))],
      ['stacked cell text', () => undefined, workbook({ cellXfs: [{ numFmtId: 0, fontId: 0, textRotation: 255 }] })],
      ['RTL reading order', () => undefined, workbook({ cellXfs: [{ numFmtId: 0, fontId: 0, readingOrder: 2 }] })],
    ];
    for (const [label, add, wb] of cases) {
      const demand = new XlsxFontDemand(wb ?? workbook(), 'jp', []);
      const first = demand.take();
      add(demand);
      const scalars = wb ? first?.scalars : demand.take()?.scalars;
      expect(scalars, label).toBe('all');
    }
  });

  it('demands chart faces and theme Japanese cell faces with the sheet', () => {
    const demand = new XlsxFontDemand(workbook(), 'jp', []);
    demand.take();
    demand.worksheetParts(sheet([], { charts: [{ chart: { title: { fontFace: '+mj-lt' }, themeMajorFontLatin: 'Aptos Display' } }] } as never));
    expect(demand.take()?.families).toEqual(['Aptos Display']);
  });

  it('demands the East Asian face stacked WordArt text paints CJK characters with', () => {
    const demand = new XlsxFontDemand(workbook(), 'jp', []);
    demand.take();
    demand.worksheetParts(sheet([], { shapeGroups: [{ shapes: [{ geom: { type: 'preset', name: 'rect' }, text: {
      vert: 'wordArtVert',
      paragraphs: [{ runs: [{ type: 'text', text: '縦書き', fontFace: 'Latin Face', fontFaceEa: 'EA Face', bold: false, italic: false, size: 11 }] }],
    } }] }] as never }));
    expect(demand.take()?.families).toEqual(['Latin Face', 'EA Face']);
  });

  it('feeds number-format output into script evidence as well as scalars', () => {
    // Built-in 31 paints yyyy"年"m"月"d"日" in a numeric-only workbook.
    const demand = new XlsxFontDemand(workbook({ cellXfs: [{ numFmtId: 31, fontId: 0 }] }), 'jp', []).take();
    expect(demand?.fallbackFamilies).toEqual(['Noto Sans JP', 'Noto Serif JP']);
  });

  describe('built-in 14 under the host locale', () => {
    afterEach(() => vi.unstubAllGlobals());

    it("demands 'all' for a calendar whose output the formatter cannot enumerate", () => {
      vi.stubGlobal('navigator', { language: 'zh-TW-u-ca-roc' });
      // Serial 1 paints a Republic-of-China era prefix (民國前…).
      expect(formatLocalizedExcelShortDate(1)).toContain('前');
      const demand = new XlsxFontDemand(workbook({ cellXfs: [{ numFmtId: 14, fontId: 0 }] }), 'jp', []).take();
      expect(demand?.scalars).toBe('all');
    });

    it('stays finite for a Gregorian, ASCII-digit locale and covers its separators', () => {
      vi.stubGlobal('navigator', { language: 'de-DE' });
      const demand = new XlsxFontDemand(workbook({ cellXfs: [{ numFmtId: 14, fontId: 0 }] }), 'jp', []).take();
      expect(demand?.scalars).not.toBe('all');
      const scalars = new Set(demand?.scalars as number[]);
      for (const ch of formatLocalizedExcelShortDate(45_000)) expect(scalars.has(ch.codePointAt(0) as number), ch).toBe(true);
    });
  });

  it('proves the scripts of a locale/calendar date output it cannot enumerate', () => {
    vi.stubGlobal('navigator', { language: 'zh-TW-u-ca-roc' });
    try {
      // Numeric-only workbook: built-in 14 paints 民國…, the only CJK text.
      const demand = new XlsxFontDemand(workbook({ cellXfs: [{ numFmtId: 14, fontId: 0 }] }), 'tc', []).take();
      expect(demand?.scalars).toBe('all');
      expect(demand?.fallbackFamilies).toEqual(['Noto Sans TC', 'Noto Serif TC']);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  describe('every painted text source feeds scalar demand and its family\'s script evidence', () => {
    const SCOPED = 'Sakkal Majalla';
    const shapeWith = (text: string) => sheet([], { shapeGroups: [{ shapes: [{ geom: { type: 'preset', name: 'rect' }, text: {
      paragraphs: [{ runs: [{ type: 'text', text, fontFace: SCOPED, bold: false, italic: false, size: 11 }] }],
    } }] }] as never });
    const slicerWith = (caption: string, item: string) => sheet([], { slicers: [{
      caption, items: [{ name: item, selected: true }],
      style: { header: { fontFamily: SCOPED }, selectedItemWithData: { fontFamily: SCOPED }, unselectedItemWithData: { fontFamily: SCOPED } },
    }] } as never);
    const scopedWorkbook = (extra: Parameters<typeof workbook>[0] = {}) => workbook({
      fonts: [font('Calibri'), font(SCOPED)],
      cellXfs: [{ numFmtId: 0, fontId: 0 }, { numFmtId: 0, fontId: 1 }],
      ...extra,
    });
    // [source, demand built from a text painted in the scoped family]
    const SOURCES: Array<[string, (text: string) => XlsxFontDemand]> = [
      ['cell text', (text) => { const d = new XlsxFontDemand(scopedWorkbook(), 'jp', OFFICE_FONT_SUBSTITUTIONS); d.rows([textRow(text, 1)]); return d; }],
      ['rich run', (text) => { const d = new XlsxFontDemand(scopedWorkbook(), 'jp', OFFICE_FONT_SUBSTITUTIONS);
        d.rows([{ index: 1, height: null, cells: [{ col: 1, row: 1, styleIndex: 0, value: { type: 'text', text,
          runs: [{ text, font: { name: SCOPED, bold: false, italic: false, underline: false, strike: false } }] } }] } as never]); return d; }],
      ['shared string run', (text) => new XlsxFontDemand(scopedWorkbook({ sharedStrings: [{ text,
        runs: [{ text, font: { name: SCOPED, bold: false, italic: false } }] }] as never }), 'jp', OFFICE_FONT_SUBSTITUTIONS)],
      ['shape text', (text) => { const d = new XlsxFontDemand(scopedWorkbook(), 'jp', OFFICE_FONT_SUBSTITUTIONS); d.worksheetParts(shapeWith(text)); return d; }],
      ['slicer caption', (text) => { const d = new XlsxFontDemand(scopedWorkbook(), 'jp', OFFICE_FONT_SUBSTITUTIONS); d.worksheetParts(slicerWith(text, 'x')); return d; }],
      ['slicer item', (text) => { const d = new XlsxFontDemand(scopedWorkbook(), 'jp', OFFICE_FONT_SUBSTITUTIONS); d.worksheetParts(slicerWith('x', text)); return d; }],
      ['number format literal', (text) => new XlsxFontDemand(scopedWorkbook({
        numFmts: [{ numFmtId: 164, formatCode: `0" ${text}"` }],
        cellXfs: [{ numFmtId: 0, fontId: 0 }, { numFmtId: 164, fontId: 1 }],
      }), 'jp', OFFICE_FONT_SUBSTITUTIONS)],
    ];
    const collect = (demand: XlsxFontDemand) => {
      const deltas = [demand.take(), demand.take()].filter(Boolean);
      return {
        scalars: deltas.some((d) => d!.scalars === 'all') ? 'all' as const : new Set(deltas.flatMap((d) => d!.scalars as number[])),
        scripts: deltas.flatMap((d) => d!.substituteScripts?.['sakkal majalla'] ?? []),
      };
    };
    for (const [source, build] of SOURCES) {
      it(source, () => {
        const greek = collect(build('Ωμεγα'));
        expect(greek.scalars, source).not.toBe('all');
        for (const ch of 'Ωμεγα') expect((greek.scalars as Set<number>).has(ch.codePointAt(0) as number), `${source} ${ch}`).toBe(true);
        expect(greek.scripts, source).toEqual([]);
        expect(collect(build('مرحبا')).scripts, source).toContain('arabic');
      });
    }
    it('chart text: its scalars are unprovable (demand is all) and it paints only unscoped substitutes', () => {
      const demand = new XlsxFontDemand(scopedWorkbook(), 'jp', OFFICE_FONT_SUBSTITUTIONS);
      demand.take();
      demand.worksheetParts(sheet([], { charts: [{ chart: { title: 'مرحبا', titleFontFace: SCOPED } }] } as never));
      const delta = demand.take();
      expect(delta?.scalars).toBe('all');
      expect(delta?.substituteScripts?.['sakkal majalla'] ?? []).toEqual([]);
    });
  });
});
