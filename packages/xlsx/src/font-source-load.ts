/**
 * XLSX font-source demand (`LoadOptions.fonts`): which families and Unicode
 * scalars one workbook asks the configured sources for. Library resource
 * policy, not OOXML behaviour: SpreadsheetML names faces (ECMA-376 §18.8.22
 * `font/name`, §18.4.7 `rPr/rFont`, DrawingML §21.1.2.3.7 `latin`) and leaves
 * their provision to the host.
 *
 * The main thread imports this module lazily, only when sources are
 * configured, so the default offline path never loads the session, the CSS
 * adapters or the inert-scalar tables. Main and the render worker build the
 * same demand from the same parsed model with this one class, so both realms
 * demand, register and load identical faces.
 *
 * Demand is incremental, like the workbook itself: the workbook delta (style
 * fonts, shared strings, number formats) is ensured at load, and each
 * worksheet contributes a delta (its cells, inline and formula strings, shape
 * text, charts, slicers, theme cell fonts) before any of its rows is measured
 * or painted. Each delta names only families and scalars not demanded before.
 */
import {
  classifyCjkFont,
  isMetafileMime,
  substituteEntryCoversText,
  type CjkLang,
  type FontSubstituteScript,
} from '@silurus/ooxml-core';
import { FontDemandScalarCollector } from '@silurus/ooxml-core/internal/font-scalar-demand';
import { ScriptPreloadAccumulator } from '@silurus/ooxml-core/internal/script-preload-accumulator';
import {
  chartFontFamilies,
  fontFamilyKey,
  type FontSessionDemand,
  type FontSource,
  type FontSubstitution,
} from '@silurus/ooxml-core/internal/font-sources';
import { usesJapaneseCellFonts } from './internal/cell-font-language.js';
import { numberFormatPaintTexts } from './number-format.js';
import type { Cell, CellFont, ParsedWorkbook, Row, ShapeInfo, Worksheet } from './types.js';
import { cellXfAt, effectiveCellStyleIndex } from './internal/cell-style.js';

export { FontSourceSession, InstalledFamilyExclusion } from '@silurus/ooxml-core/internal/font-sources';

/** The substitutions every configured source declares, in option order. */
export function declaredFontSubstitutions(sources: readonly FontSource[]): readonly FontSubstitution[] {
  return Object.freeze(sources.flatMap((source) => source.substitutes ?? []));
}

/**
 * Text the renderer paints that no model text spells, measured or painted
 * with the faces of every cell and shape: printable ASCII (row and column
 * headers, General and built-in number output, booleans and error values,
 * `#` fill, `E`/`%`/sign/separator characters, the space that `?` and `_`
 * pad with) and the horizontal ellipsis of clipped slicer text.
 */
const RENDERER_SENTINEL_TEXT = `${String.fromCharCode(...Array.from({ length: 0x7f - 0x20 }, (_, i) => 0x20 + i))}…`;

/** DrawingML vertical text modes other than horizontal (§20.1.10.83). */
const HORIZONTAL_SHAPE_TEXT: ReadonlySet<string> = new Set(['horz']);

const METAFILE_PATH = /\.(?:wmf|emf|wmz|emz)$/i;

/** Stacked text (`textRotation="255"`, §18.8.1) substitutes vertical forms;
 * RTL reading order (`readingOrder="2"`) mirrors neutral text (UAX #9). Either
 * paints scalars the text does not spell. */
function stylesNeedAllScalars(workbook: ParsedWorkbook): boolean {
  return (workbook.styles?.cellXfs ?? []).some((xf) => xf.textRotation === 255 || xf.readingOrder === 2);
}

/**
 * One workbook's font demand in one realm. {@link take} returns the delta
 * accumulated since the previous call (the workbook delta first), or null when
 * nothing new was demanded.
 */
export class XlsxFontDemand {
  private readonly cjkLang: CjkLang | null;
  private readonly scopedScripts = new Map<string, Set<FontSubstituteScript>>();
  private readonly scripts: ScriptPreloadAccumulator;
  /** Normalised family key → first spelling, every family demanded so far. */
  private readonly families = new Map<string, string>();
  private readonly fallbacks = new Set<string>();
  private readonly proven = new Map<string, Set<FontSubstituteScript>>();
  private readonly committed = new Set<number>();
  private all = false;
  private collector: FontDemandScalarCollector;
  private freshFamilies = new Map<string, string>();
  private readonly scannedRows = new WeakMap<Row[], number>();
  private readonly scannedParts = new WeakSet<Worksheet>();

  constructor(
    private readonly workbook: ParsedWorkbook,
    cjkFallback: CjkLang,
    substitutes: readonly FontSubstitution[],
  ) {
    for (const substitution of substitutes) {
      if (!substitution.script) continue;
      const key = fontFamilyKey(substitution.for);
      this.scopedScripts.set(key, (this.scopedScripts.get(key) ?? new Set()).add(substitution.script));
    }
    // The workbook's CJK language for shared Han: the first CJK-classified
    // style font, else the resolved workbook fallback (as the renderer).
    let cjkLang: CjkLang | null = null;
    for (const font of workbook.styles?.fonts ?? []) cjkLang ??= classifyCjkFont(font.name);
    this.cjkLang = cjkLang ?? cjkFallback;
    this.scripts = new ScriptPreloadAccumulator(this.cjkLang);
    this.collector = this.newCollector();
    if (stylesNeedAllScalars(workbook)) this.collector.decline();
    this.collector.add(RENDERER_SENTINEL_TEXT);
    // §18.8.22: every style font, including ones no cell uses yet (a later
    // sheet may), and the shared-string table's rich-run fonts.
    for (const font of workbook.styles?.fonts ?? []) this.addCellFont(font);
    for (const shared of workbook.sharedStrings ?? []) {
      this.addText(shared.text);
      // A run's own face paints it wherever the string is used; plain shared
      // text is charged to its cell's face when the sheet's rows are scanned.
      for (const run of shared.runs ?? []) {
        if (run.font?.name) this.addFamily(run.font.name);
        this.addText(run.text, run.font?.name);
      }
      for (const phonetic of shared.phoneticRuns ?? []) this.addText(phonetic.text);
    }
    // Number-format output is painted text like any cell text, in the font of
    // the style that formats it: its scalars, its script fallback evidence (a
    // CJK date literal demands the CJK fallback) and its family's scoped
    // scripts. A conditional format's numFmt can format a cell of any style,
    // so it proves for every style font. A locale/calendar-dependent output
    // the formatter cannot enumerate demands 'all' scalars, while its sampled
    // texts still prove their scripts.
    const formats = numberFormatPaintTexts(workbook.styles);
    if (!formats.complete) this.collector.decline();
    const styleFamily = (fontId: number | undefined) =>
      workbook.styles?.fonts?.[fontId ?? 0]?.name?.trim() || 'Calibri';
    // Styles share formats and fonts: charge each (family, format) pair once.
    const charged = new Map<string, Set<readonly string[]>>();
    (workbook.styles?.cellXfs ?? []).forEach((xf, index) => {
      const texts = formats.cellXfs[index];
      if (!texts?.length) return;
      const family = styleFamily(xf.fontId);
      const done = charged.get(family) ?? new Set();
      if (done.has(texts)) return;
      done.add(texts);
      charged.set(family, done);
      for (const text of texts) this.addText(text, family);
    });
    const families = [...new Set((workbook.styles?.fonts ?? []).map((font) => font.name?.trim() || 'Calibri'))];
    for (const text of formats.dxfs) {
      this.addText(text);
      for (const family of families) this.prove(family, text);
    }
  }

  private newCollector(): FontDemandScalarCollector {
    // Computed in the realm that loads the faces: nothing is serialized, so
    // the only bound is the scalar-count ceiling inside the collector.
    return new FontDemandScalarCollector({ budget: Number.POSITIVE_INFINITY, committed: this.committed });
  }

  /** A painted text: its scalars and its script fallback evidence. */
  private addText(text: string | null | undefined, family?: string | null): void {
    if (!text) return;
    this.collector.add(text);
    this.scripts.addText([text]);
    if (family) this.prove(family, text);
  }

  /** Record the scoped-substitute scripts `text` proves for `family`. */
  private prove(family: string, text: string): void {
    const key = fontFamilyKey(family);
    const scoped = this.scopedScripts.get(key);
    if (!scoped) return;
    for (const script of scoped) {
      if (this.proven.get(key)?.has(script) || !substituteEntryCoversText({ script }, text, 'any')) continue;
      this.proven.set(key, (this.proven.get(key) ?? new Set()).add(script));
      // A widened scope re-demands the family so the session may consult the
      // newly admissible substitute.
      const name = this.families.get(key);
      if (name) this.freshFamilies.set(key, name);
    }
  }

  private addFamily(family: string | null | undefined): void {
    const name = family?.trim();
    if (!name || name.startsWith('+')) return;
    const key = fontFamilyKey(name);
    if (!key) return;
    if (!this.families.has(key)) {
      this.families.set(key, name);
      this.freshFamilies.set(key, name);
    }
  }

  /** A style or run font: an omitted name paints the workbook default chain,
   * headed by Calibri (as the renderer). A full regular name (`"Lato
   * Regular"`) also demands its CSS base family, which the renderer offers as
   * a loaded alias (`LoadedFontSources.regularAliases`). */
  private addCellFont(font: Pick<CellFont, 'name' | 'bold' | 'italic'>): void {
    const name = font.name?.trim() || 'Calibri';
    this.addFamily(name);
    const regular = /^(.*?)\s+Regular$/i.exec(name);
    if (regular?.[1]?.trim()) this.addFamily(regular[1].trim());
  }

  /** The font a cell paints with, through the renderer's effective-style rule
   * (internal/cell-style.ts: cell style, else its column's style, else Normal). */
  private fontOfCell(worksheet: Worksheet | undefined, cell: Cell): CellFont | undefined {
    const index = worksheet ? effectiveCellStyleIndex(worksheet, cell, cell.col) : cell.styleIndex ?? 0;
    const xf = this.workbook.styles ? cellXfAt(this.workbook.styles, index) : undefined;
    return xf ? this.workbook.styles.fonts?.[xf.fontId] : undefined;
  }

  /** Cells of `rows`: text, rich runs and phonetic runs with the faces that
   * paint them, and error values. Numbers paint only number-format text
   * (workbook delta) and ASCII. */
  rows(rows: readonly Row[], worksheet?: Worksheet): void {
    const japanese = worksheet ? usesJapaneseCellFonts(worksheet) : false;
    for (const row of rows) for (const cell of row.cells) {
      const value = cell.value;
      if (value.type === 'error') { this.collector.add(value.error); continue; }
      if (value.type !== 'text') continue;
      const font = this.fontOfCell(worksheet, cell);
      const family = font?.name?.trim() || 'Calibri';
      if (value.runs?.length) {
        for (const run of value.runs) {
          if (run.font?.name) this.addFamily(run.font.name);
          this.addText(run.text, run.font?.name?.trim() || family);
        }
      } else {
        this.addText(value.text, family);
      }
      if (japanese && font?.scheme) {
        const theme = font.scheme === 'major' ? worksheet?.themeJapaneseMajorFont : worksheet?.themeJapaneseMinorFont;
        if (theme) this.prove(theme, value.text);
      }
      const phoneticFont = value.phoneticPr ? this.workbook.styles.fonts?.[value.phoneticPr.fontId] : undefined;
      for (const phonetic of value.phoneticRuns ?? []) this.addText(phonetic.text, phoneticFont?.name ?? family);
    }
  }

  /** Rows of `worksheet` not scanned before (its rows array may still be
   * growing while a provisional sheet streams in). */
  newRowsOf(worksheet: Worksheet): void {
    const rows = worksheet.rows ?? [];
    const scanned = this.scannedRows.get(rows) ?? 0;
    if (rows.length > scanned) this.rows(rows.slice(scanned), worksheet);
    this.scannedRows.set(rows, rows.length);
  }

  /** The worksheet's non-cell paint, once per worksheet model: theme cell
   * fonts, shapes, charts, slicers and pictures. */
  worksheetParts(worksheet: Worksheet): void {
    if (this.scannedParts.has(worksheet)) return;
    this.scannedParts.add(worksheet);
    if (usesJapaneseCellFonts(worksheet)) {
      // Scheme cells paint with the workbook theme's Japanese faces.
      this.addFamily(worksheet.themeJapaneseMajorFont);
      this.addFamily(worksheet.themeJapaneseMinorFont);
    }
    // Charts paint their own text and number formatting (axis labels, data
    // labels, legends): the scalars cannot be proven from the model text.
    // Chart sheets carry their chart in `charts` as well.
    if ((worksheet.charts ?? []).length > 0) this.collector.decline();
    for (const anchor of worksheet.charts ?? []) {
      for (const family of chartFontFamilies(anchor.chart)) this.addFamily(family);
    }
    for (const image of worksheet.images ?? []) {
      // Metafile text records paint with document faces (core wmf-text.ts).
      if (isMetafileMime(image.mimeType) || METAFILE_PATH.test(image.imagePath ?? '')) this.collector.decline();
    }
    for (const anchor of worksheet.shapeGroups ?? []) for (const shape of anchor.shapes) this.shape(shape);
    for (const slicer of worksheet.slicers ?? []) {
      // The caption paints with the header face, items with the item-state
      // faces (renderer drawSlicerFrame); the library default face is not an
      // authored family and is not demanded.
      const header = slicer.style?.header;
      const itemStyles = [slicer.style?.selectedItemWithData, slicer.style?.unselectedItemWithData];
      for (const style of [header, ...itemStyles, slicer.style?.whole]) {
        if (style?.fontFamily) this.addFamily(style.fontFamily);
      }
      this.addText(slicer.caption, header?.fontFamily);
      for (const item of slicer.items) {
        this.addText(item.name);
        for (const style of itemStyles) if (style?.fontFamily) this.prove(style.fontFamily, item.name);
      }
    }
  }

  private shape(shape: ShapeInfo): void {
    if (shape.geom.type === 'image') {
      const geom = shape.geom as { mimeType?: string; imagePath?: string };
      if (isMetafileMime(geom.mimeType) || METAFILE_PATH.test(geom.imagePath ?? '')) this.collector.decline();
    }
    const text = shape.text;
    if (!text) return;
    // Vertical and WordArt-vertical text rotate or substitute vertical forms;
    // RTL paragraphs mirror neutral text (UAX #9).
    if (text.vert !== undefined && !HORIZONTAL_SHAPE_TEXT.has(text.vert)) this.collector.decline();
    for (const paragraph of text.paragraphs) {
      if (paragraph.rtl) this.collector.decline();
      for (const run of paragraph.runs) {
        // OMML equations paint their own operator and layout glyphs.
        if (run.type === 'math') { this.collector.decline(); continue; }
        if (run.type !== 'text') continue;
        // Every slot the shape painter can select: the Latin face, and the
        // East Asian face that stacked (WordArt-vertical) text paints CJK
        // characters with (renderer drawShapeStackedText). The complex-script
        // slot is not painted by the XLSX shape renderer.
        const family = run.fontFace?.trim() || 'Calibri';
        const eastAsian = run.fontFaceEa?.trim();
        this.addFamily(family);
        if (eastAsian) this.addFamily(eastAsian);
        this.addText(run.text, family);
        if (eastAsian) this.prove(eastAsian, run.text);
      }
    }
  }

  /** The demand accumulated since the previous call, or null when it adds
   * no family, no scalar and no scope. */
  take(): FontSessionDemand | null {
    const scalars = this.collector.result();
    const becameAll = scalars === 'all' && !this.all;
    const newScalars = scalars === 'all' ? [] : scalars;
    // Script fallback families follow the text seen so far (cumulative).
    const fallbackFamilies = this.scripts.names().filter((family) => !this.fallbacks.has(family));
    const families = [...this.freshFamilies.values()];
    if (!becameAll && newScalars.length === 0 && families.length === 0 && fallbackFamilies.length === 0) return null;
    const substituteScripts = Object.create(null) as Record<string, readonly FontSubstituteScript[]>;
    for (const key of this.freshFamilies.keys()) {
      substituteScripts[key] = Object.freeze([...(this.proven.get(key) ?? [])].sort());
    }
    const demand: FontSessionDemand = Object.freeze({
      families: Object.freeze(families),
      fallbackFamilies: Object.freeze(fallbackFamilies),
      scalars: this.all || becameAll ? 'all' : Object.freeze(newScalars),
      substituteScripts: Object.freeze(substituteScripts),
    });
    for (const family of fallbackFamilies) this.fallbacks.add(family);
    if (scalars === 'all') this.all = true;
    else for (const scalar of newScalars) this.committed.add(scalar);
    this.freshFamilies = new Map();
    if (!this.all) this.collector = this.newCollector();
    return demand;
  }
}
