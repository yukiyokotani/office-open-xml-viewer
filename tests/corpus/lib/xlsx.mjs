// SpreadsheetML specimen composer. Column widths and row heights are the grid
// (grid.mjs XLSX); a case places values and styles by cell reference and the
// composer emits only the fonts, borders, number formats and cell formats the
// sheet actually uses.

import { FONT, XLSX } from './grid.mjs';
import { contentTypes, relationships } from './opc.mjs';
import { XML_DECL, attrs, esc } from './xml.mjs';

export const MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
export const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** Fonts by role: one family, two weights, three sizes. Index 0 is the default. */
const FONTS = {
  text: { sz: XLSX.TYPE.text },
  textBold: { sz: XLSX.TYPE.text, b: true },
  caption: { sz: XLSX.TYPE.caption },
  captionBold: { sz: XLSX.TYPE.caption, b: true },
  title: { sz: XLSX.TYPE.title, b: true },
};

/** Built-in number formats (ECMA-376 Part 1, §18.8.30) used by the corpus. */
export const BUILTIN_NUMFMT = { 3: '#,##0', 10: '0.00%', 12: '# ?/?' };

const colIndex = (letters) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const splitRef = (ref) => {
  const [, col, row] = /^([A-Z]+)(\d+)$/.exec(ref);
  return { col, row: Number(row) };
};

class Styles {
  constructor() {
    this.fonts = ['text'];
    this.fills = ['', 'gray125'];
    this.dxfs = [];
    this.borders = [''];
    this.numFmts = [];
    this.xfs = [{ font: 0, fill: 0, border: 0, numFmt: 0, align: null }];
  }
  index(list, key) {
    let i = list.indexOf(key);
    if (i < 0) i = list.push(key) - 1;
    return i;
  }
  /** { font?, numFmt?: builtin id | code, align?: { horizontal?, vertical? },
   * border?: { left?, right?, top?, bottom? } (ST_BorderStyle) } → xf index. */
  xf({ font = 'text', color, themeColor, scheme, italic, underline, fill, numFmt, align, border } = {}) {
    if (!FONTS[font]) throw new Error(`unknown font role ${font}`);
    const fontKey = themeColor || scheme || italic || underline
      ? JSON.stringify({ font, color, themeColor, scheme, italic, underline })
      : color ? `${font}|${color}` : font;
    let numFmtId = 0;
    if (typeof numFmt === 'number') {
      if (!(numFmt in BUILTIN_NUMFMT)) throw new Error(`unlisted built-in number format ${numFmt}`);
      numFmtId = numFmt;
    } else if (typeof numFmt === 'string') {
      numFmtId = 164 + this.index(this.numFmts, numFmt);
    }
    const borderKey = border
      ? ['left', 'right', 'top', 'bottom'].filter((e) => border[e]).map((e) => `${e}:${border[e]}`).join(';')
      : '';
    const xf = {
      font: this.index(this.fonts, fontKey),
      fill: fill ? this.index(this.fills, fill) : 0,
      border: this.index(this.borders, borderKey),
      numFmt: numFmtId,
      align: align ? JSON.stringify(align) : null,
    };
    const found = this.xfs.findIndex((x) => x.font === xf.font && (x.fill ?? 0) === xf.fill && x.border === xf.border
      && x.numFmt === xf.numFmt && x.align === xf.align);
    return found >= 0 ? found : this.xfs.push(xf) - 1;
  }
  /** Differential format for conditional formatting: raw dxf content. */
  dxf(xml) { return this.index(this.dxfs, xml); }
  xml() {
    const fonts = this.fonts.map((key) => {
      const o = key.startsWith('{') ? JSON.parse(key) : { font: key.split('|')[0], color: key.split('|')[1] };
      const f = FONTS[o.font];
      const color = o.themeColor
        ? `<color theme="${o.themeColor[0]}"${o.themeColor[1] ? ` tint="${o.themeColor[1]}"` : ''}/>`
        : o.color ? `<color rgb="FF${o.color}"/>` : '';
      return `<font>${f.b ? '<b/>' : ''}${o.italic ? '<i/>' : ''}${o.underline ? '<u/>' : ''}<sz val="${f.sz}"/>${color}<name val="${FONT.latin}"/>`
        + `${o.scheme ? `<family val="2"/><scheme val="${o.scheme}"/>` : ''}</font>`;
    }).join('');
    const fills = this.fills.map((f) => (f === '' ? '<fill><patternFill patternType="none"/></fill>'
      : f === 'gray125' ? '<fill><patternFill patternType="gray125"/></fill>'
        : f.startsWith('theme:')
          ? `<fill><patternFill patternType="solid"><fgColor theme="${f.split(':')[1]}"${f.split(':')[2] ? ` tint="${f.split(':')[2]}"` : ''}/><bgColor indexed="64"/></patternFill></fill>`
          : `<fill><patternFill patternType="solid"><fgColor rgb="FF${f}"/><bgColor indexed="64"/></patternFill></fill>`)).join('');
    const borders = this.borders.map((key) => {
      const edges = Object.fromEntries(key ? key.split(';').map((p) => p.split(':')) : []);
      return `<border>${['left', 'right', 'top', 'bottom']
        .map((e) => (edges[e] ? `<${e} style="${edges[e]}"/>` : '')).join('')}</border>`;
    }).join('');
    const numFmts = this.numFmts.length
      ? `<numFmts count="${this.numFmts.length}">${this.numFmts
        .map((code, i) => `<numFmt numFmtId="${164 + i}" formatCode="${esc(code)}"/>`).join('')}</numFmts>`
      : '';
    const xfs = this.xfs.map((x) => {
      const a = {
        numFmtId: x.numFmt,
        fontId: x.font,
        fillId: x.fill ?? 0,
        borderId: x.border,
        xfId: 0,
        applyNumberFormat: x.numFmt ? 1 : undefined,
        applyFont: x.font ? 1 : undefined,
        applyFill: x.fill ? 1 : undefined,
        applyBorder: x.border ? 1 : undefined,
        applyAlignment: x.align ? 1 : undefined,
      };
      return x.align ? `<xf${attrs(a)}><alignment${attrs(JSON.parse(x.align))}/></xf>` : `<xf${attrs(a)}/>`;
    }).join('');
    return XML_DECL + `<styleSheet ${MAIN}>${numFmts}`
      + `<fonts count="${this.fonts.length}">${fonts}</fonts>`
      + `<fills count="${this.fills.length}">${fills}</fills>`
      + `<borders count="${this.borders.length}">${borders}</borders>`
      + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      + `<cellXfs count="${this.xfs.length}">${xfs}</cellXfs>`
      + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
      + (this.dxfs.length ? `<dxfs count="${this.dxfs.length}">${this.dxfs.map((d) => `<dxf>${d}</dxf>`).join('')}</dxfs>` : '')
      + '</styleSheet>';
  }
}

/** A worksheet's content, shared by the first and any further sheets. */
function sheetState({ name, cols, gridLines = false }) {
  return {
    name, cols, gridLines, cells: new Map(), rowHeights: new Map(), merges: [],
    conditional: [], hiddenRows: new Set(), pane: '', sheetPr: '', after: '', sheetRels: [], state: undefined,
  };
}

function sheetApi(st, styles, root) {
  const api = {
    styles,
    /** Value: string (shared string), number, null for a styled empty cell,
     * { rich: [[text, fontStyle?], ...] } for a rich-text shared string, or
     * { f, v, t? } for a formula with its cached value; style: see Styles.xf. */
    set(ref, v, style) { st.cells.set(ref, { v, s: style ? styles.xf(style) : 0 }); return api; },
    height(row, pt) { st.rowHeights.set(row, pt); return api; },
    merge(range) { st.merges.push(range); return api; },
    /** Raw <conditionalFormatting> element. */
    conditional(xml) { st.conditional.push(xml); return api; },
    hideRow(row) { st.hiddenRows.add(row); return api; },
    /** Raw <pane> (and selection) inside the sheet view. */
    pane(xml) { st.pane = xml; return api; },
    /** Raw <sheetPr> content (tab colour, page setup properties). */
    sheetPr(xml) { st.sheetPr = xml; return api; },
    /** Worksheet elements after the conditional formats, in schema order
     * (printOptions, pageMargins, pageSetup, headerFooter, drawing). */
    after(xml) { st.after += xml; return api; },
    /** Relationships of this worksheet. */
    rels(list) { st.sheetRels.push(...list); return api; },
    /** Sheet visibility (hidden, veryHidden). */
    state(value) { st.state = value; return api; },
    ...root,
  };
  return api;
}

/** A workbook under construction; the returned api writes to its first sheet. */
export function sheet({ id, title, verify, cols, gridLines = false }) {
  const styles = new Styles();
  const sheets = [sheetState({ name: id, cols, gridLines })];
  const extra = { parts: {}, overrides: {}, defaults: {}, workbookRels: [], definedNames: [], bookView: '', theme: null };
  const root = {
    /** Extra package parts, worksheet relationships (first sheet) and content types. */
    attach({ parts = {}, sheetRels = [], overrides = {}, defaults = {} }) {
      Object.assign(extra.parts, parts); sheets[0].sheetRels.push(...sheetRels);
      Object.assign(extra.overrides, overrides); Object.assign(extra.defaults, defaults);
      return api;
    },
    /** Another worksheet sharing styles and strings; returns its own api. */
    addSheet({ name, cols: sheetCols, gridLines: g = false }) {
      const st = sheetState({ name, cols: sheetCols, gridLines: g });
      sheets.push(st);
      return sheetApi(st, styles, {});
    },
    definedName(name, ref, attrs2 = '') { extra.definedNames.push(`<definedName name="${name}"${attrs2}>${esc(ref)}</definedName>`); return api; },
    /** Raw <workbookView> attributes, e.g. ' activeTab="1"'. */
    bookView(attrs2) { extra.bookView = attrs2; return api; },
    /** Include a theme part (DrawingML theme XML). */
    theme(xml) { extra.theme = xml; return api; },
    build: () => buildPackage({ styles, sheets, extra }),
  };
  const api = sheetApi(sheets[0], styles, root);
  // Header furniture: rows 2-4 of column B; row 3 holds the 24 pt title.
  api.set('B2', id, { font: 'captionBold' });
  api.set('B3', title, { font: 'title' }).height(3, XLSX.TITLE_ROW_PT);
  api.set('B4', verify, { font: 'caption' });
  return api;
}

function worksheetXml(st, sharedString) {
  const rows = new Map();
  for (const [ref, cell] of st.cells) {
    const { row } = splitRef(ref);
    if (!rows.has(row)) rows.set(row, []);
    rows.get(row).push([ref, cell]);
  }
  for (const row of [...st.rowHeights.keys(), ...st.hiddenRows]) if (!rows.has(row)) rows.set(row, []);
  const sheetData = [...rows.keys()].sort((a, b) => a - b).map((r) => {
    const rowCells = rows.get(r)
      .sort(([a], [b]) => colIndex(splitRef(a).col) - colIndex(splitRef(b).col))
      .map(([ref, { v, s }]) => {
        const sAttr = s ? ` s="${s}"` : '';
        if (v === null) return `<c r="${ref}"${sAttr}/>`;
        if (typeof v === 'number') return `<c r="${ref}"${sAttr}><v>${v}</v></c>`;
        if (v.f !== undefined) {
          const t = typeof v.v === 'string' ? ' t="str"' : '';
          return `<c r="${ref}"${sAttr}${t}><f>${esc(v.f)}</f><v>${esc(v.v)}</v></c>`;
        }
        return `<c r="${ref}"${sAttr} t="s"><v>${sharedString(v)}</v></c>`;
      }).join('');
    const ht = st.rowHeights.get(r);
    const hidden = st.hiddenRows.has(r) ? ' hidden="1"' : '';
    return `<row r="${r}"${ht ? ` ht="${ht}" customHeight="1"` : ''}${hidden}>${rowCells}</row>`;
  }).join('');
  const colsXml = `<cols>${st.cols.map(([letter, width, hidden]) => {
    const n = colIndex(letter);
    return `<col min="${n}" max="${n}" width="${width}"${hidden ? ' hidden="1"' : ''} customWidth="1"/>`;
  }).join('')}</cols>`;

  return XML_DECL + `<worksheet ${MAIN} ${R}>`
    + (st.sheetPr ? `<sheetPr>${st.sheetPr}</sheetPr>` : '')
    + `<sheetViews><sheetView${st.gridLines ? '' : ' showGridLines="0"'} workbookViewId="0"${st.pane ? `>${st.pane}</sheetView>` : '/>'}</sheetViews>`
    + `<sheetFormatPr defaultColWidth="${XLSX.MODULE}" defaultRowHeight="${XLSX.ROW_PT}" customHeight="1"/>`
    + `${colsXml}<sheetData>${sheetData}</sheetData>`
    + (st.merges.length ? `<mergeCells count="${st.merges.length}">${st.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : '')
    + st.conditional.join('')
    + st.after
    + '</worksheet>';
}

const richRun = ([text, o = {}]) => `<r><rPr>${o.b ? '<b/>' : ''}${o.i ? '<i/>' : ''}${o.u ? '<u/>' : ''}`
  + `<sz val="${FONTS[o.font ?? 'text'].sz}"/>${o.color ? `<color rgb="FF${o.color}"/>` : ''}<rFont val="${FONT.latin}"/></rPr>`
  + `<t${/^\s|\s$/.test(text) ? ' xml:space="preserve"' : ''}>${esc(text)}</t></r>`;

function buildPackage({ styles, sheets, extra }) {
  const strings = [];
  const stringIndex = new Map();
  const sharedString = (v) => {
    const key = typeof v === 'string' ? v : JSON.stringify(v);
    if (!stringIndex.has(key)) stringIndex.set(key, strings.push(v) - 1);
    return stringIndex.get(key);
  };
  const worksheets = sheets.map((st) => worksheetXml(st, sharedString));
  const sst = XML_DECL + `<sst ${MAIN} count="${strings.length}" uniqueCount="${strings.length}">`
    + strings.map((s) => (typeof s === 'string'
      ? `<si><t${/^\s|\s$/.test(s) ? ' xml:space="preserve"' : ''}>${esc(s)}</t></si>`
      : `<si>${s.rich.map(richRun).join('')}</si>`)).join('')
    + '</sst>';

  // Relationship ids: sheet 1, styles, shared strings, then further sheets, then the theme.
  const sheetRid = (i) => (i === 0 ? 'rId1' : `rId${3 + i}`);
  const themeRid = `rId${3 + sheets.length}`;
  const workbook = XML_DECL + `<workbook ${MAIN} ${R}>`
    + (extra.bookView ? `<bookViews><workbookView${extra.bookView}/></bookViews>` : '')
    + `<sheets>${sheets.map((st, i) => `<sheet name="${esc(st.name)}" sheetId="${i + 1}"${st.state ? ` state="${st.state}"` : ''} r:id="${sheetRid(i)}"/>`).join('')}</sheets>`
    + (extra.definedNames.length ? `<definedNames>${extra.definedNames.join('')}</definedNames>` : '')
    + '</workbook>';
  const WS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
  const types = contentTypes({
    'xl/workbook.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    'xl/worksheets/sheet1.xml': WS,
    'xl/styles.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
    'xl/sharedStrings.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
    ...Object.fromEntries(sheets.slice(1).map((_, i) => [`xl/worksheets/sheet${i + 2}.xml`, WS])),
    ...(extra.theme ? { 'xl/theme/theme1.xml': 'application/vnd.openxmlformats-officedocument.theme+xml' } : {}),
    ...extra.overrides,
  });
  const defaults = Object.entries(extra.defaults)
    .map(([ext, type]) => `<Default Extension="${ext}" ContentType="${type}"/>`).join('');
  const files = {
    '[Content_Types].xml': defaults ? types.replace('<Default Extension="rels"', `${defaults}<Default Extension="rels"`) : types,
    '_rels/.rels': relationships([['rId1', 'officeDocument', 'xl/workbook.xml']]),
    'xl/_rels/workbook.xml.rels': relationships([
      ['rId1', 'worksheet', 'worksheets/sheet1.xml'],
      ['rId2', 'styles', 'styles.xml'],
      ['rId3', 'sharedStrings', 'sharedStrings.xml'],
      ...sheets.slice(1).map((_, i) => [sheetRid(i + 1), 'worksheet', `worksheets/sheet${i + 2}.xml`]),
      ...(extra.theme ? [[themeRid, 'theme', 'theme/theme1.xml']] : []),
    ]),
    'xl/workbook.xml': workbook,
    'xl/styles.xml': styles.xml(),
    'xl/sharedStrings.xml': sst,
    ...(extra.theme ? { 'xl/theme/theme1.xml': extra.theme } : {}),
    ...extra.parts,
  };
  sheets.forEach((st, i) => {
    files[`xl/worksheets/sheet${i + 1}.xml`] = worksheets[i];
    if (st.sheetRels.length) files[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = relationships(st.sheetRels);
  });
  return files;
}
