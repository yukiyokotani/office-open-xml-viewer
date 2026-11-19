// WordprocessingML specimen composer. A case lists paragraphs, tables, section
// starts and column breaks, with the baseline-grid line a block starts on; the
// composer turns the distance to that line into paragraph spacing, so every
// block lands on the 12 pt grid. Paragraphs given as `text` wrap naturally:
// their height is unknown, and no later block may then ask for a line.

import { COLOR, DOCX, FONT, LANG } from './grid.mjs';
import { contentTypes, relationships } from './opc.mjs';
import { XML_DECL, attrs, esc } from './xml.mjs';

export const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
export const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const STYLE_ID = { caption: 'Caption', title: 'Title' };
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

/** Text with '\t' → <w:t>/<w:tab/> sequence (no run wrapper). */
function textXml(text) {
  return text.split('\t').map((part, i) => {
    const tab = i > 0 ? '<w:tab/>' : '';
    if (part === '') return tab;
    const space = /^\s|\s$/.test(part) ? ' xml:space="preserve"' : '';
    return `${tab}<w:t${space}>${esc(part)}</w:t>`;
  }).join('');
}

/** A run: rPr is raw run-property XML (in schema order) or ''. */
export const run = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${textXml(text)}</w:r>`;

/** Column edges as indents: cols [c1, c2] → { left, right } in twips. */
const colInd = ([c1, c2], width = DOCX.LIVE_W) => {
  const ind = {};
  if (DOCX.colStart(c1) !== 0) ind['w:left'] = DOCX.colStart(c1);
  if (width - DOCX.colEnd(c2) !== 0) ind['w:right'] = width - DOCX.colEnd(c2);
  return ind;
};

function paragraphXml(p, before, after, sectPr) {
  const style = STYLE_ID[p.style ?? 'text'];
  const ind = { ...(p.cols ? colInd(p.cols) : {}), ...(p.ind ?? {}) };
  const pPr = [
    style ? `<w:pStyle w:val="${style}"/>` : '',
    p.pStyle ? `<w:pStyle w:val="${p.pStyle}"/>` : '',
    p.numPr ?? '',
    p.pBdr ?? '',
    p.shd ?? '',
    p.tabs ? `<w:tabs>${p.tabs.map(([val, pos]) => `<w:tab w:val="${val}" w:pos="${pos}"/>`).join('')}</w:tabs>` : '',
    p.pPrMid ?? '',
    before || after || p.lineAuto
      ? `<w:spacing${attrs({ 'w:before': before || undefined, 'w:after': after || undefined, 'w:line': p.lineAuto ? DOCX.LINE : undefined, 'w:lineRule': p.lineAuto ? 'auto' : undefined })}/>`
      : '',
    Object.keys(ind).length ? `<w:ind${attrs(ind)}/>` : '',
    p.jc ? `<w:jc w:val="${p.jc}"/>` : '',
    p.markRPr ? `<w:rPr>${p.markRPr}</w:rPr>` : '',
    sectPr ?? '',
  ].join('');
  const rPr = p.bold ? '<w:b/>' : (p.rPr ?? '');
  const brk = p.columnBreakBefore ? '<w:r><w:br w:type="column"/></w:r>'
    : p.pageBreakBefore ? '<w:r><w:br w:type="page"/></w:r>' : '';
  let content;
  if (p.runs) content = p.runs.join('');
  else if (p.text !== undefined) content = run(p.text, rPr);
  else {
    const body = p.lines.map(textXml).join('<w:br/>');
    content = body ? `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${body}</w:r>` : '';
  }
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${brk}${content}</w:p>`;
}

function border(name, b) {
  return `<w:${name} w:val="single" w:sz="${b}" w:space="0" w:color="${COLOR.ink}"/>`;
}

function tableXml(t) {
  const total = t.grid.reduce((a, b) => a + b, 0);
  const tblBorders = Object.entries(t.borders ?? {}).map(([name, sz]) => border(name, sz)).join('');
  const rows = t.rows.map((row) => {
    let col = 0;
    const cells = row.cells.map((cell) => {
      const span = cell.gridSpan ?? 1;
      const width = t.grid.slice(col, col + span).reduce((a, b) => a + b, 0);
      col += span;
      const tcPr = [
        `<w:tcW w:w="${width}" w:type="dxa"/>`,
        span > 1 ? `<w:gridSpan w:val="${span}"/>` : '',
        cell.vMerge ? `<w:vMerge${cell.vMerge === 'restart' ? ' w:val="restart"' : ''}/>` : '',
        cell.borders ? `<w:tcBorders>${Object.entries(cell.borders).map(([n, sz]) => border(n, sz)).join('')}</w:tcBorders>` : '',
        cell.shd ?? '',
        cell.marRight !== undefined ? `<w:tcMar><w:right w:w="${cell.marRight}" w:type="dxa"/></w:tcMar>` : '',
        cell.vAlign ? `<w:vAlign w:val="${cell.vAlign}"/>` : '',
      ].join('');
      const p = cell.runs ? { ...cell, text: undefined } : { ...cell, lines: [cell.text ?? ''], text: undefined };
      return `<w:tc><w:tcPr>${tcPr}</w:tcPr>${paragraphXml(p, 0, 0)}</w:tc>`;
    }).join('');
    return `<w:tr><w:trPr><w:trHeight w:val="${row.lines * DOCX.LINE}" w:hRule="exact"/></w:trPr>${cells}</w:tr>`;
  }).join('');
  const mar = t.cellMar ?? {};
  return '<w:tbl><w:tblPr>'
    + (t.style ? `<w:tblStyle w:val="${t.style}"/>` : '')
    + `<w:tblW w:w="${total}" w:type="dxa"/>`
    + (t.ind ? `<w:tblInd w:w="${t.ind}" w:type="dxa"/>` : '')
    + (tblBorders ? `<w:tblBorders>${tblBorders}</w:tblBorders>` : '')
    + '<w:tblLayout w:type="fixed"/>'
    + (t.cellMar === null ? '' : `<w:tblCellMar><w:left w:w="${mar.left ?? 0}" w:type="dxa"/><w:right w:w="${mar.right ?? 0}" w:type="dxa"/></w:tblCellMar>`)
    + (t.look ?? '')
    + `</w:tblPr><w:tblGrid>${t.grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`;
}

const leadingLines = (style) => DOCX.TYPE[style ?? 'text'].leading / DOCX.LINE;

/** Section columns: { num, space } for equal columns, or { widths: [w...],
 * space } for unequal ones (space = gutter after each column but the last). */
export function colsXml(cols) {
  if (!cols) return '';
  if (cols.raw) return cols.raw;
  if (cols.widths) {
    const n = cols.widths.length;
    return `<w:cols w:num="${n}" w:space="${cols.space}" w:equalWidth="0">${cols.widths
      .map((w, i) => `<w:col w:w="${w}"${i < n - 1 ? ` w:space="${cols.space}"` : ''}/>`).join('')}</w:cols>`;
  }
  return `<w:cols w:num="${cols.num}" w:space="${cols.space}"/>`;
}

/** sectPr body. section: { type?, cols?, textDirection?, docGrid?, refs?,
 * pgNumType?, titlePg? } */
export function sectPrXml(section = {}) {
  const { PAGE_W, PAGE_H, MARGIN } = DOCX;
  return `<w:sectPr>${section.refs ?? ''}`
    + (section.type ? `<w:type w:val="${section.type}"/>` : '')
    + `<w:pgSz w:w="${PAGE_W}" w:h="${PAGE_H}"/>`
    + `<w:pgMar w:top="${MARGIN.top}" w:right="${MARGIN.right}" w:bottom="${MARGIN.bottom}" w:left="${MARGIN.left}" w:header="${section.headerDist ?? 0}" w:footer="${section.footerDist ?? 0}" w:gutter="0"/>`
    + (section.pgNumType ?? '')
    + colsXml(section.cols)
    + (section.titlePg ? '<w:titlePg/>' : '')
    + (section.textDirection ? `<w:textDirection w:val="${section.textDirection}"/>` : '')
    + (section.docGrid ?? '')
    + '</w:sectPr>';
}

/** Blocks:
 *  { kind: 'p', at?, before?, after? (grid lines), style?, bold?, rPr?, runs? |
 *    text? | lines, cols?, ind?, tabs?, jc?, numPr?, pBdr?, shd?, pPrMid?,
 *    height? (lines, for runs), extraLines? (lines added by borders),
 *    columnBreakBefore?, pageBreakBefore?, lineAuto? (single, not exact),
 *    markRPr? (paragraph-mark run properties) }
 *  { kind: 'table', at?, grid, rows: [{ lines, cells }] }
 *  { kind: 'raw', xml, at?, height? (lines; unknown when absent) }   body-level
 *    XML written as is (tables the composer cannot express: autofit, nested,
 *    repeated header rows); offsets before it go to the previous paragraph.
 *  { kind: 'section', ...section }   ends the running section; the next
 *    section takes these properties (its sectPr is written by the composer).
 * `at` is the 1-based grid line the block starts on. Distances become space
 * before, or space after the previous paragraph across a section start, a
 * column break or before a table. Returns { body, lastSection }. */
export function composeBody(blocks, firstSection = {}) {
  let line = 1;
  let sectionTop = 1;
  let sectionBottom = 1;
  let current = firstSection;
  let boundary = false;
  const out = [];
  blocks.forEach((block, index) => {
    if (block.kind === 'section') {
      const prev = out[out.length - 1];
      if (!prev || prev.kind !== 'p') throw new Error('a section must end with a paragraph');
      prev.sectPr = sectPrXml(current);
      current = block;
      line = Number.isNaN(line) ? NaN : Math.max(line, sectionBottom);
      sectionTop = line;
      sectionBottom = line;
      boundary = true;
      return;
    }
    if (block.columnBreakBefore) {
      sectionBottom = Math.max(sectionBottom, line);
      line = sectionTop;
      boundary = true;
    }
    if (block.pageBreakBefore) {
      line = 1;
      sectionTop = 1;
      sectionBottom = 1;
      boundary = true;
    }
    const own = block.before ?? 0;
    const at = block.at ?? line + own;
    if (block.at !== undefined && Number.isNaN(line)) throw new Error(`block ${index}: 'at' after a naturally wrapped paragraph`);
    const gap = Number.isNaN(line) ? 0 : at - line - own;
    if (gap < 0) throw new Error(`block ${index} at line ${at} overlaps the previous block (next free line ${line})`);
    const height = block.kind === 'raw' ? block.height ?? NaN
      : block.kind === 'table'
      ? block.rows.reduce((sum, row) => sum + row.lines, 0)
      : block.text !== undefined ? NaN
        : block.runs ? block.height ?? NaN
          : block.lines.length * leadingLines(block.style) + (block.extraLines ?? 0);
    const after = block.after ?? 0;
    if (at + height + after - 1 > DOCX.LINES) throw new Error(`block ${index} runs past line ${DOCX.LINES}`);
    let before = own;
    if (gap > 0) {
      const prev = out[out.length - 1];
      if (block.kind === 'table' || block.kind === 'raw' || boundary) {
        if (!prev || prev.kind !== 'p') throw new Error(`block ${index}: no paragraph to carry the offset`);
        prev.after += gap;
      } else before += gap;
    }
    out.push(block.kind === 'table' || block.kind === 'raw' ? { kind: block.kind, block } : { kind: 'p', block, before, after });
    boundary = false;
    line = at + height + after;
    sectionBottom = Math.max(sectionBottom, line);
  });
  const body = out.map((item) => (item.kind === 'raw' ? item.block.xml : item.kind === 'table'
    ? tableXml(item.block)
    : paragraphXml(item.block, item.before * DOCX.LINE, item.after * DOCX.LINE, item.sectPr))).join('');
  return { body, lastSection: current };
}

/** Header furniture: id, title and what to verify, in the top-left field
 * (row 1; the title may run into column 3). Content starts at row 2. */
export function header({ id, title, verify }) {
  return [
    { kind: 'p', style: 'caption', bold: true, lines: [id], cols: [1, 2] },
    { kind: 'p', style: 'title', lines: [title], cols: [1, 3] },
    { kind: 'p', style: 'caption', lines: verify, cols: [1, 2], before: 1 },
  ];
}

/** Header furniture in its own single-column section, for documents whose
 * body is set in columns. The body section starts continuously at row 2. */
export function headerSection(furniture, bodySection) {
  const blocks = header(furniture);
  const used = 1 + 2 + 1 + furniture.verify.length;
  blocks[blocks.length - 1].after = DOCX.rowLine(2) - 1 - used;
  return [...blocks, { kind: 'section', type: 'continuous', ...bodySection }];
}

const STYLE_TEXT = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/>'
  + `<w:rPr><w:sz w:val="${DOCX.TYPE.caption.sz}"/></w:rPr></w:style>`
  + '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>'
  + `<w:pPr><w:spacing w:line="${DOCX.TYPE.title.leading}" w:lineRule="exact"/></w:pPr>`
  + `<w:rPr><w:b/><w:sz w:val="${DOCX.TYPE.title.sz}"/></w:rPr></w:style>`;

const stylesXml = (extra = '') => XML_DECL + `<w:styles ${W}>`
  + '<w:docDefaults>'
  + `<w:rPrDefault><w:rPr><w:rFonts w:ascii="${FONT.latin}" w:eastAsia="${FONT.eastAsia}" w:hAnsi="${FONT.latin}" w:cs="${FONT.latin}"/>`
  + `<w:sz w:val="${DOCX.TYPE.text.sz}"/><w:lang w:val="${LANG.latin}" w:eastAsia="${LANG.eastAsia}"/></w:rPr></w:rPrDefault>`
  + `<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="${DOCX.TYPE.text.leading}" w:lineRule="exact"/></w:pPr></w:pPrDefault>`
  + '</w:docDefaults>'
  + STYLE_TEXT + extra
  + '</w:styles>';

// compatibilityMode 15 selects current Word layout instead of the Word 2007
// compatibility mode a settings-less document opens in.
const settingsXml = (extra = '', after = '') => XML_DECL + `<w:settings ${W}>${extra}<w:compat>`
  + '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
  + `</w:compat>${after}</w:settings>`;

/** Build a package. extra: { styles?: xml, settings?: xml (before w:compat),
 * settingsAfter?: xml (after w:compat: themeFontLang, clrSchemeMapping, ...),
 * firstSection?: section, parts?: { path: xml|bytes }, rels?: [[id, type,
 * target]] (from word/document.xml), overrides?: { part: contentType },
 * externalRels?: [[id, type, url]], defaults?: { ext: contentType }, ns?: ['xmlns:x="…"', …] } */
export function docxPackage(blocks, extra = {}) {
  const { body, lastSection } = composeBody(blocks, extra.firstSection);
  const ns = [W, ...(extra.ns ?? [])].join(' ');
  const document = XML_DECL + `<w:document ${ns}><w:body>${body}${sectPrXml(lastSection)}</w:body></w:document>`;
  const types = contentTypes({
    'word/document.xml': `${WML}.document.main+xml`,
    'word/styles.xml': `${WML}.styles+xml`,
    'word/settings.xml': `${WML}.settings+xml`,
    ...(extra.overrides ?? {}),
  });
  const defaults = Object.entries(extra.defaults ?? {})
    .map(([ext, type]) => `<Default Extension="${ext}" ContentType="${type}"/>`).join('');
  return {
    '[Content_Types].xml': defaults ? types.replace('<Default Extension="rels"', `${defaults}<Default Extension="rels"`) : types,
    '_rels/.rels': relationships([['rId1', 'officeDocument', 'word/document.xml']]),
    'word/_rels/document.xml.rels': relationships([
      ['rId1', 'styles', 'styles.xml'],
      ['rId2', 'settings', 'settings.xml'],
      ...(extra.rels ?? []),
      ...(extra.externalRels ?? []).map(([id, type, target]) => [id, type, target, 'External']),
    ]),
    'word/document.xml': document,
    'word/styles.xml': stylesXml(extra.styles),
    'word/settings.xml': settingsXml(extra.settings, extra.settingsAfter),
    ...(extra.parts ?? {}),
  };
}

export { WML };
