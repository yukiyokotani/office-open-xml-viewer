// Office Math Markup (ECMA-376 Part 1, §22.1) builders for the WordprocessingML
// specimens. Each builder returns OMML for one object, with children in schema
// order; arguments that are strings are math text, arrays are concatenated.

import { esc } from './xml.mjs';

export const M_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

const v = (name, val) => `<m:${name} m:val="${esc(String(val))}"/>`;
const arg = (x) => (Array.isArray(x) ? x.map(arg).join('') : x.startsWith('<') ? x : r(x));
const el = (name, x) => `<m:${name}>${arg(x)}</m:${name}>`;
const pr = (name, props) => (props ? `<m:${name}Pr>${props}</m:${name}Pr>` : '');

/** Math run. opts: { sty: p|b|i|bi, scr: roman|script|fraktur|double-struck|sans-serif|monospace, nor } */
export function r(text, opts = {}) {
  const rPr = (opts.nor ? '<m:nor/>' : '') + (opts.scr ? v('scr', opts.scr) : '') + (opts.sty ? v('sty', opts.sty) : '');
  const wrPr = MATH.sz ? `<w:rPr><w:sz w:val="${MATH.sz}"/></w:rPr>` : '';
  return `<m:r>${rPr ? `<m:rPr>${rPr}</m:rPr>` : ''}${wrPr}<m:t xml:space="preserve">${esc(text)}</m:t></m:r>`;
}

/** Module setting: math run size in half-points (w:rPr w:sz), or undefined. */
export const MATH = { sz: undefined };
/** Em space: math ignores ordinary spaces between objects. */
export const SP = '\u2003';
export const f = (num, den, type) => `<m:f>${pr('f', type ? v('type', type) : '')}${el('num', num)}${el('den', den)}</m:f>`;
export const sub = (e, s) => `<m:sSub>${el('e', e)}${el('sub', s)}</m:sSub>`;
export const sup = (e, s) => `<m:sSup>${el('e', e)}${el('sup', s)}</m:sSup>`;
export const subSup = (e, s, p) => `<m:sSubSup>${el('e', e)}${el('sub', s)}${el('sup', p)}</m:sSubSup>`;
export const pre = (s, p, e) => `<m:sPre>${el('sub', s)}${el('sup', p)}${el('e', e)}</m:sPre>`;
export const rad = (e, deg) => `<m:rad>${pr('rad', deg === undefined ? v('degHide', 1) : '')}${el('deg', deg ?? [])}${el('e', e)}</m:rad>`;
/** opts: { limLoc: undOvr|subSup, sub, sup (omit to hide), grow } */
export function nary(chr, e, opts = {}) {
  const props = v('chr', chr) + (opts.limLoc ? v('limLoc', opts.limLoc) : '') + (opts.grow ? v('grow', 1) : '')
    + (opts.sub === undefined ? v('subHide', 1) : '') + (opts.sup === undefined ? v('supHide', 1) : '');
  return `<m:nary>${pr('nary', props)}${el('sub', opts.sub ?? [])}${el('sup', opts.sup ?? [])}${el('e', e)}</m:nary>`;
}
/** Delimiter. es: one element or an array of separated elements. */
export function d(es, beg = '(', end = ')', opts = {}) {
  const props = (beg !== '(' ? v('begChr', beg) : '') + (opts.sep ? v('sepChr', opts.sep) : '') + (end !== ')' ? v('endChr', end) : '')
    + (opts.grow === false ? v('grow', 0) : '') + (opts.shp ? v('shp', opts.shp) : '');
  const list = Array.isArray(es) && es.some(Array.isArray) ? es : [es];
  return `<m:d>${pr('d', props)}${list.map((e) => el('e', e)).join('')}</m:d>`;
}
export const acc = (chr, e) => `<m:acc>${pr('acc', v('chr', chr))}${el('e', e)}</m:acc>`;
export const bar = (e, pos = 'top') => `<m:bar>${pr('bar', v('pos', pos))}${el('e', e)}</m:bar>`;
export const groupChr = (chr, e, pos = 'bot', vertJc) => `<m:groupChr>${pr('groupChr', v('chr', chr) + v('pos', pos) + (vertJc ? v('vertJc', vertJc) : ''))}${el('e', e)}</m:groupChr>`;
export const func = (name, e) => `<m:func>${el('fName', name.startsWith('<') ? name : r(name, { sty: 'p' }))}${el('e', e)}</m:func>`;
export const limLow = (e, lim) => `<m:limLow>${el('e', e)}${el('lim', lim)}</m:limLow>`;
export const limUpp = (e, lim) => `<m:limUpp>${el('e', e)}${el('lim', lim)}</m:limUpp>`;
/** Matrix: rows of cells; jc per column (left|center|right). */
export function matrix(rows, jc) {
  const mcs = jc ? `<m:mcs>${jc.map((j) => `<m:mc><m:mcPr>${v('count', 1)}${v('mcJc', j)}</m:mcPr></m:mc>`).join('')}</m:mcs>` : '';
  return `<m:m>${pr('m', mcs)}${rows.map((row) => `<m:mr>${row.map((c) => el('e', c)).join('')}</m:mr>`).join('')}</m:m>`;
}
export const eqArr = (rows) => `<m:eqArr>${rows.map((e) => el('e', e)).join('')}</m:eqArr>`;
export const box = (e) => `<m:box>${el('e', e)}</m:box>`;
export const borderBox = (e) => `<m:borderBox>${el('e', e)}</m:borderBox>`;
/** Phantom: occupies the space of e without drawing it. */
export const phant = (e) => `<m:phant>${pr('phant', v('show', 0))}${el('e', e)}</m:phant>`;
/** Alignment point inside an equation array row. */
export const aln = () => `<m:r><m:rPr><m:aln/></m:rPr>${MATH.sz ? `<w:rPr><w:sz w:val="${MATH.sz}"/></w:rPr>` : ''}<m:t>=</m:t></m:r>`;

export const oMath = (...xs) => `<m:oMath>${xs.map(arg).join('')}</m:oMath>`;
/** Display equations, one line each. jc: left|right|center|centerGroup */
export const oMathPara = (lines, jc) => `<m:oMathPara>${jc ? `<m:oMathParaPr>${v('jc', jc)}</m:oMathParaPr>` : ''}${lines.map((l) => oMath(l)).join('')}</m:oMathPara>`;
