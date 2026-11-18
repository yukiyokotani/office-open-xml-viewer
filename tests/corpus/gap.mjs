#!/usr/bin/env node
// Coverage gap between the public specimen corpus and a local private corpus.
//
//   node tests/corpus/gap.mjs [--format docx|xlsx|pptx] [--part REGEX] [--top N] [--list-all] [--all]
//
// By default, generator artefacts the corpus omits on purpose (BOILERPLATE
// below) are excluded from both the totals and the listing; --all includes
// them. --part restricts the listing (not the totals) to part kinds matching
// REGEX. --top N limits the listing to N groups; --list-all lists every group.
//
// Private samples (packages/<fmt>/public/private/<fmt>/, top level only) are
// reduced to STRUCTURAL schema vocabulary with tests/visual/corpus-features.mjs:
// element nesting, attribute names, relationship types and binary part
// extensions. Every `attr=value` feature is dropped so nothing content-like is
// reported, and no file name is ever printed. Output goes to stdout only.
// Third parties without a private corpus get a short notice and exit 0.

import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { listCorpus } from '../visual/corpus-features.mjs';
import { FORMATS, REPO_ROOT, buildCase, loadCases } from './lib/cases.mjs';
import { packageFeatures } from './lib/features.mjs';

const args = process.argv.slice(2);
const opt = (name) => { const i = args.indexOf(name); return i < 0 ? undefined : args[i + 1]; };
const formats = opt('--format') ? [opt('--format')] : FORMATS;
const top = args.includes('--list-all') ? Infinity : Number(opt('--top') ?? 40);
const includeBoilerplate = args.includes('--all');

// Vocabulary the corpus leaves out on purpose because no renderer draws it:
// editing-session ids, proofing and view state, document properties, data the
// application keeps for its own UI. Matched against `part|selector`. Keep this
// list short and conservative: anything that can change what is drawn (fonts,
// themes, styles, settings such as compat, defaultTabStop, characterSpacing-
// Control) stays in.
const BOILERPLATE = [
  [/^docprops\//, 'document properties (title, author, statistics, thumbnail)'],
  [/^customxml\//, 'custom XML data parts'],
  [/^word\/(websettings|people|stylesWithEffects|commentsids|commentsextensible|intelligence\w*)\.xml\|/i,
    'web settings, co-author list, Word 2010 style copy, comment ids, Editor data'],
  [/^word\/glossary\//, 'building-block glossary (not part of the document body)'],
  [/^ppt\/(viewprops|presprops|commentauthors|revisioninfo)\.xml\|/, 'view and presentation properties, comment authors, revision info'],
  [/^ppt\/(changesinfos|webextensions)\//, 'co-authoring change records, web add-ins'],
  [/^xl\/(calcchain\.xml|printersettings\/)/, 'calculation chain, printer driver blobs'],
  [/[}@]rsid\w*($|[=>@])|\}rsids?>|\}rsidRoot/, 'revision-save ids (w:rsid*)'],
  [/\}(paraId|textId|docId|creationId|modId|uid|revisionPtr|absPath|anchorId|editId)($|[=@>])/, 'paragraph, document, object and revision identifiers'],
  [/\}(latentStyles|lsdException)/, 'latent style table (UI defaults for unused styles)'],
  [/\}proofErr|\}proofState|\}noProof|\}lastRenderedPageBreak|\}(zoom|decimalSymbol|listSeparator|fileVersion|calcPr)($|[=@>])/, 'proofing marks, Word\'s last-layout page-break hints, zoom, list/decimal symbols, file version, calc settings'],
  [/\}themeFamily|@\{http:\/\/schemas\.openxmlformats\.org\/markup-compatibility\/2006\}Ignorable/, 'theme family id, mc:Ignorable declarations'],
];
const isBoilerplate = (f) => BOILERPLATE.some(([re]) => re.test(f));
const partFilter = opt('--part') ? new RegExp(opt('--part')) : null;

/** Nesting, attribute-name, relationship and binary-part features only. */
export const structural = (feature) => {
  const selector = feature.slice(feature.indexOf('|') + 1);
  const at = selector.lastIndexOf('@');
  return !(at >= 0 && selector.indexOf('=', at) >= 0);
};

const local = (qname) => qname.replace(/^\{[^}]*\}/, '');
/** Element family: the element a feature is about (child of a nesting, owner of an attribute). */
function family(feature) {
  const selector = feature.slice(feature.indexOf('|') + 1);
  if (selector.startsWith('rel>')) return 'relationships';
  if (selector.startsWith('*.')) return 'binary parts';
  const at = selector.indexOf('@');
  if (at >= 0) return local(selector.slice(0, at));
  return local(selector.slice(selector.indexOf('>') + 1));
}
const NS_SHORT = [
  ['http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w'],
  ['http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'x'],
  ['http://schemas.openxmlformats.org/drawingml/2006/main', 'a'],
  ['http://schemas.openxmlformats.org/presentationml/2006/main', 'p'],
  ['http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'r'],
  ['http://schemas.openxmlformats.org/drawingml/2006/chart', 'c'],
  ['http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'wp'],
  ['http://schemas.openxmlformats.org/drawingml/2006/picture', 'pic'],
  ['http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'xdr'],
];
const short = (text) => NS_SHORT.reduce((s, [uri, p]) => s.split(`{${uri}}`).join(`${p}:`), text);

const cases = await loadCases();
for (const format of formats) {
  const dir = join(REPO_ROOT, 'packages', format, 'public', 'private', format);
  if (!existsSync(dir)) {
    console.log(`${format}: no private corpus at packages/${format}/public/private/${format}/ — nothing to compare.`);
    continue;
  }
  const files = listCorpus(dir, format);
  if (files.length === 0) { console.log(`${format}: private corpus is empty.`); continue; }
  const usage = new Map(); // feature → number of private files using it
  let unreadable = 0;
  for (const file of files) {
    let features;
    try { features = packageFeatures(format, readFileSync(join(dir, file))); } catch { unreadable += 1; continue; }
    for (const f of features) {
      if (structural(f) && (includeBoilerplate || !isBoilerplate(f))) usage.set(f, (usage.get(f) ?? 0) + 1);
    }
  }
  const publicFeatures = new Set();
  for (const c of cases.filter((x) => x.format === format)) {
    for (const f of packageFeatures(format, buildCase(c))) if (structural(f)) publicFeatures.add(f);
  }
  const covered = [...usage.keys()].filter((f) => publicFeatures.has(f)).length;
  const weightTotal = [...usage.values()].reduce((a, b) => a + b, 0);
  const weightCovered = [...usage].filter(([f]) => publicFeatures.has(f)).reduce((a, [, n]) => a + n, 0);
  console.log(`\n== ${format}: ${files.length - unreadable} private files${unreadable ? ` (${unreadable} unreadable)` : ''}`
    + (includeBoilerplate ? ', boilerplate included' : ', boilerplate excluded (--all to include)'));
  console.log(`   structural features covered ${covered} / ${usage.size}`
    + `  (${(100 * covered / usage.size).toFixed(1)} %; weighted by file use ${(100 * weightCovered / weightTotal).toFixed(1)} %)`);

  // Group uncovered features by part kind and element family; rank groups by
  // the most-used feature in them.
  const groups = new Map();
  for (const [f, n] of usage) {
    if (publicFeatures.has(f)) continue;
    if (partFilter && !partFilter.test(f.slice(0, f.indexOf('|')))) continue;
    const key = `${f.slice(0, f.indexOf('|'))}  ${family(f)}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push([f, n]);
  }
  const ranked = [...groups].map(([key, list]) => [key, list.sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : 1))])
    .sort((a, b) => b[1][0][1] - a[1][0][1] || b[1].length - a[1].length || (a[0] < b[0] ? -1 : 1));
  for (const [key, list] of ranked.slice(0, top)) {
    console.log(`   ${String(list[0][1]).padStart(4)}  ${key}  (${list.length} features)`);
    for (const [f, n] of list.slice(0, 4)) console.log(`         ${String(n).padStart(4)}  ${short(f.slice(f.indexOf('|') + 1))}`);
  }
  if (ranked.length > top) console.log(`   … ${ranked.length - top} more groups (--all)`);
}
