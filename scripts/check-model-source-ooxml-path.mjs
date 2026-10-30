#!/usr/bin/env node
// Compare the ordinary OOXML await path with the merge base using the AST.
// Source-only branches are excluded; every remaining await must keep its
// original order and callee, with the bounded opt-in font-owner refactor below.
// The XLSX render worker also retains one host.run
// around archive construction and parse, as in the previous renderer.
// This regression check covers accidental edits, not intentionally hostile code.
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parse } from '@babel/parser';

const git = (args) => execFileSync('git', args, { encoding: 'utf8' }).trimEnd();
const base = git(['merge-base', 'HEAD', 'origin/main']);
// The opt-in source option keys: the former list, the legacy readers and the
// OpenDocument readers.
const SOURCE_KEY = '(?:modelSources|legacyDoc|legacyXls|legacyPpt|odt|ods|odp)';
const SOURCE_ABSENT = new RegExp(`^options\\.${SOURCE_KEY}===undefined$`);
// A gate may name several source options of its format together:
// `(opts.legacyDoc !== undefined || opts.odt !== undefined)`.
const SOURCE_PRESENT_TERM = `(?:opts|options)\\.${SOURCE_KEY}!==undefined`;
const SOURCE_PRESENT = new RegExp(
  `^(?:__OOXML_MODEL_SOURCES__&&)?(?:${SOURCE_PRESENT_TERM}|\\(${SOURCE_PRESENT_TERM}(?:\\|\\|${SOURCE_PRESENT_TERM})+\\))$`,
);
const cases = [
  ['packages/docx/src/internal/node-acquisition.ts', 'acquireDocxNodeDocument'],
  ['packages/xlsx/src/internal/node-acquisition.ts', 'acquireXlsxNodeSession'],
  ['packages/pptx/src/internal/node-acquisition.ts', 'acquirePptxNodeSession'],
  ['packages/node/src/docx.ts', 'openDocxDocument'],
  ['packages/node/src/docx.ts', 'materializeDocxDocument'],
  ['packages/node/src/xlsx.ts', 'openXlsxWorkbook'],
  ['packages/node/src/pptx.ts', 'openPptxPresentationImpl'],
  ['packages/docx/src/document.ts', 'load'],
  ['packages/xlsx/src/workbook.ts', 'load'],
  ['packages/pptx/src/presentation.ts', 'load'],
];

function children(node) {
  return Object.values(node).flatMap((value) =>
    Array.isArray(value) ? value.filter((child) => child?.type)
      : value?.type ? [value] : []);
}

function findFunction(ast, name) {
  let found;
  function walk(node) {
    if (found) return;
    if ((node.type === 'FunctionDeclaration' && node.id?.name === name)
      || ((node.type === 'ClassMethod' || node.type === 'ClassPrivateMethod') && node.key?.name === name)) {
      found = node;
      return;
    }
    children(node).forEach(walk);
  }
  walk(ast.program);
  if (!found) throw new Error(`Missing function ${name}`);
  return found;
}

function callName(node) {
  if (!node) return '';
  if (node.type === 'CallExpression' || node.type === 'OptionalCallExpression') return callName(node.callee);
  if (node.type === 'MemberExpression' || node.type === 'OptionalMemberExpression') {
    return `${callName(node.object)}.${callName(node.property)}`;
  }
  if (node.type === 'Identifier') return node.name;
  if (node.type === 'Import') return 'import';
  if (node.type === 'NewExpression') return `new ${callName(node.callee)}`;
  return node.type;
}

function requiresGoogleFonts(node) {
  if (node.type === 'LogicalExpression' && node.operator === '&&') {
    return requiresGoogleFonts(node.left) || requiresGoogleFonts(node.right);
  }
  return node.type === 'MemberExpression' && !node.computed
    && node.object.type === 'Identifier' && node.object.name === 'opts'
    && node.property.type === 'Identifier' && node.property.name === 'useGoogleFonts';
}

function ooxmlAwaits(node, code, pptxFontOwnerRefactor = false) {
  const awaits = [];
  function walk(current, googleFontsRequired = false) {
    if (current.type === 'BlockStatement') {
      for (const statement of current.body) {
        if (statement.type === 'IfStatement') {
          const test = code.slice(statement.test.start, statement.test.end).replace(/\s+/g, '');
          if (SOURCE_ABSENT.test(test)) {
            walk(statement.consequent, googleFontsRequired);
            return;
          }
        }
        walk(statement, googleFontsRequired);
      }
      return;
    }
    if (current.type === 'IfStatement') {
      const test = code.slice(current.test.start, current.test.end).replace(/\s+/g, '');
      // The internal comparison-build flag can only narrow the existing
      // selected-source branch; the ordinary OOXML await path is unchanged.
      if (SOURCE_PRESENT.test(test)) {
        if (current.alternate) walk(current.alternate, googleFontsRequired);
        return;
      }
      if (SOURCE_ABSENT.test(test)) {
        walk(current.consequent, googleFontsRequired);
        return;
      }
      if (/^!sourceLoad$/.test(test)) {
        walk(current.consequent, googleFontsRequired);
        return;
      }
      walk(current.test, googleFontsRequired);
      walk(current.consequent, googleFontsRequired || requiresGoogleFonts(current.test));
      if (current.alternate) walk(current.alternate, googleFontsRequired);
      return;
    }
    if (current.type === 'ConditionalExpression') {
      const test = code.slice(current.test.start, current.test.end).replace(/\s+/g, '');
      if (SOURCE_ABSENT.test(test)) {
        walk(current.consequent, googleFontsRequired);
        return;
      }
      if (test === 'sourceLoad' || test === 'this._sourceLoad' || test.startsWith('sourceLoad&&')) {
        walk(current.alternate, googleFontsRequired);
        return;
      }
    }
    if (current.type === 'AwaitExpression') {
      let name = callName(current.argument);
      // The PPTX font barrier moved from a per-call array owner to the shared
      // presentation lease. Normalize only this known callee transition under
      // its positive opt-in AND gate. Every await is still visited, so additions
      // or reordering inside the gate fail; unguarded/OR/alternate calls and
      // every other package/helper retain their original identity checks.
      if (pptxFontOwnerRefactor && googleFontsRequired && name === 'pres._ensureGoogleFonts') {
        name = 'preloadGoogleFonts';
      }
      awaits.push(name);
    }
    children(current).forEach(child => walk(child, googleFontsRequired));
  }
  walk(node.body);
  return awaits;
}

const NATIVE_ROUTE_BARRIER = 'nativeRoutesPending';

export function auditAwaitCase(file, name, previous, current) {
  const pptxFontOwnerRefactor = file === 'packages/pptx/src/presentation.ts' && name === 'load';
  const baseline = ooxmlAwaits(findFunction(parse(previous, { sourceType: 'module', plugins: ['typescript', 'jsx'] }), name), previous, pptxFontOwnerRefactor);
  let candidate = ooxmlAwaits(findFunction(parse(current, { sourceType: 'module', plugins: ['typescript', 'jsx'] }), name), current, pptxFontOwnerRefactor);
  // DOCX issue #1843 adds one font-preparation barrier (awaiting the native
  // route line measurement started before font loading) between math
  // preparation and the first publication. Accept
  // exactly that insertion at exactly that position while the merge base
  // predates it; any other addition, removal or reordering still fails.
  if (file === 'packages/docx/src/document.ts' && name === 'load'
    && !baseline.includes(NATIVE_ROUTE_BARRIER)) {
    const index = candidate.indexOf(NATIVE_ROUTE_BARRIER);
    if (index > 0 && candidate[index - 1] === 'prepareMathRuns'
      && candidate.lastIndexOf(NATIVE_ROUTE_BARRIER) === index) {
      candidate = [...candidate.slice(0, index), ...candidate.slice(index + 1)];
    }
  }
  // DOCX font sources (LoadOptions.fonts) replace the Google-only preload
  // barrier with the generic font-source barrier at the same position. Accept
  // exactly that one rename while the merge base still awaits the old callee;
  // any other addition, removal or reordering still fails.
  if (file === 'packages/docx/src/document.ts' && name === 'load') {
    const index = baseline.indexOf('preloadGoogleFonts');
    if (index >= 0 && candidate[index] === 'doc._loadFontSources'
      && candidate.filter((entry) => entry === 'doc._loadFontSources').length === 1) {
      candidate = [...candidate.slice(0, index), 'preloadGoogleFonts', ...candidate.slice(index + 1)];
    }
  }
  // PPTX font sources (LoadOptions.fonts) replace the Google-only presentation
  // lease barrier with the generic font-source barrier at the same position
  // (its opt-in gate is now `opts.fonts`, so the old callee normalization does
  // not apply). Accept exactly that one rename while the merge base still
  // awaits the old callee; any other addition, removal or reordering fails.
  if (pptxFontOwnerRefactor) {
    const index = baseline.indexOf('preloadGoogleFonts');
    if (index >= 0 && candidate[index] === 'pres._ensureFontSources'
      && candidate.filter((entry) => entry === 'pres._ensureFontSources').length === 1) {
      candidate = [...candidate.slice(0, index), 'preloadGoogleFonts', ...candidate.slice(index + 1)];
    }
  }
  if (JSON.stringify(candidate) !== JSON.stringify(baseline)) {
    throw new Error(`${file} ${name}: OOXML awaits changed\nmain ${JSON.stringify(baseline)}\nhead ${JSON.stringify(candidate)}`);
  }
  return candidate.length;
}

export function hasCombinedXlsxHostRun(code) {
  const ast = parse(code, { sourceType: 'module', plugins: ['typescript', 'jsx'] });
  let found = false;
  function walk(node) {
    if (node.type === 'CallExpression' && callName(node.callee) === 'host.run') {
      const body = code.slice(node.start, node.end);
      if (/new XlsxArchive\(/.test(body) && /archive\.parse\(\)/.test(body)) found = true;
    }
    children(node).forEach(walk);
  }
  walk(ast.program);
  return found;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  for (const [file, name] of cases) {
    const count = auditAwaitCase(file, name, git(['show', `${base}:${file}`]), readFileSync(file, 'utf8'));
    console.log(`${file} ${name}: ${count} OOXML awaits, unchanged`);
  }
  const xlsxRender = readFileSync('packages/xlsx/src/render-worker.ts', 'utf8');
  if (!hasCombinedXlsxHostRun(xlsxRender)) {
    throw new Error('XLSX render worker splits OOXML construction and parse across host.run calls');
  }
  console.log('XLSX render worker keeps construction and parse in one host.run.');
}
