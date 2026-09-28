#!/usr/bin/env node
// Compare the ordinary OOXML await path with the merge base using the AST.
// Source-only branches and the explicitly opted-in DOCX bundled-font load are
// excluded; every remaining await must keep its original order and callee.
// The XLSX render worker also retains one host.run
// around archive construction and parse, as in the previous renderer.
// This regression check covers accidental edits, not intentionally hostile code.
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parse } from '@babel/parser';

const git = (args) => execFileSync('git', args, { encoding: 'utf8' }).trimEnd();
const base = git(['merge-base', 'HEAD', 'origin/main']);
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

function ooxmlAwaits(node, code) {
  const awaits = [];
  function walk(current) {
    if (current.type === 'BlockStatement') {
      for (const statement of current.body) {
        if (statement.type === 'IfStatement') {
          const test = code.slice(statement.test.start, statement.test.end).replace(/\s+/g, '');
          if (test === 'options.modelSources===undefined') {
            walk(statement.consequent);
            return;
          }
        }
        walk(statement);
      }
      return;
    }
    if (current.type === 'IfStatement') {
      const test = code.slice(current.test.start, current.test.end).replace(/\s+/g, '');
      // The internal comparison-build flag can only narrow the existing
      // selected-source branch; the ordinary OOXML await path is unchanged.
      if (/^(?:__OOXML_MODEL_SOURCES__&&)?(?:opts|options)\.modelSources!==undefined$/.test(test)) {
        if (current.alternate) walk(current.alternate);
        return;
      }
      if (/^options\.modelSources===undefined$/.test(test)) {
        walk(current.consequent);
        return;
      }
      if (/^!sourceLoad$/.test(test)) {
        walk(current.consequent);
        return;
      }
    }
    if (current.type === 'ConditionalExpression') {
      const test = code.slice(current.test.start, current.test.end).replace(/\s+/g, '');
      // This font acquisition runs only when a caller explicitly opts in.
      // The default OOXML path still has the exact await sequence from main.
      if (test === "doc._mode==='main'&&opts.useBundledOfficeFonts") {
        walk(current.alternate);
        return;
      }
      if (test === 'options.modelSources===undefined') {
        walk(current.consequent);
        return;
      }
      if (test === 'sourceLoad' || test === 'this._sourceLoad' || test.startsWith('sourceLoad&&')) {
        walk(current.alternate);
        return;
      }
    }
    if (current.type === 'AwaitExpression') awaits.push(callName(current.argument));
    children(current).forEach(walk);
  }
  walk(node.body);
  return awaits;
}

export function auditAwaitCase(file, name, previous, current) {
  const baseline = ooxmlAwaits(findFunction(parse(previous, { sourceType: 'module', plugins: ['typescript', 'jsx'] }), name), previous);
  const candidate = ooxmlAwaits(findFunction(parse(current, { sourceType: 'module', plugins: ['typescript', 'jsx'] }), name), current);
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
