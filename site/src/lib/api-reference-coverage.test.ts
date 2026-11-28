// Keeps the hand-written API reference (the /api/* pages, /api/*.md and
// llms-full.txt) in step with the real public API. The checked-in
// packages/*/api/public-api-baseline.d.ts files are regenerated whenever the
// public declarations change (scripts/check-public-api.mjs), so a new, renamed
// or removed option or member fails here until api-reference.ts follows.
import { fileURLToPath } from 'node:url';
import ts from 'typescript-compiler-api';
import { describe, expect, it } from 'vitest';
import { apiReference, type ApiClass } from './api-reference';

const baseline = (format: string): string =>
  fileURLToPath(new URL(`../../../packages/${format}/api/public-api-baseline.d.ts`, import.meta.url));

const formats = Object.keys(apiReference) as (keyof typeof apiReference)[];
const program = ts.createProgram(formats.map(baseline), {
  noEmit: true,
  skipLibCheck: true,
  lib: ['lib.esnext.d.ts', 'lib.dom.d.ts'],
});
const checker = program.getTypeChecker();

/**
 * Public members the reference intentionally leaves out. Each needs a reason;
 * prefer documenting a member over listing it here.
 */
const undocumentedMembers: Readonly<Record<string, readonly string[]>> = {};

/**
 * Options the reference intentionally leaves out of a class, with the reason.
 * `xlsxWorksheetLimits` sits on the shared LoadOptions but only XLSX reads it
 * (packages/core/src/types/load-options.ts); DOCX and PPTX ignore it.
 */
const ignoredOptions: Readonly<Record<'docx' | 'xlsx' | 'pptx', readonly string[]>> = {
  docx: ['xlsxWorksheetLimits'],
  pptx: ['xlsxWorksheetLimits'],
  xlsx: [],
};

function classSymbol(format: string, name: string): ts.Symbol {
  const source = program.getSourceFile(baseline(format));
  const moduleSymbol = source && checker.getSymbolAtLocation(source);
  const symbol = moduleSymbol && checker.getExportsOfModule(moduleSymbol).find((s) => s.name === name);
  if (!symbol) throw new Error(`${name} is not exported by the ${format} public API baseline`);
  return symbol;
}

const isPublic = (symbol: ts.Symbol): boolean => {
  const declaration = symbol.declarations?.[0];
  if (!declaration) return false;
  const flags = ts.getCombinedModifierFlags(declaration);
  return !(flags & (ts.ModifierFlags.Private | ts.ModifierFlags.Protected))
    && !symbol.name.startsWith('#')
    && !symbol.name.startsWith('__');
};

function actualApi(format: string, name: string): { options: string[]; members: string[] } {
  const symbol = classSymbol(format, name);
  const declaration = symbol.valueDeclaration ?? symbol.declarations![0]!;
  const staticSide = checker.getTypeOfSymbolAtLocation(symbol, declaration);
  const instanceSide = checker.getDeclaredTypeOfSymbol(symbol);
  const members = [
    ...checker.getPropertiesOfType(instanceSide),
    ...checker.getPropertiesOfType(staticSide).filter((s) => s.name !== 'prototype'),
  ].filter(isPublic).map((s) => s.name);

  // Engines document `static load(source, options)`; Viewers their constructor.
  const load = staticSide.getProperty('load');
  const signatures = load
    ? checker.getTypeOfSymbolAtLocation(load, declaration).getCallSignatures()
    : staticSide.getConstructSignatures();
  const optionsParameter = signatures.flatMap((signature) => signature.parameters.slice(-1))
    .find((parameter) => parameter.name === 'options' || parameter.name === 'opts');
  const options = optionsParameter
    ? checker.getPropertiesOfType(checker.getNonNullableType(
      checker.getTypeOfSymbolAtLocation(optionsParameter, declaration),
    )).map((s) => s.name)
    : [];
  return { options, members };
}

// Rows such as `zoomMin / zoomMax` document a pair of options together.
const documentedOptions = (apiClass: ApiClass): string[] =>
  (apiClass.options ?? []).flatMap((option) => option.name.split(' / '));

// The member name is the first identifier after modifiers: `static load(…)`,
// `get pageCount: number`, `await doc.toMarkdown()` all name their member.
const documentedMembers = (apiClass: ApiClass): string[] =>
  apiClass.methods.map((method) => {
    const name = method.sig
      .replace(/^(?:(?:static|readonly|get|set|async|await)\s+)+/, '')
      .replace(/^[A-Za-z]+\./, '')
      .match(/^[A-Za-z_$][\w$]*/)?.[0];
    if (!name) throw new Error(`Cannot read a member name from ${method.sig}`);
    return name;
  });

describe('official-site API reference coverage', () => {
  for (const format of formats) {
    for (const apiClass of apiReference[format]) {
      it(`${apiClass.name} documents exactly its public options`, () => {
        const { options } = actualApi(format, apiClass.name);
        expect(options.length, `${apiClass.name} options did not resolve`).toBeGreaterThan(0);
        const documented = new Set(documentedOptions(apiClass));
        const ignored = new Set(ignoredOptions[format]);
        expect(options.filter((o) => !documented.has(o) && !ignored.has(o)).sort().join(', '), 'undocumented options').toBe('');
        expect([...documented].filter((o) => !options.includes(o)).sort().join(', '), 'stale options').toBe('');
      });

      it(`${apiClass.name} documents every public member`, () => {
        const { members } = actualApi(format, apiClass.name);
        const documented = new Set(documentedMembers(apiClass));
        const skipped = new Set(undocumentedMembers[apiClass.name] ?? []);
        expect(members.filter((m) => !documented.has(m) && !skipped.has(m)).sort().join(', '), 'undocumented members').toBe('');
        expect([...documented].filter((m) => !members.includes(m)).sort().join(', '), 'stale members').toBe('');
      });
    }
  }
});
