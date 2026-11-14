import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { test } from 'node:test';

import { checkFormat, scanNullProvenance } from './check-model-wire-drift.mjs';

// A serialize-mode schema in the shape schemars emits for:
//   struct Doc { body: Vec<Block>, #[skip_serializing_if] title: Option<String>,
//                color: Option<String>, kind: Kind, size: [f64; 2] }
//   #[serde(tag = "type")] enum Block { Para(Para), Rule { width: f64 } }
//   enum Kind { A, B }
const schema = {
  title: 'Doc',
  type: 'object',
  properties: {
    body: { type: 'array', items: { $ref: '#/$defs/Block' } },
    title: { type: ['string', 'null'] },
    color: { type: ['string', 'null'] },
    kind: { $ref: '#/$defs/Kind' },
    size: { type: 'array', items: { type: 'number' }, minItems: 2, maxItems: 2 },
  },
  required: ['body', 'color', 'kind', 'size'],
  $defs: {
    Block: {
      oneOf: [
        { type: 'object', properties: { type: { type: 'string', const: 'para' } }, $ref: '#/$defs/Para', required: ['type'] },
        {
          type: 'object',
          properties: { type: { type: 'string', const: 'rule' }, width: { type: 'number' } },
          required: ['type', 'width'],
        },
      ],
    },
    Para: {
      type: 'object',
      properties: { text: { type: 'string' }, children: { type: 'array', items: { $ref: '#/$defs/Block' } } },
      required: ['text', 'children'],
    },
    Kind: { oneOf: [{ type: 'string', enum: ['a'] }, { type: 'string', const: 'b', description: 'documented' }] },
  },
};

const inSync = `
export type Block = ({ type: 'para' } & Para) | { type: 'rule'; width: number };
export interface Para { text: string; children: Block[] }
export interface Doc {
  body: Block[];
  title?: string;
  color: string | null;
  kind: 'a' | 'b';
  size: [number, number];
}
`;

function run(source, allowlist = {}, rustSchema = schema) {
  const dir = mkdtempSync(path.join(tmpdir(), 'model-wire-drift-'));
  try {
    writeFileSync(path.join(dir, 'tsconfig.json'), JSON.stringify({ compilerOptions: { strict: true } }));
    writeFileSync(path.join(dir, 'types.ts'), source);
    writeFileSync(path.join(dir, 'schema.json'), JSON.stringify(rustSchema));
    return checkFormat('fixture', {
      schemaFile: path.join(dir, 'schema.json'),
      format: { tsconfig: path.join(dir, 'tsconfig.json'), tsFile: path.join(dir, 'types.ts'), tsRoot: 'Doc' },
      allowlist,
    });
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

test('matching recursive models report no difference', () => {
  const result = run(inSync);
  assert.deepEqual(result.findings, []);
  assert.deepEqual(result.stale, []);
});

test('each kind of drift is reported once with a stable key', () => {
  const drifted = inSync
    .replace("{ type: 'rule'; width: number }", "{ type: 'rule'; width: number } | { type: 'image' }")
    .replace('children: Block[] }', 'children: Block[]; depth: number }')
    .replace('title?: string;', 'title: string;')
    .replace('color: string | null;', 'color?: string;')
    .replace("kind: 'a' | 'b';", "kind: 'a';")
    .replace('size: [number, number];', 'size: [number, number, number];');
  const keys = run(drifted).findings.map((f) => f.key);
  assert.deepEqual(keys, [
    'Doc.body<type=image>: variant-never-emitted',
    'Doc.color: rust-null',
    'Doc.kind: literal-missing-in-ts ["b"]',
    'Doc.size: type-mismatch (Rust length 2..2 vs TS length 3..3)',
    'Doc.title: rust-optional',
    // The recursive union is checked again where Para nests it.
    'Para<type=para>.children<type=image>: variant-never-emitted',
    'Para<type=para>.depth: ts-only',
  ]);
});

test('a field Rust emits but TS lacks is reported', () => {
  const keys = run(inSync.replace('text: string; ', '')).findings.map((f) => f.key);
  assert.deepEqual(keys, ['Para<type=para>.text: missing-in-ts']);
});

test('allowlist entries suppress known drift and stale entries are reported', () => {
  const drifted = inSync.replace('title?: string;', 'title: string;');
  const result = run(drifted, { 'Doc.title: rust-optional': 'known', 'Doc.kind: ts-only': 'gone' });
  assert.deepEqual(result.unexpected, []);
  assert.deepEqual(result.stale, ['Doc.kind: ts-only']);
});

const keysOf = (source, rustSchema) => run(source, {}, rustSchema).findings.map((f) => f.key);
const single = (properties, required = Object.keys(properties), $defs = {}) => ({
  title: 'Doc', type: 'object', properties, required, $defs,
});

test('untyped and never TS types do not silence the comparison', () => {
  assert.deepEqual(keysOf(inSync.replace('body: Block[];', 'body: any;')), ['Doc.body: ts-untyped']);
  assert.deepEqual(keysOf(inSync.replace('body: Block[];', 'body: never;')), [
    'Doc.body: type-mismatch (Rust array vs TS never)',
  ]);
  assert.deepEqual(keysOf('export interface Doc { v: string }', single({ v: true })), ['Doc.v: rust-untyped']);
});

test('an allowlisted literal difference cannot silently grow', () => {
  const drifted = inSync.replace("kind: 'a' | 'b';", "kind: 'a';");
  assert.deepEqual(run(drifted, { 'Doc.kind: literal-missing-in-ts ["b"]': 'known' }).unexpected, []);
  const grown = structuredClone(schema);
  grown.$defs.Kind.oneOf[0].enum.push('c');
  assert.deepEqual(
    run(drifted, { 'Doc.kind: literal-missing-in-ts ["b"]': 'known' }, grown).unexpected.map((f) => f.key),
    ['Doc.kind: literal-missing-in-ts ["b","c"]'],
  );
  // Narrowing a Rust String in TS is pinned by its literals too.
  assert.deepEqual(keysOf("export interface Doc { v: 'x' | 'y' }", single({ v: { type: 'string' } })), [
    'Doc.v: ts-narrower-literal ["x","y"]',
  ]);
});

test('union members are paired by compatibility, not declaration order', () => {
  const rust = single({
    v: { anyOf: [{ type: 'array', items: { type: 'number' } }, { type: 'array', items: { type: 'string' } }] },
  });
  assert.deepEqual(keysOf('export interface Doc { v: string[] | number[] }', rust), []);
  for (const ts of ['string[] | boolean[]', 'boolean[] | string[]']) {
    assert.deepEqual(keysOf(`export interface Doc { v: ${ts} }`, rust), [
      'Doc.v: type-mismatch (Rust number[] vs TS boolean[] | string[])',
      'Doc.v: variant-never-emitted (TS boolean[])',
    ]);
  }
  // Two TS members sharing a tag value: the compatible one is used.
  const tagged = single({
    v: { oneOf: [{ type: 'object', properties: { t: { const: 'a' }, n: { type: 'number' } }, required: ['t', 'n'] }] },
  });
  // ...and the other one is reported as an alternative Rust never emits.
  assert.deepEqual(
    keysOf("export interface Doc { v: { t: 'a'; s: string } | { t: 'a'; n: number } }", tagged),
    ['Doc.v<t=a>: variant-never-emitted (object { t: "a"; s: string; })'],
  );
});

test('named fields and index signatures are both compared', () => {
  const rust = single({ id: { type: 'string' } });
  rust.additionalProperties = { type: 'number' };
  assert.deepEqual(keysOf('export interface Doc { id: string; [key: string]: string }', rust), [
    'Doc.[key]: type-mismatch (Rust number vs TS string)',
  ]);
  assert.deepEqual(keysOf('export interface Doc { id: string }', rust), ['Doc.[key]: missing-in-ts']);
});

test('an omittable literal-valued field still needs an optional TS field', () => {
  const rust = single({ v: { const: 'fixed' } }, []);
  assert.deepEqual(keysOf("export interface Doc { v: 'fixed' }", rust), ['Doc.v: rust-optional']);
});

test('fields whose omission does not exclude null are rejected', () => {
  const dir = mkdtempSync(path.join(tmpdir(), 'model-wire-drift-rs-'));
  try {
    writeFileSync(path.join(dir, 'model.rs'), [
      'pub struct A {',
      '    #[serde(skip_serializing_if = "Option::is_none")]',
      '    pub fine: Option<String>,',
      '    /// Doc comment between attribute and field.',
      '    #[serde(',
      '        rename = "x",',
      '        skip_serializing_if = "is_blank"',
      '    )]',
      '    pub other_predicate: Option<String>,',
      '    #[serde(skip_serializing_if = "Option::is_none")]',
      '    pub nested: Option<Option<u32>>,',
      '    #[serde(skip_serializing_if = "Option::is_none")]',
      '    pub boxed: Option<Box<Option<String>>>,',
      '    #[serde(skip_serializing_if = "Option::is_none")]',
      '    pub aliased: Option<Maybe>,',
      '    #[serde(skip_serializing_if = "Option::is_none")]',
      '    pub unit: Option<Marker>,',
      '}',
      'type Maybe = Option<String>;',
      'struct Marker;',
      'pub struct G<T: Clone> {',
      '    #[serde(skip_serializing_if = "Option::is_none")]',
      '    pub generic: Option<T>,',
      '}',
    ].join('\n'));
    const problems = scanNullProvenance([dir]).map((p) => p.replace(/^.*model\.rs:/, ''));
    assert.deepEqual(problems, [
      '9 other_predicate: Option skipped by "is_blank" can still serialize null',
      '11 nested: Option<Option<u32>> can serialize null while present (Option)',
      '13 boxed: Option<Box<Option<String>>> can serialize null while present (Option)',
      '15 aliased: Option<Maybe> can serialize null while present (Maybe)',
      '17 unit: Option<Marker> can serialize null while present (Marker)',
      '23 generic: Option<T> can serialize null while present (T)',
    ]);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('recursive untagged unions terminate', () => {
  const node = (own) => ({
    type: 'object',
    properties: { child: { anyOf: [{ $ref: '#/$defs/A' }, { $ref: '#/$defs/B' }] }, [own]: { type: 'number' } },
    required: ['child', own],
  });
  const rust = { title: 'Doc', $ref: '#/$defs/A', $defs: { A: node('a'), B: node('b') } };
  const source = `
    export interface A { child: A | B; a: number }
    export interface B { child: A | B; b: number }
    export type Doc = A;`;
  assert.deepEqual(keysOf(source, rust), []);
});

test('numeric and boolean literal narrowing is reported', () => {
  const rust = single({ n: { type: 'number' }, b: { type: 'boolean' }, k: { const: 2 } });
  assert.deepEqual(keysOf('export interface Doc { n: 1 | 2; b: true; k: 2 | 3 }', rust), [
    'Doc.b: ts-narrower-literal [true]',
    'Doc.k: literal-never-emitted [3]',
    'Doc.n: ts-narrower-literal [1,2]',
  ]);
});

test('TS-only alternatives of untagged unions are reported', () => {
  const rust = single({ v: { type: ['number', 'string'] } });
  assert.deepEqual(keysOf('export interface Doc { v: number | string | boolean }', rust), [
    'Doc.v: variant-never-emitted (TS boolean)',
  ]);
});

test('required TS fields are not satisfied by extra keys', () => {
  const rust = single({ id: { type: 'number' } });
  rust.additionalProperties = { type: 'number' };
  assert.deepEqual(keysOf('export interface Doc { id: number; mustExist: number; [key: string]: number }', rust), [
    'Doc.mustExist: rust-optional',
  ]);
});

test('one Rust tagged type compared with two TS projections keeps separate keys', () => {
  const variant = { type: 'object', properties: { type: { const: 'a' }, n: { type: 'number' } }, required: ['type', 'n'] };
  const rust = single({ x: { $ref: '#/$defs/R' }, y: { $ref: '#/$defs/R' } }, ['x', 'y'], { R: { oneOf: [variant] } });
  const keys = keysOf(`
    export type TA = { type: 'a' };
    export type TB = { type: 'a' };
    export interface Doc { x: TA; y: TB }`, rust);
  assert.deepEqual(keys, ['R(ts TA).n: missing-in-ts', 'R(ts TB).n: missing-in-ts']);
});

test('TS tuples with rest elements accept fixed-size Rust arrays', () => {
  const rust = single({ v: { type: 'array', items: { type: 'number' }, minItems: 3, maxItems: 3 } });
  assert.deepEqual(keysOf('export interface Doc { v: [number, ...number[]] }', rust), []);
  assert.deepEqual(keysOf('export interface Doc { v: [number, number, number, number?] }', rust), []);
  assert.deepEqual(keysOf('export interface Doc { v: [number, number] }', rust), [
    'Doc.v: type-mismatch (Rust length 3..3 vs TS length 2..2)',
  ]);
});

test('widening a tagged member to an extra discriminant value is reported', () => {
  const widened = inSync.replace("{ type: 'rule'; width: number }", "{ type: 'rule' | 'hr'; width: number }");
  assert.deepEqual(keysOf(widened), [
    'Doc.body<type=hr>: variant-never-emitted',
    'Para<type=para>.children<type=hr>: variant-never-emitted',
  ]);
});

test('an empty TS object does not erase a Rust map', () => {
  const rust = single({ m: { type: 'object', additionalProperties: { type: 'string' } } });
  assert.deepEqual(keysOf('export interface Doc { m: {} }', rust), ['Doc.m: type-mismatch (Rust {[key]: string} vs TS object {})']);
});

test('null provenance follows outer aliases and wrappers', () => {
  const dir = mkdtempSync(path.join(tmpdir(), 'model-wire-drift-rs-'));
  try {
    writeFileSync(path.join(dir, 'model.rs'), [
      'type MaybeError = Option<Option<String>>;',
      'type Plain = Option<String>;',
      'type Maybe<T> = Option<T>;',
      'type Identity<T> = T;',
      'pub struct A {',
      '    #[serde(skip_serializing_if = "Option::is_none")]',
      '    pub aliased: MaybeError,',
      '    #[serde(skip_serializing_if = "Option::is_none")]',
      '    pub boxed: Box<Option<Option<String>>>,',
      '    #[serde(skip_serializing_if = "Option::is_none")]',
      '    pub fine: Plain,',
      '    #[serde(skip_serializing_if = "Option::is_none")]',
      '    pub generic_outer: Maybe<Option<String>>,',
      '    #[serde(skip_serializing_if = "Option::is_none")]',
      '    pub generic_inner: Option<Identity<Option<String>>>,',
      '}',
    ].join('\n'));
    const problems = scanNullProvenance([dir]).map((p) => p.replace(/^.*model\.rs:/, ''));
    assert.deepEqual(problems, [
      '7 aliased: Option<Option<String>> can serialize null while present (Option)',
      '9 boxed: Option<Option<String>> can serialize null while present (Option)',
      '13 generic_outer: skipped Maybe<Option<String>> uses generic alias Maybe; null exclusion cannot be established',
      '15 generic_inner: skipped Option<Identity<Option<String>>> uses generic alias Identity; null exclusion cannot be established',
    ]);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('TS tuples with trailing fixed elements', () => {
  const rust = single({ v: { type: 'array', prefixItems: [{ type: 'number' }, { type: 'string' }], minItems: 2, maxItems: 2 } });
  assert.deepEqual(keysOf('export interface Doc { v: [...number[], string] }', rust), []);
});
