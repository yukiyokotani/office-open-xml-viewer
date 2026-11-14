import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { test } from 'node:test';

import { checkFormat } from './check-model-wire-drift.mjs';

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

function run(source, allowlist = {}) {
  const dir = mkdtempSync(path.join(tmpdir(), 'model-wire-drift-'));
  try {
    writeFileSync(path.join(dir, 'tsconfig.json'), JSON.stringify({ compilerOptions: { strict: true } }));
    writeFileSync(path.join(dir, 'types.ts'), source);
    writeFileSync(path.join(dir, 'schema.json'), JSON.stringify(schema));
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
    'Doc.kind: literal-missing-in-ts',
    'Doc.size: type-mismatch',
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
  const result = run(drifted, { 'Doc.title: rust-optional': 'known', 'Doc.*: ts-only': 'gone' });
  assert.deepEqual(result.unexpected, []);
  assert.deepEqual(result.stale, ['Doc.*: ts-only']);
});
