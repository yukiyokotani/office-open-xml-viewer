import assert from 'node:assert/strict';
import { cpSync, mkdtempSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

const repositoryRoot = resolve(import.meta.dirname, '..');
const checker = resolve(import.meta.dirname, 'check-docx-compatibility-evidence.mjs');

function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'docx-compatibility-evidence-'));
  cpSync(
    resolve(repositoryRoot, 'packages/docx/src/layout/compatibility.ts'),
    resolve(root, 'packages/docx/src/layout/compatibility.ts'),
    { recursive: true },
  );
  const references = [
    ['packages/docx/src/layout/coordinate-space.test.ts', ['maps Transitional text direction %s to %s']],
    ['packages/docx/src/float-line-start-one-inch.test.ts', [
      'keeps an anchor-host metric-only line on the paragraph-mark threshold',
    ]],
    ['packages/docx/src/layout/floats.test.ts', ['keeps observed different-paragraph displacement on exclusion bounds']],
    ['packages/docx/src/float-table-page-fit.test.ts', ['(g) DEFERS a page-anchored floating table when its raw band intersects an existing table float']],
  ];
  for (const [path, titles] of references) {
    write(
      root,
      path,
      titles.map((title) => `it(${JSON.stringify(title)}, () => {});`).join('\n'),
    );
  }
  write(
    root,
    'scripts/docx-compatibility-microsoft-evidence.json',
    readFileSync(
      resolve(repositoryRoot, 'scripts/docx-compatibility-microsoft-evidence.json'),
      'utf8',
    ),
  );
  return root;
}

function write(root, path, source) {
  const absolute = resolve(root, path);
  mkdirSync(dirname(absolute), { recursive: true });
  writeFileSync(absolute, source);
}

function run(root, ...args) {
  return spawnSync(process.execPath, [checker, '--root', root, ...args], {
    cwd: repositoryRoot,
    encoding: 'utf8',
  });
}

test('accepts literal rules with live regression evidence', () => {
  const result = run(fixture());
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /verified \(5 rules\)/);
});

test('rejects a stale regression test title', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/floats.test.ts', 'it("renamed", () => {});\n');
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /COMPATIBILITY_EVIDENCE_STALE/);
});

test('rejects a rule declaration outside the closed authority set', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/table.ts', `
    import { defineCompatibilityRule } from './compatibility.js';
    export const WORD_INLINE = defineCompatibilityRule({
      id: 'word-inline',
      evidence: {
        kind: 'regression-test',
        reference: 'packages/docx/src/layout/floats.test.ts#keeps observed different-paragraph displacement on exclusion bounds',
      },
      description: 'Inline rule',
    });
  `);
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /COMPATIBILITY_FACTORY_ACCESS|COMPATIBILITY_DECLARATION_AUTHORITY/);
});

test('rejects duplicate rule ids across compatibility modules', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/table-compatibility.ts', `
    import { defineCompatibilityRule } from './compatibility.js';
    export const WORD_DUPLICATE = defineCompatibilityRule({
      id: 'word-float-gap-flow',
      evidence: { kind: 'microsoft-note', reference: '[MS-OI29500] §2.1.120' },
      description: 'Duplicate rule',
    });
  `);
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /COMPATIBILITY_DUPLICATE_ID/);
});

test('rejects an aliased factory import', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/table-compatibility.ts', `
    import { defineCompatibilityRule as defineRule } from './compatibility.js';
    export const WORD_ALIAS = defineRule({
      id: 'word-alias',
      evidence: { kind: 'microsoft-note', reference: '[MS-OI29500] §2.1.120' },
      description: 'Aliased factory',
    });
  `);
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /COMPATIBILITY_FACTORY_ACCESS/);
});

test('rejects namespace and binding access to the factory', () => {
  for (const source of [
    `
      import * as compatibility from './compatibility.js';
      export const WORD_NAMESPACE = compatibility.defineCompatibilityRule({
        id: 'word-namespace',
        evidence: { kind: 'microsoft-note', reference: '[MS-OI29500] §2.1.120' },
        description: 'Namespace factory',
      });
    `,
    `
      import { defineCompatibilityRule } from './compatibility.js';
      const defineRule = defineCompatibilityRule;
      export const WORD_BINDING = defineRule({
        id: 'word-binding',
        evidence: { kind: 'microsoft-note', reference: '[MS-OI29500] §2.1.120' },
        description: 'Bound factory',
      });
    `,
  ]) {
    const root = fixture();
    write(root, 'packages/docx/src/layout/table-compatibility.ts', source);
    const result = run(root);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /COMPATIBILITY_FACTORY_ACCESS/);
  }
});

test('rejects factory re-export laundering', () => {
  for (const source of [
    "export { defineCompatibilityRule as defineRule } from './compatibility.js';\n",
    "export * from './compatibility.js';\n",
  ]) {
    const root = fixture();
    write(root, 'packages/docx/src/layout/factory-launder.ts', source);
    const result = run(root);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /COMPATIBILITY_FACTORY_ACCESS/);
  }
});

test('rejects dynamic and CommonJS loading of the compatibility owner', () => {
  for (const [path, source] of [
    [
      'packages/docx/src/layout/dynamic-factory.mts',
      "export const factory = await import('./compatibility.js');\n",
    ],
    [
      'packages/docx/src/layout/commonjs-factory.cjs',
      "module.exports = require('./compatibility.js');\n",
    ],
    [
      'packages/docx/src/layout/template-factory.mts',
      `
        const compatibility = await import(\`./compatibility.js\`);
        export const WORD_TEMPLATE = compatibility['defineCompatibilityRule']({
          id: 'word-template',
          evidence: {
            kind: 'regression-test',
            reference: 'packages/docx/src/layout/coordinate-space.test.ts#maps Transitional text direction %s to %s',
          },
          description: 'Template-loaded factory',
        });
      `,
    ],
  ]) {
    const root = fixture();
    write(root, path, source);
    const result = run(root);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /COMPATIBILITY_FACTORY_ACCESS/);
  }
});

test('scans production-importable mts and test-support modules', () => {
  for (const path of [
    'packages/docx/src/layout/hidden-rule.mts',
    'packages/docx/src/layout/hidden-rule.test-support.ts',
  ]) {
    const root = fixture();
    write(root, path, `
      import { defineCompatibilityRule } from './compatibility.js';
      export const WORD_HIDDEN = defineCompatibilityRule({
        id: 'word-hidden',
        evidence: { kind: 'microsoft-note', reference: '[MS-OI29500] §2.1.120' },
        description: 'Hidden rule',
      });
    `);
    const result = run(root);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /COMPATIBILITY_FACTORY_ACCESS|COMPATIBILITY_DECLARATION_AUTHORITY/);
  }
});

test('rejects an unstructured Microsoft note reference', () => {
  const root = fixture();
  const source = readFileSync(
    resolve(root, 'packages/docx/src/layout/compatibility.ts'),
    'utf8',
  ).replace(
    "kind: 'regression-test',\n    reference: 'packages/docx/src/layout/coordinate-space.test.ts#maps Transitional text direction %s to %s'",
    "kind: 'microsoft-note',\n    reference: 'MS note somewhere'",
  );
  write(root, 'packages/docx/src/layout/compatibility.ts', source);
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /COMPATIBILITY_MICROSOFT_REFERENCE/);
});

test('rejects a structured but uncatalogued Microsoft note reference', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/table-compatibility.ts', `
    import { defineCompatibilityRule } from './compatibility.js';
    export const WORD_UNKNOWN_NOTE = defineCompatibilityRule({
      id: 'word-unknown-note',
      evidence: { kind: 'microsoft-note', reference: '[MS-OI29500] §2.1.999999' },
      description: 'Unknown Microsoft evidence',
    });
  `);
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /COMPATIBILITY_MICROSOFT_EVIDENCE/);
});

test('rejects a new inline Office observation outside compatibility modules', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/table.ts',
    '// Word uses an undocumented magic branch here.\nexport const value = 1;\n');
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /INLINE_COMPATIBILITY_OBSERVATION/);
});

test('detects varied Office observation language outside compatibility modules', () => {
  const claims = [
    'Word renders an undocumented branch here.',
    'Word draws an undocumented branch here.',
    'Word fills an undocumented branch here.',
    'Word resolves an undocumented branch here.',
    'Word collapses an undocumented branch here.',
    'Word ground truth requires an undocumented branch here.',
    'Word-observed behavior requires an undocumented branch here.',
    'Observed Windows Word behavior requires an undocumented branch here.',
    'Observed Word offsets require an undocumented branch here.',
    "Word's hierarchy default is undocumented.",
    "This moves exactly like Word's pen.",
    'No fixture pins where Word clamps the box.',
    'Word — which paints the whole merged span.',
    'Word pins the baseline.',
    'This was measured against the Word PDFs.',
    'This is matching Word.',
  ];
  for (const claim of claims) {
    const root = fixture();
    write(root, 'packages/docx/src/layout/table.ts',
      `// ${claim}\nexport const value = 1;\n`);
    const result = run(root);
    assert.equal(result.status, 1);
    assert.match(result.stderr, /INLINE_COMPATIBILITY_OBSERVATION/);
  }
});

test('does not treat a lowercase generic word as a compatibility claim', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/table.ts',
    '// The next word uses the inherited style.\nexport const value = 1;\n');
  assert.equal(run(root).status, 0);
});

test('does not exempt an unreviewed file merely because its name ends in compatibility', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/rogue-compatibility.ts',
    '// Word uses an unevidenced branch here.\nexport const value = 1;\n');
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /INLINE_COMPATIBILITY_OBSERVATION/);
});

test('rejects any transitional observation baseline in the final architecture', () => {
  const root = fixture();
  write(root, 'scripts/docx-compatibility-observation-baseline.json',
    '{ "version": 1, "observations": [] }\n');
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /FINAL_COMPATIBILITY_OBSERVATION_BASELINE/);
});

test('rejects the removed transitional observation inventory argument', () => {
  const root = fixture();
  const result = run(root, '--print-observations');
  assert.equal(result.status, 1);
  assert.match(result.stderr, /UNKNOWN_ARGUMENT: --print-observations/);
});

// The AST gate must validate the same fields that the runtime object exposes.
// Later entries can replace an otherwise valid id or evidence reference.
test('rejects rule entries that can override validated literal fields', () => {
  for (const override of [
    "'id': 'overridden-rule',",
    "...{ id: 'overridden-rule' },",
    "['id']: 'overridden-rule',",
    "get id() { return 'overridden-rule'; },",
  ]) {
    const root = fixture();
    write(root, 'packages/docx/src/layout/table-compatibility.ts', `
      import { defineCompatibilityRule } from './compatibility.js';
      export const WORD_OVERRIDE = defineCompatibilityRule({
        id: 'word-literal-override',
        evidence: { kind: 'microsoft-note', reference: '[MS-OI29500] §2.1.120' },
        description: 'Literal override regression',
        ${override}
      });
    `);
    const result = run(root);
    assert.equal(result.status, 1, `accepted rule override: ${override}`);
    assert.match(result.stderr, /COMPATIBILITY_LITERAL_SCHEMA/);
  }
});

test('rejects evidence entries that can override validated literal fields', () => {
  for (const override of [
    "reference: 'unverified reference',",
    "...{ reference: 'unverified reference' },",
    "['reference']: 'unverified reference',",
    "reference() { return 'unverified reference'; },",
  ]) {
    const root = fixture();
    write(root, 'packages/docx/src/layout/table-compatibility.ts', `
      import { defineCompatibilityRule } from './compatibility.js';
      export const WORD_OVERRIDE = defineCompatibilityRule({
        id: 'word-literal-override',
        evidence: {
          kind: 'microsoft-note',
          reference: '[MS-OI29500] §2.1.120',
          ${override}
        },
        description: 'Literal override regression',
      });
    `);
    const result = run(root);
    assert.equal(result.status, 1, `accepted evidence override: ${override}`);
    assert.match(result.stderr, /COMPATIBILITY_LITERAL_SCHEMA/);
  }
});

test('accepts quoted literal keys and extra static metadata', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/table-compatibility.ts', `
    import { defineCompatibilityRule } from './compatibility.js';
    export const WORD_QUOTED = defineCompatibilityRule({
      'id': 'word-quoted-literal',
      'evidence': {
        'kind': 'microsoft-note',
        'reference': '[MS-OI29500] §2.1.120',
        'metadata': 'Static evidence metadata',
      },
      'description': 'Quoted literal regression',
      'metadata': 'Static rule metadata',
    });
  `);
  const result = run(root);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /verified \(6 rules\)/);
});
