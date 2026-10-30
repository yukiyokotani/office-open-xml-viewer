import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { chmodSync, copyFileSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';

function fixture(t) {
  const root = mkdtempSync(join(tmpdir(), 'ooxml-companion-release-'));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  for (const directory of ['scripts', 'src', 'packages/legacy-office', 'packages/opendocument', 'bin']) mkdirSync(join(root, directory), { recursive: true });
  for (const name of ['sync-companion-version.mjs', 'publish-packages.mjs']) {
    copyFileSync(new URL(name, import.meta.url), join(root, 'scripts', name));
  }
  writeFileSync(join(root, 'package.json'), JSON.stringify({ name: '@silurus/ooxml', version: '1.2.3' }));
  const writeLegacy = (version, peer) => writeFileSync(join(root, 'packages/legacy-office/package.json'), JSON.stringify({
    name: '@silurus/ooxml-legacy-office', version, peerDependencies: { '@silurus/ooxml': peer },
  }));
  writeLegacy('1.2.3', '1.2.3');
  const writeOpenDocument = (version, peer) => writeFileSync(join(root, 'packages/opendocument/package.json'), JSON.stringify({
    name: '@silurus/ooxml-opendocument', version, peerDependencies: { '@silurus/ooxml': peer },
  }));
  writeOpenDocument('1.2.3', '1.2.3');
  const run = (name, args = [], env = {}) => spawnSync(process.execPath, [join(root, 'scripts', name), ...args], {
    cwd: root, encoding: 'utf8', env: { ...process.env, ...env },
  });
  assert.equal(run('sync-companion-version.mjs', ['--write']).status, 0);
  return { root, run, writeLegacy, writeOpenDocument };
}

test('release check rejects stale versions, ranged peers, and a stale compiled ABI; sync repairs them', t => {
  const { root, run, writeLegacy, writeOpenDocument } = fixture(t);
  assert.equal(run('sync-companion-version.mjs').status, 0);
  for (const write of [writeLegacy, writeOpenDocument]) {
    for (const [version, peer] of [['1.2.2', '1.2.3'], ['1.2.3', '^1.2.3']]) {
      write(version, peer);
      assert.notEqual(run('sync-companion-version.mjs').status, 0);
      assert.equal(run('sync-companion-version.mjs', ['--write']).status, 0);
      assert.equal(run('sync-companion-version.mjs').status, 0);
    }
  }
  writeFileSync(join(root, 'src/companion-version.ts'), "export const COMPANION_VERSION = '1.2.2';\n");
  assert.notEqual(run('sync-companion-version.mjs').status, 0);
});

// Fake registry and publisher executables exercise retry ordering and failure
// policy without contacting npm or authorizing a real publication.
for (const scenario of ['fresh', 'root-published', 'both-published', 'registry-failure']) {
  test(`publishing ${scenario} preserves viewer-before-companion ordering and fails closed`, t => {
    const { root, run } = fixture(t);
    const executable = (name, code) => {
      const file = join(root, 'bin', name);
      writeFileSync(file, `#!${process.execPath}\n${code}`);
      chmodSync(file, 0o755);
    };
    executable('npm', `
const scenario = process.env.SCENARIO;
const root = process.argv[3].startsWith('@silurus/ooxml@');
if (scenario === 'registry-failure') { console.log(JSON.stringify({error:{code:'E503'}})); process.exit(1); }
if (scenario === 'both-published' || (scenario === 'root-published' && root)) { console.log(JSON.stringify('1.2.3')); }
else { console.log(JSON.stringify({error:{code:'E404'}})); process.exit(1); }
`);
    executable('pnpm', `
const fs = require('node:fs');
const pkg = JSON.parse(fs.readFileSync('package.json', 'utf8'));
fs.appendFileSync(process.env.PUBLISH_LOG, pkg.name + '\\n');
`);
    const log = join(root, 'published');
    const result = run('publish-packages.mjs', [], {
      PATH: `${join(root, 'bin')}:${process.env.PATH}`, SCENARIO: scenario, PUBLISH_LOG: log,
    });
    assert.equal(result.status, scenario === 'registry-failure' ? 1 : 0, result.stderr);
    const expected = scenario === 'fresh' ? ['@silurus/ooxml', '@silurus/ooxml-legacy-office', '@silurus/ooxml-opendocument']
      : scenario === 'root-published' ? ['@silurus/ooxml-legacy-office', '@silurus/ooxml-opendocument'] : [];
    let published = [];
    try { published = readFileSync(log, 'utf8').trim().split('\n'); } catch {}
    assert.deepEqual(published, expected);
  });
}
