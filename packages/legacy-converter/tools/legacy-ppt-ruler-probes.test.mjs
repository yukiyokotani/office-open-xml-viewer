import assert from 'node:assert/strict';
import { test } from 'node:test';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { rulerProbeCases, rulerProbeManifest } from './legacy-ppt-ruler-probes.mjs';

test('stable IDs, baseline and unchanged repeat define 29 paired conditions', () => {
  const cases = rulerProbeCases();
  assert.equal(cases.length, 29);
  assert.deepEqual(cases.map(c => c.id), Array.from({ length: 29 }, (_, i) => `T${String(i + 1).padStart(3, '0')}`));
  assert.equal(cases[0].parent, null);
  assert.ok(cases.slice(1).every(c => c.parent === cases[0].id));
  assert.deepEqual(cases[0].first, cases[0].second);
  assert.deepEqual(cases.at(-1).second, cases[0].second);
});

test('each changed condition changes exactly one paragraph property', () => {
  const cases = rulerProbeCases();
  for (const [index, c] of cases.entries()) {
    assert.deepEqual(c.first, cases[0].first);
    const keys = Object.keys(c.first).filter(key => JSON.stringify(c.first[key]) !== JSON.stringify(c.second[key]));
    assert.equal(keys.length, index === 0 || index === cases.length - 1 ? 0 : 1, c.id);
    assert.deepEqual(keys.sort(), Object.keys(c.changes).sort());
    assert.deepEqual(c.second, { ...c.first, ...c.changes });
  }
});

test('all coordinates and text properties fit their DrawingML schema types', () => {
  for (const c of rulerProbeCases()) for (const p of [c.first, c.second]) {
    assert.ok(Number.isInteger(p.defTabSz) && p.defTabSz >= -(2 ** 31) && p.defTabSz < 2 ** 31);
    assert.ok(Number.isInteger(p.marL) && p.marL >= 0 && p.marL <= 51206400);
    assert.ok(Number.isInteger(p.indent) && Math.abs(p.indent) <= 51206400);
    assert.ok(Number.isInteger(p.lvl) && p.lvl >= 0 && p.lvl <= 8);
    assert.ok(p.rtl === 0 || p.rtl === 1);
    for (const [i, tab] of p.tabs.entries()) {
      assert.ok(Number.isInteger(tab.pos) && tab.pos >= -(2 ** 31) && tab.pos < 2 ** 31);
      assert.ok(['l', 'ctr', 'r', 'dec'].includes(tab.algn));
      if (i) assert.ok(p.tabs[i - 1].pos < tab.pos);
    }
  }
});

test('boundary, empty-list and alignment controls cannot silently disappear', () => {
  const values = key => new Set(rulerProbeCases().map(c => c.second[key]));
  for (const n of [-1587, 0, 1587, 912812, 914400, 915988]) assert.ok(values('defTabSz').has(n));
  const tabs = rulerProbeCases().map(c => c.second.tabs);
  assert.ok(tabs.some(t => t.length === 0));
  assert.ok(tabs.some(t => t.length === 2));
  for (const algn of ['l', 'ctr', 'r', 'dec']) assert.ok(tabs.some(t => t.some(tab => tab.algn === algn)));
  assert.ok(values('rtl').has(1));
  assert.ok(values('lvl').has(1));
});

test('nested parameter objects never alias controls, repeats or later invocations', () => {
  const cases = rulerProbeCases();
  const expected = rulerProbeCases();
  cases[0].first.tabs[0].pos = 1;
  cases[1].second.tabs.push({ pos: 9999999, algn: 'r' });
  cases[15].changes.tabs[0].pos = 2;
  assert.deepEqual(cases[0].second, expected[0].second);
  assert.deepEqual(cases[1].first, expected[1].first);
  assert.deepEqual(cases[15].second, expected[15].second);
  assert.deepEqual(cases.at(-1), expected.at(-1));
  assert.deepEqual(rulerProbeCases(), expected);
});

test('manifest contains source conditions but no expected converter or Office result', () => {
  const manifest = rulerProbeManifest();
  assert.equal(manifest.units, 'EMU');
  assert.equal(manifest.text, 'AAAA\tBBBB\t123.45');
  assert.match(manifest.origin, /not an observed Office result/);
  for (const condition of manifest.conditions) {
    assert.deepEqual(Object.keys(condition).sort(), ['changes', 'first', 'id', 'name', 'parent', 'second']);
  }
});

test('command output is the same deterministic manifest and requires no Office access', () => {
  const output = execFileSync(process.execPath, [fileURLToPath(new URL('./legacy-ppt-ruler-probes.mjs', import.meta.url))], {
    encoding: 'utf8', timeout: 10000, maxBuffer: 1024 * 1024,
  });
  assert.deepEqual(JSON.parse(output), rulerProbeManifest());
});
