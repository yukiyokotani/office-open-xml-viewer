import assert from 'node:assert/strict';
import { test } from 'node:test';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { textProbeManifest } from './legacy-ppt-text-probes.mjs';
import { rulerProbeManifest } from './legacy-ppt-ruler-probes.mjs';

test('all original ruler conditions survive, each in its own text body', () => {
  const manifest = textProbeManifest();
  assert.equal(manifest.arrangement, 'two-separate-text-bodies');
  const original = rulerProbeManifest();
  const rulers = manifest.conditions.filter(c => c.family === 'separate-rulers');
  assert.equal(rulers.length, 29);
  for (const [i, condition] of rulers.entries()) {
    assert.equal(condition.sourceCondition, original.conditions[i].id);
    assert.equal(condition.bodies.length, 2);
    assert.deepEqual(condition.bodies.map(b => b.properties), [original.conditions[i].first, original.conditions[i].second]);
    for (const b of condition.bodies) assert.deepEqual(b.paragraphs, [{ runs: [{ text: original.text }] }]);
  }
});

test('text-slot counterexamples keep formatting fixed and expose UTF-16 ambiguity', () => {
  const slots = textProbeManifest().conditions.filter(c => c.family === 'text-slots');
  assert.equal(slots.length, 16);
  const get = id => slots.find(c => c.id === id).bodies[1].paragraphs;
  for (const c of slots) {
    assert.deepEqual(c.bodies[0].properties, c.bodies[1].properties);
    assert.deepEqual(c.bodies[0].paragraphs, [{ runs: [{ text: 'ABCD' }] }]);
  }
  const strings = ['M001', 'M002', 'M003', 'M004', 'M007', 'M008', 'M009'].map(id => get(id)[0].runs[0].text);
  assert.ok(strings.every(s => s.length === 4));
  assert.equal(new Set(strings).size, strings.length);
  assert.equal([...get('M008')[0].runs[0].text].length, 3);
  assert.equal(get('M010').length, 3);
  assert.deepEqual(get('M011').map(p => p.runs[0].text.length), [1, 3]);
  assert.deepEqual(get('M012').map(p => p.runs[0].text.length), [3, 1]);
  assert.equal(get('M013')[0].runs[1].bold, true);
  assert.deepEqual(get('M014')[0].runs[1], { break: true });
  assert.equal(get('M015')[0].runs[1].text, '');
  assert.deepEqual(get('M001'), get('M016'));
});

test('source manifests are independent, passive, deterministic and not outcome assertions', () => {
  const first = textProbeManifest(), second = textProbeManifest();
  assert.deepEqual(first, second);
  assert.equal(new Set(first.conditions.map(c => c.id)).size, 45);
  assert.match(first.origin, /not an observed Office result/);
  first.conditions[0].bodies[0].properties.tabs[0].pos = 1;
  first.conditions[0].bodies[0].paragraphs[0].runs[0].text = 'changed';
  assert.deepEqual(first.conditions[0].bodies[1], second.conditions[0].bodies[1]);
  assert.deepEqual(textProbeManifest(), second);
  for (const c of second.conditions) for (const b of c.bodies) for (const p of b.paragraphs) for (const run of p.runs) {
    assert.ok(Object.keys(run).every(key => ['text', 'bold', 'break'].includes(key)));
    assert.ok(run.break === true || typeof run.text === 'string');
  }
});

test('the standalone command emits only the deterministic source manifest', () => {
  const output = execFileSync(process.execPath, [fileURLToPath(new URL('./legacy-ppt-text-probes.mjs', import.meta.url))], {
    encoding: 'utf8', timeout: 10000, maxBuffer: 1024 * 1024,
  });
  assert.deepEqual(JSON.parse(output), textProbeManifest());
});
