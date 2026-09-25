// Controlled source conditions, not expected Office behavior.
// All coordinates are EMU. See the accompanying experiment protocol.
import { pathToFileURL } from 'node:url';

export function rulerProbeCases() {
  const baseline = {
    defTabSz: 914400, marL: 0, indent: 0, lvl: 0, rtl: 0,
    tabs: [{ pos: 1828800, algn: 'l' }],
  };
  const conditions = [{ name: 'Baseline', changes: {} }];
  const add = (name, key, values) => {
    for (const value of values) conditions.push({ name: `${name} ${value}`, changes: { [key]: value } });
  };
  add('Default interval', 'defTabSz', [-1587, 0, 1587, 912812, 915988, 457200, 1828800]);
  add('Left margin', 'marL', [1587, 228600, 457200]);
  add('First line', 'indent', [-228600, -1587, 1587, 228600]);
  for (const pos of [-1587, 0, 914400, 1827212, 1830388, 2743200]) {
    conditions.push({ name: `Tab position ${pos}`, changes: { tabs: [{ pos, algn: 'l' }] } });
  }
  for (const algn of ['ctr', 'r', 'dec']) {
    conditions.push({ name: `Tab alignment ${algn}`, changes: { tabs: [{ pos: 1828800, algn }] } });
  }
  conditions.push(
    { name: 'Explicit empty tabs', changes: { tabs: [] } },
    { name: 'Two explicit tabs', changes: { tabs: [{ pos: 1828800, algn: 'l' }, { pos: 3657600, algn: 'r' }] } },
    { name: 'Level one', changes: { lvl: 1 } },
    { name: 'RTL paragraph', changes: { rtl: 1 } },
    { name: 'Unchanged repeat', changes: {} },
  );
  return conditions.map((condition, index) => ({
    id: `T${String(index + 1).padStart(3, '0')}`,
    parent: index === 0 ? null : 'T001',
    name: condition.name,
    changes: structuredClone(condition.changes),
    first: structuredClone(baseline),
    second: { ...structuredClone(baseline), ...structuredClone(condition.changes) },
  }));
}

export function rulerProbeManifest() {
  return {
    schemaVersion: 1,
    units: 'EMU',
    text: 'AAAA\tBBBB\t123.45',
    origin: 'synthetic authored OOXML; not an observed Office result',
    conditions: rulerProbeCases(),
  };
}

// Keep authoring separate from inference. This command writes only a manifest
// to stdout; it never opens Office, samples, or an output file.
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  process.stdout.write(`${JSON.stringify(rulerProbeManifest(), null, 2)}\n`);
}
