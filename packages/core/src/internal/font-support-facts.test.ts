import { expect, it } from 'vitest';
import { parseFontSupportFacts, type FontTable } from './font-support-facts.js';

const tag = (s: string) => [...s].reduce((n, c) => n * 256 + c.charCodeAt(0), 0);
function inspect(subtables: number[][], types: number[], flags = 0, zeroClass = 0, repeat = 1, glyphCount = 8, version = 0x10000, classSource: 'absent' | 'null' | 'empty' = 'absent') {
  const bytes = new Uint8Array(Math.max(4096, 100 + subtables.reduce((n, s) => n + s.length * 2 + 8, 0) + subtables.length * repeat * 2)), v = new DataView(bytes.buffer);
  const tables = new Map<number, FontTable>();
  v.setUint32(0, 0x10000); v.setUint16(4, glyphCount); tables.set(tag('maxp'), { offset: 0, length: 6 });
  v.setUint32(32, version); v.setUint16(36, 10); v.setUint16(38, 12); v.setUint16(40, 14);
  v.setUint16(46, subtables.length * repeat);
  let at = Math.max(64, 48 + subtables.length * repeat * 2);
  for (let i = 0; i < subtables.length; i++) {
    for (let j = i; j < subtables.length * repeat; j += subtables.length) v.setUint16(48 + j * 2, at - 46);
    v.setUint16(at, types[i]); v.setUint16(at + 2, flags); v.setUint16(at + 4, 1); v.setUint16(at + 6, 8);
    subtables[i].forEach((value, j) => v.setUint16(at + 8 + j * 2, value));
    at += 8 + subtables[i].length * 2;
  }
  tables.set(tag('GSUB'), { offset: 32, length: at - 32 });
  if (zeroClass || classSource !== 'absent') {
    v.setUint32(at, 0x10000); v.setUint16(at + 4, classSource === 'null' ? 0 : 12);
    [1, 0, classSource === 'empty' ? 0 : 1, zeroClass].forEach((n, i) => v.setUint16(at + 12 + i * 2, n));
    tables.set(tag('GDEF'), { offset: at, length: 20 });
  }
  return parseFontSupportFacts(v, tables);
}

it('distinguishes safe substitutions from missing outputs with the same cmap inputs', () => {
  // SingleSubst format 2, coverage glyph 1, replacing it by glyph 2 or zero.
  expect(inspect([[2, 8, 1, 2, 1, 1, 1]], [1])).toMatchObject({ nonzeroPreserved: true, missingIsolated: true });
  expect(inspect([[2, 8, 1, 0, 1, 1, 1]], [1])).toMatchObject({ nonzeroPreserved: false, missingIsolated: true });
  // Delta uses unsigned 16-bit arithmetic; 1 + 65535 becomes missing.
  expect(inspect([[1, 6, 65535, 1, 1, 1]], [1]).nonzeroPreserved).toBe(false);
  expect(inspect([[2, 8, 1, 8, 1, 1, 1]], [1]).reason).toBe('malformed');
});

it('checks all multiple/alternate outputs and rejects deletion', () => {
  for (const type of [2, 3]) {
    expect(inspect([[1, 12, 1, 8, 1, 2, 1, 1, 1]], [type]).nonzeroPreserved).toBe(true);
    expect(inspect([[1, 14, 1, 8, 2, 2, 0, 1, 1, 1]], [type]).nonzeroPreserved).toBe(false);
    expect(inspect([[1, 10, 1, 8, 0, 1, 1, 1]], [type])).toMatchObject({ nonzeroPreserved: true, noErasure: type !== 2 });
  }
});

it('keeps missing-glyph isolation separate from preservation, including skip flags', () => {
  expect(inspect([[2, 8, 1, 2, 1, 1, 0]], [1])).toMatchObject({ nonzeroPreserved: true, missingIsolated: false });
  expect(inspect([[2, 8, 1, 2, 1, 1, 1]], [1], 8, 1)).toMatchObject({ nonzeroPreserved: true, missingIsolated: true });
  expect(inspect([[2, 8, 1, 2, 1, 1, 1]], [1], 8, 3)).toMatchObject({ nonzeroPreserved: true, missingIsolated: false });
  const single = [[2, 8, 1, 2, 1, 1, 1]];
  for (const source of ['absent', 'null'] as const) {
    expect(inspect(single, [1], 8, 0, 1, 8, 0x10000, source)).toMatchObject({ nonzeroPreserved: true, missingIsolated: false });
    expect(inspect(single, [1], 2, 0, 1, 8, 0x10000, source).missingIsolated).toBe(false);
    expect(inspect(single, [1], 0x100, 0, 1, 8, 0x10000, source).missingIsolated).toBe(false);
    expect(inspect(single, [1], 4, 0, 1, 8, 0x10000, source).missingIsolated).toBe(true);
  }
  expect(inspect(single, [1], 8, 0, 1, 8, 0x10000, 'empty').missingIsolated).toBe(true);
});

it('follows contextual references and rejects cycles without treating budget failures as absence', () => {
  // Context format 3, one covered input, a lookup record applying lookup 1.
  const context = [3, 1, 1, 12, 0, 1, 1, 1, 1];
  expect(inspect([context, [2, 8, 1, 0, 1, 1, 1]], [5, 1]).nonzeroPreserved).toBe(false);
  const cycle = [...context]; cycle[5] = 0;
  expect(inspect([cycle], [5])).toMatchObject({ nonzeroPreserved: undefined, missingIsolated: undefined, reason: 'cycle' });
  expect(inspect([[99]], [1])).toMatchObject({ nonzeroPreserved: undefined, missingIsolated: undefined });
});

it('accepts single-component ligature records and checks extension targets', () => {
  // Ligature format 1, coverage glyph 1, one component replaced by glyph 2.
  const ligature = [1, 16, 1, 8, 1, 4, 2, 1, 1, 1, 1];
  expect(inspect([ligature], [4])).toMatchObject({ nonzeroPreserved: true, missingIsolated: true });
  expect(inspect([[1, 4, 0, 8, ...ligature]], [7]).nonzeroPreserved).toBe(true);
});

it('bounds repeated lookup aliases and glyph-domain visits with explicit unknown exits', () => {
  const single = [2, 8, 1, 2, 1, 1, 1];
  expect(inspect([single], [1], 0, 0, 4096).nonzeroPreserved).toBe(true);
  expect(inspect([single], [1], 0, 0, 4097)).toMatchObject({ reason: 'budget', nonzeroPreserved: undefined, missingIsolated: undefined });
  const broad = [1, 6, 0, 2, 1, 1, 65534, 0];
  expect(inspect([broad, broad, broad, broad, broad], [1, 1, 1, 1, 1], 0, 0, 1, 65535)).toMatchObject({ reason: 'budget', nonzeroPreserved: undefined, missingIsolated: undefined });
});

it('audits consuming actions rather than treating read-only class-zero contexts as repairs', () => {
  // Context format2 with an additional class-zero input and safe nested action.
  const context = [2, 38, 30, 1, 10, 1, 4, 2, 1, 0, 0, 1, 0, 0, 0, 1, 1, 1, 0, 1, 1, 1];
  expect(inspect([context, [2, 8, 1, 2, 1, 1, 1]], [5, 1])).toMatchObject({ nonzeroPreserved: true, missingIsolated: true });
  expect(inspect([context, [2, 8, 1, 2, 1, 1, 0]], [5, 1])).toMatchObject({ nonzeroPreserved: true, missingIsolated: false });
});

it('distinguishes profile-inactive majors from active unsafe bodies and unsupported headers', () => {
  const unsafe = [2, 8, 1, 0, 1, 1, 1];
  expect(inspect([unsafe], [1], 0, 0, 1, 8, 0x1000)).toMatchObject({
    nonzeroPreserved: true, missingIsolated: true,
    gsubDisposition: 'profile-inactive-major', profile: 'canonical-static-v1',
  });
  expect(inspect([unsafe], [1]).nonzeroPreserved).toBe(false);
  expect(inspect([unsafe], [1], 0, 0, 1, 8, 0x20000)).toMatchObject({ reason: 'unsupported', nonzeroPreserved: undefined });
  const v = new DataView(new ArrayBuffer(16)); v.setUint32(0, 0x5000); v.setUint16(4, 8);
  expect(parseFontSupportFacts(v, new Map([[tag('maxp'), { offset: 0, length: 6 }], [tag('GSUB'), { offset: 8, length: 3 }]])))
    .toMatchObject({ reason: 'malformed', nonzeroPreserved: undefined });
});

it('keeps original coverage indices when out-of-domain inputs are unreachable', () => {
  // Two original coverage indices: valid input1 ->2, unreachable9 ->0.
  expect(inspect([[2, 10, 2, 2, 0, 1, 2, 1, 9]], [1])).toMatchObject({ nonzeroPreserved: true, missingIsolated: true });
  // Skipping index1 must not hide the unsafe reachable index0 output.
  expect(inspect([[2, 10, 2, 0, 2, 1, 2, 1, 9]], [1]).nonzeroPreserved).toBe(false);
  expect(inspect([[2, 8, 1, 9, 1, 1, 1]], [1])).toMatchObject({ reason: 'malformed', nonzeroPreserved: undefined });
  // A ligature cannot run with an unreachable consumed component.
  expect(inspect([[1, 18, 1, 8, 1, 4, 0, 2, 9, 1, 1, 1]], [4]).nonzeroPreserved).toBe(true);
});

it('admits nullable context ClassDefs under the declared profile and bounds graph depth', () => {
  expect(inspect([[2, 12, 0, 0, 0, 0, 1, 1, 1]], [6]).nonzeroPreserved).toBe(true);
  for (const count of [64, 65]) {
    const tables = Array.from({ length: count - 1 }, (_, index) => [3, 1, 1, 12, 0, index + 1, 1, 1, 1]);
    tables.push([2, 8, 1, 2, 1, 1, 1]);
    const result = inspect(tables, [...Array(count - 1).fill(5), 1]);
    expect(result.nonzeroPreserved).toBe(count === 64 ? true : undefined);
    if (count === 65) expect(result.reason).toBe('budget');
  }
});

it('separates read-only back/look predicates from reverse-substitution targets', () => {
  const context = [3, 1, 20, 1, 26, 1, 20, 1, 0, 1, 1, 1, 0, 1, 1, 1];
  expect(inspect([context, [2, 8, 1, 2, 1, 1, 1]], [6, 1])).toMatchObject({ nonzeroPreserved: true, missingIsolated: true });
  const reverse = [1, 16, 1, 22, 1, 22, 1, 2, 1, 1, 1, 1, 1, 0];
  expect(inspect([reverse], [8])).toMatchObject({ nonzeroPreserved: true, missingIsolated: true });
  reverse[10] = 0;
  expect(inspect([reverse], [8])).toMatchObject({ nonzeroPreserved: true, missingIsolated: false });
  expect(inspect([[2, 8, 1, 2, 1, 1, 1]], [1], 2, 1).missingIsolated).toBe(false);
});

it('admits nullable top-level lists while retaining zero-lookup script-selection facts', () => {
  const bytes = new Uint8Array(64), view = new DataView(bytes.buffer);
  view.setUint32(0, 0x5000); view.setUint16(4, 8);
  view.setUint32(16, 0x10000); view.setUint16(20, 10); // ScriptList only
  view.setUint16(26, 1); view.setUint32(28, tag('dev3')); view.setUint16(32, 8);
  view.setUint16(34, 4); view.setUint16(40, 0xffff); // default LangSys, no features
  const tables = new Map([[tag('maxp'), { offset: 0, length: 6 }], [tag('GSUB'), { offset: 16, length: 28 }]]);
  expect(parseFontSupportFacts(view, tables)).toMatchObject({ nonzeroPreserved: true, noErasure: true,
    gsubDisposition: 'active-open-type', gsubLookupCount: 0, anyIndic3ScriptPresent: true });
  view.setUint16(24, 63); // Non-NULL LookupList outside the table
  expect(parseFontSupportFacts(view, tables)).toMatchObject({ reason: 'malformed', nonzeroPreserved: undefined });
  tables.set(tag('GSUB'), { offset: 16, length: 9 });
  expect(parseFontSupportFacts(view, tables)).toMatchObject({ reason: 'malformed', nonzeroPreserved: undefined });
  expect(inspect([[1, 10, 1, 8, 0, 1, 1, 1]], [3])).toMatchObject({ nonzeroPreserved: true, noErasure: true });
  expect(inspect([[1, 10, 1, 8, 0, 1, 1, 1]], [4])).toMatchObject({ nonzeroPreserved: true, noErasure: true });
  const ligature = [1, 16, 1, 8, 1, 4, 0, 1, 1, 1, 1];
  expect(inspect([ligature], [4]).nonzeroPreserved).toBe(false);
});
