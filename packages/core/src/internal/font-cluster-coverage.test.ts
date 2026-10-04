import { expect, it } from 'vitest';
import { analyzeFontResourceSupport, fontResourceCoversCluster } from './font-cluster-coverage.js';
import { canonicalClusterPairs } from '../test-fixtures/canonical-clusters.js';

it('requires all marks or a supported canonical equivalent from the same resource', () => {
  const covered = (points: number[]) => (cp: number) => points.includes(cp);
  expect(fontResourceCoversCluster('A\u0301', covered([0x41]))).toBe(false);
  expect(fontResourceCoversCluster('A\u0301', covered([0x41, 0x301]))).toBe(true);
  expect(fontResourceCoversCluster('A\u0301', covered([0xc1]))).toBe(true);
  expect(fontResourceCoversCluster('\u1100\u1161\u11a8', covered([0xac01]))).toBe(true);
  expect(fontResourceCoversCluster('A\u0301', (cp) => cp === 0x41 ? true : undefined)).toBeUndefined();
});

it('certifies the same resource for generated canonical equivalents with remaining marks', () => {
  for (const [a, b] of canonicalClusterPairs()) {
    const points = new Set([...a.normalize('NFC')].map((ch) => ch.codePointAt(0)));
    const covers = (cp: number) => points.has(cp);
    expect([fontResourceCoversCluster(a, covers), fontResourceCoversCluster(b, covers)],
      [...a].map((ch) => ch.codePointAt(0)?.toString(16)).join(' ')).toEqual([true, true]);
  }
});

it('handles reordered marks, canonical singletons and decomposed-only resources consistently', () => {
  for (const [a, b] of [['A\u0301\u0323', 'A\u0323\u0301'], ['\u212b', 'Å'], ['\u2126', 'Ω']]) {
    for (const form of ['NFC', 'NFD'] as const) {
      const points = new Set([...a.normalize(form)].map((ch) => ch.codePointAt(0)));
      expect([fontResourceCoversCluster(a, (cp) => points.has(cp)),
        fontResourceCoversCluster(b, (cp) => points.has(cp))]).toEqual([true, true]);
    }
  }
  // Compatibility decomposition does not confer canonical ownership.
  expect(fontResourceCoversCluster('ﬀ', (cp) => cp === 0x66)).toBe(false);
});

it('accepts resource-supported intermediate compositions and respects canonical blocking', () => {
  const covered = new Set([0xc5, 0x301, 0x307]); // Å + acute + dot; neither complete NFC nor NFD.
  for (const text of ['A\u030a\u0301\u0307', 'Å\u0301\u0307', 'Ǻ\u0307']) {
    expect(fontResourceCoversCluster(text, (cp) => covered.has(cp))).toBe(true);
  }
  // Acute and ring have the same class; acute blocks A+ring composition.
  expect(fontResourceCoversCluster('A\u0301\u030a', (cp) => covered.has(cp))).toBe(false);
  // A lower-class retained mark allows composition across it.
  expect(fontResourceCoversCluster('A\u0323\u030a', (cp) => cp === 0xc5 || cp === 0x323)).toBe(false);
  expect(fontResourceCoversCluster('A\u0323\u030a', (cp) => cp === 0x41 || cp === 0xc5 || cp === 0x323)).toBe(true);
  // Bengali class-zero spacing marks compose with one another, not the base.
  expect(fontResourceCoversCluster('ক\u09c7\u09be', (cp) => cp === 0x995 || cp === 0x9cb)).toBe(true);
});

it('keeps sequence-only ownership unknown even if every scalar has a cmap entry', () => {
  for (const cluster of ['§\ufe0e', '§\ufe0f', '👩\u200d💻']) {
    expect(fontResourceCoversCluster(cluster, () => true)).toBeUndefined();
    expect(fontResourceCoversCluster(cluster, () => false)).toBeUndefined();
  }
});

// Independent browser resources distinguish the recursive cut from a full-NFD
// search: the supported parent is reachable without its unsupported child.
it('reaches a directly supported Greek decomposition cut from actual NFC', () => {
  const points = new Set([0x1f0c, 0x345]);
  for (const text of ['Α\u0313\u0301\u0345', '\u1f8c', '\u1f0c\u0345']) {
    expect(fontResourceCoversCluster(text, (cp) => points.has(cp))).toBe(true);
  }
});

it('does not certify a multi-atom Hangul grapheme as an absent resource', () => {
  expect(fontResourceCoversCluster('각ᆨ', (cp) => cp === 0xac01)).toBeUndefined();
  expect(fontResourceCoversCluster('각ᆨ', (cp) => cp === 0x11a8)).toBeUndefined();
});

it('keeps resource attribution bounded for a long mixed Hangul shaping span', () => {
  // A mapped Latin prefix before the syllable must not be rescanned for every
  // trailing jamo. The production paragraph path validates this whole span.
  const text = 'A'.repeat(64_000) + '각' + '\u11a8'.repeat(64_000);
  const started = performance.now();
  const support = analyzeFontResourceSupport(text,
    cp => cp === 0x41 || cp === 0xac01 || cp === 0x11a8,
    { schema: 'ot-definedness-1', glyphCount: 4, nonzeroPreserved: true,
      missingIsolated: true, noErasure: true, anyIndic3ScriptPresent: false });
  expect(support.kind).toBe('complete');
  expect(performance.now() - started).toBeLessThan(5_000);
}, 60_000);

it('admits pinned simple syllables while declining unimplemented script preprocessing', () => {
  for (const text of ['कि', 'ကေ', 'কো']) expect(fontResourceCoversCluster(text, () => true)).toBe(true);
  for (const text of ['กํา', 'កេ', 'කි', 'ש', 'क्', 'क्क', 'က\u1039က', 'अा']) {
    expect(fontResourceCoversCluster(text, () => true), text).toBeUndefined();
  }
  expect(fontResourceCoversCluster('\ud800', () => true)).toBeUndefined();
});

it('requires erasure safety for an inserted dotted circle without turning uncertainty into absence', () => {
  const facts = { schema: 'ot-definedness-1' as const, glyphCount: 8, nonzeroPreserved: true,
    missingIsolated: true, noErasure: false, anyIndic3ScriptPresent: false,
    erasureSafeRanges: [[0x301, 0x301]] as const };
  expect(analyzeFontResourceSupport('\u0301', cp => cp === 0x301, facts).kind).toBe('complete');
  expect(analyzeFontResourceSupport('\u0301', cp => cp === 0x301 || cp === 0x25cc, facts).kind).toBe('unknown');
  expect(analyzeFontResourceSupport('\u0301', cp => cp === 0x25cc,
    { ...facts, erasureSafeRanges: [[0x25cc, 0x25cc]] }).kind).toBe('absent');
  expect(analyzeFontResourceSupport('क', () => true, { ...facts, noErasure: true, anyIndic3ScriptPresent: true }).kind).toBe('unknown');
});
