/** Representative canonical pairs: Latin/Greek composition, algorithmic
 * Hangul, and excluded singleton decompositions with a retained mark. Whole
 * Unicode scalar sweeps only repeated the full-NFC resource case; the recursive
 * and opposite resource cuts have distinct regressions at the support boundary. */
export function canonicalClusterPairs(): [string, string][] {
  return [
    ['À\u0307', 'A\u0300\u0307'],
    ['Ὄ\u0307', 'Ο\u0313\u0301\u0307'],
    ['가\u0307', '\u1100\u1161\u0307'],
    ['Ω\u0307', 'Ω\u0307'],
    ['Å\u0307', 'A\u030a\u0307'],
  ];
}
