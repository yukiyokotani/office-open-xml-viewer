// Independently decoded synthetic Windows electronic-distribution tagged PDF.
// Checked all 864 authored text/lang/size/face assignments against PPTX XML,
// bound case labels and probe MCIDs by source geometry, decoded ToUnicode, and
// identified families from embedded font name tables. No renderer/classifier
// participates in the extraction. A cycle settles only when all three rotated
// assignments select one authored slot with the exact scalar. Fixed-family
// cycles, substitution and absent tagged glyphs cannot establish slot routing.
// These are font-slot observations, not pixel-fidelity or shaping measurements.
export const POWERPOINT_BOUNDARY_FONT_SLOT_EVIDENCE = {
  kind: 'office-observation',
  syntheticFixtureId: 'pptx-boundary-fix-1653',
  application: 'Microsoft PowerPoint for Windows',
  version: 'not recorded',
  installedFontVersions: 'not recorded on export host; authoring cmap coverage does not prove availability',
  engine: 'electronic-distribution tagged PDF',
  date: '2026-10-01',
  deckSha256: '16f69fa063eefe190d06b3f966259426f953305ca2831bdfe4ce76d68084b87f',
  pdfSha256: '2969ffae0bb822e2d8c0cfb6faaab934da2f7ce8da3b4527809970299e04271e',
  pages: 54,
  cases: 864,
  cycles: 288,
  acceptanceGroups: { latin: 42, ea: 30, inconsistent: 138, substitution: 60, unextractable: 18 },
  languages: ['en-US', 'ja-JP'],
  sizesPt: [10, 18, 32],
  // Every cell below repeats identically at all three sizes. For each endpoint:
  // triple-group index, then enUs / jaJp, each with three variants' c0/c1/c2.
  // 'substitution' means exact tagged scalar but no authored family; 'unextractable'
  // means no exact tagged scalar. Mixed slot vectors are rejected, even if every
  // emitted family was authored (a fixed-family fallback rotates the apparent slot).
  // Rotation cN assigns latin=triple[N], ea=triple[(N+1)%3], cs=triple[(N+2)%3].
  substitutedFamilies: ['Calibri', 'Arial'],
  tripleGroups: [
    [
      ['Calibri', 'Cambria', 'Arial'],
      ['Calibri', 'Cambria', 'Times New Roman'],
      ['Calibri', 'Arial', 'Times New Roman'],
    ],
    [
      ['Segoe UI Symbol', 'STIX Two Math', 'Yu Gothic'],
      ['Segoe UI Symbol', 'STIX Two Math', 'Yu Mincho'],
      ['Segoe UI Symbol', 'Yu Gothic', 'Yu Mincho'],
    ],
    [
      ['Segoe UI Symbol', 'Apple Symbols', 'Menlo'],
      ['Segoe UI Symbol', 'Apple Symbols', 'Unifont'],
      ['Segoe UI Symbol', 'Menlo', 'Unifont'],
    ],
    [
      ['Segoe UI Symbol', 'Apple Symbols', 'Menlo'],
      ['Segoe UI Symbol', 'Apple Symbols', 'STIX Two Math'],
      ['Segoe UI Symbol', 'Menlo', 'STIX Two Math'],
    ],
  ],
  observations: [
    { codePoint: 0x201D, tripleGroup: 0,
      enUs: [['latin', 'latin', 'latin'], ['latin', 'latin', 'latin'], ['latin', 'latin', 'latin']],
      jaJp: [['ea', 'ea', 'ea'], ['ea', 'ea', 'ea'], ['ea', 'ea', 'ea']] },
    { codePoint: 0x201E, tripleGroup: 0,
      enUs: [['latin', 'latin', 'latin'], ['latin', 'latin', 'latin'], ['latin', 'latin', 'latin']],
      jaJp: [['ea', 'ea', 'ea'], ['ea', 'ea', 'ea'], ['ea', 'ea', 'ea']] },
    { codePoint: 0x201F, tripleGroup: 0,
      enUs: [['latin', 'latin', 'latin'], ['latin', 'latin', 'latin'], ['latin', 'latin', 'latin']],
      jaJp: [['latin', 'latin', 'latin'], ['latin', 'latin', 'latin'], ['latin', 'latin', 'latin']] },
    { codePoint: 0x24FE, tripleGroup: 1,
      enUs: [['latin', 'ea', 'ea'], ['latin', 'ea', 'ea'], ['ea', 'ea', 'ea']],
      jaJp: [['latin', 'ea', 'ea'], ['latin', 'ea', 'ea'], ['ea', 'ea', 'ea']] },
    { codePoint: 0x24FF, tripleGroup: 1,
      enUs: [['substitution', 'ea', 'ea'], ['substitution', 'ea', 'ea'], ['ea', 'ea', 'ea']],
      jaJp: [['substitution', 'ea', 'ea'], ['substitution', 'ea', 'ea'], ['ea', 'ea', 'ea']] },
    { codePoint: 0x2500, tripleGroup: 1,
      enUs: [['latin', 'substitution', 'latin'], ['latin', 'substitution', 'latin'], ['latin', 'latin', 'latin']],
      jaJp: [['latin', 'substitution', 'latin'], ['latin', 'substitution', 'latin'], ['latin', 'latin', 'latin']] },
    { codePoint: 0x259F, tripleGroup: 2,
      enUs: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']],
      jaJp: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']] },
    { codePoint: 0x25A0, tripleGroup: 2,
      enUs: [['substitution', 'substitution', 'ea'], ['substitution', 'substitution', 'ea'], ['substitution', 'substitution', 'ea']],
      jaJp: [['substitution', 'substitution', 'ea'], ['substitution', 'substitution', 'ea'], ['substitution', 'substitution', 'ea']] },
    { codePoint: 0x25A1, tripleGroup: 2,
      enUs: [['substitution', 'substitution', 'ea'], ['substitution', 'substitution', 'ea'], ['substitution', 'substitution', 'ea']],
      jaJp: [['substitution', 'substitution', 'ea'], ['substitution', 'substitution', 'ea'], ['substitution', 'substitution', 'ea']] },
    { codePoint: 0x2618, tripleGroup: 3,
      enUs: [['unextractable', 'unextractable', 'unextractable'], ['unextractable', 'unextractable', 'unextractable'], ['unextractable', 'unextractable', 'unextractable']],
      jaJp: [['unextractable', 'unextractable', 'unextractable'], ['unextractable', 'unextractable', 'unextractable'], ['unextractable', 'unextractable', 'unextractable']] },
    { codePoint: 0x2619, tripleGroup: 3,
      enUs: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']],
      jaJp: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']] },
    { codePoint: 0x261A, tripleGroup: 3,
      enUs: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']],
      jaJp: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']] },
    { codePoint: 0x266F, tripleGroup: 3,
      enUs: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']],
      jaJp: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']] },
    { codePoint: 0x2670, tripleGroup: 3,
      enUs: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']],
      jaJp: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']] },
    { codePoint: 0x2671, tripleGroup: 3,
      enUs: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']],
      jaJp: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']] },
    { codePoint: 0x2672, tripleGroup: 3,
      enUs: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']],
      jaJp: [['latin', 'cs', 'ea'], ['latin', 'cs', 'ea'], ['latin', 'cs', 'ea']] },
  ],
  // U+201D/E/F settle across three triples × three sizes × two languages.
  // Japanese U+201F is an observed deviation; U+201D/E follow ECMA-376
  // §21.1.2.3 (latin for en-US, ea for ja-JP). Other quote-language IDs still
  // have only their prior single-triple evidence, not this independence claim.
  quoteCoverage: { 'en-us': [[0x201D, 0x201F, 'latin']],
    'ja-jp': [[0x201D, 0x201E, 'ea'], [0x201F, 0x201F, 'latin']] },
  // Withdrawal is library policy under incomplete/contradictory evidence, not
  // a claim that Office always uses the normative slot. U+24FF has six complete
  // ea cycles, contradicting the former latin override; its other 12 cycles
  // substitute. U+259F/2619/2670/2671 have 18 inconsistent cycles each.
  withdrawnOverrides: [0x24FF, 0x259F, 0x2619, 0x2670, 0x2671],
  normativeFallback: [[0x24FF, 0x24FF, 'ea'], [0x259F, 0x259F, 'ea'],
    [0x2619, 0x2619, 'ea'], [0x2670, 0x2671, 'cs']],
  evidenceGaps: {
    symbols: 'No endpoint validates a face-independent symbol deviation. U+24FE/24FF settle ea for one triple; U+2500 settles latin for one triple. Other triples are inconsistent or substituted; U+25A0/25A1 substitute throughout; U+2618 has no tagged glyph. Other symbol overrides retain their prior evidence/routing; do not infer a block or a font-name rule.',
    contexts: 'These controls contain isolated scalars only; no neighbour, cluster or shaping rule follows.',
  },
} as const;
