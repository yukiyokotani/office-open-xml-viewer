// Independently extracted Windows electronic-distribution tagged PDFs of
// synthetic controls (2026-10-01). No viewer classifier was used to derive slots.
// Each PPTX shape/text/lang/size/slot triple was checked against its manifest;
// PDF MCID text and embedded FontFile2/FontFile3 name tables identified families.
// A group settles only when all three cyclic assignments follow the same slot,
// with exact authored/display scalar multisets (bidi order may differ). Fixed
// fallback to an authored family is NOT sufficient: it changes slot on rotation.
// Missing/substituted/rasterized glyphs, uncovered neighbours, nonexact symbol
// mappings and layout controls are excluded. Exact tagged U+2047/U+2048 and
// standalone Myanmar marks have complete cycles and are included. Base-cluster
// and seam data are recorded separately; a split-font cluster cannot prove base-slot inheritance.
// coverage contains only scalar/language pairs consistent across settled contexts;
// contextualPunctuation records the isolated/Latin-neighbour routing contract.
// No pixel fidelity or font/size independence is claimed for incomplete cycles.
import type { SlotCoverageRange } from './font-slot-evidence.js';

const AR_EG: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x0660, 0x0669, 'cs'],
  [0x201F, 0x201F, 'latin'],
];

const AR_SA: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x0660, 0x0669, 'cs'],
  [0x201F, 0x201F, 'latin'],
];

const EN_US: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'latin'],
  [0x005C, 0x005C, 'latin'],
  [0x00BB, 0x00BB, 'latin'],
  [0x00D7, 0x00D7, 'ea'],
  [0x00F7, 0x00F7, 'ea'],
  [0x1022, 0x1022, 'cs'],
  [0x1028, 0x1028, 'cs'],
  [0x105A, 0x105D, 'cs'],
  [0x1061, 0x1061, 'cs'],
  [0x1065, 0x1066, 'cs'],
  [0x106E, 0x1070, 'cs'],
  [0x1075, 0x1081, 'cs'],
  [0x108E, 0x108E, 'cs'],
  [0x1090, 0x1099, 'cs'],
  [0x109E, 0x109F, 'cs'],
  [0x2018, 0x201F, 'latin'],
  [0x2047, 0x2048, 'ea'],
  [0x2160, 0x216B, 'latin'],
  [0x2170, 0x217B, 'latin'],
  [0x217F, 0x217F, 'latin'],
  [0x3298, 0x3298, 'ea'],
  [0x329A, 0x32B0, 'ea'],
  [0xA9E0, 0xA9FE, 'ea'],
  [0xAA60, 0xAA7F, 'ea'],
];

const FA_IR: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x06F4, 0x06F6, 'cs'],
  [0x201F, 0x201F, 'latin'],
];

const HE: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x201F, 0x201F, 'latin'],
];

const HE_IL: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x201F, 0x201F, 'latin'],
];

const HI_IN: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x201F, 0x201F, 'latin'],
];

const II_CN: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'latin'],
  [0x00BB, 0x00BB, 'latin'],
  [0x00D7, 0x00D7, 'ea'],
  [0x00F7, 0x00F7, 'ea'],
  [0x2018, 0x201E, 'ea'],
  [0x201F, 0x201F, 'latin'],
  [0x2047, 0x2048, 'ea'],
];

const JA_JP: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'latin'],
  [0x005C, 0x005C, 'latin'],
  [0x00BB, 0x00BB, 'latin'],
  [0x00D7, 0x00D7, 'ea'],
  [0x00F7, 0x00F7, 'ea'],
  [0x1022, 0x1022, 'cs'],
  [0x1028, 0x1028, 'cs'],
  [0x105A, 0x105D, 'cs'],
  [0x1061, 0x1061, 'cs'],
  [0x1065, 0x1066, 'cs'],
  [0x106E, 0x1070, 'cs'],
  [0x1075, 0x1081, 'cs'],
  [0x108E, 0x108E, 'cs'],
  [0x1090, 0x1099, 'cs'],
  [0x109E, 0x109F, 'cs'],
  [0x2018, 0x201E, 'ea'],
  [0x201F, 0x201F, 'latin'],
  [0x2047, 0x2048, 'ea'],
  [0x2160, 0x216B, 'latin'],
  [0x2170, 0x217B, 'latin'],
  [0x217F, 0x217F, 'latin'],
  [0x3298, 0x3298, 'ea'],
  [0x329A, 0x32B0, 'ea'],
  [0xA9E0, 0xA9FE, 'ea'],
  [0xAA60, 0xAA7F, 'ea'],
];

const KO_KR: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'latin'],
  [0x00BB, 0x00BB, 'latin'],
  [0x00D7, 0x00D7, 'ea'],
  [0x00F7, 0x00F7, 'ea'],
  [0x2018, 0x201E, 'ea'],
  [0x201F, 0x201F, 'latin'],
  [0x2047, 0x2048, 'ea'],
];

const MY_MM: readonly SlotCoverageRange[] = [
  [0x1022, 0x1022, 'cs'],
  [0x1028, 0x1028, 'cs'],
  [0x105A, 0x105D, 'cs'],
  [0x1061, 0x1061, 'cs'],
  [0x1065, 0x1066, 'cs'],
  [0x106E, 0x1070, 'cs'],
  [0x1075, 0x1081, 'cs'],
  [0x108E, 0x108E, 'cs'],
  [0x1090, 0x1099, 'cs'],
  [0x109E, 0x109F, 'cs'],
  [0xA9E0, 0xA9FE, 'ea'],
  [0xAA60, 0xAA7F, 'ea'],
];

const SYR_SY: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x201F, 0x201F, 'latin'],
];

const TH_TH: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x201F, 0x201F, 'latin'],
];

const UG_CN: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x0660, 0x0669, 'cs'],
  [0x201F, 0x201F, 'latin'],
];

const UR_IN: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x0660, 0x0669, 'cs'],
  [0x06F4, 0x06F6, 'cs'],
  [0x201F, 0x201F, 'latin'],
];

const UR_PK: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x0660, 0x0669, 'cs'],
  [0x06F4, 0x06F6, 'cs'],
  [0x201F, 0x201F, 'latin'],
];

const YI_001: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'cs'],
  [0x201F, 0x201F, 'latin'],
];

const ZH_CN: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'latin'],
  [0x00BB, 0x00BB, 'latin'],
  [0x00D7, 0x00D7, 'ea'],
  [0x00F7, 0x00F7, 'ea'],
  [0x2018, 0x201E, 'ea'],
  [0x201F, 0x201F, 'latin'],
  [0x2047, 0x2048, 'ea'],
];

const ZH_HK: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'latin'],
  [0x00BB, 0x00BB, 'latin'],
  [0x00D7, 0x00D7, 'ea'],
  [0x00F7, 0x00F7, 'ea'],
  [0x2018, 0x201E, 'ea'],
  [0x201F, 0x201F, 'latin'],
  [0x2047, 0x2048, 'ea'],
];

const ZH_MO: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'latin'],
  [0x00BB, 0x00BB, 'latin'],
  [0x00D7, 0x00D7, 'ea'],
  [0x00F7, 0x00F7, 'ea'],
  [0x2018, 0x201E, 'ea'],
  [0x201F, 0x201F, 'latin'],
  [0x2047, 0x2048, 'ea'],
];

const ZH_SG: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'latin'],
  [0x00BB, 0x00BB, 'latin'],
  [0x00D7, 0x00D7, 'ea'],
  [0x00F7, 0x00F7, 'ea'],
  [0x2018, 0x201E, 'ea'],
  [0x201F, 0x201F, 'latin'],
  [0x2047, 0x2048, 'ea'],
];

const ZH_TW: readonly SlotCoverageRange[] = [
  [0x0030, 0x0039, 'latin'],
  [0x00BB, 0x00BB, 'latin'],
  [0x00D7, 0x00D7, 'ea'],
  [0x00F7, 0x00F7, 'ea'],
  [0x2018, 0x201E, 'ea'],
  [0x201F, 0x201F, 'latin'],
  [0x2047, 0x2048, 'ea'],
];

export const POWERPOINT_EXTRA_FONT_SLOT_EVIDENCE = {
  kind: 'office-observation',
  syntheticFixtureId: 'pptx-font-slots-extra-1653',
  application: 'Microsoft PowerPoint for Windows',
  version: 'not recorded',
  installedFontVersions: 'not recorded on export host; authoring inventory is not export-host provenance',
  unicodeVersion: '13.0.0',
  engine: 'electronic-distribution tagged PDF',
  date: '2026-10-01',
  cases: 59613,
  pages: 5085,
  decks: [
    { id: 1, pages: 500, deckSha256: '3e6d7cd13b16b3e8422a2ee600bd387610cc886377326d4c46cc9ef0cc14490c',
      pdfSha256: '8c29a2711bfdb6cc849173aab1920d7c100836a88c7ace36d556078a45b4f3f3' },
    { id: 2, pages: 146, deckSha256: 'a7432a49191fb313441f5d65e0c46bab88a7e6c826e37e1189f4d439172be645',
      pdfSha256: 'ad1d9a40a005b0698a09d25e941b1a0c38fdc56c255515ba4ec7ee8b487c623f' },
    { id: 3, pages: 500, deckSha256: 'a9d9c46fb8fef1c2fd02da57d35902f37b2e44bb8b014c7af1a50e874225fac3',
      pdfSha256: '9a9ea4add7ed14da8fa7dc2f6c9c462806d602c0d0c07966ef061a7332320bd4' },
    { id: 4, pages: 500, deckSha256: '9bf21ec83898ba6c860a18d67ee81fb3889ec41b12612aea2e58b83bcb911ca3',
      pdfSha256: '68c7cee461c454023be7f923f17f560a77b6d6aa5cfb02b4c1666fc5bfdfff2b' },
    { id: 5, pages: 500, deckSha256: '46ebbf13b1ff2b0d23e265110ce7ce78f64ccd486dd97fc159863fddd303ca41',
      pdfSha256: 'df5528302a382c4b2808ae592190a488d6493d03a9eeb3285984a022ce7236fc' },
    { id: 6, pages: 500, deckSha256: '2196c78f45a97c5cbb1fbc30ed3e09df97a551527a5255cc4ca410c58ae2c400',
      pdfSha256: '71d0d64c841e72a14f5d55daa46c276913b41e09a552c98e0040e2605fc368ab' },
    { id: 7, pages: 500, deckSha256: '6b73d5c448169c842b415566a87807b9764dbdf7959f7e6399af480c2f614028',
      pdfSha256: 'fbbcc45d267043434eb8feb409b9c1097ef509e93f595410593965ee0f62aa4e' },
    { id: 8, pages: 162, deckSha256: '89047f6eef0256277f6e3543ae087940b1bc3ae414d5c27c0d097591c2ede3da',
      pdfSha256: '244116b5a93bddaa0ffdb3b9d331ab0daa2b1c2b6effd82babccbbca64a22690' },
    { id: 9, pages: 500, deckSha256: 'facb19dc21e49f26c0dfd8ef8ad1cddb126dd6a6ff6c04f9b0322844daab6b58',
      pdfSha256: 'a24957accb34e2684c51d0e1c6767c6bb7457dcd5716195e9cdd9fd58990dafb' },
    { id: 10, pages: 500, deckSha256: '88eb457e0e3a29d3d75a5a5db1ab09ff2fe8dabcfc313adefc83960fcf7b3f91',
      pdfSha256: '5247a884beb9178fa0addcfe05ac066838c5f6d8986a123c2e65c12cef91701f' },
    { id: 11, pages: 184, deckSha256: '0893745f2af074cfc80d813383e6044c90acc7d359dd72ab089d3069742181ea',
      pdfSha256: '489f104e7db34314b4f2f9f1767842ae926b1ca57e9eb506a934b6dc331960fd' },
    { id: 12, pages: 500, deckSha256: 'a37cbecb74e30c488091b6bf78d25423adf5cd551a5a0e14af7828476a7c8147',
      pdfSha256: 'd7e72da9c25fbee7332b1e32de8a3d1b799b60b7080d839030f56f2d54fa0afa' },
    { id: 13, pages: 93, deckSha256: '6a1b1c77eb182faf5cecbd28fa82f3c1cddd3c93127cc1f90bc3e872ab9972c9',
      pdfSha256: '4d2f77a04ea5dd5991b1880c4c7cab54ae1fb1500b71912e0c4c3d0bf065d0c7' },
  ],
  // Restore 178 complete exact cycles formerly excluded by Unicode category:
  // 166 U+2047/U+2048 cycles (contextual and ea/altLang counterexamples), plus
  // 12 standalone marks. Nonexact/substituted observations remain ineligible.
  acceptanceGroups: {"settled": 3054, "inconsistent-cycle": 9294, "fallback-or-unextractable": 4565, "ineligible": 3044},
  coverage: {
    'ar-eg': AR_EG,
    'ar-sa': AR_SA,
    'en-us': EN_US,
    'fa-ir': FA_IR,
    'he': HE,
    'he-il': HE_IL,
    'hi-in': HI_IN,
    'ii-cn': II_CN,
    'ja-jp': JA_JP,
    'ko-kr': KO_KR,
    'my-mm': MY_MM,
    'syr-sy': SYR_SY,
    'th-th': TH_TH,
    'ug-cn': UG_CN,
    'ur-in': UR_IN,
    'ur-pk': UR_PK,
    'yi-001': YI_001,
    'zh-cn': ZH_CN,
    'zh-hk': ZH_HK,
    'zh-mo': ZH_MO,
    'zh-sg': ZH_SG,
    'zh-tw': ZH_TW,
  },
  // The same triple at 20 pt: » × ÷, U+2018–201E and U+2047/U+2048 follow cs alone,
  // latin between A/B. Native neighbours/ascending sequences are uncovered,
  // so they cannot establish a broader itemization rule. Paragraph routing
  // implements only isolated and Latin-neighbour contexts; other contexts
  // keep their prior scalar routing, including the original he/ar » result.
  contextualPunctuation: {
    languages: ["ar-eg", "ar-sa", "fa-ir", "he", "he-il", "hi-in", "syr-sy", "th-th", "ug-cn", "ur-in", "ur-pk", "yi-001"],
    ranges: [[0x00BB, 0x00BB], [0x00D7, 0x00D7], [0x00F7, 0x00F7], [0x2018, 0x201E], [0x2047, 0x2048]],
    // 48 U+2047/U+2048 cycles (12 languages × two scalars × two contexts);
    // 40 disagreed with the prior paragraph adapter. Exact cmap/ToUnicode
    // scalars in embedded authored families distinguish these from substitutions.
    alone: 'cs',
    latinBetween: 'latin',
    triple: ['Segoe UI Symbol', 'Meiryo UI', 'Microsoft Sans Serif'],
    sizePt: 20,
  },
  altLangControls: {
    primaryLanguage: 'en-US',
    alternateLanguages: ['ar-EG', 'ar-SA', 'fa-IR', 'he', 'he-IL', 'hi-IN', 'ii-CN',
      'ja-JP', 'ko-KR', 'syr-SY', 'th-TH', 'ug-CN', 'ur-IN', 'ur-PK', 'yi-001',
      'zh-CN', 'zh-HK', 'zh-MO', 'zh-SG', 'zh-TW'],
    europeanDigits: [0x0030, 0x0039],
    // 80 exact cycles: these scalars retain ea under every listed altLang,
    // alone and between A/B. Together with the 38 ordinary ea cycles, these
    // bound the 48 contextual cs/latin cycles above (166 complete in total).
    punctuation: { codePoints: [0x2047, 0x2048], slot: 'ea', settledGroups: 80 },
    contexts: ['alone', 'latin-between'],
    slot: 'latin',
  },
  clusterControls: {
    script: 'Myanmar',
    uniformCsGroups: 216,
    splitCsEaGroups: 24,
    contexts: ['base-mark', 'base-mark-seam'],
    // This is a split-font observation, deliberately absent from scalar coverage.
    splitCodePoints: [0x1000, 0xA9E5, 0xAA7B, 0xAA7C, 0xAA7D],
    languages: ['en-US', 'my-MM', 'ja-JP'],
    baseSlot: 'cs',
    markSlot: 'ea',
    sizePt: 20,
  },
  scriptControls: {
    'Ethiopic': { scalarRanges: [
    ], settledGroups: 0, gaps: {"inconsistent-cycle": 1558, "ineligible": 9},
      authoringExclusions: {"unassigned": 65} },
    'Mongolian': { scalarRanges: [
    ], settledGroups: 0, gaps: {"inconsistent-cycle": 528, "ineligible": 27, "fallback-or-unextractable": 6},
      authoringExclusions: {"unassigned": 38, "fewer-than-three-covering-families": 13} },
    'Yi': { scalarRanges: [
    ], settledGroups: 0, gaps: {"inconsistent-cycle": 3666},
      authoringExclusions: {"unassigned": 12} },
    'Cherokee': { scalarRanges: [
    ], settledGroups: 0, gaps: {"fallback-or-unextractable": 516},
      authoringExclusions: {"fewer-than-three-covering-families": 2, "unassigned": 4} },
    'Armenian': { scalarRanges: [
    ], settledGroups: 0, gaps: {"fallback-or-unextractable": 304},
      authoringExclusions: {"unassigned": 5, "fewer-than-three-covering-families": 2} },
    'Georgian': { scalarRanges: [
    ], settledGroups: 0, gaps: {"fallback-or-unextractable": 286},
      authoringExclusions: {"fewer-than-three-covering-families": 80, "unassigned": 18} },
    'Tibetan': { scalarRanges: [
    ], settledGroups: 0, gaps: {"inconsistent-cycle": 814, "ineligible": 231, "fallback-or-unextractable": 142},
      authoringExclusions: {"unassigned": 45, "fewer-than-three-covering-families": 2} },
    'Myanmar': { scalarRanges: [
      [0x1022, 0x1022, 'cs'],
      [0x1028, 0x1028, 'cs'],
      [0x105A, 0x105D, 'cs'],
      [0x1061, 0x1061, 'cs'],
      [0x1065, 0x1066, 'cs'],
      [0x106E, 0x1070, 'cs'],
      [0x1075, 0x1081, 'cs'],
      [0x108E, 0x108E, 'cs'],
      [0x1090, 0x1099, 'cs'],
      [0x109E, 0x109F, 'cs'],
      [0xA9E0, 0xA9FE, 'ea'],
      [0xAA60, 0xAA7F, 'ea'],
    ], settledGroups: 587, gaps: {"inconsistent-cycle": 340, "ineligible": 174, "fallback-or-unextractable": 18},
      authoringExclusions: {"unassigned": 1} },
    'Syriac': { scalarRanges: [
    ], settledGroups: 0, gaps: {"inconsistent-cycle": 176, "ineligible": 87, "fallback-or-unextractable": 168},
      authoringExclusions: {"unassigned": 8, "fewer-than-three-covering-families": 11} },
    'Tai Le': { scalarRanges: [
    ], settledGroups: 0, gaps: {"inconsistent-cycle": 111},
      authoringExclusions: {"unassigned": 13} },
    'New Tai Lue': { scalarRanges: [
    ], settledGroups: 0, gaps: {"inconsistent-cycle": 277},
      authoringExclusions: {"unassigned": 13} },
    'Devanagari': { scalarRanges: [
    ], settledGroups: 0, gaps: {"ineligible": 102, "fallback-or-unextractable": 78, "inconsistent-cycle": 438},
      authoringExclusions: {"fewer-than-three-covering-families": 32} },
  },
  evidenceGaps: {
    symbolBoundaries: {
      codePoints: [0x24FE, 0x24FF, 0x259F, 0x25A0, 0x2619, 0x2670, 0x201E, 0x201F],
      sizesPt: [10, 18, 32],
      settledGroups: 0,
      reason: 'This original matrix has no complete cycles. Replacement evidence is in POWERPOINT_BOUNDARY_FONT_SLOT_EVIDENCE: quote independence settles, while tested symbol overrides are withdrawn because cycles remain inconsistent/substituted or contradict latin.',
    },
    scripts: 'Only Myanmar has complete scalar cycles. Other scripts retain previous routing; no block-level extrapolation across missing/mark/unassigned scalars.',
    clusters: 'Myanmar base U+1000 plus U+A9E5/U+AA7B/U+AA7C/U+AA7D selects cs/ea in single-run and same-language seam controls. Standalone marks independently select ea in all 12 exact-scalar cycles (four scalars × three languages at 20 pt), with no dotted circle or extra extracted scalar. Other bases, longer clusters and language-changing seams retain previous inheritance pending broader controls.',
    context: 'Measured standalone cs and Latin-surrounded latin punctuation are implemented for the 12 exact language IDs. Native/ascending controls with uncovered neighbours remain inconclusive and retain previous routing.',
    sizeAndFaces: 'Script/digit controls are at 20 pt. Digit/punctuation scalar cycles use one distinguishable triple; Myanmar extension cycles use Myanmar Text/Noto Sans Myanmar/Noto Serif Myanmar. Replacement boundary controls settle quote face/size independence; symbol variation remains inconclusive (see POWERPOINT_BOUNDARY_FONT_SLOT_EVIDENCE).',
  },
} as const;
