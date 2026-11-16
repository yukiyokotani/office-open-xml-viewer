// Burn-down list of #2263: the private-corpus inputs the legacy readers still
// reject, each with the native rejection it stops at. `private-corpus.test.ts`
// accepts a rejection only when it is listed here with that reason, and fails
// when a listed input renders, so an entry is removed by the change that makes
// its input render (or its reason is updated when the input then stops at a
// later construct).
//
// Inputs are identified by the SHA-256 of their bytes; the corpus is local and
// its file names are not published. Run the corpus test with
// OOXML_LEGACY_CORPUS_DETAILS=1 to print the local paths.
//
// The list ends with only the classes the readers reject by policy: `corrupt`,
// `limit`, `encryption` and `version` (pre-97 formats). `pending` marks an
// input that must render and does not yet.

export type LegacyCorpusRejectionClass = 'pending' | 'corrupt' | 'limit' | 'encryption' | 'version';

export interface LegacyCorpusRejection {
  readonly format: 'doc' | 'xls' | 'ppt';
  readonly class: LegacyCorpusRejectionClass;
  /** The first native rejection of the load, marker included. */
  readonly reason: string;
  /** Lowercase hex SHA-256 of each input that stops at `reason`. */
  readonly sha256: readonly string[];
}

export const LEGACY_CORPUS_REJECTIONS: readonly LegacyCorpusRejection[] = [
  {
    format: 'xls',
    class: 'pending',
    reason: 'UNSUPPORTED:BIFF picture object without a BLIP reference',
    sha256: [
      '0ddede6e6e90d4b821059aa887f9313f531edb65c82958fe9f4dae36ce2808b7',
    ],
  },
  {
    format: 'xls',
    class: 'pending',
    reason: 'UNSUPPORTED:XLS PivotTable style element is not representable',
    sha256: [
      '25987c6f20a747e15bd7c573c12278127c516150ceb5f646f552e477cc2f8228',
      '2c4bd594c119b8e696474a9b9db9f3c802445e0f28455a702d15062347a92a14',
    ],
  },
  {
    format: 'xls',
    class: 'pending',
    reason: 'UNSUPPORTED:XLS shape text linked to a formula is not projected',
    sha256: [
      '54b57c1ded6f280dbb8fa44775ccd015afc6002bd335956d19acf436e6c7eb0b',
    ],
  },
  {
    format: 'ppt',
    class: 'pending',
    reason: 'UNSUPPORTED:PowerPoint alternative shape XML custom geometry cannot be compared with the binary shape',
    sha256: [
      'da839619e16c35e281c6e66831c5fe0d12640608ddbfd7137bcad1d2afde4206',
    ],
  },
  {
    format: 'ppt',
    class: 'pending',
    reason: 'UNSUPPORTED:PowerPoint alternative shape XML fill cannot be compared with the binary shape',
    sha256: [
      '1180be153c16c3e444edc565b4a5a8ea810e9d805035eb91243eee0d5f1a4902',
      '2bdd29055fbc977180c8631409d3489e6536af01002a28954b58f0f6f3044ae5',
      '2dd5f2477c2018c42459f98593628b1465d0574c31a46c2ba675426a2ce96366',
      '4029e8998391b1942bea93eaa77ab602f66681e2e76b2b55f8caa6b7da68150a',
      '5c4cb512fc80825032ae22b0f999134b3ed378bb0ea6d3bed434e8a7307a66b2',
      '735e88e51c18c8bc6c0fea612cbbe990cf65667b7dc1d9dac07fda0d0317c86f',
      'a8175cc42a961767bb38d30bc610a8c75237c6eafe28fe803b6719fa68321b2c',
    ],
  },
];
