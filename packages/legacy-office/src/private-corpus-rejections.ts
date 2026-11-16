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
    format: 'doc',
    class: 'pending',
    reason: 'UNSUPPORTED:Word aligned drawing disagrees with its horizontal SPA origin',
    sha256: [
      'c31eb2dd550d184e387eeeaad5b328a08359ccb6ff16400dbd21b2ceb388eaad',
    ],
  },
  {
    format: 'doc',
    class: 'pending',
    reason: 'UNSUPPORTED:Word drawing shape complex property 0x00c0 is not supported',
    sha256: [
      '5a2b7047b16ddd1ecaafde6b091c4375a1df52f080ba8a42f0b8531b8c5b4a0c',
    ],
  },
  {
    format: 'doc',
    class: 'pending',
    reason: 'UNSUPPORTED:Word non-solid drawing fills are not supported',
    sha256: [
      '23e199353de6798ae1772c25b78031148f7c87eac8f92da111c888f232929b90',
      '71218c7385aa2df67e1a1afd82b04f51d6821ede5d5eed28aed0e80fd590fffd',
    ],
  },
];
