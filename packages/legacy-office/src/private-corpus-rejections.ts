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

export const LEGACY_CORPUS_REJECTIONS: readonly LegacyCorpusRejection[] = [];
