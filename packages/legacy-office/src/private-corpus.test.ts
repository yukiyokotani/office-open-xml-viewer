import { createHash } from 'node:crypto';
import { readdir, readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it, vi } from 'vitest';
import {
  materializeDocxDocument,
  materializePptxPresentation,
  materializeXlsxWorkbook,
  skia,
  skiaFactory,
} from './node/node-facade.js';
import type { ModelSource, ModelSourceTarget } from '@silurus/ooxml-core';
import {
  LEGACY_CORPUS_REJECTIONS,
  type LegacyCorpusRejection,
  type LegacyCorpusRejectionClass,
} from './private-corpus-rejections.js';
import { testDocSource, testPptSource, testXlsSource } from './test-sources.js';

// Opt-in walk of the local, uncommitted Office-produced corpus through the Node
// facade with the direct legacy readers. A sample passes only when it
// materializes. A rejection passes only when `private-corpus-rejections.ts`
// lists the sample with the native rejection the load stops at; a listed
// sample that materializes, or is absent from the corpus, fails, so the list
// cannot go stale. Anything else
// fails too: a sample the legacy source does not claim (it would silently take
// the OOXML path), a WASM trap, a poisoned runtime, a non-Error rejection or a
// size-limit error.
// Output is aggregate counts only; a failure is identified by its index in the
// sorted walk and a digest prefix, never by its name, path or content, unless
// OOXML_LEGACY_CORPUS_DETAILS=1 asks for the local path and native diagnostic.
const runPrivateCorpus = process.env.OOXML_LEGACY_CORPUS === '1';
const corpusRoot = process.env.OOXML_LEGACY_CORPUS_ROOT;
const details = process.env.OOXML_LEGACY_CORPUS_DETAILS === '1';

// The public error carries no native diagnostic by design. The source modules
// run in this realm under Node, so the walk observes the strings the native
// readers return at the boundary that decodes them.
const nativeFailures = vi.hoisted(() => [] as string[]);
vi.mock('./native-rejection.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./native-rejection.js')>();
  return {
    ...actual,
    nativeRejection(error: unknown, format: Parameters<typeof actual.nativeRejection>[1]) {
      if (typeof error === 'string') nativeFailures.push(error);
      return actual.nativeRejection(error, format);
    },
  };
});

/** Native markers of a rejected input, by the classes that may carry them. */
const REJECTION_MARKERS = {
  'UNSUPPORTED:': { classes: ['pending', 'corrupt', 'version'], code: 'unsupported-content' },
  'LEGACY_UNSUPPORTED_ENCRYPTION:': { classes: ['encryption'], code: 'unsupported-encryption' },
  // A reader limit has no public code: its Error carries the native diagnostic.
  'LEGACY_RESOURCE_LIMIT:': { classes: ['limit'], code: undefined },
  'OOXML_RESOURCE_LIMIT:': { classes: ['limit'], code: 'ooxml-resource-limit' },
} as const satisfies Record<string, { classes: readonly LegacyCorpusRejectionClass[]; code: string | undefined }>;

function rejectionMarker(native: string): keyof typeof REJECTION_MARKERS | undefined {
  return (Object.keys(REJECTION_MARKERS) as (keyof typeof REJECTION_MARKERS)[])
    .find((marker) => native.startsWith(marker));
}

/**
 * The native rejection of a load: the first native failure that carries a
 * rejection marker. Other native answers are protocol signals, not rejections
 * of the input: a reader without a per-slide ledger answers the PPTX cursor's
 * usage query with "slide cursor usage is unavailable" before and after the
 * rejecting pull, and a closed cursor refuses the calls that follow it.
 */
function nativeRejectionOf(failures: readonly string[]): string | undefined {
  return failures.find((failure) => rejectionMarker(failure) !== undefined);
}

/** Wrap a source so the walk can tell whether it claimed the last input. */
function observed<T extends ModelSourceTarget>(source: ModelSource<T>): { source: ModelSource<T>; claimed(): boolean } {
  let claimed = false;
  return {
    source: {
      target: source.target,
      claim: (bytes) => (claimed = source.claim(bytes)),
      beginLoad: () => source.beginLoad(),
    },
    claimed: () => claimed,
  };
}

const formats = [
  {
    from: 'doc',
    to: 'docx',
    directory: new URL('../public/private/doc/', import.meta.url),
    open: (bytes: Uint8Array, source = observed(testDocSource())) => ({
      claimed: source.claimed,
      done: materializeDocxDocument(bytes, { legacyDoc: source.source }),
    }),
  },
  {
    from: 'xls',
    to: 'xlsx',
    directory: new URL('../public/private/xls/', import.meta.url),
    // The factory canvas measures the Normal font; without it drawings are omitted.
    open: (bytes: Uint8Array, source = observed(testXlsSource())) => ({
      claimed: source.claimed,
      done: materializeXlsxWorkbook(bytes, {
        legacyXls: source.source,
        ...(skia ? { factory: skiaFactory() } : {}),
      }),
    }),
  },
  {
    from: 'ppt',
    to: 'pptx',
    directory: new URL('../public/private/ppt/', import.meta.url),
    open: (bytes: Uint8Array, source = observed(testPptSource())) => ({
      claimed: source.claimed,
      done: materializePptxPresentation(bytes, { legacyPpt: source.source }),
    }),
  },
] as const;

async function corpusFiles(format: (typeof formats)[number]): Promise<{ directory: string; names: string[] }> {
  const directory = corpusRoot
    ? join(corpusRoot, 'packages', 'legacy-office', 'public', 'private', format.from)
    : fileURLToPath(format.directory);
  const names: string[] = [];
  // Retain relative directories: equal basenames are distinct corpus inputs.
  async function walk(relative: string): Promise<void> {
    for (const entry of await readdir(join(directory, relative), { withFileTypes: true })) {
      if (entry.name.startsWith('.') || entry.name.startsWith('~$')) continue;
      const path = join(relative, entry.name);
      if (entry.isDirectory()) await walk(path);
      else if (entry.isFile() && extension(path) === format.from) names.push(path);
      // Do not follow symlinks out of the selected corpus.
    }
  }
  await walk('');
  return { directory, names: names.sort() };
}

/** A content-free label for a rejection: its class and code only. */
function rejectionKind(error: unknown): string {
  if (!(error instanceof Error)) return `non-Error ${typeof error}`;
  const code = (error as { code?: unknown }).code;
  return typeof code === 'string' ? `${error.name}:${code}` : error.name;
}

type Outcome =
  | { readonly materialized: true }
  | { readonly materialized: false; readonly error: unknown; readonly native: string | undefined };

/** Why a load breaks the corpus contract, or undefined when it keeps it. */
function violation(outcome: Outcome, listed: LegacyCorpusRejection | undefined): string | undefined {
  if (outcome.materialized) {
    return listed ? 'materializes but is listed as a rejection: remove it from the list' : undefined;
  }
  const kind = rejectionKind(outcome.error);
  if (!listed) return `${kind} is not a listed rejection`;
  const marker = outcome.native === undefined ? undefined : rejectionMarker(outcome.native);
  if (!marker) return `${kind} without a native rejection`;
  const { code } = REJECTION_MARKERS[marker];
  const error = outcome.error;
  if (!(error instanceof Error) || (code === undefined
    ? error.message !== outcome.native
    : (error as { code?: unknown }).code !== code)) {
    return `${kind} does not carry the code of its native rejection`;
  }
  if (outcome.native !== listed.reason) return `${kind} stops at another rejection than the listed one`;
  return undefined;
}

describe('legacy corpus rejection list', () => {
  it('names every input once, with a reason of its class', () => {
    const digests = LEGACY_CORPUS_REJECTIONS.flatMap((entry) => entry.sha256.map((digest) => `${entry.format}:${digest}`));
    expect(new Set(digests).size).toBe(digests.length);
    for (const entry of LEGACY_CORPUS_REJECTIONS) {
      expect(entry.sha256.length).toBeGreaterThan(0);
      for (const digest of entry.sha256) expect(digest).toMatch(/^[0-9a-f]{64}$/u);
      const marker = rejectionMarker(entry.reason);
      expect(marker, entry.reason).toBeDefined();
      expect(REJECTION_MARKERS[marker!].classes, entry.reason).toContain(entry.class);
    }
  });

  it('accepts a rejection only for a listed input that stops at the listed reason', () => {
    const listed: LegacyCorpusRejection = {
      format: 'ppt', class: 'pending', reason: 'UNSUPPORTED:listed', sha256: ['0'.repeat(64)],
    };
    const rejected = (native: string | undefined, code = 'unsupported-content'): Outcome => ({
      materialized: false,
      error: Object.assign(new Error('rejected'), { code }),
      native,
    });
    expect(violation({ materialized: true }, undefined)).toBeUndefined();
    expect(violation(rejected('UNSUPPORTED:listed'), listed)).toBeUndefined();
    expect(violation({ materialized: true }, listed)).toMatch(/remove it from the list/u);
    expect(violation(rejected('UNSUPPORTED:listed'), undefined)).toMatch(/not a listed rejection/u);
    expect(violation(rejected('UNSUPPORTED:other'), listed)).toMatch(/another rejection/u);
    expect(violation(rejected(undefined), listed)).toMatch(/without a native rejection/u);
    expect(violation(rejected('UNSUPPORTED:listed', 'invalid-file'), listed)).toMatch(/code of its native rejection/u);
    // A limit is accepted only as the Error that carries its diagnostic.
    const limit: LegacyCorpusRejection = { ...listed, class: 'limit', reason: 'LEGACY_RESOURCE_LIMIT:budget' };
    const failed = (error: unknown): Outcome => ({ materialized: false, error, native: limit.reason });
    expect(violation(failed(new Error(limit.reason)), limit)).toBeUndefined();
    expect(violation(failed(limit.reason), limit)).toMatch(/code of its native rejection/u);
    expect(violation(failed(new Error('WASM runtime trapped')), limit)).toMatch(/code of its native rejection/u);
    // The PPTX cursor's usage query is answered before the rejecting pull.
    expect(nativeRejectionOf([
      'slide cursor usage is unavailable',
      'UNSUPPORTED:listed',
      'PowerPoint direct cursor is closed',
    ])).toBe('UNSUPPORTED:listed');
    expect(nativeRejectionOf(['slide cursor usage is unavailable'])).toBeUndefined();
  });
});

describe.skipIf(!runPrivateCorpus)('local Office-produced legacy corpus through the direct readers', () => {
  it.each(formats)('materializes every .$from sample except the listed rejections', async (format) => {
    const { directory, names } = await corpusFiles(format);
    expect(names.length).toBeGreaterThan(0);
    const listedByDigest = new Map(LEGACY_CORPUS_REJECTIONS
      .filter((entry) => entry.format === format.from)
      .flatMap((entry) => entry.sha256.map((digest) => [digest, entry] as const)));
    let materialized = 0;
    const rejected: Record<LegacyCorpusRejectionClass, number> = {
      pending: 0, corrupt: 0, limit: 0, encryption: 0, version: 0,
    };
    const failures: string[] = [];
    const unseen = new Set(listedByDigest.keys());
    for (const [index, name] of names.entries()) {
      const bytes = new Uint8Array(await readFile(join(directory, name)));
      const digest = createHash('sha256').update(bytes).digest('hex');
      const listed = listedByDigest.get(digest);
      unseen.delete(digest);
      nativeFailures.length = 0;
      const { claimed, done } = format.open(bytes);
      let outcome: Outcome;
      try {
        await done;
        outcome = { materialized: true };
      } catch (error) {
        outcome = { materialized: false, error, native: nativeRejectionOf(nativeFailures) };
      }
      const broken = claimed() ? violation(outcome, listed) : 'not claimed by the legacy source';
      if (broken) {
        failures.push(`#${index} sha256:${digest.slice(0, 12)} ${broken}`
          + (details ? ` [${name}] [${outcome.materialized ? '' : outcome.native ?? ''}]` : ''));
      } else if (outcome.materialized) materialized += 1;
      else rejected[listed!.class] += 1;
    }
    // A listed input that left the corpus would keep its exception forever.
    for (const digest of unseen) failures.push(`sha256:${digest.slice(0, 12)} is listed but not in the corpus`);
    console.info(`legacy ${format.from} corpus: ${names.length} samples, ${materialized} materialized, `
      + `${Object.entries(rejected).map(([kind, count]) => `${count} ${kind}`).join(', ')} listed rejections, `
      + `${failures.length} failed`);
    expect(failures).toEqual([]);
  }, 1_800_000);
});

function extension(name: string): string {
  return name.slice(name.lastIndexOf('.') + 1).toLowerCase();
}
