import { createHash } from 'node:crypto';
import { readdir, readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  materializePptxPresentation,
  materializeXlsxWorkbook,
  skia,
  skiaFactory,
} from './node/node-facade.js';
import type { ModelSource, ModelSourceTarget } from '@silurus/ooxml-core';
import { testPptSource, testXlsSource } from './test-sources.js';

// Opt-in walk of the local, uncommitted Office-produced corpus through the Node
// facade with the direct legacy readers. A sample passes when it either
// materializes or is rejected fail-closed by the direct reader (an Error whose
// message carries `UNSUPPORTED:`). Anything else fails: a sample the legacy
// source does not claim (it would silently take the OOXML path), a WASM trap,
// a poisoned runtime, a non-Error rejection or a size-limit error.
// Output is aggregate counts only; a failure is identified by its index in the
// sorted walk and a digest prefix, never by its name, path or content.
const runPrivateCorpus = process.env.OOXML_LEGACY_CORPUS === '1';
const corpusRoot = process.env.OOXML_LEGACY_CORPUS_ROOT;

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
    from: 'xls',
    to: 'xlsx',
    directory: new URL('../../xlsx/public/private/', import.meta.url),
    // The factory canvas measures the Normal font; without it drawings are omitted.
    open: (bytes: Uint8Array, source = observed(testXlsSource())) => ({
      claimed: source.claimed,
      done: materializeXlsxWorkbook(bytes, {
        modelSources: [source.source],
        ...(skia ? { factory: skiaFactory() } : {}),
      }),
    }),
  },
  {
    from: 'ppt',
    to: 'pptx',
    directory: new URL('../../pptx/public/private/', import.meta.url),
    open: (bytes: Uint8Array, source = observed(testPptSource())) => ({
      claimed: source.claimed,
      done: materializePptxPresentation(bytes, { modelSources: [source.source] }),
    }),
  },
] as const;

async function corpusFiles(format: (typeof formats)[number]): Promise<{ directory: string; names: string[] }> {
  const directory = corpusRoot
    ? join(corpusRoot, 'packages', format.to, 'public', 'private')
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

function isFailClosedRejection(error: unknown): boolean {
  return error instanceof Error && /\bUNSUPPORTED:/.test(error.message);
}

/** A content-free label for an unexpected rejection: its class and code only. */
function rejectionKind(error: unknown): string {
  if (!(error instanceof Error)) return `non-Error ${typeof error}`;
  const code = (error as { code?: unknown }).code;
  return typeof code === 'string' ? `${error.name}:${code}` : error.name;
}

describe.skipIf(!runPrivateCorpus)('local Office-produced legacy corpus through the direct readers', () => {
  it.each(formats)('materializes or fail-closed rejects every .$from sample', async (format) => {
    const { directory, names } = await corpusFiles(format);
    expect(names.length).toBeGreaterThan(0);
    let materialized = 0;
    let unsupported = 0;
    const failures: string[] = [];
    for (const [index, name] of names.entries()) {
      const bytes = new Uint8Array(await readFile(join(directory, name)));
      const digest = createHash('sha256').update(bytes).digest('hex').slice(0, 12);
      const { claimed, done } = format.open(bytes);
      let outcome: 'materialized' | 'unsupported' | string;
      try {
        await done;
        outcome = 'materialized';
      } catch (error) {
        outcome = isFailClosedRejection(error) ? 'unsupported' : rejectionKind(error);
      }
      if (!claimed()) outcome = 'not claimed by the legacy source';
      if (outcome === 'materialized') materialized += 1;
      else if (outcome === 'unsupported') unsupported += 1;
      else failures.push(`#${index} sha256:${digest} ${outcome}`);
    }
    console.info(`legacy ${format.from} corpus: ${names.length} samples, ${materialized} materialized, `
      + `${unsupported} fail-closed UNSUPPORTED, ${failures.length} failed`);
    expect(failures).toEqual([]);
  }, 1_800_000);
});

function extension(name: string): string {
  return name.slice(name.lastIndexOf('.') + 1).toLowerCase();
}
