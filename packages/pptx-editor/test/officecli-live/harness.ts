import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

import type { Presentation } from '@maxgent/ooxml/pptx';

// @ts-ignore — wasm-pack generated JavaScript is local build output.
import * as pptxWasm from '../../../pptx/src/wasm/pptx_parser.js';

import { createElementRef } from '../../src/adapters/pptx-json-adapter';
import type { ElementRef } from '../../src/domain/mutation';
import { OFFICECLI_VERSION } from '../../src/transport/officecli/constants';
import type { OfficeCliBatch } from '../../src/transport/officecli/types';

/** Structured node returned by `officecli get --json`. */
export interface LiveNode {
  readonly path: string;
  readonly type: string;
  readonly text?: string;
  readonly format: Readonly<Record<string, unknown>>;
  readonly children: readonly LiveNode[];
}

interface OfficeCliJsonResult {
  readonly success: boolean;
  readonly data?: unknown;
  readonly message?: string;
  readonly error?: { readonly error: string; readonly code: string };
}

const EMU_PER_UNIT: Readonly<Record<string, number>> = Object.freeze({
  emu: 1,
  cm: 360_000,
  mm: 36_000,
  in: 914_400,
  pt: 12_700,
});

/**
 * Fails loudly (never skips) when the live environment is unusable: a green
 * live run must always mean the OfficeCLI contract actually executed.
 */
export function assertLiveOfficeCli(): void {
  if (process.env.OFFICECLI_LIVE !== '1') {
    throw new Error(
      'Live OfficeCLI tests require OFFICECLI_LIVE=1 '
      + '(run: OFFICECLI_LIVE=1 pnpm test:officecli-live)',
    );
  }
  const probe = spawnSync('officecli', ['--version'], { encoding: 'utf8' });
  if (probe.error || probe.status !== 0) {
    throw new Error(
      `Live OfficeCLI tests require the officecli binary on PATH: ${probe.error?.message ?? probe.stderr}`,
    );
  }
  const version = probe.stdout.trim();
  if (version !== OFFICECLI_VERSION) {
    throw new Error(
      `Installed officecli ${version} does not match the translator contract `
      + `version ${OFFICECLI_VERSION}; align the binary or the constant before trusting these tests`,
    );
  }
}

export function createLiveWorkspace(prefix: string): string {
  return mkdtempSync(join(tmpdir(), `pptx-editor-${prefix}-`));
}

/** Closes any residents and removes the workspace; safe to call in afterAll. */
export function destroyLiveWorkspace(dir: string, pptxPaths: readonly string[] = []): void {
  for (const pptxPath of pptxPaths) {
    // Ignore failures: the resident may already have exited.
    spawnSync('officecli', ['close', pptxPath, '--json'], { encoding: 'utf8' });
  }
  rmSync(dir, { recursive: true, force: true });
}

function runOfficeCli(args: readonly string[]): OfficeCliJsonResult {
  const result = spawnSync('officecli', [...args, '--json'], { encoding: 'utf8' });
  if (result.error) {
    throw new Error(`officecli ${args[0]} failed to spawn: ${result.error.message}`);
  }
  let parsed: OfficeCliJsonResult;
  try {
    parsed = JSON.parse(result.stdout) as OfficeCliJsonResult;
  } catch {
    throw new Error(
      `officecli ${args.join(' ')} produced non-JSON output (exit ${result.status}):\n`
      + `${result.stdout}\n${result.stderr}`,
    );
  }
  return parsed;
}

function runOfficeCliOk(args: readonly string[]): OfficeCliJsonResult {
  const parsed = runOfficeCli(args);
  if (!parsed.success) {
    throw new Error(`officecli ${args.join(' ')} failed: ${JSON.stringify(parsed.error)}`);
  }
  return parsed;
}

export function createDeck(pptxPath: string): void {
  runOfficeCliOk(['create', pptxPath]);
}

export function addSlide(pptxPath: string, props: Readonly<Record<string, string>> = {}): void {
  runOfficeCliOk(['add', pptxPath, '/', '--type', 'slide', ...propArgs(props)]);
}

/** Adds a shape and returns the canonical stable path reported by OfficeCLI. */
export function addShape(
  pptxPath: string,
  slidePath: string,
  props: Readonly<Record<string, string>>,
): string {
  return addSlideElement(pptxPath, slidePath, 'shape', props);
}

export function addSlideElement(
  pptxPath: string,
  slidePath: string,
  type: string,
  props: Readonly<Record<string, string>>,
): string {
  const result = runOfficeCliOk(['add', pptxPath, slidePath, '--type', type, ...propArgs(props)]);
  const message = typeof result.data === 'string' ? result.data : result.message ?? '';
  const match = message.match(/at (\/\S+)/);
  if (!match) throw new Error(`Cannot extract the canonical ${type} path from: ${message}`);
  return match[1];
}

function propArgs(props: Readonly<Record<string, string>>): string[] {
  return Object.entries(props).flatMap(([key, value]) => ['--prop', `${key}=${value}`]);
}

/**
 * OfficeCLI keeps a resident process holding edits in memory; flush before
 * any non-officecli reader (such as the WASM parser) touches the file.
 */
export function flushDeck(pptxPath: string): void {
  runOfficeCliOk(['close', pptxPath]);
}

let wasmInitialized = false;

/** Parses the on-disk deck with the repository WASM parser (flush first). */
export function parseDeck(pptxPath: string): Presentation {
  if (!wasmInitialized) {
    const wasmBinaryPath = fileURLToPath(
      new URL('../../../pptx/src/wasm/pptx_parser_bg.wasm', import.meta.url),
    );
    let wasmBytes: Uint8Array<ArrayBuffer>;
    try {
      wasmBytes = new Uint8Array(readFileSync(wasmBinaryPath));
    } catch (cause) {
      throw new Error(
        `Live tests need the pptx WASM parser; build it first with pnpm build:wasm (${String(cause)})`,
      );
    }
    (pptxWasm as {
      initSync: (init: { module: WebAssembly.Module }) => unknown;
    }).initSync({ module: new WebAssembly.Module(wasmBytes) });
    wasmInitialized = true;
  }
  const json = (pptxWasm as {
    parse_pptx: (bytes: Uint8Array) => Uint8Array;
  }).parse_pptx(new Uint8Array(readFileSync(pptxPath)));
  return JSON.parse(new TextDecoder().decode(json)) as Presentation;
}

let batchCounter = 0;

/**
 * Executes the native command array of a translated batch exactly as the
 * production sender must: the product envelope is stripped and only
 * `batch.commands` reaches `officecli batch --input`. Asserts per-command
 * success rather than trusting the process exit code, then flushes.
 */
export function runBatch(pptxPath: string, batch: OfficeCliBatch): void {
  batchCounter += 1;
  const inputPath = join(pptxPath, `../batch-${batchCounter}.json`);
  writeFileSync(inputPath, JSON.stringify(batch.commands));
  const parsed = runOfficeCliOk(['batch', pptxPath, '--input', inputPath]);
  const data = parsed.data as {
    results: readonly { index: number; success: boolean; output?: string; error?: string }[];
    summary: { total: number; succeeded: number; failed: number; skipped: number };
  };
  const failed = data.results.filter((result) => !result.success);
  if (failed.length > 0 || data.summary.failed > 0 || data.summary.skipped > 0) {
    throw new Error(
      `officecli batch ${batch.commandId} did not apply every command: ${JSON.stringify(data)}`,
    );
  }
  if (data.summary.succeeded !== batch.commands.length) {
    throw new Error(
      `officecli batch ${batch.commandId} applied ${data.summary.succeeded} of `
      + `${batch.commands.length} commands: ${JSON.stringify(data.summary)}`,
    );
  }
  flushDeck(pptxPath);
}

export function getNode(pptxPath: string, path: string, depth = 1): LiveNode {
  const node = tryGetNode(pptxPath, path, depth);
  if (!node) throw new Error(`officecli get found no node at ${path}`);
  return node;
}

/** Returns undefined when the path resolves to nothing (e.g. after remove). */
export function tryGetNode(pptxPath: string, path: string, depth = 1): LiveNode | undefined {
  const parsed = runOfficeCli(['get', pptxPath, path, '--depth', String(depth)]);
  if (!parsed.success) {
    if (parsed.error?.code === 'not_found') return undefined;
    throw new Error(`officecli get ${path} failed: ${JSON.stringify(parsed.error)}`);
  }
  const data = parsed.data as { matches: number; results: readonly LiveNode[] };
  if (data.matches === 0) return undefined;
  if (data.matches !== 1) {
    throw new Error(`officecli get ${path} matched ${data.matches} nodes; expected exactly one`);
  }
  return data.results[0];
}

/**
 * OfficeCLI readback normalizes lengths to "nice" units (914400emu reads
 * back as "2.54cm"), so assertions must compare in EMUs.
 */
export function lengthToEmu(value: unknown): number {
  if (typeof value === 'number') return value;
  if (typeof value !== 'string') {
    throw new TypeError(`Cannot interpret length readback: ${JSON.stringify(value)}`);
  }
  const match = value.match(/^(-?\d+(?:\.\d+)?)(emu|cm|mm|in|pt)$/);
  if (!match) throw new TypeError(`Cannot interpret length readback: ${value}`);
  return Math.round(Number(match[1]) * EMU_PER_UNIT[match[2]]);
}

/** Absent rotation/flip keys in readback mean "not rotated / not flipped". */
export function normalizeShapeFormat(format: Readonly<Record<string, unknown>>): {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
} {
  return {
    x: lengthToEmu(format.x),
    y: lengthToEmu(format.y),
    width: lengthToEmu(format.width),
    height: lengthToEmu(format.height),
    rotation: Number(format.rotation ?? 0),
    flipH: Boolean(format.flipH),
    flipV: Boolean(format.flipV),
  };
}

/** Extracts the numeric OOXML id from a canonical `shape[@id=N]` path. */
export function elementIdOfPath(path: string): string {
  const match = path.match(/@id=(\d+)/);
  if (!match) throw new Error(`Path ${path} carries no stable @id`);
  return match[1];
}

/** Builds the editable reference for an element of the first slide. */
export function refForElementId(presentation: Presentation, elementId: string): ElementRef {
  const slide = presentation.slides[0];
  const index = slide.elements.findIndex(
    (element) => (element as { id?: string }).id === elementId,
  );
  if (index < 0) throw new Error(`No element with id ${elementId} on the first slide`);
  return createElementRef(slide, slide.elements[index], index);
}

/** Reads a zip-internal part from a flushed .pptx (e.g. `ppt/slides/slide1.xml`). */
export function readPptxPart(pptxPath: string, partName: string): string {
  const result = spawnSync('unzip', ['-p', pptxPath, partName], { encoding: 'utf8' });
  if (result.error) {
    throw new Error(`unzip -p failed to spawn: ${result.error.message}`);
  }
  if (result.status !== 0) {
    throw new Error(
      `unzip -p ${pptxPath} ${partName} failed (exit ${result.status}): ${result.stderr}`,
    );
  }
  return result.stdout;
}
