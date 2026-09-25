#!/usr/bin/env node
/**
 * Local-only exact comparison for the bounded legacy DOC direct-model stream.
 * Builds an immutable baseline and the supplied candidate without writing build
 * products into either source tree. This does not validate Word display fidelity.
 */
import { createHash } from 'node:crypto';
import { constants as fsConstants } from 'node:fs';
import {
  access, lstat, mkdir, open, readlink, realpath, stat, symlink, unlink, writeFile,
} from 'node:fs/promises';
import { basename, delimiter, dirname, isAbsolute, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { spawn } from 'node:child_process';

const SHA1 = /^[0-9a-f]{40}$/;
const CASE_ID = /^[A-Za-z0-9._-]{1,128}$/;
const MAX_MANIFEST_BYTES = 1024 * 1024;
const MAX_CASES = 10_000;
const HARD_MAX_SOURCE_BYTES = 256 * 1024 * 1024;
const HARD_MAX_CHUNK_BYTES = 64 * 1024 * 1024;
const HARD_MAX_STREAM_BYTES = 512 * 1024 * 1024;
const HARD_MAX_CHUNKS = 1_000_000;
const MAX_COMMAND_OUTPUT_BYTES = 8 * 1024 * 1024;
const MAX_REPORT_BYTES = 16 * 1024 * 1024;
const MAX_ERROR_BYTES = 64 * 1024;
const MAX_WASM_BYTES = 256 * 1024 * 1024;
const MAX_GLUE_BYTES = 16 * 1024 * 1024;
const MAX_SOURCE_MANIFEST_FILES = 20_000;
const MAX_SOURCE_MANIFEST_BYTES = 256 * 1024 * 1024;
const MAX_SOURCE_MANIFEST_FILE_BYTES = 64 * 1024 * 1024;

const DEFAULT_LIMITS = Object.freeze({
  maxSourceBytes: HARD_MAX_SOURCE_BYTES,
  chunkCreditBytes: HARD_MAX_CHUNK_BYTES,
  maxStreamBytes: HARD_MAX_STREAM_BYTES,
  maxChunks: HARD_MAX_CHUNKS,
  modelBudget: HARD_MAX_SOURCE_BYTES,
});

class HarnessLimitError extends Error {}

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

function errorText(error) {
  return String(error instanceof Error ? error.message : error);
}

function boundedErrorText(error) {
  const text = errorText(error);
  if (Buffer.byteLength(text) > MAX_ERROR_BYTES) throw new HarnessLimitError('converter error payload limit exceeded');
  return text;
}

function checkedInteger(value, fallback, maximum, name) {
  const selected = value ?? fallback;
  if (!Number.isSafeInteger(selected) || selected <= 0 || selected > maximum) {
    throw new Error(`${name} must be a positive integer no greater than ${maximum}`);
  }
  return selected;
}

async function readBoundedFile(path, maximumBytes) {
  const flags = fsConstants.O_RDONLY | (fsConstants.O_NOFOLLOW ?? 0);
  const file = await open(path, flags);
  try {
    const info = await file.stat();
    if (!info.isFile() || info.size > maximumBytes) throw new Error(`invalid or oversized file: ${path}`);
    const bytes = Buffer.allocUnsafe(info.size + 1);
    let offset = 0;
    while (offset < bytes.length) {
      const result = await file.read(bytes, offset, bytes.length - offset, offset);
      if (result.bytesRead === 0) break;
      offset += result.bytesRead;
    }
    if (offset > maximumBytes) throw new Error(`file grew beyond its byte limit: ${path}`);
    const finalInfo = await file.stat();
    if (offset !== info.size || finalInfo.size !== info.size) {
      throw new Error(`file changed while reading: ${path}`);
    }
    return bytes.subarray(0, offset);
  } finally {
    await file.close();
  }
}

async function readJson(path, maximumBytes = MAX_MANIFEST_BYTES) {
  return JSON.parse((await readBoundedFile(path, maximumBytes)).toString('utf8'));
}

export async function loadManifest(path) {
  const requestedManifest = resolve(path);
  const manifestInfo = await lstat(requestedManifest);
  if (!manifestInfo.isFile() || manifestInfo.isSymbolicLink()) {
    throw new Error('manifest must be a regular non-symlink file');
  }
  const manifestPath = await realpath(requestedManifest);
  const raw = await readJson(manifestPath);
  if (raw?.version !== 1 || !Array.isArray(raw.inputs) || raw.inputs.length === 0 || raw.inputs.length > MAX_CASES) {
    throw new Error('manifest must be version 1 with 1..10000 inputs');
  }
  const limits = {
    maxSourceBytes: checkedInteger(raw.limits?.maxSourceBytes, DEFAULT_LIMITS.maxSourceBytes, HARD_MAX_SOURCE_BYTES, 'maxSourceBytes'),
    chunkCreditBytes: checkedInteger(raw.limits?.chunkCreditBytes, DEFAULT_LIMITS.chunkCreditBytes, HARD_MAX_CHUNK_BYTES, 'chunkCreditBytes'),
    maxStreamBytes: checkedInteger(raw.limits?.maxStreamBytes, DEFAULT_LIMITS.maxStreamBytes, HARD_MAX_STREAM_BYTES, 'maxStreamBytes'),
    maxChunks: checkedInteger(raw.limits?.maxChunks, DEFAULT_LIMITS.maxChunks, HARD_MAX_CHUNKS, 'maxChunks'),
    modelBudget: checkedInteger(raw.limits?.modelBudget, DEFAULT_LIMITS.modelBudget, HARD_MAX_SOURCE_BYTES, 'modelBudget'),
  };
  const ids = new Set();
  const paths = new Set();
  const inputs = [];
  for (const entry of raw.inputs) {
    if (!CASE_ID.test(entry?.id ?? '') || ids.has(entry.id) || typeof entry.path !== 'string') {
      throw new Error('manifest input IDs must be unique safe names with string paths');
    }
    const candidate = resolve(dirname(manifestPath), entry.path);
    const info = await lstat(candidate);
    if (!info.isFile() || info.isSymbolicLink() || info.size > limits.maxSourceBytes) {
      throw new Error(`input must be a bounded regular non-symlink file: ${candidate}`);
    }
    const path = await realpath(candidate);
    if (paths.has(path)) throw new Error(`duplicate input path: ${path}`);
    ids.add(entry.id);
    paths.add(path);
    inputs.push({ id: entry.id, path });
  }
  return { version: 1, inputs, limits };
}

export async function streamDocument(document, limits) {
  const digest = createHash('sha256');
  let chunks = 0;
  let bytes = 0;
  document.open_document_cursor(1, 1);
  for (let sequence = 0; ; sequence += 1) {
    if (chunks >= limits.maxChunks) throw new HarnessLimitError('model stream chunk limit exceeded');
    let chunk;
    try {
      chunk = document.pull_document_chunk(sequence, 1, 1, limits.chunkCreditBytes);
    } catch (error) {
      if (errorText(error).startsWith('OOXML_INSUFFICIENT_CREDIT:')) {
        throw new HarnessLimitError('model stream chunk credit exceeded');
      }
      throw error;
    }
    if (!(chunk instanceof Uint8Array) || chunk.byteLength > limits.chunkCreditBytes) {
      throw new HarnessLimitError('invalid or oversized model stream chunk');
    }
    bytes += chunk.byteLength;
    if (!Number.isSafeInteger(bytes) || bytes > limits.maxStreamBytes) {
      throw new HarnessLimitError('model stream byte limit exceeded');
    }
    const length = Buffer.allocUnsafe(8);
    length.writeBigUInt64LE(BigInt(chunk.byteLength));
    digest.update(length);
    digest.update(chunk);
    chunks += 1;
    const done = document.document_chunk_done();
    document.acknowledge_document_chunk(sequence, 1, 1);
    if (done) break;
  }
  document.assert_healthy();
  return { digest: digest.digest('hex'), chunks, bytes };
}

async function captureBuild({ gluePath, wasmPath, manifestPath, resultPath }) {
  const manifest = await loadManifest(manifestPath);
  const wasm = await readBoundedFile(wasmPath, MAX_WASM_BYTES);
  const glueInfo = await lstat(gluePath);
  if (!glueInfo.isFile() || glueInfo.isSymbolicLink() || glueInfo.size > MAX_GLUE_BYTES) {
    throw new HarnessLimitError('generated glue is invalid or oversized');
  }
  const glue = await import(`${pathToFileURL(gluePath).href}?capture=${Date.now()}`);
  await glue.default({ module_or_path: wasm });
  if (typeof glue.LegacyDocDocument !== 'function') throw new Error('generated glue lacks LegacyDocDocument');
  const cases = [];
  let retainedBytes = 0;
  for (const input of manifest.inputs) {
    const source = await readBoundedFile(input.path, manifest.limits.maxSourceBytes);
    const entry = { id: input.id, sourceSha256: sha256(source) };
    let document;
    let result;
    try {
      document = new glue.LegacyDocDocument(source, manifest.limits.modelBudget);
    } catch (error) {
      result = { ...entry, status: 'error', stage: 'admission', error: boundedErrorText(error) };
    }
    if (document) {
      try {
        const stream = await streamDocument(document, manifest.limits);
        result = { ...entry, status: 'admitted', ...stream };
      } catch (error) {
        if (error instanceof HarnessLimitError) throw error;
        result = { ...entry, status: 'error', stage: 'stream', error: boundedErrorText(error) };
      } finally {
        try { document.close_document_session(); } finally { document.free(); }
      }
    }
    retainedBytes += Buffer.byteLength(JSON.stringify(result));
    if (retainedBytes > MAX_REPORT_BYTES) throw new HarnessLimitError('capture report limit exceeded');
    cases.push(result);
  }
  const result = JSON.stringify({ version: 1, cases });
  if (Buffer.byteLength(result) > MAX_REPORT_BYTES) throw new HarnessLimitError('capture report limit exceeded');
  await writeFile(resultPath, `${result}\n`, { flag: 'wx' });
}

async function resolveTool(value) {
  const candidates = isAbsolute(value)
    ? [value]
    : (process.env.PATH ?? '').split(delimiter).filter(Boolean).map(directory => resolve(directory, value));
  for (const candidate of candidates) {
    try {
      await access(candidate, fsConstants.X_OK);
      const resolved = await realpath(candidate);
      if ((await stat(resolved)).isFile()) return { invocation: candidate, resolved };
    } catch {}
  }
  throw new Error(`executable is unavailable: ${value}`);
}

async function resolvedTool(value, resolver) {
  const tool = await resolver(value);
  return typeof tool === 'string' ? { invocation: tool, resolved: tool } : tool;
}

async function captureCommand(command, args, options = {}) {
  return await new Promise((resolvePromise, reject) => {
    const child = spawn(command, args, { cwd: options.cwd, env: options.env, stdio: ['ignore', 'pipe', 'pipe'] });
    const chunks = [];
    let size = 0;
    let exceeded = false;
    for (const stream of [child.stdout, child.stderr]) stream.on('data', chunk => {
      size += chunk.length;
      if (size > MAX_COMMAND_OUTPUT_BYTES) {
        exceeded = true;
        child.kill('SIGKILL');
      } else chunks.push(chunk);
    });
    child.on('error', reject);
    child.on('close', code => {
      const output = Buffer.concat(chunks).toString('utf8');
      if (exceeded) reject(new Error(`command output limit exceeded: ${command}`));
      else if (code !== 0) reject(new Error(`command failed (${code}): ${command} ${args.join(' ')}\n${output}`));
      else resolvePromise(output.trim());
    });
  });
}

async function loggedCommand(command, args, { cwd, env, logPath }) {
  const log = await open(logPath, 'wx');
  try {
    await new Promise((resolvePromise, reject) => {
      const child = spawn(command, args, { cwd, env, stdio: ['ignore', log.fd, log.fd] });
      child.on('error', reject);
      child.on('close', code => code === 0
        ? resolvePromise()
        : reject(new Error(`command failed (${code}): ${command} ${args.join(' ')}`)));
    });
  } finally {
    await log.close();
  }
}

function nestedWithin(path, parent) {
  return path === parent || path.startsWith(`${parent}/`);
}

function isBuildSource(relative) {
  const name = basename(relative);
  return relative.endsWith('.rs')
    || name === 'Cargo.toml'
    || name === 'Cargo.lock'
    || relative.startsWith('.cargo/')
    || name === 'rust-toolchain'
    || name === 'rust-toolchain.toml';
}

async function sourceManifest(candidate, git, run = captureCommand) {
  const output = await run(git, [
    'ls-files', '-z', '--cached', '--others', '--exclude-standard',
  ], { cwd: candidate });
  if (output.includes('\ufffd')) throw new Error('git returned a non-UTF-8 build-source path');
  const paths = output.split('\0').filter(isBuildSource).sort();
  if (paths.length === 0 || paths.length > MAX_SOURCE_MANIFEST_FILES) {
    throw new Error('candidate build-source file count is invalid or excessive');
  }
  const digest = createHash('sha256');
  let bytes = 0;
  for (const relative of paths) {
    if (isAbsolute(relative) || relative.split('/').some(part => part === '..')) {
      throw new Error('git returned an unsafe build-source path');
    }
    const path = join(candidate, relative);
    const info = await lstat(path);
    if (!info.isFile() || info.isSymbolicLink() || info.size > MAX_SOURCE_MANIFEST_FILE_BYTES) {
      throw new Error(`unsupported build-source file: ${relative}`);
    }
    bytes += info.size;
    if (!Number.isSafeInteger(bytes) || bytes > MAX_SOURCE_MANIFEST_BYTES) {
      throw new Error('candidate build-source payload budget exceeded');
    }
    const content = await readBoundedFile(path, MAX_SOURCE_MANIFEST_FILE_BYTES);
    if (content.byteLength !== info.size) throw new Error(`build-source file changed while reading: ${relative}`);
    digest.update(`${Buffer.byteLength(relative)}:${relative}:${info.mode & 0o7777}:${content.byteLength}:`);
    digest.update(content);
  }
  return { sha256: digest.digest('hex'), files: paths.length, bytes };
}

function compareCases(baseline, candidate) {
  if (baseline.version !== 1 || candidate.version !== 1 || baseline.cases.length !== candidate.cases.length) {
    throw new Error('capture result schemas or case counts differ');
  }
  return baseline.cases.map((left, index) => {
    const right = candidate.cases[index];
    if (left.id !== right.id || left.sourceSha256 !== right.sourceSha256) {
      throw new Error(`input identity changed between captures at index ${index}`);
    }
    const fields = left.status === 'admitted' && right.status === 'admitted'
      ? ['status', 'digest', 'chunks', 'bytes']
      : ['status', 'stage', 'error'];
    const differences = fields.filter(field => left[field] !== right[field]);
    return {
      id: left.id,
      sourceSha256: left.sourceSha256,
      equal: differences.length === 0,
      differences,
      baseline: Object.fromEntries(fields.map(field => [field, left[field]])),
      candidate: Object.fromEntries(fields.map(field => [field, right[field]])),
    };
  });
}

export async function runHarness(options, dependencies = {}) {
  const deps = {
    resolveTool: dependencies.resolveTool ?? resolveTool,
    captureCommand: dependencies.captureCommand ?? captureCommand,
    loggedCommand: dependencies.loggedCommand ?? loggedCommand,
    captureBuild: dependencies.captureBuild ?? captureBuild,
  };
  if (!SHA1.test(options.baseline ?? '')) throw new Error('baseline must be a full lowercase 40-character commit ID');
  const requestedCandidate = resolve(options.candidate);
  const candidateInfo = await lstat(requestedCandidate);
  if (!candidateInfo.isDirectory() || candidateInfo.isSymbolicLink()) {
    throw new Error('candidate must be a non-symlink directory');
  }
  const candidate = await realpath(requestedCandidate);
  const manifest = await loadManifest(options.manifest);
  const requestedOutput = resolve(options.output);
  const outputParent = await realpath(dirname(requestedOutput));
  const output = join(outputParent, basename(requestedOutput));
  try { await lstat(output); throw new Error('output directory must not already exist'); } catch (error) {
    if (error?.code !== 'ENOENT') throw error;
  }
  if (nestedWithin(output, candidate) || nestedWithin(candidate, output)) {
    throw new Error('output and candidate source directories must not overlap');
  }
  const tools = {
    git: await resolvedTool(options.git ?? 'git', deps.resolveTool),
    wasmPack: await resolvedTool(options.wasmPack ?? 'wasm-pack', deps.resolveTool),
    node: await resolvedTool(options.node ?? process.execPath, deps.resolveTool),
    rustc: await resolvedTool('rustc', deps.resolveTool),
    cargo: await resolvedTool('cargo', deps.resolveTool),
  };
  const top = await deps.captureCommand(tools.git.invocation, ['rev-parse', '--show-toplevel'], { cwd: candidate });
  if (await realpath(top) !== candidate) throw new Error('candidate must be the root of its git worktree');
  const baselineSha = await deps.captureCommand(tools.git.invocation, ['rev-parse', '--verify', `${options.baseline}^{commit}`], { cwd: candidate });
  if (baselineSha !== options.baseline) throw new Error('baseline did not resolve to the requested immutable commit');
  const candidateSha = await deps.captureCommand(tools.git.invocation, ['rev-parse', 'HEAD'], { cwd: candidate });
  if (!SHA1.test(candidateSha)) throw new Error('candidate HEAD is invalid');
  const candidateSource = await sourceManifest(candidate, tools.git.invocation, deps.captureCommand);
  const versions = {
    wasmPack: await deps.captureCommand(tools.wasmPack.invocation, ['--version']),
    rustc: await deps.captureCommand(tools.rustc.invocation, ['--version', '--verbose']),
    cargo: await deps.captureCommand(tools.cargo.invocation, ['--version', '--verbose']),
    node: await deps.captureCommand(tools.node.invocation, ['--version']),
  };

  let candidateSpec;
  try {
    const requestedSpec = join(candidate, 'spec');
    const specInfo = await lstat(requestedSpec);
    if (!specInfo.isSymbolicLink() || !(await stat(requestedSpec)).isDirectory()) {
      throw new Error('candidate spec must be a symlink to a directory when present');
    }
    candidateSpec = await realpath(requestedSpec);
  } catch (error) {
    if (error?.code !== 'ENOENT') throw error;
  }

  await mkdir(output);
  const paths = Object.fromEntries([
    'baselineWorktree', 'baseline-target', 'candidate-target', 'baseline-wasm', 'candidate-wasm',
  ].map(name => [name.replace(/-([a-z])/g, (_, letter) => letter.toUpperCase()), join(output, name)]));
  const manifestPath = join(output, 'manifest.json');
  await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, { flag: 'wx' });
  const baselineResult = join(output, 'baseline-capture.json');
  const candidateResult = join(output, 'candidate-capture.json');
  let worktreeAdded = false;
  let baselineSpecAdded = false;
  let primaryError;
  try {
    await deps.loggedCommand(tools.git.invocation, ['worktree', 'add', '--detach', paths.baselineWorktree, baselineSha], {
      cwd: candidate, logPath: join(output, 'baseline-worktree.log'),
    });
    worktreeAdded = true;
    const clean = await deps.captureCommand(tools.git.invocation, ['status', '--porcelain', '--untracked-files=all'], { cwd: paths.baselineWorktree });
    if (clean !== '') throw new Error('baseline worktree is not clean');
    if (candidateSpec) {
      await symlink(candidateSpec, join(paths.baselineWorktree, 'spec'), 'dir');
      baselineSpecAdded = true;
    }
    const builds = [
      ['baseline', paths.baselineWorktree, paths.baselineTarget, paths.baselineWasm, baselineResult],
      ['candidate', candidate, paths.candidateTarget, paths.candidateWasm, candidateResult],
    ];
    for (const [label, source, target, wasm, result] of builds) {
      const args = ['build', join(source, 'packages/legacy-converter/parser'), '--mode', 'no-install', '--release', '--target', 'web', '--out-dir', wasm, '--out-name', 'legacy_doc_direct', '--features', 'direct-doc', '--locked', '--offline'];
      await deps.loggedCommand(tools.wasmPack.invocation, args, {
        cwd: source,
        env: { ...process.env, CARGO_TARGET_DIR: target, CARGO_NET_OFFLINE: 'true' },
        logPath: join(output, `${label}-build.log`),
      });
      await deps.captureBuild({
        gluePath: join(wasm, 'legacy_doc_direct.js'),
        wasmPath: join(wasm, 'legacy_doc_direct_bg.wasm'),
        manifestPath,
        resultPath: result,
        node: tools.node.invocation,
        logPath: join(output, `${label}-capture.log`),
      });
    }
    const baselineCapture = await readJson(baselineResult, MAX_REPORT_BYTES);
    const candidateCapture = await readJson(candidateResult, MAX_REPORT_BYTES);
    const cases = compareCases(baselineCapture, candidateCapture);
    const buildConfiguration = {
      profile: 'release', target: 'web', mode: 'no-install', features: ['direct-doc'], locked: true, offline: true,
      environment: {
        CARGO_NET_OFFLINE: 'true',
        CARGO_TARGET_DIR: 'separate per build under output',
        RUSTFLAGS: process.env.RUSTFLAGS ?? null,
      },
      tools, versions,
    };
    const report = {
      version: 1,
      comparison: 'length-framed exact streamed-model bytes',
      limits: manifest.limits,
      baseline: {
        requested: options.baseline, resolved: baselineSha,
        wasmSha256: sha256(await readBoundedFile(join(paths.baselineWasm, 'legacy_doc_direct_bg.wasm'), MAX_WASM_BYTES)),
        build: buildConfiguration,
      },
      candidate: {
        path: candidate, resolved: candidateSha, sourceManifest: candidateSource,
        wasmSha256: sha256(await readBoundedFile(join(paths.candidateWasm, 'legacy_doc_direct_bg.wasm'), MAX_WASM_BYTES)),
        build: buildConfiguration,
      },
      equal: cases.every(entry => entry.equal),
      cases,
    };
    const candidateSourceAfter = await sourceManifest(candidate, tools.git.invocation, deps.captureCommand);
    if (JSON.stringify(candidateSourceAfter) !== JSON.stringify(candidateSource)) {
      throw new Error('candidate build sources changed during comparison');
    }
    const serialized = JSON.stringify(report, null, 2);
    if (Buffer.byteLength(serialized) > MAX_REPORT_BYTES) throw new HarnessLimitError('comparison report limit exceeded');
    await writeFile(join(output, 'report.json'), `${serialized}\n`, { flag: 'wx' });
    return report;
  } catch (error) {
    primaryError = error;
  } finally {
    if (worktreeAdded) {
      const cleanupErrors = [];
      if (baselineSpecAdded) {
        try {
          const baselineSpec = join(paths.baselineWorktree, 'spec');
          const info = await lstat(baselineSpec);
          if (!info.isSymbolicLink() || await readlink(baselineSpec) !== candidateSpec) {
            throw new Error('owned baseline spec symlink changed during comparison');
          }
          await unlink(baselineSpec);
        } catch (error) {
          cleanupErrors.push(error);
        }
      }
      try {
        const clean = await deps.captureCommand(tools.git.invocation, ['status', '--porcelain', '--untracked-files=all'], { cwd: paths.baselineWorktree });
        if (clean !== '') throw new Error('baseline worktree changed during comparison');
      } catch (error) {
        cleanupErrors.push(error);
      }
      if (cleanupErrors.length === 0) {
        try {
          await deps.loggedCommand(tools.git.invocation, ['worktree', 'remove', paths.baselineWorktree], {
            cwd: candidate, logPath: join(output, 'baseline-worktree-cleanup.log'),
          });
        } catch (cleanupError) {
          cleanupErrors.push(cleanupError);
        }
      }
      if (cleanupErrors.length > 0) {
        if (primaryError) {
          throw new AggregateError([primaryError, ...cleanupErrors], 'comparison failed and baseline cleanup was incomplete');
        }
        throw new AggregateError(cleanupErrors, 'baseline cleanup was incomplete');
      }
    }
  }
  if (primaryError) throw primaryError;
}

function parseArgs(args) {
  const values = {};
  for (const argument of args) {
    const match = /^--([a-z-]+)=(.+)$/.exec(argument);
    if (!match) throw new Error('arguments must use --name=value');
    const key = match[1].replace(/-([a-z])/g, (_, letter) => letter.toUpperCase());
    if (values[key] !== undefined) throw new Error(`duplicate argument: ${match[1]}`);
    values[key] = match[2];
  }
  return values;
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  for (const name of ['baseline', 'candidate', 'manifest', 'output']) {
    if (!options[name]) throw new Error('usage: legacy-doc-model-compare --baseline=FULL_SHA --candidate=WORKTREE --manifest=JSON --output=NEW_DIR [--git=PATH --wasm-pack=PATH --node=PATH]');
  }
  // Capture runs in a fresh process so each generated wasm-bindgen module is
  // initialized exactly once and baseline/candidate runtimes cannot alias.
  const externalCapture = async ({ gluePath, wasmPath, manifestPath, resultPath, node, logPath }) => {
    await loggedCommand(node, [resolve(process.argv[1]), '--capture-mode=1', `--glue=${gluePath}`, `--wasm=${wasmPath}`, `--manifest=${manifestPath}`, `--result=${resultPath}`], {
      cwd: dirname(resultPath), logPath,
    });
  };
  const report = await runHarness(options, { captureBuild: externalCapture });
  console.log(JSON.stringify({ output: resolve(options.output), equal: report.equal, cases: report.cases.length }));
  if (!report.equal) process.exitCode = 1;
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const options = parseArgs(process.argv.slice(2));
  if (options.captureMode === '1') {
    await captureBuild({ gluePath: options.glue, wasmPath: options.wasm, manifestPath: options.manifest, resultPath: options.result });
  } else {
    await main();
  }
}
