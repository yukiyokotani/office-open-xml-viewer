import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { lstat, mkdir, mkdtemp, readFile, realpath, rm, symlink, unlink, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import test from 'node:test';
import {
  loadManifest, runHarness, streamDocument,
} from './legacy-doc-model-compare.mjs';

async function fixture() {
  const directory = await mkdtemp(join(tmpdir(), 'legacy-doc-model-compare-'));
  const input = join(directory, 'input.doc');
  const manifest = join(directory, 'manifest.json');
  await writeFile(input, 'doc');
  await writeFile(manifest, JSON.stringify({
    version: 1,
    inputs: [{ id: 'one', path: input }],
    limits: { maxSourceBytes: 32, chunkCreditBytes: 16, maxStreamBytes: 32, maxChunks: 4, modelBudget: 32 },
  }));
  return { directory, input, manifest };
}

test('manifest validation resolves bounded regular inputs and rejects aliases', async t => {
  const value = await fixture();
  t.after(() => rm(value.directory, { recursive: true, force: true }));
  const manifest = await loadManifest(value.manifest);
  assert.equal(manifest.inputs.length, 1);
  assert.equal(manifest.limits.maxChunks, 4);

  const alias = join(value.directory, 'alias.doc');
  await symlink(value.input, alias);
  await writeFile(value.manifest, JSON.stringify({ version: 1, inputs: [{ id: 'alias', path: alias }] }));
  await assert.rejects(loadManifest(value.manifest), /non-symlink/);
});

class FakeDocument {
  constructor(chunks) { this.chunks = chunks.map(value => new TextEncoder().encode(value)); this.index = 0; }
  open_document_cursor(operation, generation) { assert.deepEqual([operation, generation], [1, 1]); }
  pull_document_chunk(sequence, operation, generation, credit) {
    assert.deepEqual([sequence, operation, generation], [this.index, 1, 1]);
    if (this.chunks[this.index].byteLength > credit) throw 'OOXML_INSUFFICIENT_CREDIT:{}';
    return this.chunks[this.index];
  }
  document_chunk_done() { return this.index === this.chunks.length - 1; }
  acknowledge_document_chunk(sequence) { assert.equal(sequence, this.index); this.index += 1; }
  assert_healthy() { this.healthy = true; }
}

test('stream digest covers every length-framed chunk and enforces aggregate limits', async () => {
  const limits = { chunkCreditBytes: 8, maxStreamBytes: 8, maxChunks: 3 };
  const result = await streamDocument(new FakeDocument(['ab', 'c']), limits);
  const digest = createHash('sha256');
  for (const value of ['ab', 'c']) {
    const bytes = new TextEncoder().encode(value);
    const length = Buffer.alloc(8); length.writeBigUInt64LE(BigInt(bytes.length));
    digest.update(length); digest.update(bytes);
  }
  assert.deepEqual(result, { digest: digest.digest('hex'), chunks: 2, bytes: 3 });
  await assert.rejects(streamDocument(new FakeDocument(['ab', 'c']), { ...limits, maxChunks: 1 }), /chunk limit/);
  await assert.rejects(streamDocument(new FakeDocument(['ab', 'c']), { ...limits, maxStreamBytes: 2 }), /byte limit/);
  await assert.rejects(streamDocument(new FakeDocument(['oversized']), { ...limits, chunkCreditBytes: 2 }), /chunk credit/);
});

async function harnessFixture() {
  const value = await fixture();
  const candidate = join(value.directory, 'candidate');
  await mkdir(join(candidate, '.git'), { recursive: true });
  await mkdir(join(candidate, '.cargo'), { recursive: true });
  await mkdir(join(candidate, 'packages/legacy-converter/parser/src'), { recursive: true });
  const specs = join(value.directory, 'spec-source');
  await mkdir(specs);
  await symlink(specs, join(candidate, 'spec'), 'dir');
  await writeFile(join(candidate, 'Cargo.toml'), '[workspace]\n');
  await writeFile(join(candidate, 'Cargo.lock'), 'version = 4\n');
  await writeFile(join(candidate, '.cargo/config.toml'), '[net]\noffline = true\n');
  await writeFile(join(candidate, 'packages/legacy-converter/parser/Cargo.toml'), '[package]\nname = "fixture"\n');
  await writeFile(join(candidate, 'packages/legacy-converter/parser/src/lib.rs'), 'pub fn candidate() {}\n');
  return { ...value, candidate, output: join(value.directory, 'result') };
}

function mockDependencies(value, {
  mutateCandidate = false, cleanupFailure = false, mutateBaselineSpec = false,
} = {}) {
  const baseline = 'a'.repeat(40);
  const calls = { builds: [], captures: [], removed: false, removeArgs: undefined, sourceListings: 0, specLinked: false };
  const resolveTool = async name => `/tools/${name}`;
  const captureCommand = async (_command, args, options = {}) => {
    if (args[0] === 'rev-parse' && args[1] === '--show-toplevel') return value.candidate;
    if (args[0] === 'rev-parse' && args[1] === '--verify') return baseline;
    if (args[0] === 'rev-parse' && args[1] === 'HEAD') return 'b'.repeat(40);
    if (args[0] === 'ls-files') {
      calls.sourceListings += 1;
      return [
        '.cargo/config.toml', 'Cargo.toml', 'Cargo.lock',
        'packages/legacy-converter/parser/Cargo.toml',
        'packages/legacy-converter/parser/src/lib.rs', 'README.md', '',
      ].join('\0');
    }
    if (args[0] === 'status') return '';
    if (args[0] === '--version') return `${_command} 1`;
    throw new Error(`unexpected capture command: ${args.join(' ')}, cwd=${options.cwd}`);
  };
  const loggedCommand = async (command, args, options) => {
    await writeFile(options.logPath, `${command} ${args.join(' ')}\n`, { flag: 'wx' });
    if (args[0] === 'worktree' && args[1] === 'add') {
      await mkdir(join(args[3], 'packages/legacy-converter/parser'), { recursive: true });
    } else if (args[0] === 'worktree' && args[1] === 'remove') {
      calls.removeArgs = args;
      if (cleanupFailure) throw new Error('simulated cleanup failure');
      calls.removed = true;
      await rm(args[2], { recursive: true, force: true });
    } else if (args[0] === 'build') {
      const out = args[args.indexOf('--out-dir') + 1];
      if (options.cwd.includes('baselineWorktree')) {
        calls.specLinked = await realpath(join(options.cwd, 'spec')) === await realpath(join(value.candidate, 'spec'));
        if (mutateBaselineSpec) {
          await unlink(join(options.cwd, 'spec'));
          await mkdir(join(options.cwd, 'spec'));
        }
      }
      await mkdir(out);
      await writeFile(join(out, 'legacy_office_converter.js'), 'export default async()=>{};\n');
      await writeFile(join(out, 'legacy_office_converter_bg.wasm'), command);
      calls.builds.push({ out, target: options.env.CARGO_TARGET_DIR, args });
    } else throw new Error(`unexpected logged command: ${args.join(' ')}`);
  };
  const captureBuild = async options => {
    calls.captures.push(options);
    await writeFile(options.resultPath, JSON.stringify({
      version: 1,
      cases: [{ id: 'one', sourceSha256: createHash('sha256').update('doc').digest('hex'), status: 'admitted', digest: 'model', chunks: 2, bytes: 8 }],
    }));
    if (mutateCandidate && options.resultPath.includes('candidate-')) {
      await writeFile(join(value.candidate, 'packages/legacy-converter/parser/src/lib.rs'), 'changed\n');
    }
  };
  return { baseline, calls, dependencies: { resolveTool, captureCommand, loggedCommand, captureBuild } };
}

test('orchestrator isolates mocked build and capture directories and records provenance', async t => {
  const value = await harnessFixture();
  t.after(() => rm(value.directory, { recursive: true, force: true }));
  const mock = mockDependencies(value);
  const report = await runHarness({
    baseline: mock.baseline, candidate: value.candidate, manifest: value.manifest, output: value.output,
  }, mock.dependencies);

  assert.equal(report.equal, true);
  assert.equal(report.candidate.sourceManifest.files, 5);
  assert.equal(mock.calls.builds.length, 2);
  assert.notEqual(mock.calls.builds[0].out, mock.calls.builds[1].out);
  assert.notEqual(mock.calls.builds[0].target, mock.calls.builds[1].target);
  const actualOutput = await realpath(value.output);
  assert.ok(mock.calls.builds.every(call => call.out.startsWith(`${actualOutput}/`) && call.target.startsWith(`${actualOutput}/`)));
  assert.ok(mock.calls.builds.every(call => call.args.includes('--locked') && call.args.includes('--offline')));
  assert.ok(mock.calls.builds.every(call => call.args.includes('--mode') && call.args.includes('no-install')));
  assert.equal(mock.calls.captures.length, 2);
  assert.equal(mock.calls.removed, true);
  assert.equal(mock.calls.removeArgs.includes('--force'), false);
  assert.equal(mock.calls.specLinked, true);
  assert.equal(mock.calls.sourceListings, 2);
  await readFile(join(value.output, 'report.json'));
  await assert.rejects(readFile(join(value.output, 'baselineWorktree')));
});

test('source mutation fails comparison after cleaning only the owned worktree', async t => {
  const value = await harnessFixture();
  t.after(() => rm(value.directory, { recursive: true, force: true }));
  const mock = mockDependencies(value, { mutateCandidate: true });
  await assert.rejects(runHarness({
    baseline: mock.baseline, candidate: value.candidate, manifest: value.manifest, output: value.output,
  }, mock.dependencies), /build sources changed/);
  assert.equal(mock.calls.removed, true);
  await readFile(join(value.output, 'baseline-build.log'));
  await readFile(join(value.output, 'candidate-wasm/legacy_office_converter_bg.wasm'));
});

test('baseline cleanup failure is surfaced without deleting comparison artifacts', async t => {
  const value = await harnessFixture();
  t.after(() => rm(value.directory, { recursive: true, force: true }));
  const mock = mockDependencies(value, { cleanupFailure: true });
  await assert.rejects(
    runHarness({
      baseline: mock.baseline, candidate: value.candidate, manifest: value.manifest, output: value.output,
    }, mock.dependencies),
    /baseline cleanup was incomplete/,
  );
  assert.equal(mock.calls.removeArgs.includes('--force'), false);
  await readFile(join(value.output, 'report.json'));
  await readFile(join(value.output, 'baseline-worktree-cleanup.log'));
});

test('changed owned spec link preserves the baseline worktree without attempting removal', async t => {
  const value = await harnessFixture();
  t.after(() => rm(value.directory, { recursive: true, force: true }));
  const mock = mockDependencies(value, { mutateBaselineSpec: true });
  await assert.rejects(
    runHarness({
      baseline: mock.baseline, candidate: value.candidate, manifest: value.manifest, output: value.output,
    }, mock.dependencies),
    /baseline cleanup was incomplete/,
  );
  assert.equal(mock.calls.removeArgs, undefined);
  assert.equal((await lstat(join(value.output, 'baselineWorktree/spec'))).isDirectory(), true);
});
