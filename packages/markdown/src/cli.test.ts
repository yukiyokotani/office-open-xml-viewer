import { describe, it, expect } from 'vitest';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * Smoke coverage for the `ooxml-md` CLI: run the real bin against a committed
 * demo sample and assert it prints markdown to stdout. This exercises the whole
 * path — extension detection, `resolveWasm` (the `./wasm-binary` export lookup),
 * WASM init, and conversion — the same way a `npx ooxml-md file.docx` invocation
 * would.
 *
 * Gated on the parser WASM + sample being present (git-ignored build output;
 * CI builds it before `pnpm test`) so a pre-build local run SKIPS rather than
 * failing.
 */

const root = new URL('../../..', import.meta.url);
const bin = fileURLToPath(new URL('../bin/ooxml-md.mjs', import.meta.url));
const sample = fileURLToPath(new URL('packages/docx/public/demo/sample-1.docx', root));
const wasm = fileURLToPath(new URL('packages/docx/src/wasm/docx_parser_bg.wasm', root));

const ready = existsSync(sample) && existsSync(wasm);

describe('ooxml-md CLI', () => {
  it.skipIf(!ready)('prints markdown to stdout for a .docx', () => {
    const out = execFileSync('node', [bin, sample], { encoding: 'utf8' });
    expect(out.length).toBeGreaterThan(0);
    expect(out).toContain('CANOPY');
  });

  const parsersReady = ['docx', 'xlsx', 'pptx'].every((format) =>
    existsSync(fileURLToPath(new URL(`packages/${format}/src/wasm/${format}_parser_bg.wasm`, root))));

  it.skipIf(!parsersReady)('exits 3 with a readable message for non-OOXML input', () => {
    const directory = mkdtempSync(join(tmpdir(), 'ooxml-md-'));
    try {
      for (const extension of ['docx', 'xlsx', 'pptx']) {
        const file = join(directory, `garbage.${extension}`);
        writeFileSync(file, new Uint8Array([1, 2, 3]));
        let code = 0;
        let stderr = '';
        try {
          execFileSync('node', [bin, file], { encoding: 'utf8', stdio: 'pipe' });
        } catch (err) {
          const e = err as { status: number; stderr: string };
          code = e.status;
          stderr = e.stderr;
        }
        expect(code, extension).toBe(3);
        expect(stderr, extension).toContain('is not an Office Open XML document');
      }
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  it('prints usage and exits non-zero with no arguments', () => {
    let code = 0;
    let stdout = '';
    try {
      stdout = execFileSync('node', [bin], { encoding: 'utf8' });
    } catch (err) {
      const e = err as { status: number; stdout: string };
      code = e.status;
      stdout = e.stdout;
    }
    expect(code).not.toBe(0);
    expect(stdout).toContain('ooxml-md');
  });
});
