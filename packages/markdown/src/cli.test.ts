import { describe, it, expect } from 'vitest';
import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
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
