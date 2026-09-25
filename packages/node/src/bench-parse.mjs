/**
 * WASM-boundary parse benchmark.
 *
 * Measures the end-to-end cost of parsing one OOXML file through the WASM
 * boundary as the main thread actually pays it: the WASM call itself + whatever
 * marshalling the parser return needs + `JSON.parse` to materialize the model.
 *
 * It deliberately spans the whole "bytes in → model object out" path so the
 * before/after numbers for the boundary-protocol change (String return +
 * JSON.parse  vs.  Vec<u8> return + TextDecoder.decode + JSON.parse) are a fair,
 * apples-to-apples comparison. The script auto-detects the parser's return type
 * (string vs. Uint8Array), so the SAME file measures both protocol variants
 * without edits — only the WASM behind it changes.
 *
 * Usage:
 *   node packages/node/src/bench-parse.mjs <file> [iterations]
 *
 * Example:
 *   node packages/node/src/bench-parse.mjs packages/docx/public/private/docx/sample-10.docx
 *
 * Requires the WASM artifacts to be freshly built (`pnpm build:wasm`) so it
 * measures the current parser source, not a stale committed pkg.
 */
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve, extname, basename } from 'node:path';

const HERE = dirname(fileURLToPath(import.meta.url));
const WARMUP = 1;
const DEFAULT_ITERS = 5;

/** Load a wasm-pack `--target web` module and synchronously initialize it from
 *  disk bytes (mirrors packages/node/src/wasm-loader.ts). No OffscreenCanvas
 *  shim is needed — parsing never touches a canvas. */
function loadParser(relJs, relWasm) {
  const wasmPath = resolve(HERE, relWasm);
  const bytes = readFileSync(wasmPath);
  const module = new WebAssembly.Module(bytes);
  return import(resolve(HERE, relJs)).then((mod) => {
    mod.initSync({ module });
    return mod;
  });
}

/** Decode whatever the parser returned into the model object, timing the
 *  marshalling + parse the same way the real main-thread receiver does.
 *  - String return  → JSON.parse(str)                     (old protocol)
 *  - Uint8Array     → JSON.parse(TextDecoder.decode(u8))  (new protocol) */
const decoder = new TextDecoder();
function toModel(ret) {
  if (typeof ret === 'string') return JSON.parse(ret);
  if (ret instanceof Uint8Array) return JSON.parse(decoder.decode(ret));
  throw new Error(`unexpected parser return type: ${Object.prototype.toString.call(ret)}`);
}

/** Pick the parser call for a file extension. Returns { module, call, label }.
 *  `call(mod, bytes)` runs the WASM parse and returns its raw result (string or
 *  Uint8Array). For xlsx we additionally parse sheet 0 so the measured work is
 *  representative of what the viewer does on open (index + first sheet). */
async function selectParser(ext) {
  switch (ext) {
    case '.docx': {
      const mod = await loadParser(
        '../../docx/src/wasm/docx_parser.js',
        '../../docx/src/wasm/docx_parser_bg.wasm',
      );
      return {
        label: 'parse_docx',
        run: (bytes) => mod.parse_docx(bytes, undefined),
      };
    }
    case '.pptx': {
      const mod = await loadParser(
        '../../pptx/src/wasm/pptx_parser.js',
        '../../pptx/src/wasm/pptx_parser_bg.wasm',
      );
      return {
        label: 'parse_pptx',
        run: (bytes) => mod.parse_pptx(bytes, undefined),
      };
    }
    case '.xlsx': {
      const mod = await loadParser(
        '../../xlsx/src/wasm/xlsx_parser.js',
        '../../xlsx/src/wasm/xlsx_parser_bg.wasm',
      );
      return {
        label: 'parse_xlsx',
        run: (bytes) => mod.parse_xlsx(bytes, undefined),
      };
    }
    default:
      throw new Error(`unsupported extension: ${ext} (expected .docx/.pptx/.xlsx)`);
  }
}

function stats(samples) {
  const sorted = [...samples].sort((a, b) => a - b);
  const n = sorted.length;
  const min = sorted[0];
  const mean = sorted.reduce((s, x) => s + x, 0) / n;
  const median =
    n % 2 === 1 ? sorted[(n - 1) / 2] : (sorted[n / 2 - 1] + sorted[n / 2]) / 2;
  return { min, median, mean };
}

async function main() {
  const [, , file, itersArg] = process.argv;
  if (!file) {
    console.error('usage: node packages/node/src/bench-parse.mjs <file> [iterations]');
    process.exit(1);
  }
  const iters = itersArg ? Number(itersArg) : DEFAULT_ITERS;
  const ext = extname(file).toLowerCase();
  const bytes = readFileSync(resolve(process.cwd(), file));
  const { label, run } = await selectParser(ext);

  // Warmup (JIT + WASM linear-memory growth) — not counted.
  let returnKind = 'unknown';
  for (let i = 0; i < WARMUP; i++) {
    const ret = run(bytes);
    returnKind = typeof ret === 'string' ? 'string' : 'Uint8Array';
    toModel(ret); // include decode so warmup exercises the same path
  }

  const samples = [];
  for (let i = 0; i < iters; i++) {
    const t0 = performance.now();
    const ret = run(bytes);
    const model = toModel(ret);
    const t1 = performance.now();
    // Touch the model so a clever engine can't dead-code-eliminate the parse.
    if (model == null) throw new Error('parse produced null model');
    samples.push(t1 - t0);
  }

  const { min, median, mean } = stats(samples);
  const fmt = (x) => x.toFixed(2).padStart(9);
  console.log(`file      : ${basename(file)} (${(bytes.length / 1_000_000).toFixed(2)} MB)`);
  console.log(`parser    : ${label}  return=${returnKind}  warmup=${WARMUP} iters=${iters}`);
  console.log('           ' + ['min', 'median', 'mean'].map((h) => h.padStart(9)).join(' '));
  console.log(`  ms      :${fmt(min)} ${fmt(median)} ${fmt(mean)}`);
  // Machine-readable line for scripted collection.
  console.log(
    `RESULT\t${basename(file)}\t${returnKind}\t${min.toFixed(2)}\t${median.toFixed(2)}\t${mean.toFixed(2)}`,
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
