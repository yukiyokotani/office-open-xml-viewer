// Pre-bundle the MathJax v4 + STIX Two Math converter into a single opaque
// asset (`assets/mathjax-stix2.js`). The renderer loads this asset lazily;
// keeping it pre-bundled (rather than importing @mathjax/src directly) stops the
// consuming app's bundler from re-bundling — and over-including — the MathJax
// source.
//
// The generated bundle is ~4 MB of minified IIFE, so — like the WASM parsers —
// it is NOT committed (see .gitignore) and is instead regenerated from source.
// This script runs automatically on install via core's `prepare` script; run it
// by hand with `pnpm --filter @silurus/ooxml-core build:mathjax`.
import { build } from 'esbuild';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const dir = path.dirname(fileURLToPath(import.meta.url));

await build({
  entryPoints: [path.join(dir, 'stix2-entry.mjs')],
  bundle: true,
  minify: true,
  format: 'iife',
  target: 'es2020',
  // 'linked' preserves any `@license`/`@preserve`-tagged comment esbuild finds
  // while bundling (extracted to a sibling `mathjax-stix2.js.LEGAL.txt`,
  // referenced from a top-of-file comment) instead of the previous 'none',
  // which silently discarded them. @mathjax/src and
  // @mathjax/mathjax-stix2-font (both Apache-2.0) currently carry no
  // per-file license banners in their compiled .mjs output — verified by
  // diffing the bundle across all four `legalComments` modes, which produced
  // byte-identical output and no .LEGAL.txt in every case — so this is a
  // no-op today. It is still the correct setting: if upstream ever adds
  // banners, they survive the build automatically instead of being stripped.
  // The actual Apache-2.0 attribution for MathJax/STIX2 lives in
  // THIRD_PARTY_NOTICES.md (repo root), which is bundled into the npm
  // tarball regardless of what this bundler step finds.
  legalComments: 'linked',
  outfile: path.join(dir, '../assets/mathjax-stix2.js'),
});

console.log('built assets/mathjax-stix2.js');
