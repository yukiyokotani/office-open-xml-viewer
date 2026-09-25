// Opt-in math engine entry point: `@silurus/ooxml/math`.
//
// Importing this module pulls in the MathJax v4 + STIX Two Math engine asset
// (~4 MB). It is intentionally a SEPARATE bundle entry so that docx/pptx
// viewers stay lean by default — equations only render when the consumer
// explicitly wires the engine in via a named import:
//
//   import { DocxViewer } from '@silurus/ooxml/docx';
//   import { math } from '@silurus/ooxml/math';
//   new DocxViewer(canvas, { math });
//
// `math` is a `MathRenderer` — the contract the viewers' `math` option expects.
import {
  loadMathJax,
  mathMLToSvg,
  resolveMathJaxAssetUrl,
} from '../packages/core/src/math/engine.js';
import { registerBuiltinWorkerRenderer } from '../packages/core/src/worker/renderer-module-contract.js';
import type { MathRenderer } from '../packages/core/src/math/mathjax.js';

/**
 * The OMML equation engine (MathJax + STIX Two Math). Pass it to a viewer's
 * `math` option to enable equation rendering. Self-contained: no network, no
 * cross-origin requests.
 */
export const math: MathRenderer = registerBuiltinWorkerRenderer({
  loadMathJax,
  mathMLToSvg,
}, 'math', { engineAssetUrl: resolveMathJaxAssetUrl() });

export type { MathSvg, MathRenderer } from '../packages/core/src/math/mathjax.js';
