// Light, dependency-free helpers for the math feature.
//
// IMPORTANT: this module must stay free of the heavy MathJax engine asset (that
// lives in `./engine`, referenced as a separate asset URL). The docx/pptx renderers
// import ONLY from here (extents math, recolor, the `MathRenderer` contract),
// so the ~4 MB engine asset tree-shakes out of their bundles. The engine is
// injected at runtime via the `math` viewer option — see `MathRenderer`.

export interface MathSvg {
  /** standalone `<svg>…</svg>` markup. */
  svg: string;
  /** extents in em (the SVG viewBox uses 1em = 1000 units). */
  widthEm: number;
  ascentEm: number;
  descentEm: number;
}

/**
 * The math engine contract a viewer needs to render equations. Satisfied by the
 * `math` named export of the separate `@silurus/ooxml/math` entry point, which
 * the consumer opts into:
 *
 * ```ts
 * import { DocxViewer } from '@silurus/ooxml/docx';
 * import { math } from '@silurus/ooxml/math';
 * new DocxViewer(canvas, { math });
 * ```
 *
 * Omit it and the equation engine (MathJax + STIX Two Math, ~4 MB) is not
 * fetched or evaluated. The self-contained worker asset retains only the
 * worker-side loader until the built-in is supplied and a document uses math.
 */
export interface MathRenderer {
  /** Preload the engine. Called once before converting equations. */
  loadMathJax(): Promise<void>;
  /** MathML string → standalone SVG + baseline-relative em extents. */
  mathMLToSvg(mathml: string): Promise<MathSvg>;
}

const UNITS_PER_EM = 1000;

/** Parse the MathJax SVG viewBox into baseline-relative em extents. */
export function svgExtents(svg: string): { widthEm: number; ascentEm: number; descentEm: number } {
  const m = /viewBox="([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+)"/.exec(svg);
  if (!m) return { widthEm: 0, ascentEm: 0, descentEm: 0 };
  const minY = parseFloat(m[2]);
  const w = parseFloat(m[3]);
  const h = parseFloat(m[4]);
  // The output's top <g> applies scale(1,-1): content rises to -minY above the
  // baseline and falls to (minY + h) below it. A viewBox can sit wholly on one
  // side of the baseline, so the opposite extent is zero rather than negative.
  return {
    widthEm: w / UNITS_PER_EM,
    ascentEm: Math.max(0, -minY / UNITS_PER_EM),
    descentEm: Math.max(0, (minY + h) / UNITS_PER_EM),
  };
}

/** Replace MathJax's `currentColor` placeholders with an explicit color (for raster). */
export function recolorSvg(svg: string, color: string): string {
  return svg.replace(/currentColor/g, color);
}
