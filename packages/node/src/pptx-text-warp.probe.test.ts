/**
 * Headless geometry probe for WordArt text warp (ECMA-376 §20.1.9.19,
 * `<a:prstTxWarp>`).
 *
 * Canvas 2D cannot deform glyph outlines, so the pptx renderer approximates a
 * text warp PER GLYPH: it lays the run out flat, then places each glyph against
 * the preset's envelope (a top + bottom boundary curve). This probe renders a
 * synthetic in-memory Presentation (no .pptx, no WASM) carrying one text-bearing
 * shape, and MEASURES the resulting ink to confirm the glyphs actually follow
 * the envelope rather than sitting on a flat baseline.
 *
 * Assertions, all purely geometric (no eyeballing):
 *   - textArchUp: Follow Path (issue #846) — the word keeps its NATURAL width
 *     and its paragraph alignment selects where that compact segment sits on
 *     the path. The centred fixture therefore stays compact around the apex
 *     instead of stretching across the entire arch.
 *   - textPlain (control): the ink stays flat — the topmost inked row is roughly
 *     constant across the width. This guards against the warp path accidentally
 *     bending un-warpable ("identity") text.
 *
 * CI-safe: gated on skia-canvas via the shared test helper (skip locally, fail
 * under OOXML_REQUIRE_SKIA=1).
 */
import { describe, it, expect } from 'vitest';
import type { Presentation, Slide, ShapeElement, TextBody, Paragraph } from '@silurus/ooxml-pptx';
import { importForTests, loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;

const renderMod = await importForTests(() => import('./render.ts'), './render.ts');
const { renderSlideNode } = (renderMod ?? {}) as typeof import('./render.ts');
type NodeCanvasFactory = import('./render.ts').NodeCanvasFactory;
// A canvas factory so renderSlideNode installs the OffscreenCanvas shim. The
// paired-edge strip draw itself no longer needs it (each strip is a clipped
// vector fillText — no auxiliary canvas anywhere), but the factory keeps the
// rest of the renderer's aux-canvas paths (effects, patterns) live, matching
// how the browser runs.
const warpFactory: NodeCanvasFactory | undefined = skia
  ? {
      createCanvas: (w, h) => new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
      loadImage: (b) => loadImage(b as Buffer) as unknown as ReturnType<NodeCanvasFactory['loadImage']>,
    }
  : undefined;

const EMU_PER_PX = 9525; // 96 dpi
const px = (n: number) => Math.round(n * EMU_PER_PX);

function warpShape(preset: string, color = '000000'): ShapeElement {
  const para: Paragraph = {
    alignment: 'ctr',
    marL: 0,
    marR: 0,
    indent: 0,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl: 0,
    bullet: { type: 'none' },
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs: [
      {
        type: 'text',
        text: 'WARP',
        bold: true,
        italic: null,
        underline: false,
        strikethrough: false,
        fontSize: 40,
        color,
        fontFamily: 'Arial',
      } as Paragraph['runs'][number],
    ],
  };
  const textBody: TextBody = {
    verticalAnchor: 'ctr',
    paragraphs: [para],
    defaultFontSize: 40,
    defaultBold: null,
    defaultItalic: null,
    lIns: 0,
    rIns: 0,
    tIns: 0,
    bIns: 0,
    wrap: 'none',
    vert: 'horz',
    autoFit: 'none',
    textWarp: { preset },
  } as TextBody;
  return {
    type: 'shape',
    x: px(40),
    y: px(40),
    width: px(320),
    height: px(160),
    rotation: 0,
    flipH: false,
    flipV: false,
    geometry: 'rect',
    fill: null,
    stroke: null,
    textBody,
    defaultTextColor: null,
    custGeom: null,
    adj: null,
    adj2: null,
    adj3: null,
    adj4: null,
    adj5: null,
    adj6: null,
    adj7: null,
    adj8: null,
    shadow: null,
  };
}

function buildSlide(preset: string, color = '000000'): Presentation {
  const slide: Slide = {
    index: 0,
    slideNumber: 1,
    background: null,
    elements: [warpShape(preset, color)],
  };
  return {
    slideWidth: px(400),
    slideHeight: px(240),
    slides: [slide],
    defaultTextColor: null,
    majorFont: null,
    minorFont: null,
  };
}

async function renderInk(preset: string): Promise<{
  topRow: (col: number) => number | null;
  bottomRow: (col: number) => number | null;
  width: number;
  height: number;
}> {
  const width = 400;
  const height = 240;
  const dpr = 1;
  const canvas = new Canvas(width * dpr, height * dpr);
  await renderSlideNode(canvas as unknown as Parameters<typeof renderSlideNode>[0], buildSlide(preset), 0, {
    width,
    dpr,
    factory: warpFactory,
  });
  const ctx = canvas.getContext('2d');
  const img = ctx.getImageData(0, 0, width * dpr, height * dpr);
  const data = img.data as unknown as Uint8ClampedArray;
  const W = width * dpr;
  const H = height * dpr;
  const inked = (col: number, y: number): boolean => {
    const i = (y * W + col) * 4;
    // Dark ink on the default (transparent/white) background: low luminance
    // AND non-transparent. skia clears to transparent; text is opaque black.
    const a = data[i + 3];
    const lum = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
    return a > 40 && lum < 128;
  };
  // For a given column, the first (topmost) row whose pixel is inked (dark).
  const topRow = (col: number): number | null => {
    for (let y = 0; y < H; y++) if (inked(col, y)) return y;
    return null;
  };
  // …and the last (bottommost) inked row.
  const bottomRow = (col: number): number | null => {
    for (let y = H - 1; y >= 0; y--) if (inked(col, y)) return y;
    return null;
  };
  return { topRow, bottomRow, width: W, height: H };
}

/**
 * Luminance histogram of a TRANSLUCENT warped WordArt render (issue #879).
 *
 * The paired-edge warp draws each glyph as overlapping strips (a 1-device-px
 * clip overlap hides the shared-edge AA seam). Painting an opaque colour over
 * itself in that band is idempotent, but a TRANSLUCENT fill composed twice
 * DARKENS it: over the slide's opaque white background a single composite of a
 * fill at alpha `a` yields luminance L1 = 255 − a·(255 − Lfill), while the
 * double-composed band yields L2 = 255 − a(2−a)·(255 − Lfill) < L1 — a visible
 * darker pinstripe (Inflate ≈28%, CirclePour ≈71% of ink columns in the issue's
 * browser measurement). Antialiased glyph edges only ever composite LESS ink, so
 * they are LIGHTER than L1; the sole source of a pixel darker than L1 is the
 * double composition. The #879 fix paints the strips opaque into a per-glyph
 * device layer and composites it once, so every ink pixel lands at L1.
 *
 * Renders `<colorHex>` (8-digit RRGGBBAA) at dpr 2 and returns, over the solid
 * ink pixels, the fraction darker than the L1/L2 midpoint (the double-composed
 * band fraction) and the darkest ink luminance seen. band ≈ 0 / minLum ≈ L1 when
 * the composite is uniform; band large / minLum ≈ L2 when the overlap darkens.
 */
async function warpTranslucentBand(
  preset: string,
  colorHex: string,
): Promise<{ ink: number; bandFraction: number; minLum: number; l1: number; l2: number }> {
  const width = 400;
  const height = 240;
  const dpr = 2;
  const canvas = new Canvas(width * dpr, height * dpr);
  await renderSlideNode(
    canvas as unknown as Parameters<typeof renderSlideNode>[0],
    buildSlide(preset, colorHex),
    0,
    { width, dpr, factory: warpFactory },
  );
  const ctx = canvas.getContext('2d');
  const W = width * dpr;
  const H = height * dpr;
  const data = ctx.getImageData(0, 0, W, H).data as unknown as Uint8ClampedArray;
  const lumAt = (p: number): number =>
    0.299 * data[p * 4] + 0.587 * data[p * 4 + 1] + 0.114 * data[p * 4 + 2];
  // Fill colour + alpha from the 8-digit hex; single (L1) and double (L2)
  // composite luminance over the white background.
  const fr = parseInt(colorHex.slice(0, 2), 16);
  const fg = parseInt(colorHex.slice(2, 4), 16);
  const fb = parseInt(colorHex.slice(4, 6), 16);
  const a = parseInt(colorHex.slice(6, 8), 16) / 255;
  const lFill = 0.299 * fr + 0.587 * fg + 0.114 * fb;
  const l1 = 255 - a * (255 - lFill); // single composite
  const l2 = 255 - a * (2 - a) * (255 - lFill); // double composite (band)
  const bandCut = (l1 + l2) / 2; // darker than this ⇒ double-composed
  const inkCut = (l1 + 255) / 2; // darker than this ⇒ solid ink (not AA fringe / bg)
  let ink = 0;
  let band = 0;
  let minLum = 255;
  for (let p = 0; p < W * H; p++) {
    const l = lumAt(p);
    if (l >= inkCut) continue; // background or faint AA edge
    ink++;
    if (l < minLum) minLum = l;
    if (l < bandCut) band++;
  }
  return { ink, bandFraction: ink > 0 ? band / ink : 0, minLum, l1, l2 };
}

/** Min top-inked row across a column band [c0, c1). */
function minTopInBand(topRow: (c: number) => number | null, c0: number, c1: number): number | null {
  let best: number | null = null;
  for (let c = c0; c < c1; c++) {
    const t = topRow(c);
    if (t != null && (best == null || t < best)) best = t;
  }
  return best;
}

describe.skipIf(!skia)('node WordArt text-warp geometry (prstTxWarp)', () => {
  it('textArchUp follows the path at natural width with paragraph alignment (Follow Path, #846)', async () => {
    // Shape box: x=40, w=320 (canvas px). The arch baseline is the top half of
    // the box ellipse, running clockwise from stAng=180° (LEFT end, mid-height)
    // over the apex (270°) to the right end. Follow Path lays "WARP" along only
    // its natural arc length, while the centred paragraph positions that short
    // segment around the apex. The pre-#846 renderer stretched the run across
    // the whole 180° arc (ink ≈[40,360]).
    const { topRow, width } = await renderInk('textArchUp');
    const shapeX = 40;
    const shapeW = 320;

    // Locate the inked column range.
    let minC: number | null = null;
    let maxC: number | null = null;
    for (let c = 0; c < width; c++) {
      if (topRow(c) != null) {
        if (minC == null) minC = c;
        maxC = c;
      }
    }
    expect(minC).not.toBeNull();
    expect(maxC).not.toBeNull();

    // Natural width, not full-path stretch: the ink span stays well under the
    // box width (a full-arch distribution spans ~the whole 320px).
    expect(maxC! - minC!).toBeLessThan(shapeW * 0.75);

    // The authored centre alignment is preserved on the path rather than being
    // pinned to stAng. Font rasterisation may shift either edge slightly, so
    // compare the span centre to the shape centre instead of exact columns.
    const inkCentre = (minC! + maxC!) / 2;
    expect(Math.abs(inkCentre - (shapeX + shapeW / 2))).toBeLessThan(shapeW * 0.05);

    // Within its span the middle of the word rises toward the arch apex. This
    // distinguishes the curved placement from a compact but flat text run.
    const span = maxC! - minC!;
    const startTop = minTopInBand(topRow, minC!, minC! + Math.max(1, Math.floor(span * 0.15)));
    const endTop = minTopInBand(topRow, maxC! - Math.max(1, Math.floor(span * 0.15)), maxC! + 1);
    const centreTop = minTopInBand(topRow, minC! + Math.floor(span * 0.4), minC! + Math.ceil(span * 0.6));
    expect(startTop).not.toBeNull();
    expect(centreTop).not.toBeNull();
    expect(endTop).not.toBeNull();
    expect(centreTop!).toBeLessThan(Math.min(startTop!, endTop!));
  });

  it('textPlain leaves the ink flat (control)', async () => {
    const { topRow, width } = await renderInk('textPlain');
    const third = Math.floor(width / 3);
    const leftTop = minTopInBand(topRow, Math.floor(width * 0.18), Math.floor(width * 0.32));
    const centreTop = minTopInBand(topRow, third, 2 * third);
    const rightTop = minTopInBand(topRow, Math.floor(width * 0.68), Math.floor(width * 0.82));
    expect(leftTop).not.toBeNull();
    expect(centreTop).not.toBeNull();
    expect(rightTop).not.toBeNull();
    // Flat: the three bands' top rows agree within a small tolerance (glyph
    // ascenders vary a little, but there is no arch rise).
    const spread = Math.max(leftTop!, centreTop!, rightTop!) - Math.min(leftTop!, centreTop!, rightTop!);
    expect(spread).toBeLessThan(12);
  });

  it('textInflate stretches glyph ink to fill the envelope gap (G1 guard)', async () => {
    // PowerPoint stretches the flat text's ink rectangle to the shape box
    // before warping — verified against PowerPoint's PDF of the warp fixture
    // (ink height 1.48in in a 1.5in-tall shape). The shape here is 160px tall
    // and textInflate's envelope spans the full box height at its centre, so
    // the centre glyphs' ink must span well over half the box. Before the fix
    // glyphs kept their natural ~38px size (vScale collapsed to ≈1).
    const { topRow, bottomRow, width } = await renderInk('textInflate');
    const boxHpx = 160; // shape height in px at this render scale
    let maxSpan = 0;
    for (let c = Math.floor(width * 0.35); c < Math.floor(width * 0.65); c++) {
      const t = topRow(c);
      const b = bottomRow(c);
      if (t != null && b != null) maxSpan = Math.max(maxSpan, b - t + 1);
    }
    expect(maxSpan).toBeGreaterThan(boxHpx * 0.5);
  });

  it('textWave1 shears glyphs along the wave slope, not a rigid rotation', async () => {
    // ECMA-376 §20.1.9.19: the envelope map's local Jacobian is a rotate+shear,
    // not a plain rotate+scale — on the wave's rising/falling flanks the glyph
    // em-box skews into a parallelogram (vertical strokes track the vertical
    // edge-gap while the baseline follows the slope). PowerPoint's own PDF of the
    // warp fixture shows the "Wave One" glyphs leaning italic-forward.
    //
    // Pixel signature of a SHEAR (as opposed to a rigid rotation): within a
    // narrow vertical slab, the ink's horizontal CENTROID shifts between the
    // upper and lower parts of the glyph band — the top slides forward on an
    // upslope and back on a downslope. A rigid rotation keeps the slab's ink
    // vertically aligned (top and bottom centroids agree). We sample several
    // slabs across the word and require a MEANINGFUL top↔bottom shift that also
    // CHANGES SIGN across the wave (proof the skew tracks the local slope rather
    // than being a single global tilt).
    const width = 400;
    const height = 240;
    const canvas = new Canvas(width, height);
    await renderSlideNode(
      canvas as unknown as Parameters<typeof renderSlideNode>[0],
      buildSlide('textWave1'),
      0,
      { width, dpr: 1, factory: warpFactory },
    );
    const ctx = canvas.getContext('2d');
    const data = ctx.getImageData(0, 0, width, height).data as unknown as Uint8ClampedArray;
    const inked = (x: number, y: number): boolean => {
      const i = (y * width + x) * 4;
      const a = data[i + 3];
      const lum = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
      return a > 40 && lum < 128;
    };
    // Ink bbox of the whole warped word.
    let minX = width, maxX = -1, minY = height, maxY = -1;
    for (let y = 0; y < height; y++)
      for (let x = 0; x < width; x++)
        if (inked(x, y)) {
          if (x < minX) minX = x;
          if (x > maxX) maxX = x;
          if (y < minY) minY = y;
          if (y > maxY) maxY = y;
        }
    expect(maxX).toBeGreaterThan(minX);
    expect(maxY).toBeGreaterThan(minY);

    // For a vertical slab [xc-6, xc+6], the mean-x of ink in the TOP third vs the
    // BOTTOM third of the glyph band. Positive shift = top sits right of bottom.
    const slabShift = (xc: number): number | null => {
      const bandTop = minY + (maxY - minY) * 0.15;
      const bandBot = minY + (maxY - minY) * 0.85;
      const third = (bandBot - bandTop) / 3;
      let txSum = 0, tN = 0, bxSum = 0, bN = 0;
      for (let x = Math.max(0, xc - 6); x <= Math.min(width - 1, xc + 6); x++) {
        for (let y = Math.floor(bandTop); y < bandTop + third; y++) if (inked(x, y)) { txSum += x; tN++; }
        for (let y = Math.floor(bandBot - third); y < bandBot; y++) if (inked(x, y)) { bxSum += x; bN++; }
      }
      if (tN < 4 || bN < 4) return null;
      return txSum / tN - bxSum / bN;
    };
    const shifts: number[] = [];
    for (let f = 0.15; f <= 0.85; f += 0.05) {
      const s = slabShift(Math.round(minX + (maxX - minX) * f));
      if (s != null) shifts.push(s);
    }
    expect(shifts.length).toBeGreaterThan(4);
    // The lean is substantial somewhere (a rigid rotate keeps |shift| small since
    // the whole glyph turns together; the shear slides the top several px).
    const maxAbs = Math.max(...shifts.map((s) => Math.abs(s)));
    expect(maxAbs).toBeGreaterThan(4);
    // …and it reverses across the wave (upslope vs downslope) — a single global
    // tilt could not do that.
    expect(Math.max(...shifts)).toBeGreaterThan(2);
    expect(Math.min(...shifts)).toBeLessThan(-2);
  });

  it('textInflate stretches a single glyph MORE on its centre-facing edge (piecewise-affine strips)', async () => {
    // The definitive strip-subdivision invariant. On Inflate the envelope gap
    // (top→bottom) is TALLEST at the box centre and shrinks toward the ends, and
    // the gap stays VERTICAL everywhere (top/bottom mirror-symmetric) — so shear
    // and angle are ~0 and the only warp is a vertical stretch that VARIES with x.
    // A single per-glyph affine (#866) samples that stretch at ONE u (the glyph
    // centre) and applies it uniformly, leaving a wide glyph the SAME height on
    // its left and right edges. Mapping the glyph outline continuously (as
    // PowerPoint does) makes the edge nearer the box centre taller. We render a
    // single wide glyph sitting in the LEFT half of the box and require its RIGHT
    // (centre-facing) edge to be meaningfully taller than its LEFT edge. Strips
    // reproduce this; the single affine cannot (ratio ≈ 1).
    const width = 400;
    const height = 240;
    const canvas = new Canvas(width, height);
    // A wide glyph then filler so the measured glyph lands left-of-centre.
    const slide = buildSlide('textInflate');
    // Swap the run text to "W....." (wide W, dotted filler pushes it left).
    (slide.slides[0].elements[0] as ShapeElement).textBody!.paragraphs[0].runs[0] = {
      type: 'text', text: 'W.....', bold: true, italic: null, underline: false,
      strikethrough: false, fontSize: 40, color: '000000', fontFamily: 'Arial',
    } as Paragraph['runs'][number];
    await renderSlideNode(
      canvas as unknown as Parameters<typeof renderSlideNode>[0],
      slide,
      0,
      { width, dpr: 1, factory: warpFactory },
    );
    const ctx = canvas.getContext('2d');
    const data = ctx.getImageData(0, 0, width, height).data as unknown as Uint8ClampedArray;
    const inked = (x: number, y: number): boolean => {
      const i = (y * width + x) * 4;
      const a = data[i + 3];
      const lum = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
      return a > 40 && lum < 128;
    };
    const span = (x: number): number | null => {
      let t = -1, b = -1;
      for (let y = 0; y < height; y++) if (inked(x, y)) { if (t < 0) t = y; b = y; }
      return t < 0 ? null : b - t + 1;
    };
    // Leftmost inked column, then the first glyph block (until an >8px empty gap).
    let minX = width;
    for (let x = 0; x < width; x++) if (span(x) != null) { minX = x; break; }
    let gx1 = minX, gap = 0;
    for (let x = minX; x < width; x++) {
      if (span(x) != null) { gx1 = x; gap = 0; }
      else { gap++; if (gap > 8 && gx1 > minX + 4) break; }
    }
    const gw = gx1 - minX;
    expect(gw).toBeGreaterThan(10);
    const edgeH = (f: number): number => {
      const c = Math.round(minX + gw * f);
      let s = 0, n = 0;
      for (let x = c - 1; x <= c + 1; x++) { const v = span(x); if (v != null) { s += v; n++; } }
      return n ? s / n : 0;
    };
    const leftH = edgeH(0.15);
    const rightH = edgeH(0.85);
    expect(leftH).toBeGreaterThan(0);
    // Centre-facing (right) edge is at least 12% taller than the outer (left) edge.
    // Measured: single-affine ≈1.00 (fails), strips ≈1.30 (passes).
    expect(rightH / leftH).toBeGreaterThan(1.12);
  });

  it('high-curvature envelopes (rings/pours) stay contiguous — no radial slivers', async () => {
    // Regression guard for the ADAPTIVE strip count. On ring/pour envelopes the
    // edge tangent sweeps up to ~345° across the width, so with a fixed
    // width-only strip count adjacent strips rotate apart faster than the ±1px
    // overlap can bridge: each glyph shatters into a fan of disconnected radial
    // slivers (measured: "WARP" in textRingOutside split into dozens of
    // 8-connected ink components; a contiguous render has one component per
    // detached glyph part). The deviation-driven subdivision keeps every strip
    // within 1 device px of the exact envelope map, so the ink stays connected.
    // "WARP" is four single-piece letters — allow a small margin for warped
    // letters splitting at hairline joins, but a sliver fan (>>8) must fail.
    for (const preset of ['textCirclePour', 'textRingOutside', 'textCanUp']) {
      const width = 400;
      const height = 240;
      const canvas = new Canvas(width, height);
      await renderSlideNode(
        canvas as unknown as Parameters<typeof renderSlideNode>[0],
        buildSlide(preset),
        0,
        { width, dpr: 1, factory: warpFactory },
      );
      const ctx = canvas.getContext('2d');
      const data = ctx.getImageData(0, 0, width, height).data as unknown as Uint8ClampedArray;
      const grid = new Uint8Array(width * height);
      let inkPx = 0;
      for (let p = 0; p < width * height; p++) {
        const i = p * 4;
        const lum = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
        if (data[i + 3] > 40 && lum < 128) { grid[p] = 1; inkPx++; }
      }
      expect(inkPx).toBeGreaterThan(100);
      // 8-connected component count over the ink mask.
      const seen = new Uint8Array(width * height);
      const stack: number[] = [];
      let comps = 0;
      for (let start = 0; start < width * height; start++) {
        if (!grid[start] || seen[start]) continue;
        comps++;
        stack.length = 0;
        stack.push(start);
        seen[start] = 1;
        while (stack.length) {
          const p = stack.pop() as number;
          const x = p % width, y = (p / width) | 0;
          for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
            if (!dx && !dy) continue;
            const nx = x + dx, ny = y + dy;
            if (nx < 0 || ny < 0 || nx >= width || ny >= height) continue;
            const np = ny * width + nx;
            if (grid[np] && !seen[np]) { seen[np] = 1; stack.push(np); }
          }
        }
      }
      // eslint-disable-next-line no-console
      console.log(`${preset}: inkPx=${inkPx} components=${comps}`);
      expect(comps).toBeLessThanOrEqual(8);
    }
  });

  it('translucent warp fill lands at a uniform alpha — no strip-overlap band (#879)', async () => {
    // A 50%-alpha fill (RRGGBBAA = 1F4E7980, alpha byte 0x80) through the strip
    // overlap band double-composed to ~75% opacity before #879, striping the ink
    // with higher-opacity pinstripes (Inflate ≈28%, CirclePour ≈71% of ink
    // columns in the issue's browser measurement). Two presets — a low-curvature
    // (Inflate) and a high-curvature (CirclePour, many strips ⇒ many seams) —
    // must now composite to a uniform single-composite alpha: essentially no ink
    // pixel exceeds the single/double midpoint, and the peak alpha stays near the
    // single level rather than the doubled one.
    for (const preset of ['textInflate', 'textCirclePour']) {
      const { ink, bandFraction, minLum, l1, l2 } = await warpTranslucentBand(preset, '1F4E7980');
      // eslint-disable-next-line no-console
      console.log(
        `${preset}: ink=${ink} band=${(bandFraction * 100).toFixed(2)}% minLum=${minLum.toFixed(1)} (L1=${l1.toFixed(1)} L2=${l2.toFixed(1)})`,
      );
      expect(ink, `${preset}: has translucent ink`).toBeGreaterThan(2000);
      // Double-composed band pixels: a whole-glyph raster would stripe a large
      // fraction; the layer composite drives this to ~0. A hair of slack absorbs
      // the odd self-overlapping outline corner.
      expect(bandFraction, `${preset}: double-composed band fraction`).toBeLessThan(0.01);
      // Darkest ink stays near the single composite L1, nowhere near the doubled
      // L2 the overlap band would reach.
      expect(minLum, `${preset}: darkest ink near single composite`).toBeGreaterThan((l1 + l2) / 2);
    }
  });
});
