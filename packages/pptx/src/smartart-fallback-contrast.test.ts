import { describe, it, expect } from 'vitest';
import {
  backgroundLuminance,
  isSmartArtFallbackShape,
  smartArtFallbackTextColor,
} from './smartart-fallback-contrast';
import { renderSlide, type PptxTextRunInfo } from './renderer';
import type { Fill } from '@silurus/ooxml-core';
import type { Paragraph, ShapeElement, Slide, TextBody } from './types';

/**
 * Issue #805: the SmartArt data-model fallback (a synthetic shape emitted by
 * the Rust parser when no drawing part is stored) renders its node texts with
 * `color: null`, falling to the renderer's theme default (dk1 — dark). On a
 * dark slide background the list is invisible. The renderer now derives a
 * contrast-aware default for that synthetic shape from the slide-background
 * luminance. Ordinary shapes are untouched.
 */

// ── pure helpers ───────────────────────────────────────────────────────────

describe('backgroundLuminance', () => {
  it('is dark for a solid navy background and light for solid white', () => {
    const navy: Fill = { fillType: 'solid', color: '1F3864' };
    const white: Fill = { fillType: 'solid', color: 'FFFFFF' };
    expect(backgroundLuminance(navy)).not.toBeNull();
    expect(backgroundLuminance(navy)!).toBeLessThan(0.5);
    expect(backgroundLuminance(white)).toBeCloseTo(1, 5);
  });

  it('integrates gradient stops piecewise-linearly over [0,1]', () => {
    // black@0 → white@1: mean luma of the linear ramp is 0.5.
    const ramp: Fill = {
      fillType: 'gradient',
      gradType: 'linear',
      angle: 90,
      stops: [
        { position: 0, color: '000000' },
        { position: 1, color: 'FFFFFF' },
      ],
    };
    expect(backgroundLuminance(ramp)).toBeCloseTo(0.5, 5);
    // Both stops dark navy → dark; end stops extend to the 0/1 edges.
    const navyRamp: Fill = {
      fillType: 'gradient',
      gradType: 'linear',
      angle: 90,
      stops: [
        { position: 0.2, color: '1F3864' },
        { position: 0.8, color: '0A1430' },
      ],
    };
    expect(backgroundLuminance(navyRamp)!).toBeLessThan(0.3);
  });

  it('composites an 8-char RRGGBBAA solid over the white canvas base', () => {
    // Black at 0% alpha is effectively the white base.
    const clear: Fill = { fillType: 'solid', color: '00000000' };
    expect(backgroundLuminance(clear)).toBeCloseTo(1, 5);
    // Black at full alpha stays black.
    const opaque: Fill = { fillType: 'solid', color: '000000FF' };
    expect(backgroundLuminance(opaque)).toBeCloseTo(0, 5);
  });

  it('returns null (unknown) for image/pattern/none/absent backgrounds', () => {
    expect(backgroundLuminance(null)).toBeNull();
    expect(backgroundLuminance({ fillType: 'none' })).toBeNull();
    expect(
      backgroundLuminance({
        fillType: 'image',
        imagePath: 'ppt/media/image1.png',
        mimeType: 'image/png',
      } as Fill),
    ).toBeNull();
    expect(
      backgroundLuminance({
        fillType: 'pattern',
        fg: '000000',
        bg: 'FFFFFF',
        preset: 'pct50',
      }),
    ).toBeNull();
  });
});

describe('smartArtFallbackTextColor', () => {
  const navy: Fill = { fillType: 'solid', color: '1F3864' };
  const white: Fill = { fillType: 'solid', color: 'FFFFFF' };

  it('turns white on a dark background when the theme default is dark', () => {
    expect(smartArtFallbackTextColor(navy, '#383838')).toBe('#FFFFFF');
  });

  it('keeps the existing default on a light background', () => {
    expect(smartArtFallbackTextColor(white, '#383838')).toBeNull();
  });

  it('keeps the existing default when the background is unknown', () => {
    expect(smartArtFallbackTextColor(null, '#383838')).toBeNull();
  });

  it('keeps a theme default that is already light (legible on dark)', () => {
    expect(smartArtFallbackTextColor(navy, '#EEEEEE')).toBeNull();
  });
});

describe('isSmartArtFallbackShape', () => {
  it('matches the parser-synthesized fingerprint (name, no cNvPr id)', () => {
    expect(isSmartArtFallbackShape({ name: 'SmartArt' })).toBe(true);
  });

  it('rejects real shapes (cNvPr id present or another name)', () => {
    expect(isSmartArtFallbackShape({ name: 'SmartArt', id: '4' })).toBe(false);
    expect(isSmartArtFallbackShape({ name: 'Rectangle 5' })).toBe(false);
    expect(isSmartArtFallbackShape({})).toBe(false);
  });
});

// ── renderSlide integration ────────────────────────────────────────────────

interface FillTextCall {
  text: string;
  fillStyle: string;
}

/** A minimal recording 2D context capturing fillStyle at each fillText. */
function recordingCtx(): { ctx: CanvasRenderingContext2D; fillTexts: FillTextCall[] } {
  const fillTexts: FillTextCall[] = [];
  const noop = () => {};
  const ctx: Record<string, unknown> = {
    // state + transforms
    save: noop,
    restore: noop,
    scale: noop,
    translate: noop,
    rotate: noop,
    setTransform: noop,
    transform: noop,
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
    // path building
    beginPath: noop,
    closePath: noop,
    moveTo: noop,
    lineTo: noop,
    bezierCurveTo: noop,
    quadraticCurveTo: noop,
    arc: noop,
    arcTo: noop,
    ellipse: noop,
    rect: noop,
    clip: noop,
    // paint
    fill: noop,
    stroke: noop,
    fillRect: noop,
    strokeRect: noop,
    clearRect: noop,
    strokeText: noop,
    drawImage: noop,
    setLineDash: noop,
    createLinearGradient: () => ({ addColorStop: noop }),
    createRadialGradient: () => ({ addColorStop: noop }),
    createPattern: () => null,
    // text
    fillText(text: string) {
      fillTexts.push({ text, fillStyle: String(this.fillStyle) });
    },
    measureText: (t: string) => ({
      width: t.length * 6,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 9,
      fontBoundingBoxDescent: 3,
    }),
    // style props (assigned by the renderer)
    fillStyle: '',
    strokeStyle: '',
    lineWidth: 1,
    lineCap: 'butt',
    lineJoin: 'miter',
    miterLimit: 10,
    globalAlpha: 1,
    font: '',
    textAlign: 'start',
    textBaseline: 'alphabetic',
    letterSpacing: '0px',
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fillTexts };
}

function stubCanvas(ctx: CanvasRenderingContext2D): HTMLCanvasElement {
  const canvas = {
    width: 0,
    height: 0,
    style: {} as CSSStyleDeclaration,
    offsetWidth: 960,
    getContext: () => ctx,
  } as unknown as HTMLCanvasElement;
  // Back-reference used by the renderer's effect-canvas sizing.
  (ctx as unknown as { canvas: HTMLCanvasElement }).canvas = canvas;
  return canvas;
}

function paragraph(text: string): Paragraph {
  return {
    alignment: 'l',
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
        text,
        bold: null,
        italic: null,
        underline: false,
        strikethrough: false,
        fontSize: 18,
        color: null,
        fontFamily: null,
      },
    ],
  };
}

function textBody(text: string): TextBody {
  return {
    verticalAnchor: 't',
    paragraphs: [paragraph(text)],
    defaultFontSize: 18,
    defaultBold: null,
    defaultItalic: null,
    lIns: 91440,
    rIns: 91440,
    tIns: 45720,
    bIns: 45720,
    wrap: 'square',
    vert: 'horz',
    autoFit: 'none',
    numCol: 1,
    spcCol: 0,
  };
}

/** The shape the Rust SmartArt fallback synthesizes (see smartart_fallback.rs
 *  `text_list_shape`): name "SmartArt", no cNvPr id, no style-derived default
 *  text colour. */
function fallbackShape(text: string, overrides: Partial<ShapeElement> = {}): ShapeElement {
  return {
    type: 'shape',
    x: 914400,
    y: 914400,
    width: 4572000,
    height: 2286000,
    rotation: 0,
    flipH: false,
    flipV: false,
    geometry: 'rect',
    fill: null,
    stroke: null,
    textBody: textBody(text),
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
    name: 'SmartArt',
    ...overrides,
  };
}

function slideWith(background: Fill | null, elements: ShapeElement[]): Slide {
  return {
    index: 0,
    slideNumber: 1,
    background,
    elements,
  };
}

const DARK_BG: Fill = { fillType: 'solid', color: '1F3864' };

async function renderAndCollect(slide: Slide): Promise<FillTextCall[]> {
  const { ctx, fillTexts } = recordingCtx();
  const canvas = stubCanvas(ctx);
  await renderSlide(canvas, slide, 9_144_000, 6_858_000, {
    width: 960,
    dpr: 1,
    defaultTextColor: '383838',
  });
  return fillTexts;
}

async function renderAndCollectTextRuns(slide: Slide): Promise<PptxTextRunInfo[]> {
  const { ctx } = recordingCtx();
  const canvas = stubCanvas(ctx);
  const runs: PptxTextRunInfo[] = [];
  await renderSlide(canvas, slide, 9_144_000, 6_858_000, {
    width: 960,
    dpr: 1,
    defaultTextColor: '383838',
  }, (run) => runs.push(run));
  return runs;
}

describe('renderSlide text-run shape identity', () => {
  it('surfaces the source cNvPr id on every run of a file-authored shape', async () => {
    const runs = await renderAndCollectTextRuns(
      slideWith(null, [fallbackShape('Editable title', { id: '42', name: 'Title 1' })]),
    );

    expect(runs).not.toHaveLength(0);
    expect(runs.every((run) => run.shapeId === '42')).toBe(true);
  });

  it('leaves identity absent for parser-synthesized SmartArt fallback shapes', async () => {
    const runs = await renderAndCollectTextRuns(
      slideWith(null, [fallbackShape('Synthetic item')]),
    );

    expect(runs).not.toHaveLength(0);
    expect(runs.every((run) => run.shapeId === undefined)).toBe(true);
  });
});

describe('renderSlide SmartArt fallback contrast (issue #805)', () => {
  it('renders null-colour fallback runs white on a dark background', async () => {
    const calls = await renderAndCollect(slideWith(DARK_BG, [fallbackShape('Task one')]));
    const run = calls.find((c) => c.text.includes('Task one'));
    expect(run).toBeDefined();
    expect(run!.fillStyle).toBe('#FFFFFF');
  });

  it('keeps the theme default on a light background (unchanged behaviour)', async () => {
    const calls = await renderAndCollect(
      slideWith({ fillType: 'solid', color: 'FFFFFF' }, [fallbackShape('Task one')]),
    );
    const run = calls.find((c) => c.text.includes('Task one'));
    expect(run).toBeDefined();
    expect(run!.fillStyle).toBe('#383838');
  });

  it('does not touch ordinary shapes on a dark background', async () => {
    const calls = await renderAndCollect(
      slideWith(DARK_BG, [fallbackShape('Plain box', { name: 'Rectangle 5', id: '4' })]),
    );
    const run = calls.find((c) => c.text.includes('Plain box'));
    expect(run).toBeDefined();
    expect(run!.fillStyle).toBe('#383838');
  });

  it('leaves explicit run colours alone on the fallback shape', async () => {
    const shape = fallbackShape('Orange label');
    const runData = shape.textBody!.paragraphs[0].runs[0];
    if (runData.type === 'text') runData.color = 'ED7D31';
    const calls = await renderAndCollect(slideWith(DARK_BG, [shape]));
    const run = calls.find((c) => c.text.includes('Orange label'));
    expect(run).toBeDefined();
    expect(run!.fillStyle).not.toBe('#FFFFFF');
  });
});
