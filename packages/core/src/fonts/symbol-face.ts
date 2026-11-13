/**
 * Which authored symbol codes the painting font draws itself (symbol-font.ts).
 *
 * Nothing here models font matching. The engine that paints is asked, in the
 * measuring/painting context and with the Canvas font the kept text paints
 * with, so `unicode-range`, weight/style matching and synthesis, family-name
 * syntax, local() aliases and installed faces are all the engine's own.
 *
 * Sentinel mode (the route sentinels of route-sentinels.ts are
 * loaded in the context's realm). A sentinel maps every scalar, so appending
 * one to the font's named families ends font fallback there: a character the
 * named families do not draw is drawn by the sentinel, never by a generic or
 * system fallback face. The two sentinels differ in advance and ink, hence
 *
 *   drawn(x)  :=  `named, A` and `named, B` measure x identically
 *
 * (after a control showing that a sentinel alone draws x at all). Then
 *
 *   U+F000+c is native  :=  drawn(U+F000+c)
 *   bare c is native    :=  drawn(c), drawn(U+F000+c), and both measure
 *                           identically in `named, A`
 *
 * A named family that maps U+F0xx is taken to draw the symbol glyph there. A
 * bare code is a symbol glyph only through the symbol-cmap remapping of the
 * shaper, which resolves it to the very glyph of U+F000+c; a Unicode face
 * draws a Latin character for it, or nothing. Kept text is painted with the
 * same named families in the same order, so the face judged is the face that
 * paints.
 *
 * Plain mode (a realm without any FontFaceSet: a Node canvas, whose fonts are
 * host fonts only). No sentinel can be registered, and an absent family's
 * fallback is no evidence (system fallback depends on the primary face), so
 * a code is native only on positive proof: the bare code and U+F000+c measure
 * identically in the painting font, and each differs from its measurement in
 * an absent family. A face that maps only the PUA form converts there.
 *
 * Otherwise (a FontFaceSet without working sentinels, a context that does
 * not measure) nothing is judged and the 0.89 subset converts.
 */
import { createAuxCanvasForContext } from '../canvas/aux-canvas.js';
import { quoteCssFontFamily } from '../internal/css-font-family.js';
import { SENTINEL_FAMILIES, parseFamilyList, sentinelsWork } from './route-sentinels.js';
import { symbolTextToUnicodeSegments, type SymbolTextSegment } from './symbol-font.js';

/** The part of a 2D context the measurement needs. */
export interface SymbolProbeContext {
  font: string;
  direction?: CanvasDirection;
  letterSpacing?: string;
  measureText(text: string): Pick<TextMetrics, 'width'> & Partial<Pick<TextMetrics,
    'actualBoundingBoxLeft' | 'actualBoundingBoxRight' | 'actualBoundingBoxAscent' | 'actualBoundingBoxDescent'>>;
}

const PROBE_PX = 64;
/** No host installs this family. */
const ABSENT = '"__ooxml_symbol_probe_absent_7c41__"';
const [SENTINEL_A, SENTINEL_B] = SENTINEL_FAMILIES.map(quoteCssFontFamily) as [string, string];
/** DELETE and SOFT HYPHEN are default-ignorable as bare characters: no bare
 * measurement can show the identity of plain mode for them. */
const BARE_IGNORED: ReadonlySet<number> = new Set([0x7f, 0xad]);

const same = (a: readonly number[], b: readonly number[]) =>
  a.every((value, index) => Number.isFinite(value) && value === b[index]);

interface ProbeCanvas {
  getContext(type: '2d', options?: CanvasRenderingContext2DSettings): unknown;
}

/** Allocates a canvas of the backend that measures and paints the text. */
export type SymbolProbeCanvasFactory = (width: number, height: number) => ProbeCanvas | null | undefined;

/**
 * A scratch canvas of the backend (and so the fonts) that paints for `owner`.
 * `factory` is the owner's own canvas factory: the caller that created the
 * measuring/painting canvas allocates the probe's the same way (a Node
 * session through its canvas factory). Only without one is the backend
 * inferred: the owner canvas's document, then the realm's global helpers
 * (OffscreenCanvas, `document`, the owner's class), which in a realm with
 * several canvas backends need not be the one that paints.
 */
function scratchCanvas(owner: unknown, factory?: SymbolProbeCanvasFactory, size = 48): ProbeCanvas | null {
  try {
    if (factory) return factory(size, size) ?? null;
    const element = (owner as { canvas?: { ownerDocument?: Document } } | null)?.canvas;
    const doc = element?.ownerDocument;
    if (doc && typeof doc.createElement === 'function') {
      const canvas = doc.createElement('canvas');
      canvas.width = size;
      canvas.height = size;
      return canvas;
    }
    return createAuxCanvasForContext(owner as CanvasRenderingContext2D, size, size);
  } catch { /* no scratch canvas */ }
  return null;
}

/**
 * Whether `font` (a CSS font shorthand: the route converted text paints with)
 * paints `text` in the text colour in the realm of `owner` (a
 * measuring/painting 2D context): drawn in red and then in blue, a glyph
 * that takes the fill leaves only pixels of that colour, while a colour
 * glyph paints its own (a fixed black palette entry included). Undefined
 * (inconclusive) when no scratch canvas exists, it cannot be read back (or
 * its readback is not the canvas drawn: a cleared canvas must read back all
 * zero), or the drawing left no ink to judge. Results live as long as the probe.
 */
export function createSymbolMonochromeProbe(
  owner: unknown,
  defaultFont = '32px sans-serif',
  factory?: SymbolProbeCanvasFactory,
): (text: string, font?: string) => boolean | undefined {
  let ctx: CanvasRenderingContext2D | null | undefined;
  let created = false;
  const results = new Map<string, boolean | undefined>();
  return (text: string, font = defaultFont) => {
    const key = `${font}\u0000${text}`;
    if (results.has(key)) return results.get(key);
    if (!created) {
      created = true;
      try {
        ctx = scratchCanvas(owner, factory)?.getContext('2d', { willReadFrequently: true }) as CanvasRenderingContext2D | null | undefined;
      } catch { ctx = null; }
    }
    let monochrome: boolean | undefined;
    if (ctx) {
      const context = ctx;
      // Whether every solid pixel of the text drawn in `fill` has that
      // colour (`channel` full, the other two empty). Faint antialiasing
      // pixels carry too little colour to judge, so a drawing without one
      // solid pixel (no ink, or a wholly translucent glyph) is inconclusive.
      const follows = (fill: string, channel: number): boolean | undefined => {
        context.clearRect(0, 0, 48, 48);
        // A readback of the canvas just cleared that is not all zero (an
        // anti-fingerprinting engine adds noise to getImageData or answers
        // a constant image) proves nothing about the glyph: inconclusive,
        // never "not monochrome".
        if (context.getImageData(0, 0, 48, 48).data.some((value) => value !== 0)) return undefined;
        context.globalAlpha = 1;
        context.globalCompositeOperation = 'source-over';
        context.fillStyle = fill;
        context.textBaseline = 'middle';
        context.font = font;
        context.fillText(text, 4, 24);
        const data = context.getImageData(0, 0, 48, 48).data;
        let judged = false;
        for (let index = 0; index < data.length; index += 4) {
          if (data[index + 3]! < 64) continue;
          judged = true;
          const own = data[index + channel]!;
          if (own < 200 || data[index]! + data[index + 1]! + data[index + 2]! - own > 48) return false;
        }
        return judged ? true : undefined;
      };
      try {
        monochrome = follows('#f00', 0);
        if (monochrome) monochrome = follows('#00f', 2);
      } catch {
        monochrome = undefined;
      }
    }
    if (results.size >= 1024) results.clear();
    results.set(key, monochrome);
    return monochrome;
  };
}

/** The FontFaceSet of the realm `owner` (a 2D context) measures in. */
function realmFontSet(owner: unknown): FontFaceSet | null {
  const element = (owner as { canvas?: { ownerDocument?: Document } } | null)?.canvas;
  if (element?.ownerDocument?.fonts) return element.ownerDocument.fonts;
  const scope = globalThis as { document?: Document; fonts?: FontFaceSet };
  return scope.document?.fonts ?? scope.fonts ?? null;
}

export interface SymbolFaceOracleOptions {
  /** The real 2D context when `context` wraps it: its realm is the one whose
   * fonts are judged, and colour probes draw on a scratch canvas of it.
   * Defaults to `context`. */
  readonly probeOwner?: unknown;
  /** The canvas factory of the backend that measures and paints: colour
   * probes draw on a canvas it allocates. Absent: inferred from the owner
   * (see {@link SymbolProbeCanvasFactory}). */
  readonly createCanvas?: SymbolProbeCanvasFactory;
}

/** The CSS font converted text paints with, at a probe size, for one
 * converted string. */
export type SymbolFallbackFont = (text: string) => string;

export interface SymbolFaceOracle {
  /**
   * Whether `font` draws the authored symbol code `code` (one character: a
   * bare code 0x20..0xFF or U+F020..U+F0FF) as its own symbol glyph. `font`
   * is the Canvas font the kept text paints with, as the library builds it:
   * `[style] [weight] <size>px <family list>`. Undefined: not judged.
   */
  draws(code: string, font: string): boolean | undefined;
  /** Whether `font` (a Canvas font) paints `text` in the text colour
   * ({@link createSymbolMonochromeProbe}); undefined: inconclusive. */
  monochrome(text: string, font: string): boolean | undefined;
  /** {@link symbolTextToUnicodeSegments} of `text` in the symbol family
   * `fontFamily`, each code judged by {@link draws} in `font`. */
  segments(
    text: string,
    fontFamily: string | null | undefined,
    font: string,
    fallbackFont?: SymbolFallbackFont,
    keepsShifted?: (shifted: string) => boolean,
  ): SymbolTextSegment[];
}

/**
 * Symbol decisions in one measuring/painting context (see the module notes).
 * Verdicts are kept for the oracle's lifetime, which callers bound by the
 * font resources they lay out or render with (a DOCX text service, a PPTX
 * document while its realm's faces are unchanged:
 * {@link createSymbolFaceOracleCache}), exactly like the advances measured in the same context. A kept
 * verdict is answered without touching the context; an unjudged code (the
 * sentinels do not work yet) is not kept, so it is judged once they do.
 */
export function createSymbolFaceOracle(
  context: SymbolProbeContext | null | undefined,
  options: SymbolFaceOracleOptions = {},
): SymbolFaceOracle {
  const probeOwner = options.probeOwner ?? context;
  const monochrome = createSymbolMonochromeProbe(probeOwner, undefined, options.createCanvas);
  const verdicts = new Map<string, boolean>();
  let font = '';
  const measure = (probeFont: string, text: string): number[] => {
    if (font !== probeFont) {
      context!.font = probeFont;
      font = probeFont;
    }
    const m = context!.measureText(text);
    return [m.width, m.actualBoundingBoxLeft ?? 0, m.actualBoundingBoxRight ?? 0,
      m.actualBoundingBoxAscent ?? 0, m.actualBoundingBoxDescent ?? 0];
  };
  /** Run `probe` left to right and without spacing (a bare "(" is mirrored
   * in a right-to-left context, its PUA form is not), restoring the context. */
  const probing = <T>(probe: () => T): T | undefined => {
    if (!context) return undefined;
    const previous = context.font;
    const direction = context.direction;
    const spacing = context.letterSpacing;
    try {
      if (direction !== undefined) context.direction = 'ltr';
      if (spacing !== undefined) context.letterSpacing = '0px';
      font = '';
      return probe();
    } catch {
      return undefined;
    } finally {
      try {
        context.font = previous;
        if (direction !== undefined) context.direction = direction;
        if (spacing !== undefined) context.letterSpacing = spacing;
      } catch { /* restored best-effort */ }
    }
  };
  const judge = (mode: 'sentinel' | 'plain', pua: boolean, code: number, prefix: string, familyList: string): boolean | undefined => {
    const shifted = String.fromCharCode(0xf000 + code);
    const bare = String.fromCharCode(code);
    const size = `${prefix}${PROBE_PX}px `;
    if (mode === 'sentinel') {
      const families = parseFamilyList(familyList);
      if (!families) return undefined;
      const generic = families.findIndex((entry) => entry.generic);
      const named = (generic < 0 ? families : families.slice(0, generic)).map((entry) => `${entry.text}, `).join('');
      const drawn = (text: string) =>
        !same(measure(`${size}${SENTINEL_A}`, text), measure(`${size}${SENTINEL_B}`, text))
        && same(measure(`${size}${named}${SENTINEL_A}`, text), measure(`${size}${named}${SENTINEL_B}`, text));
      if (!drawn(shifted)) return false;
      return pua || (drawn(bare)
        && same(measure(`${size}${named}${SENTINEL_A}`, bare), measure(`${size}${named}${SENTINEL_A}`, shifted)));
    }
    const painted = `${size}${familyList}`;
    const absent = `${size}${ABSENT}`;
    // A context whose text geometry does not vary with the text (a stub, a
    // canvas without fonts) cannot judge.
    if (same(measure(absent, 'i'), measure(absent, 'W'))) return undefined;
    const differs = (text: string) => !same(measure(painted, text), measure(absent, text));
    const identity = (probe: number) => {
      const b = String.fromCharCode(probe);
      const s = String.fromCharCode(0xf000 + probe);
      return same(measure(painted, s), measure(painted, b)) && differs(s) && differs(b);
    };
    // The PUA form of a default-ignorable code stays in a face that another
    // code ("A") proves symbol-encoded; its bare form is never native.
    if (BARE_IGNORED.has(code)) return pua && identity(0x41) && differs(shifted);
    return identity(code);
  };
  const decide = (character: string, paintFont: string): boolean | undefined => {
    const cp = character.codePointAt(0) ?? 0;
    const pua = cp >= 0xf020 && cp <= 0xf0ff;
    if (character.length !== 1 || (!pua && (cp < 0x20 || cp > 0xff))) return false;
    // The size is the first `<number>px` token: style, weight and width
    // keywords before it never contain one, and the library writes no
    // line-height.
    const split = /^(.*?)\d*\.?\d+px\s+(.+)$/su.exec(paintFont);
    if (!split) return undefined;
    const key = `${split[1]}\u0000${split[2]}\u0000${character}`;
    const kept = verdicts.get(key);
    if (kept !== undefined) return kept;
    return probing(() => {
      // A realm has a FontFaceSet or not for good, so a kept verdict never
      // outlives its mode.
      const mode = sentinelsWork(context as never) ? 'sentinel' : realmFontSet(probeOwner) ? null : 'plain';
      if (!mode) return undefined;
      font = '';
      const verdict = judge(mode, pua, pua ? cp - 0xf000 : cp, split[1]!, split[2]!);
      if (verdict !== undefined) {
        if (verdicts.size >= 4096) verdicts.clear();
        verdicts.set(key, verdict);
      }
      return verdict;
    });
  };
  return Object.freeze({
    draws: decide,
    monochrome: (text: string, paintFont: string) => monochrome(text, paintFont),
    segments(text: string, fontFamily: string | null | undefined, paintFont: string, fallbackFont?: SymbolFallbackFont,
      keepsShifted?: (shifted: string) => boolean) {
      return symbolTextToUnicodeSegments(text, fontFamily, {
        ...(keepsShifted ? { keepsShifted } : {}),
        native: (code) => decide(code, paintFont),
        // A shaper may paint the VS15 apart from its base (another script
        // slot), so the base alone must be monochrome in its route too.
        monochrome: (converted) => {
          const base = converted.replace(/︎/gu, '');
          const withSelector = monochrome(converted, fallbackFont?.(converted));
          if (withSelector !== true || base === converted) return withSelector;
          return monochrome(base, fallbackFont?.(base));
        },
      });
    },
  });
}

/** One document's {@link SymbolFaceOracle} across its renders in a realm. */
export interface SymbolFaceOracleCache {
  /** The oracle for a render painting with `context`: the kept one while the
   * realm's font faces are unchanged since it was made. */
  oracleFor(context: SymbolProbeContext | null | undefined): SymbolFaceOracle;
}

/**
 * Keep one oracle across a document's renders (PPTX slides), instead of one
 * per render: its verdicts (sentinel measurements) and colour proofs are the
 * same while the faces it was judged with are. Verdicts depend only on the
 * Canvas font string (kept per verdict) and on the faces of the realm's
 * FontFaceSet, so the oracle is replaced whenever that set changes: a face
 * added or deleted (a delta of a font source, an Office fallback route, an
 * embedded face), a face whose status changed (a subset that finished
 * loading) or one whose matching descriptors were rewritten. It measures on a 1x1 scratch canvas of the painting context's
 * realm, so it retains no painting canvas; without one each render gets its
 * own oracle, as before. A realm without a FontFaceSet (a Node canvas) has
 * fixed host fonts: its oracle is kept for good.
 */
export function createSymbolFaceOracleCache(factory?: SymbolProbeCanvasFactory): SymbolFaceOracleCache {
  let oracle: SymbolFaceOracle | undefined;
  let fontSet: FontFaceSet | null = null;
  let faces: FontFace[] = [];
  let states: string[] = [];
  // What font matching reads of a face (CSS Fonts 4 §5): its status and its
  // matching descriptors, which a page may mutate on a registered face.
  const state = (face: FontFace) =>
    `${face.status}\u0000${face.family}\u0000${face.style}\u0000${face.weight}\u0000${face.stretch}\u0000${face.unicodeRange}`;
  const unchanged = (set: FontFaceSet | null): boolean => {
    if (set !== fontSet) return false;
    if (!set) return true;
    let index = 0;
    for (const face of set as unknown as Iterable<FontFace>) {
      if (faces[index] !== face || states[index] !== state(face)) return false;
      index++;
    }
    return index === faces.length;
  };
  const record = (set: FontFaceSet | null) => {
    fontSet = set;
    faces = [];
    states = [];
    if (set) {
      for (const face of set as unknown as Iterable<FontFace>) {
        faces.push(face);
        states.push(state(face));
      }
    }
  };
  return {
    oracleFor(context) {
      const options = factory ? { createCanvas: factory } : {};
      if (!context) return createSymbolFaceOracle(context, options);
      let set: FontFaceSet | null;
      try {
        set = realmFontSet(context);
        if (oracle && unchanged(set)) return oracle;
      } catch {
        return createSymbolFaceOracle(context, options);
      }
      const scratch = scratchCanvas(context, factory, 1);
      let probe: SymbolProbeContext | null = null;
      try {
        probe = scratch?.getContext('2d') as SymbolProbeContext | null ?? null;
      } catch { probe = null; }
      if (!probe) {
        oracle = undefined;
        return createSymbolFaceOracle(context, options);
      }
      record(set);
      oracle = createSymbolFaceOracle(probe, options);
      return oracle;
    },
  };
}
