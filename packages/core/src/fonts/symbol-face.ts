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

/** A scratch canvas in the realm (and so the FontFaceSet) of `owner`. */
function scratchCanvas(owner: unknown): ProbeCanvas | null {
  try {
    const element = (owner as { canvas?: { ownerDocument?: Document } } | null)?.canvas;
    const doc = element?.ownerDocument;
    if (doc && typeof doc.createElement === 'function') {
      const canvas = doc.createElement('canvas');
      canvas.width = 48;
      canvas.height = 48;
      return canvas;
    }
    // A detached canvas (OffscreenCanvas, or a Node canvas implementation such
    // as @napi-rs/canvas or skia-canvas): another canvas of its own class.
    const Ctor = (element as { constructor?: unknown } | undefined)?.constructor;
    if (typeof Ctor === 'function' && Ctor !== Object) {
      try {
        const canvas = new (Ctor as new (width: number, height: number) => ProbeCanvas)(48, 48);
        if (typeof canvas.getContext === 'function') return canvas;
      } catch { /* not constructible this way */ }
    }
    if (typeof OffscreenCanvas !== 'undefined') return new OffscreenCanvas(48, 48);
  } catch { /* no scratch canvas */ }
  return null;
}

/**
 * Whether `font` (a CSS font shorthand: the route converted text paints with)
 * paints `text` without colour in the realm of `owner` (a measuring/painting
 * 2D context): drawn in opaque black, a monochrome glyph leaves only black
 * pixels, while a colour emoji face paints its own colours. Undefined
 * (inconclusive) when no scratch canvas exists, it cannot be read back, or
 * the drawing left no ink to judge. Results live as long as the probe.
 */
export function createSymbolMonochromeProbe(
  owner: unknown,
  defaultFont = '32px sans-serif',
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
        ctx = scratchCanvas(owner)?.getContext('2d', { willReadFrequently: true }) as CanvasRenderingContext2D | null | undefined;
      } catch { ctx = null; }
    }
    let monochrome: boolean | undefined;
    if (ctx) {
      try {
        ctx.clearRect(0, 0, 48, 48);
        ctx.globalAlpha = 1;
        ctx.globalCompositeOperation = 'source-over';
        ctx.fillStyle = '#000';
        ctx.textBaseline = 'middle';
        ctx.font = font;
        ctx.fillText(text, 4, 24);
        const data = ctx.getImageData(0, 0, 48, 48).data;
        let ink = false;
        let colour = false;
        for (let index = 0; index < data.length; index += 4) {
          if (data[index + 3]! === 0) continue;
          ink = true;
          if (data[index]! + data[index + 1]! + data[index + 2]! > 24) { colour = true; break; }
        }
        monochrome = ink ? !colour : undefined;
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
  /** {@link symbolTextToUnicodeSegments} of `text` in the symbol family
   * `fontFamily`, each code judged by {@link draws} in `font`. */
  segments(
    text: string,
    fontFamily: string | null | undefined,
    font: string,
    fallbackFont?: SymbolFallbackFont,
  ): SymbolTextSegment[];
}

/**
 * Symbol decisions in one measuring/painting context (see the module notes).
 * Verdicts are kept for the oracle's lifetime, which callers bound by the
 * font resources they lay out or render with (a DOCX text service, a PPTX
 * render), exactly like the advances measured in the same context; whether
 * the sentinels work is re-read with every {@link SymbolFaceOracle.segments}
 * and {@link SymbolFaceOracle.draws} call.
 */
export function createSymbolFaceOracle(
  context: SymbolProbeContext | null | undefined,
  options: SymbolFaceOracleOptions = {},
): SymbolFaceOracle {
  const probeOwner = options.probeOwner ?? context;
  const monochrome = createSymbolMonochromeProbe(probeOwner);
  const verdicts = new Map<string, boolean | undefined>();
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
  type Mode = 'sentinel' | 'plain' | null;
  const readMode = (): Mode =>
    sentinelsWork(context as never) ? 'sentinel' : realmFontSet(probeOwner) ? null : 'plain';
  /** `state.mode` is read once per public call, before its first verdict. */
  const decide = (character: string, paintFont: string, state: { mode?: Mode }): boolean | undefined => {
    const cp = character.codePointAt(0) ?? 0;
    const pua = cp >= 0xf020 && cp <= 0xf0ff;
    if (character.length !== 1 || (!pua && (cp < 0x20 || cp > 0xff))) return false;
    // The size is the first `<number>px` token: style, weight and width
    // keywords before it never contain one, and the library writes no
    // line-height.
    const split = /^(.*?)\d*\.?\d+px\s+(.+)$/su.exec(paintFont);
    if (!split) return undefined;
    return probing(() => {
      if (state.mode === undefined) state.mode = readMode();
      const mode = state.mode;
      if (!mode) return undefined;
      const key = `${mode}\u0000${split[1]}\u0000${split[2]}\u0000${character}`;
      if (verdicts.has(key)) return verdicts.get(key);
      font = '';
      const verdict = judge(mode, pua, pua ? cp - 0xf000 : cp, split[1]!, split[2]!);
      if (verdicts.size >= 4096) verdicts.clear();
      verdicts.set(key, verdict);
      return verdict;
    });
  };
  return Object.freeze({
    draws: (code: string, paintFont: string) => decide(code, paintFont, {}),
    segments(text: string, fontFamily: string | null | undefined, paintFont: string, fallbackFont?: SymbolFallbackFont) {
      const state = {};
      return symbolTextToUnicodeSegments(text, fontFamily, {
        native: (code) => decide(code, paintFont, state),
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
