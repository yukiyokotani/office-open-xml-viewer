/**
 * One document's symbol oracle (symbol-face.ts) across its renders. Kept in
 * its own module: only PPTX keeps an oracle across renders (a DOCX text
 * service already lives as long as its fonts), so the other formats neither
 * bundle nor run it.
 */
import {
  createSymbolFaceOracle,
  realmFontSet,
  scratchCanvas,
  type SymbolFaceOracle,
  type SymbolProbeCanvasFactory,
  type SymbolProbeContext,
} from './symbol-face.js';

/** One document's {@link SymbolFaceOracle} across its renders in a realm. */
export interface SymbolFaceOracleCache {
  /** The oracle for a render painting with `context`: the kept one while the
   * realm's font faces are unchanged since it was made. */
  oracleFor(context: SymbolProbeContext | null | undefined): SymbolFaceOracle;
  /** Drop the kept oracle and the realm's face snapshot (the document is
   * gone). An oracle handed out before still answers, judging afresh in its
   * own context, so a painting context that outlives the document retains
   * no font face through it. */
  dispose(): void;
}

/**
 * Keep one oracle across a document's renders (PPTX slides), instead of one
 * per render: its verdicts (sentinel measurements) and colour proofs are the
 * same while the faces it was judged with are. Verdicts depend only on the
 * Canvas font string (kept per verdict) and on the faces of the realm's
 * FontFaceSet, so the oracle is replaced whenever that set changes: a face
 * added or deleted (a delta of a font source, an Office fallback route, an
 * embedded face), a face whose status changed (a subset that finished
 * loading) or one whose descriptors were rewritten. It measures on a 1x1 scratch canvas of the painting context's
 * realm, so it retains no painting canvas; without one each render gets its
 * own oracle, as before. A realm without a FontFaceSet (a Node canvas) has
 * fixed host fonts: its oracle is kept for good.
 */
export function createSymbolFaceOracleCache(factory?: SymbolProbeCanvasFactory): SymbolFaceOracleCache {
  let oracle: SymbolFaceOracle | undefined;
  let fontSet: FontFaceSet | null = null;
  let faces: FontFace[] = [];
  let states: string[] = [];
  let disposed = false;
  // What font matching and shaping read of a face (CSS Fonts 4 §4.5, §5):
  // its status and descriptors, which a page may rewrite on a registered face.
  const state = (face: FontFace) => [face.status, face.family, face.style, face.weight, face.stretch,
    face.unicodeRange, face.featureSettings, (face as { variationSettings?: string }).variationSettings].join('\u0000');
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
  const options = factory ? { createCanvas: factory } : {};
  /** The kept oracle when the realm's faces are those it was made with, a
   * new one otherwise; undefined when the realm cannot be keyed or has no
   * scratch canvas. */
  const current = (context: SymbolProbeContext): SymbolFaceOracle | undefined => {
    if (disposed) return undefined;
    try {
      const set = realmFontSet(context);
      if (oracle && unchanged(set)) return oracle;
      const probe = scratchCanvas(context, factory, 1)?.getContext('2d') as SymbolProbeContext | null | undefined;
      oracle = undefined;
      if (!probe) return undefined;
      record(set);
      oracle = createSymbolFaceOracle(probe, options);
      return oracle;
    } catch {
      oracle = undefined;
      return undefined;
    }
  };
  return {
    oracleFor(context) {
      if (!context) return createSymbolFaceOracle(context, options);
      // A render awaits (images, fonts) between its layout and paint steps,
      // and the faces may change meanwhile: the oracle is looked up again in
      // each synchronous step that consults it (once per step, as faces
      // change only between tasks), never kept across an await.
      let step: SymbolFaceOracle | undefined;
      let own: SymbolFaceOracle | undefined;
      const now = (): SymbolFaceOracle => {
        if (!step) {
          step = current(context) ?? (own ??= createSymbolFaceOracle(context, options));
          queueMicrotask(() => { step = undefined; });
        }
        return step;
      };
      return Object.freeze({
        draws: (code: string, font: string) => now().draws(code, font),
        monochrome: (text: string, font: string) => now().monochrome(text, font),
        segments: (...args: Parameters<SymbolFaceOracle['segments']>) => now().segments(...args),
      });
    },
    dispose() {
      disposed = true;
      oracle = undefined;
      fontSet = null;
      faces = [];
      states = [];
    },
  };
}
