// From-scratch UAX#9 engine — the swappable implementation behind BidiEngine.
// The rule set lives in ./rules.ts (conformance-tested against Unicode's
// BidiCharacterTest.txt / BidiTest.txt). This file adapts the code-point-space
// core to the string / code-unit surface renderers use.

import type { BidiEngine } from '../engine.js';
import { REMOVED_LEVEL } from '../types.js';
import type { BaseDirection, BidiClass, BidiLevels } from '../types.js';
import { mirror as mirrorGlyph } from '../char-data.js';
import { resolveLevels, reorderByLevels, REMOVED } from './rules.js';

/** Decode a UTF-16 string into code points, their code-unit widths, and the
 *  code-unit offset at which each code point begins. */
function decode(text: string): { cps: number[]; units: number[]; starts: number[] } {
  const cps: number[] = [];
  const units: number[] = [];
  const starts: number[] = [];
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    const len = cp > 0xffff ? 2 : 1;
    cps.push(cp);
    units.push(len);
    starts.push(i);
    i += len;
  }
  return { cps, units, starts };
}

class Uax9Engine implements BidiEngine {
  computeLevels(
    text: string,
    base: BaseDirection,
    classOverride?: ReadonlyArray<BidiClass | null | undefined>,
  ): { levels: BidiLevels; paragraphLevel: number } {
    const { cps, units, starts } = decode(text);
    // Collapse the per-code-UNIT override the renderer supplies into the
    // per-code-POINT space the rule engine works in (read it at the code
    // point's leading unit). Skipped entirely when no override is passed.
    const cpOverride = classOverride
      ? cps.map((_, i) => classOverride[starts[i]] ?? null)
      : undefined;
    const { levels: cpLevels, paragraphLevel } = resolveLevels(cps, base, cpOverride);

    // Expand per-code-point levels back to per-code-unit (both surrogate halves
    // share the level), mapping the REMOVED sentinel to 255.
    const out = new Uint8Array(text.length);
    let u = 0;
    for (let i = 0; i < cpLevels.length; i++) {
      const v = cpLevels[i] === REMOVED ? REMOVED_LEVEL : cpLevels[i];
      for (let k = 0; k < units[i]; k++) out[u++] = v;
    }
    return { levels: out, paragraphLevel };
  }

  reorderVisual(levels: BidiLevels, start: number, end: number): number[] {
    return reorderByLevels(levels, start, end);
  }

  getMirror(codePoint: number): number | null {
    return mirrorGlyph(codePoint);
  }
}

// Weak branding certifies only factory-created, unmodified native algorithms;
// forwarding adapters/proxies and copied methods do not inherit this authority.
const nativeEngines = new WeakSet<BidiEngine>();
const nativeComputeLevels = Uax9Engine.prototype.computeLevels;
const nativeReorderVisual = Uax9Engine.prototype.reorderVisual;

export function isUnmodifiedNativeBidiEngine(engine: BidiEngine): boolean {
  return nativeEngines.has(engine) && engine.computeLevels === nativeComputeLevels
    && engine.reorderVisual === nativeReorderVisual;
}

export function createUax9Engine(): BidiEngine {
  const engine = new Uax9Engine();
  nativeEngines.add(engine);
  return engine;
}
