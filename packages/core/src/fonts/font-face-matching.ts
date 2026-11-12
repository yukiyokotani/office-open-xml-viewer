// CSS font matching as the symbol-face oracle reads it: the Canvas `font`
// shorthand and CSS Fonts 4 §5.2 face selection (its own module, so formats
// that only need the descriptor predicates do not ship it).
import { parseFamilyList, type FamilyEntry } from './css-font-shorthand.js';

/** A Canvas `font` value as CSS font matching reads it (CSS Fonts 4 §2.8
 * shorthand: style, weight and stretch before the size, then the list). */
export interface CanvasFontParts {
  /** CSS Fonts 4 §2.4: oblique counts as italic for face selection here. */
  readonly style: 'normal' | 'italic';
  /** 1..1000 (§2.2; bolder/lighter relative to normal). */
  readonly weight: number;
  /** Percentage (§2.3). */
  readonly stretch: number;
  /** The family list exactly as written. */
  readonly familyList: string;
  readonly families: readonly FamilyEntry[];
}

const STRETCH_KEYWORDS: Readonly<Record<string, number>> = {
  'ultra-condensed': 50, 'extra-condensed': 62.5, condensed: 75, 'semi-condensed': 87.5, normal: 100,
  'semi-expanded': 112.5, expanded: 125, 'extra-expanded': 150, 'ultra-expanded': 200,
};

/** Parse a Canvas `font` value; null when it is not a font shorthand. */
export function parseCanvasFont(font: string): CanvasFontParts | null {
  const match = /^\s*((?:[a-z-]+|\d+(?:\.\d+)?)\s+)*?(\d*\.?\d+(?:px|pt|em|rem|%))(?:\s*\/\s*\S+)?\s+(\S[\s\S]*)$/iu.exec(font);
  if (!match) return null;
  const prefix = font.slice(0, font.indexOf(match[2]!)).trim();
  let style: CanvasFontParts['style'] = 'normal';
  let weight = 400;
  let stretch = 100;
  for (const token of prefix ? prefix.toLowerCase().split(/\s+/u) : []) {
    if (token === 'italic' || token === 'oblique') style = 'italic';
    else if (token === 'bold' || token === 'bolder') weight = 700;
    else if (token === 'lighter') weight = 100;
    else if (/^\d+(?:\.\d+)?$/u.test(token)) weight = Math.min(1000, Math.max(1, Number(token)));
    else if (token in STRETCH_KEYWORDS) stretch = STRETCH_KEYWORDS[token]!;
    else if (token !== 'normal' && token !== 'small-caps') return null;
  }
  const familyList = match[3]!.trim();
  const families = parseFamilyList(familyList);
  return families ? { style, weight, stretch, familyList, families } : null;
}

const STRETCH_PERCENT: Readonly<Record<string, number>> = {
  'ultra-condensed': 50, 'extra-condensed': 62.5, condensed: 75, 'semi-condensed': 87.5, normal: 100,
  'semi-expanded': 112.5, expanded: 125, 'extra-expanded': 150, 'ultra-expanded': 200,
};

/** A descriptor value as a [min, max] range (CSS Fonts 4 §4.3, §4.4); null
 * when not recognised. */
function descriptorRange(value: string, keywords: Readonly<Record<string, number>>, unit = ''): [number, number] | null {
  const parts = value.trim().toLowerCase().split(/\s+/u);
  const number = (part: string) => part in keywords ? keywords[part]!
    : new RegExp(`^\\d*\\.?\\d+${unit}$`, 'u').test(part) ? Number.parseFloat(part) : Number.NaN;
  const [lo, hi = parts[0]] = parts;
  const range = [number(lo ?? ''), number(hi ?? '')] as [number, number];
  if (parts.length > 2 || !range.every(Number.isFinite)) return null;
  return range[0] <= range[1] ? range : [range[1], range[0]];
}

/** CSS Fonts 4 §5.2 distance of a weight range from the desired weight:
 * 0 inside it; then the order of step 4 (400..500 look up to 500 first,
 * then down, then up; lighter look down then up; bolder up then down). */
function weightDistance([lo, hi]: [number, number], desired: number): number {
  if (lo <= desired && desired <= hi) return 0;
  const below = desired - hi;
  const above = lo - desired;
  if (desired >= 400 && desired <= 500) {
    return lo > desired && lo <= 500 ? above : hi < desired ? 1000 + below : 2000 + above;
  }
  return desired < 400 ? (hi < desired ? below : 1000 + above) : (lo > desired ? above : 1000 + below);
}

/** §5.2 step 4 for font-stretch at the desired width (normal: narrower first). */
function stretchDistance([lo, hi]: [number, number], desired: number): number {
  if (lo <= desired && desired <= hi) return 0;
  return desired <= 100
    ? (hi < desired ? desired - hi : 1000 + lo - desired)
    : (lo > desired ? lo - desired : 1000 + desired - hi);
}

/** A font-style descriptor as a slope range in degrees, as engines match it
 * (normal 0, oblique 14deg by default, italic at the italic slope 20deg,
 * CSS Fonts 4 §4.2); null when not recognised. */
function slopeRange(value: string): [number, number] | null {
  const style = value.trim().toLowerCase();
  if (style === 'normal') return [0, 0];
  if (style === 'italic') return [ITALIC_SLOPE, ITALIC_SLOPE];
  const oblique = /^oblique(?:\s+(-?\d*\.?\d+)deg(?:\s+(-?\d*\.?\d+)deg)?)?$/u.exec(style);
  if (!oblique) return null;
  const lo = oblique[1] === undefined ? 14 : Number.parseFloat(oblique[1]);
  const hi = oblique[2] === undefined ? lo : Number.parseFloat(oblique[2]);
  return lo <= hi ? [lo, hi] : [hi, lo];
}
const ITALIC_SLOPE = 20;

/** §5.2 step 4 for font-style by slope: italic (and oblique ≥ 11deg) looks at
 * steeper slopes, then shallower ones down to 0, then negative ones; normal
 * looks at slopes ≥ 0 ascending, then negative ones. */
function slopeDistance([lo, hi]: [number, number], desired: number): number {
  if (lo <= desired && desired <= hi) return 0;
  if (desired >= 11) return lo > desired ? lo - desired : hi >= 0 ? 1000 + desired - hi : 2000 + desired - hi;
  return lo > desired ? lo - desired : 1000 + desired - hi;
}

/**
 * The faces of one family that CSS font matching (CSS Fonts 4 §5.2 step 4)
 * selects for a desired width, style and weight: the faces nearest in
 * font-stretch, then among those in font-style, then in font-weight. Faces
 * with identical descriptors form one composite by unicode-range (step 5);
 * whatever their descriptors lack is synthesized (§2.7), so they paint any
 * tuple. Faces with unrecognised descriptors are left out, as CSS rejects
 * such descriptors (§4). Null when none remains, or when faces with
 * different descriptors tie (overlapping ranges): engines break such ties
 * by their own order, so the caller measures instead.
 */
export function selectMatchingFontFaces<T extends Readonly<{ weight: string; style: string; stretch?: string }>>(
  faces: readonly T[],
  desired: Readonly<{ weight: number; style: 'normal' | 'italic'; stretch?: number }>,
): T[] | null {
  let candidates = faces.flatMap((face) => {
    const weight = descriptorRange(face.weight, { normal: 400, bold: 700 });
    const stretch = descriptorRange(face.stretch ?? 'normal', STRETCH_PERCENT, '%');
    const slope = slopeRange(face.style);
    return weight && stretch && slope ? [{ face, weight, stretch, slope }] : [];
  });
  if (candidates.length === 0) return null;
  const nearest = <C>(items: C[], distance: (item: C) => number) => {
    const best = Math.min(...items.map(distance));
    return items.filter((item) => distance(item) === best);
  };
  candidates = nearest(candidates, (item) => stretchDistance(item.stretch, desired.stretch ?? 100));
  candidates = nearest(candidates, (item) => slopeDistance(item.slope, desired.style === 'italic' ? ITALIC_SLOPE : 0));
  candidates = nearest(candidates, (item) => weightDistance(item.weight, desired.weight));
  const descriptors = new Set(candidates.map((item) => [...item.stretch, ...item.slope, ...item.weight].join(',')));
  return descriptors.size === 1 ? candidates.map((item) => item.face) : null;
}
