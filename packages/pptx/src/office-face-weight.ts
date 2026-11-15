import { normalizeReferenceFontFamilyName } from '@silurus/ooxml-core';
import { OFFICE_SINGLE_WEIGHT_FACES } from './office-face-weight-index.generated.js';
import { officeLayoutFaceName, type OfficeFaceWeight } from './office-face-name.js';

// `${n | i}${normalized name}` → the single weight every catalogue profile of
// that name and style has. Parsed on first use from the generated index.
let singleWeights: ReadonlyMap<string, number> | undefined;

function singleWeightFaces(): ReadonlyMap<string, number> {
  if (singleWeights !== undefined) return singleWeights;
  const map = new Map<string, number>();
  for (const entry of OFFICE_SINGLE_WEIGHT_FACES) {
    const [group, ...names] = entry.split('|');
    const style = group[0];
    const weight = Number(group.slice(1));
    for (const name of names) map.set(`${style}${name}`, weight);
  }
  singleWeights = map;
  return map;
}

/**
 * `officeFaceWeight` (powerpoint-line-metrics.ts) without the metrics
 * catalogue, for graphs that do not otherwise bundle it (the parse worker's
 * font-demand preflight). Same result for every input:
 * `officeFaceWeight` returns the requested weight when the name has a profile
 * of it, else the name's single weight when all its profiles of the style share
 * one, else the requested weight. So a single-weight name selects that weight
 * for both requests (when it is the requested weight the first query already
 * returns it), and any other name keeps the requested weight. The generated
 * index lists exactly the single-weight names; office-face-weight.test.ts
 * compares the two over every catalogue name.
 */
export const indexedOfficeFaceWeight: OfficeFaceWeight = (family, bold, italic) => {
  const requested = bold ? 700 : 400;
  const name = officeLayoutFaceName(family);
  if (!name) return requested;
  const key = normalizeReferenceFontFamilyName(name);
  if (!key) return requested;
  return singleWeightFaces().get(`${italic ? 'i' : 'n'}${key}`) ?? requested;
};
