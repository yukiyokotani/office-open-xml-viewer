import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import type { ReferenceFontMetricProfile } from '@silurus/ooxml-core';
import referenceData from '../../core/src/fonts/reference-font-metrics-data.json';
import { OPEN_FONT_REFERENCE_PROFILES } from '../../core/src/fonts/reference-font-metrics-open.js';
// @ts-expect-error -- untyped repository script (ESM .mjs)
import { buildOfficeFaceWeightIndexSource, OUTPUT } from '../../../scripts/generate-pptx-face-weight-index.mjs';
import { OBSERVED_SUBSTITUTES } from './office-face-name.js';
import { indexedOfficeFaceWeight } from './office-face-weight.js';
import { officeFaceWeight } from './powerpoint-line-metrics.js';
import { PptxFontPreloadAccumulator, pptxSlideOfficeFontRequests, slideOfficeFontRequests } from './font-requests.js';
import type { Slide } from './types.js';

const profiles = [
  ...referenceData.profiles as unknown as readonly ReferenceFontMetricProfile[],
  ...OPEN_FONT_REFERENCE_PROFILES,
];

/** Every catalogue name and alias, plus spellings that normalize onto them. */
function familyInputs(): string[] {
  const inputs = new Set<string>([
    '', ' ', '\t', '__proto__', 'constructor', 'toString', 'hasOwnProperty', 'n400', 'i700',
    'Not A Catalogued Face', 'serif', 'sans-serif', 'Arial,Helvetica', 'Ｃａｌｉｂｒｉ　Ｌｉｇｈｔ',
    'CALİBRİ', 'calibrı light',
  ]);
  for (const key of Object.keys(OBSERVED_SUBSTITUTES)) {
    inputs.add(key);
    inputs.add(key.toUpperCase());
    inputs.add(` ${key} `);
  }
  for (const profile of profiles) {
    for (const alias of [profile.family, ...profile.aliases]) {
      inputs.add(alias);
      inputs.add(alias.toUpperCase());
      inputs.add(alias.toLowerCase());
      inputs.add(`  ${alias}\t`);
      inputs.add(alias.replace(/ /g, '  '));
      inputs.add(alias.normalize('NFD'));
      inputs.add(`${alias}x`);
    }
  }
  return [...inputs];
}

describe('PPTX office face weight index', () => {
  it('is the committed output of its generator', () => {
    expect(buildOfficeFaceWeightIndexSource()).toBe(readFileSync(OUTPUT, 'utf8'));
  });

  it('selects the catalogue weight for every catalogue name and style', () => {
    const inputs = familyInputs();
    expect(inputs.length).toBeGreaterThan(profiles.length);
    let compared = 0;
    let singleCut = 0;
    for (const family of inputs) {
      for (const bold of [false, true]) {
        for (const italic of [false, true]) {
          const expected = officeFaceWeight(family, bold, italic);
          const actual = indexedOfficeFaceWeight(family, bold, italic);
          if (actual !== expected) {
            throw new Error(`${JSON.stringify(family)} bold=${bold} italic=${italic}: ${actual} != ${expected}`);
          }
          if (expected !== (bold ? 700 : 400)) singleCut += 1;
          compared += 1;
        }
      }
    }
    expect(compared).toBe(inputs.length * 4);
    // The comparison exercises names whose single cut differs from the request.
    expect(singleCut).toBeGreaterThan(1000);
  });

  it('demands the same family names as the catalogue weights', () => {
    // Same-name families that differ only in capitalisation are one request
    // when their tuples share a weight, so the weight decides which spellings
    // the demand keeps.
    const runs: { fontFamily: string; bold: boolean; italic: boolean }[] = [];
    for (const family of familyInputs()) {
      for (const bold of [false, true]) {
        for (const italic of [false, true]) runs.push({ fontFamily: family, bold, italic });
      }
    }
    const slide = { elements: [{ type: 'shape', textBody: { paragraphs: [{
      bullet: { type: 'none' },
      runs: runs.map((run) => ({ type: 'text', text: 'x', ...run })),
    }] } }] } as unknown as Slide;
    // The exact face tuples, spelling and order included.
    const requests = pptxSlideOfficeFontRequests(slide, null, null);
    expect(slideOfficeFontRequests(slide, null, null, indexedOfficeFaceWeight)).toEqual(requests);
    expect(requests.length).toBeGreaterThan(familyInputs().length);
    const demand = (faceWeight: typeof officeFaceWeight) => {
      const accumulator = new PptxFontPreloadAccumulator(
        null, null, undefined, undefined, undefined, undefined, faceWeight);
      accumulator.addSlide(slide);
      return { names: accumulator.names(), facts: accumulator.demandFacts() };
    };
    const expected = demand(officeFaceWeight);
    expect(demand(indexedOfficeFaceWeight)).toEqual(expected);
    // Spellings that share a lower-case name and tuple collapse to one request,
    // so the grouping is exercised.
    const tuples = new Set(requests.filter((request) => !request.presenceOnly)
      .map((request) => `${request.family.toLowerCase()}:${request.weight}:${request.style}`));
    expect(tuples.size).toBe(requests.filter((request) => !request.presenceOnly).length);
    expect(tuples.size).toBeLessThan(runs.length);
  });
});
