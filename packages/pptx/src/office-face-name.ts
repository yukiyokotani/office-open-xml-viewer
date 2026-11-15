import { ownValue } from '@silurus/ooxml-core/internal/own-key';

/**
 * Faces PowerPoint substitutes for names only macOS /System/Library/Fonts
 * provides (see `powerPointAscentShare` in powerpoint-line-metrics.ts for the
 * controls behind them). Keyed by the trimmed en-US lower-case name.
 */
export const OBSERVED_SUBSTITUTES: Readonly<Record<string, string>> = {
  palatino: 'Palatino Linotype',
  helvetica: 'Arial',
  'helvetica neue': 'Arial',
};

/** The face name PowerPoint lays a run out with: an observed substitute, or the trimmed name. */
export function officeLayoutFaceName(family: string): string {
  const trimmed = family.trim();
  return ownValue(OBSERVED_SUBSTITUTES, trimmed.toLocaleLowerCase('en-US')) ?? trimmed;
}

/** The catalogued face weight a run selects: `officeFaceWeight`
 * (powerpoint-line-metrics.ts) or its generated-index form
 * `indexedOfficeFaceWeight` (office-face-weight.ts). */
export type OfficeFaceWeight = (family: string, bold: boolean, italic: boolean) => number;
