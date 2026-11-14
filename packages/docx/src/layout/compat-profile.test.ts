import { describe, expect, it } from 'vitest';
import {
  DEFAULT_DOCX_COMPAT_PROFILE,
  buildDocxCompatProfile,
} from './compat-profile.js';
import { wordContinuousSectionMayShareFootnotePage } from './body-pagination-compatibility.js';

// Modes Word writes (11, 12, 14, 15), an omitted mode, and out-of-range values
// that a document may still author.
const MODES = [undefined, 0, 9, 11, 12, 14, 15, 16] as const;

describe('DocxCompatProfile', () => {
  it.each(MODES)('resolves the body-pagination facts of compatibilityMode %s', (mode) => {
    for (const ww8 of [undefined, false, true]) {
      for (const producer of [undefined, false, true]) {
        const profile = buildDocxCompatProfile(
          { compatibilityMode: mode, footnoteLayoutLikeWW8: ww8 },
          { footnoteColumnBands: producer },
        );
        expect(profile).toEqual({
          pageWideFootnoteArea: mode === 15,
          pageWideEndnoteArea: mode === 15,
          perColumnFootnoteBands: producer === true || (mode !== undefined && mode < 15),
          continuousSectionKeepsFootnotePage: mode === 15,
          footnoteLayoutLikeWW8: ww8 === true,
          pageBreakBeforeSuppressesTopSpacing: mode === 15,
        });
        expect(Object.isFrozen(profile)).toBe(true);
        for (const hasColumnBands of [false, true]) {
          expect(wordContinuousSectionMayShareFootnotePage(profile, hasColumnBands))
            .toBe(mode === 15 || (ww8 === true && !hasColumnBands));
        }
      }
    }
  });

  it('treats a document without settings as an omitted compatibilityMode', () => {
    expect(DEFAULT_DOCX_COMPAT_PROFILE).toEqual(buildDocxCompatProfile({}));
    expect(Object.values(DEFAULT_DOCX_COMPAT_PROFILE).every((value) => value === false)).toBe(true);
  });
});
