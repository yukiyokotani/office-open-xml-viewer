import { describe, expect, it } from 'vitest';
import {
  DEFAULT_DOCX_COMPAT_PROFILE,
  buildDocxCompatProfile,
  docxCompatProfileKey,
  type DocxCompatSettingsInput,
} from './compat-profile.js';
import { wordContinuousSectionMayShareFootnotePage } from './body-pagination-compatibility.js';
import { wordCjkClosingMarkHang, wordLatinInterwordXavgFloorApplies } from './line-compatibility.js';

// Modes Word writes (11, 12, 14, 15), an omitted mode, and out-of-range values
// that a document may still author.
const MODES = [undefined, 0, 9, 11, 12, 13, 14, 15, 16] as const;
const INDENT_FLAGS = ['doNotSuppressIndentation', 'ww11IndentRules',
  'doNotUseIndentAsNumberingTabStop', 'noTabHangInd'] as const;

// The gate expressions each layout site evaluated before the profile existed;
// the profile must reproduce them for every input.
function previousGates(s: DocxCompatSettingsInput) {
  const mode = s.compatibilityMode;
  return {
    pageWideFootnoteArea: mode === 15,
    pageWideEndnoteArea: mode === 15,
    perColumnFootnoteBands: s.__footnoteColumnBands === true || (mode !== undefined && mode < 15),
    continuousSectionKeepsFootnotePage: mode === 15,
    footnoteLayoutLikeWW8: s.footnoteLayoutLikeWW8 === true,
    pageBreakBeforeSuppressesTopSpacing: mode === 15,
    tableOriginMeasuredMode: mode === 11 || mode === 12 || mode === 14 || mode === 15,
    tableOriginPlacementMode: mode === undefined || mode === 11 || mode === 12 || mode === 14 || mode === 15,
    tableMode15Geometry: mode === 15,
    autofitLeadingIndentBandMode: mode === undefined || mode === 14 || mode === 15,
    autofitOuterMarginsHangOutsideBand: mode === undefined || mode === 14,
    numberingSpaceSuffixArialMode: mode === undefined || mode === 14 || mode === 15,
    textMarkerKeepsBodyDescent: mode == null || mode === 12 || mode === 14 || mode === 15,
    textMarkerFirstLineUnionMode15: mode === 15,
    numberedWindowFrame: (mode === undefined || mode === 12 || mode === 14)
      && s.doNotSuppressIndentation !== true && s.ww11IndentRules !== true
      && s.doNotUseIndentAsNumberingTabStop !== true && s.noTabHangInd !== true,
    specifiedTextLineBoxMode: mode === 14 || mode === 15,
    specifiedTextLineBoxMode15: mode === 15,
    inlinePictureAutoLeading: mode === 15,
    lineEndJustificationBeforeMode15: mode !== 15,
    thaiDistributeProportionalSpread: (mode ?? 0) < 15,
    kashidaLineFitMode15: mode === 15,
    justifiedInterwordCompression: mode === 15,
    preservePre15AnchorLine: mode !== 15,
    mode14TightAnchor: mode === 14,
    mode14ColumnLineStartOrigin: mode === 14,
    throughWrapStepsLines: mode === 15,
    tabDisplacedReadingFrame: mode === undefined || mode < 15,
    compressedSpaceLineFitMode: mode !== undefined && mode < 15,
    firstOwnerClosingPairAdvance: ([undefined, 11, 12, 14, 15] as (number | undefined)[]).includes(mode),
    gridRightIndentPitchAlignment: mode === undefined || mode < 15,
    paragraphOffJapaneseLineEdges: mode === undefined || mode === 11 || mode === 12 || mode === 14 || mode === 15,
    paragraphOffThaiLineEdges: !(mode !== undefined && mode !== 11 && mode !== 12 && mode !== 14),
    adjustLineHeightInTable: s.adjustLineHeightInTable === true,
    useFeLayout: s.useFeLayout === true,
    balanceSingleByteDoubleByteWidth: s.balanceSingleByteDoubleByteWidth === true,
    lineWrapLikeWord6: s.lineWrapLikeWord6 === true,
    enableOpenTypeFeatures: s.enableOpenTypeFeatures === true,
    ulTrailSpace: s.ulTrailSpace === true,
    doNotExpandShiftReturn: s.doNotExpandShiftReturn === true,
    doNotSuppressIndentation: s.doNotSuppressIndentation === true,
    ww11IndentRules: s.ww11IndentRules === true,
    doNotUseIndentAsNumberingTabStop: s.doNotUseIndentAsNumberingTabStop === true,
    noTabHangInd: s.noTabHangInd === true,
  };
}

function previousXavgFloor(s: DocxCompatSettingsInput, hasExplicitKerningThreshold: boolean): boolean {
  const mode = s.compatibilityMode;
  if (mode !== undefined && mode < 15) return mode === 14 || s.enableOpenTypeFeatures !== true;
  return s.lineWrapLikeWord6 !== true && s.enableOpenTypeFeatures !== true && !hasExplicitKerningThreshold;
}

function previousClosingMarkHang(
  alignment: string, insideTableCell: boolean, paragraphKinsokuOff: boolean,
  mode: number | undefined, documentKinsokuEnabled: boolean, bounded: boolean,
) {
  if (bounded && paragraphKinsokuOff && documentKinsokuEnabled) {
    if (mode === 15) return 'none';
    if ((mode === undefined || mode === 11 || mode === 12 || mode === 14)
      && (alignment === 'left' || alignment === 'both' || alignment === 'justify' || alignment === 'distribute')) {
      return 'one-from-fitting-prefix';
    }
  }
  const measuredOnProfile = mode === undefined || mode === 11 || mode === 12 || mode === 14 || mode === 15;
  if (documentKinsokuEnabled && !paragraphKinsokuOff && measuredOnProfile
    && (alignment === 'left' || alignment === 'start' || alignment === 'center'
      || alignment === 'right' || alignment === 'end')) {
    return insideTableCell || mode === 15 ? 'none' : 'one';
  }
  switch (alignment) {
    case 'lowKashida': case 'mediumKashida': case 'highKashida': case 'thaiDistribute':
      return 'none';
    case 'left': case 'start': case 'both': case 'justify': case 'distribute':
      return insideTableCell ? 'none' : 'one';
    default:
      return undefined;
  }
}

function* settingsCases(): Generator<DocxCompatSettingsInput> {
  for (const compatibilityMode of MODES) {
    for (const flag of [undefined, ...INDENT_FLAGS]) {
      for (const enableOpenTypeFeatures of [undefined, false, true]) {
        for (const lineWrapLikeWord6 of [undefined, true]) {
          for (const shared of [false, true]) {
            yield {
              compatibilityMode,
              ...(flag ? { [flag]: true } : {}),
              enableOpenTypeFeatures,
              lineWrapLikeWord6,
              ...(shared ? {
                footnoteLayoutLikeWW8: true, __footnoteColumnBands: true, adjustLineHeightInTable: true,
                useFeLayout: true, balanceSingleByteDoubleByteWidth: true, ulTrailSpace: true,
                doNotExpandShiftReturn: true,
              } : {}),
            };
          }
        }
      }
    }
  }
}

describe('DocxCompatProfile', () => {
  it('reproduces every previous gate expression', () => {
    for (const settings of settingsCases()) {
      const profile = buildDocxCompatProfile(settings);
      const { latinInterwordXavgFloor: _xavg, cjkClosingMarkHangProfile: _hang, ...booleans } = profile;
      expect(booleans, JSON.stringify(settings)).toEqual(previousGates(settings));
      expect(Object.isFrozen(profile)).toBe(true);
      for (const explicit of [false, true]) {
        expect(wordLatinInterwordXavgFloorApplies(profile, explicit), JSON.stringify(settings))
          .toBe(previousXavgFloor(settings, explicit));
      }
      for (const hasColumnBands of [false, true]) {
        expect(wordContinuousSectionMayShareFootnotePage(profile, hasColumnBands))
          .toBe(settings.compatibilityMode === 15
            || (settings.footnoteLayoutLikeWW8 === true && !hasColumnBands));
      }
    }
  });

  it('reproduces the previous closing-mark hang policy', () => {
    const alignments = ['left', 'start', 'center', 'right', 'end', 'both', 'justify', 'distribute',
      'lowKashida', 'thaiDistribute', 'other'];
    for (const mode of MODES) {
      const { cjkClosingMarkHangProfile } = buildDocxCompatProfile({ compatibilityMode: mode });
      for (const alignment of alignments) {
        for (const insideTableCell of [false, true]) {
          for (const paragraphOff of [false, true]) {
            for (const documentOn of [false, true]) {
              for (const bounded of [false, true]) {
                expect(wordCjkClosingMarkHang(alignment, true, insideTableCell, paragraphOff,
                  cjkClosingMarkHangProfile, documentOn, bounded))
                  .toBe(previousClosingMarkHang(alignment, insideTableCell, paragraphOff, mode, documentOn, bounded));
              }
            }
          }
        }
      }
    }
  });

  it('treats a document without settings as an omitted compatibilityMode', () => {
    expect(DEFAULT_DOCX_COMPAT_PROFILE).toEqual(buildDocxCompatProfile({}));
    expect(DEFAULT_DOCX_COMPAT_PROFILE).toEqual(buildDocxCompatProfile(undefined));
  });

  it('keys equal profiles equally and distinct profiles distinctly', () => {
    expect(docxCompatProfileKey(buildDocxCompatProfile({ compatibilityMode: 15 })))
      .toBe(docxCompatProfileKey(buildDocxCompatProfile({ compatibilityMode: 15 })));
    expect(docxCompatProfileKey(buildDocxCompatProfile({ compatibilityMode: 15 })))
      .not.toBe(docxCompatProfileKey(buildDocxCompatProfile({ compatibilityMode: 14 })));
    expect(docxCompatProfileKey(buildDocxCompatProfile({ ulTrailSpace: true })))
      .not.toBe(docxCompatProfileKey(DEFAULT_DOCX_COMPAT_PROFILE));
  });
});
