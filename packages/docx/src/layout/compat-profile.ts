import type { DocSettings } from '../types.js';

/**
 * Document-level compatibility profile.
 *
 * Every document-level input of a compatibility rule (`w:compat` flags,
 * `w:compatSetting` compatibilityMode, producer layout settings) is resolved
 * here exactly once per document model into a named, immutable fact. Layout
 * sites read these facts and never re-derive them from the raw settings, so
 * this file is the single audit point for which compatibility mode or flag
 * selects which rule. A rule's description, evidence and any per-region or
 * per-paragraph condition stay with its declaration (`*-compatibility.ts`).
 *
 * Each field carries `@compatRule <rule-id>` for every rule whose document
 * gate it is, or `@compatSetting ECMA-376 [Part N ]§x.y` for a normative flag
 * that layout reads directly. scripts/check-docx-compatibility-evidence.mjs
 * rejects a field without one, and a `@compatRule` that names no declared
 * rule.
 *
 * The profile is plain frozen data (booleans and small closed string
 * unions), so it may be carried inside snapshotted layout inputs. It is
 * internal: it is not part of the public model or options. Layout code never
 * compares compatibilityMode values outside this file.
 */
export interface DocxCompatProfile {
  /**
   * Compatibility mode 15 flows one page-wide footnote area through the
   * columns of a multi-column region.
   * @compatRule word-page-wide-footnote-area
   */
  readonly pageWideFootnoteArea: boolean;
  /**
   * Compatibility mode 15 balances the final body columns and flows
   * document-end notes through the section columns.
   * @compatRule word-page-wide-endnote-area
   */
  readonly pageWideEndnoteArea: boolean;
  /**
   * Authored compatibility modes below 15, or a producer that lays notes out
   * per column, give each column of a multi-column region its own note band.
   * @compatRule word-per-column-footnote-bands
   */
  readonly perColumnFootnoteBands: boolean;
  /**
   * Compatibility mode 15 keeps the section after a continuous break on a
   * page that already holds footnotes, independent of footnoteLayoutLikeWW8.
   * @compatRule word-continuous-section-footnote-page
   */
  readonly continuousSectionKeepsFootnotePage: boolean;
  /**
   * `w:compat/w:footnoteLayoutLikeWW8` was authored true.
   * @compatSetting ECMA-376 Part 4 §14.8.3.20
   */
  readonly footnoteLayoutLikeWW8: boolean;
  /**
   * Compatibility mode 15 suppresses the space before a paragraph that
   * starts a page through pageBreakBefore.
   * @compatRule word-standalone-hard-page-break-top-spacing
   */
  readonly pageBreakBeforeSuppressesTopSpacing: boolean;

  // Tables.
  /**
   * Compatibility mode 11, 12, 14 or 15 authored explicitly.
   * @compatRule word-table-origin-compatibility
   * @compatRule word-fixed-unused-leading-grid
   */
  readonly tableOriginMeasuredMode: boolean;
  /**
   * Compatibility mode omitted (placed like 14) or 11, 12, 14 or 15.
   * @compatRule word-table-origin-compatibility
   * @compatRule word-table-nonleading-indent-scope
   * @compatRule word-table-margin-hang-scope
   */
  readonly tableOriginPlacementMode: boolean;
  /**
   * Compatibility mode 15 table origin and spacing geometry (no legacy hang).
   * @compatRule word-table-origin-compatibility
   * @compatRule word-table-cell-spacing-geometry
   * @compatRule word-table-margin-hang-scope
   */
  readonly tableMode15Geometry: boolean;
  /**
   * Compatibility mode omitted, 14 or 15.
   * @compatRule word-autofit-leading-indent-band
   */
  readonly autofitLeadingIndentBandMode: boolean;
  /**
   * Compatibility mode omitted or 14: nonleading rows keep the leading band
   * width and both outer cell margins hang outside it.
   * @compatRule word-autofit-outer-cell-margin-band
   */
  readonly autofitOuterMarginsHangOutsideBand: boolean;

  // Numbering.
  /**
   * Compatibility mode omitted, 14 or 15.
   * @compatRule word-numbering-space-suffix-arial
   */
  readonly numberingSpaceSuffixArialMode: boolean;
  /**
   * Compatibility mode omitted, 12, 14 or 15.
   * @compatRule word-numbering-marker-first-line-union
   */
  readonly textMarkerKeepsBodyDescent: boolean;
  /**
   * Compatibility mode 15 marker union eligibility (floats, atLeast).
   * @compatRule word-numbering-marker-first-line-union
   */
  readonly textMarkerFirstLineUnionMode15: boolean;
  /**
   * Compatibility mode omitted, 12 or 14 with none of doNotSuppressIndentation,
   * ww11IndentRules, doNotUseIndentAsNumberingTabStop or noTabHangInd true.
   * @compatRule word-numbered-window-frame
   */
  readonly numberedWindowFrame: boolean;

  // Paragraph line box and anchors.
  /**
   * Compatibility mode 14 or 15.
   * @compatRule word-specified-text-line-box
   */
  readonly specifiedTextLineBoxMode: boolean;
  /**
   * Compatibility mode 15 extensions of the specified line box (mixed faces,
   * numbered lines, mark-only marker ascent, exact/atLeast ruby lines).
   * @compatRule word-specified-text-line-box
   */
  readonly specifiedTextLineBoxMode15: boolean;
  /**
   * Compatibility mode 15.
   * @compatRule word-inline-picture-auto-leading
   */
  readonly inlinePictureAutoLeading: boolean;
  /**
   * Compatibility mode other than 15 (including omitted).
   * @compatRule word-line-end-justification
   */
  readonly lineEndJustificationBeforeMode15: boolean;
  /**
   * Compatibility mode omitted or below 15.
   * @compatRule word-thai-distribute-proportional-spread
   */
  readonly thaiDistributeProportionalSpread: boolean;
  /**
   * Compatibility mode 15 kashida reserve and source-space charge.
   * @compatRule word-kashida-line-fit
   */
  readonly kashidaLineFitMode15: boolean;
  /**
   * Compatibility mode 15.
   * @compatRule word-justified-interword-compression
   */
  readonly justifiedInterwordCompression: boolean;
  /**
   * Compatibility mode other than 15 (including omitted).
   * @compatRule word-overlong-head-after-anchor
   */
  readonly preservePre15AnchorLine: boolean;
  /**
   * Compatibility mode 14.
   * @compatRule word-mode14-tight-anchor-line-rewrap
   * @compatRule word-mode14-tight-anchor-top-touch
   */
  readonly mode14TightAnchor: boolean;
  /**
   * Compatibility mode 14.
   * @compatRule word-mode14-column-line-start-origin
   */
  readonly mode14ColumnLineStartOrigin: boolean;
  /**
   * Compatibility mode 15.
   * @compatRule word-through-wrap-line-step-advance
   */
  readonly throughWrapStepsLines: boolean;

  // Line breaking.
  /**
   * Compatibility mode omitted or below 15.
   * @compatRule word-float-tab-reading-frame
   */
  readonly tabDisplacedReadingFrame: boolean;
  /**
   * Document gate of the Latin interword floor: always, only for runs without
   * an explicit kerning threshold, or never.
   * @compatRule word-latin-interword-xavg-floor
   */
  readonly latinInterwordXavgFloor: 'always' | 'without-explicit-kerning' | 'never';
  /**
   * Compatibility mode authored below 15 (the characterSpacingControl gate
   * stays with the rule).
   * @compatRule word-compressed-space-line-fit
   */
  readonly compressedSpaceLineFitMode: boolean;
  /**
   * Compatibility mode omitted or 11, 12, 14 or 15.
   * @compatRule word-first-owner-closing-pair-advance
   */
  readonly firstOwnerClosingPairAdvance: boolean;

  // Document and paragraph context.
  /**
   * Compatibility mode omitted or below 15.
   * @compatRule word-grid-right-indent-pitch-alignment
   */
  readonly gridRightIndentPitchAlignment: boolean;
  /**
   * Closing-mark hang profile: compatibility mode 15, an earlier measured mode
   * (omitted, 11, 12, 14) or an unmeasured mode.
   * @compatRule word-justification-punctuation-hang
   * @compatRule word-paragraph-off-japanese-line-edges
   */
  readonly cjkClosingMarkHangProfile: 'mode15' | 'earlier' | 'unmeasured';
  /**
   * Compatibility mode omitted or 11, 12, 14 or 15.
   * @compatRule word-paragraph-off-japanese-line-edges
   */
  readonly paragraphOffJapaneseLineEdges: boolean;
  /**
   * Compatibility mode omitted or 11, 12 or 14 (Thai cut exception).
   * @compatRule word-paragraph-off-japanese-line-edges
   */
  readonly paragraphOffThaiLineEdges: boolean;

  // Settings flags read directly.
  /** @compatSetting ECMA-376 §17.15.3.1 */
  readonly adjustLineHeightInTable: boolean;
  /** @compatSetting ECMA-376 Part 4 §14.8.3.50 */
  readonly useFeLayout: boolean;
  /** @compatSetting ECMA-376 §17.15.3.3 */
  readonly balanceSingleByteDoubleByteWidth: boolean;
  /** @compatSetting ECMA-376 §17.15.3.31 */
  readonly lineWrapLikeWord6: boolean;
  /** @compatRule word-kern-threshold-authority */
  readonly enableOpenTypeFeatures: boolean;
  /** @compatSetting ECMA-376 §17.15.3.8 */
  readonly ulTrailSpace: boolean;
  /** @compatSetting ECMA-376 §17.15.3.5 */
  readonly doNotExpandShiftReturn: boolean;
  /** @compatSetting ECMA-376 Part 4 §14.8.3.12 */
  readonly doNotSuppressIndentation: boolean;
  /** @compatRule word-numbered-window-frame */
  readonly ww11IndentRules: boolean;
  /** @compatSetting ECMA-376 Part 4 §14.8.3.16 */
  readonly doNotUseIndentAsNumberingTabStop: boolean;
  /** @compatSetting ECMA-376 Part 4 §14.8.3.31 */
  readonly noTabHangInd: boolean;
}

/** The `w:settings` facts the profile reads, plus the producer (non-DOCX)
 * setting that lays footnotes out per column whatever the mode. */
export type DocxCompatSettingsInput = Partial<Pick<DocSettings,
  | 'compatibilityMode' | 'footnoteLayoutLikeWW8' | 'adjustLineHeightInTable' | 'useFeLayout'
  | 'balanceSingleByteDoubleByteWidth' | 'lineWrapLikeWord6' | 'enableOpenTypeFeatures'
  | 'ulTrailSpace' | 'doNotExpandShiftReturn' | 'doNotSuppressIndentation' | 'ww11IndentRules'
  | 'doNotUseIndentAsNumberingTabStop' | 'noTabHangInd'>>
  & { readonly __footnoteColumnBands?: boolean };

/** Build the profile of one document. Pure, O(1) and allocation-light: it
 * is called once per document at the model boundary (body input and layout
 * settings), never per paragraph or per pass. */
export function buildDocxCompatProfile(
  settings: DocxCompatSettingsInput | undefined,
): DocxCompatProfile {
  const mode = settings?.compatibilityMode;
  const omitted = mode === undefined;
  const mode14 = mode === 14;
  const mode15 = mode === 15;
  // An omitted compatibilityMode is neither mode 15 nor an authored earlier
  // mode: each rule above states how it treats the omitted value.
  const authoredPreMode15 = !omitted && mode < 15;
  const preMode15 = omitted || mode < 15;
  const measuredMode = mode === 11 || mode === 12 || mode14 || mode15;
  const earlierMeasured = omitted || mode === 11 || mode === 12 || mode14;
  const flag = (value: boolean | undefined) => value === true;
  const enableOpenTypeFeatures = flag(settings?.enableOpenTypeFeatures);
  const lineWrapLikeWord6 = flag(settings?.lineWrapLikeWord6);
  const doNotSuppressIndentation = flag(settings?.doNotSuppressIndentation);
  const ww11IndentRules = flag(settings?.ww11IndentRules);
  const doNotUseIndentAsNumberingTabStop = flag(settings?.doNotUseIndentAsNumberingTabStop);
  const noTabHangInd = flag(settings?.noTabHangInd);
  return Object.freeze({
    pageWideFootnoteArea: mode15,
    pageWideEndnoteArea: mode15,
    perColumnFootnoteBands: settings?.__footnoteColumnBands === true || authoredPreMode15,
    continuousSectionKeepsFootnotePage: mode15,
    footnoteLayoutLikeWW8: flag(settings?.footnoteLayoutLikeWW8),
    pageBreakBeforeSuppressesTopSpacing: mode15,
    tableOriginMeasuredMode: measuredMode,
    tableOriginPlacementMode: omitted || measuredMode,
    tableMode15Geometry: mode15,
    autofitLeadingIndentBandMode: omitted || mode14 || mode15,
    autofitOuterMarginsHangOutsideBand: omitted || mode14,
    numberingSpaceSuffixArialMode: omitted || mode14 || mode15,
    textMarkerKeepsBodyDescent: omitted || mode === 12 || mode14 || mode15,
    textMarkerFirstLineUnionMode15: mode15,
    numberedWindowFrame: (omitted || mode === 12 || mode14) && !doNotSuppressIndentation
      && !ww11IndentRules && !doNotUseIndentAsNumberingTabStop && !noTabHangInd,
    specifiedTextLineBoxMode: mode14 || mode15,
    specifiedTextLineBoxMode15: mode15,
    inlinePictureAutoLeading: mode15,
    lineEndJustificationBeforeMode15: !mode15,
    thaiDistributeProportionalSpread: preMode15,
    kashidaLineFitMode15: mode15,
    justifiedInterwordCompression: mode15,
    preservePre15AnchorLine: !mode15,
    mode14TightAnchor: mode14,
    mode14ColumnLineStartOrigin: mode14,
    throughWrapStepsLines: mode15,
    tabDisplacedReadingFrame: preMode15,
    // OpenType independence is measured in mode 14 only.
    latinInterwordXavgFloor: authoredPreMode15
      ? (mode14 || !enableOpenTypeFeatures ? 'always' : 'never')
      : (!lineWrapLikeWord6 && !enableOpenTypeFeatures ? 'without-explicit-kerning' : 'never'),
    compressedSpaceLineFitMode: authoredPreMode15,
    firstOwnerClosingPairAdvance: omitted || measuredMode,
    gridRightIndentPitchAlignment: preMode15,
    cjkClosingMarkHangProfile: mode15 ? 'mode15' : earlierMeasured ? 'earlier' : 'unmeasured',
    paragraphOffJapaneseLineEdges: omitted || measuredMode,
    paragraphOffThaiLineEdges: earlierMeasured,
    adjustLineHeightInTable: flag(settings?.adjustLineHeightInTable),
    useFeLayout: flag(settings?.useFeLayout),
    balanceSingleByteDoubleByteWidth: flag(settings?.balanceSingleByteDoubleByteWidth),
    lineWrapLikeWord6,
    enableOpenTypeFeatures,
    ulTrailSpace: flag(settings?.ulTrailSpace),
    doNotExpandShiftReturn: flag(settings?.doNotExpandShiftReturn),
    doNotSuppressIndentation,
    ww11IndentRules,
    doNotUseIndentAsNumberingTabStop,
    noTabHangInd,
  });
}

/** Profile of a document without settings (an omitted compatibilityMode and
 * no flags). Used when a layout input carries no profile, as for synthetic
 * inputs. */
export const DEFAULT_DOCX_COMPAT_PROFILE: DocxCompatProfile = buildDocxCompatProfile(undefined);

const profileKeys = new WeakMap<DocxCompatProfile, string>();

/** Stable value key of a profile for acquisition cache keys. */
export function docxCompatProfileKey(profile: DocxCompatProfile): string {
  let key = profileKeys.get(profile);
  if (key === undefined) {
    key = JSON.stringify(Object.values(profile));
    profileKeys.set(profile, key);
  }
  return key;
}

/** The profile carried by a measurement environment, or the default. */
export function compatOf(carrier: Readonly<{ compat?: DocxCompatProfile }>): DocxCompatProfile {
  return carrier.compat ?? DEFAULT_DOCX_COMPAT_PROFILE;
}
