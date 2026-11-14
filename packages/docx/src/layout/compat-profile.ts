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
 * gate it is, or `@compatSetting <reference>` for a normative settings flag
 * that layout reads directly. scripts/check-docx-compatibility-evidence.mjs
 * rejects a field without one, and a `@compatRule` that names no declared
 * rule.
 *
 * The profile is plain frozen data (booleans only), so it may be carried
 * inside snapshotted layout inputs. It is internal: it is not part of the
 * public model or options.
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
}

/** Producer (non-DOCX) layout settings that select a profile fact. */
export interface DocxCompatProducerSettings {
  /** The producer lays footnotes out per column whatever the mode. */
  readonly footnoteColumnBands?: boolean;
}

type DocxCompatSettingsInput = Pick<DocSettings, 'compatibilityMode' | 'footnoteLayoutLikeWW8'>;

/** Build the profile of one document. Pure, O(1) and allocation-light: it
 * is called once at the model boundary, not per paragraph or per pass. */
export function buildDocxCompatProfile(
  settings: DocxCompatSettingsInput | undefined,
  producer: DocxCompatProducerSettings = {},
): DocxCompatProfile {
  const mode = settings?.compatibilityMode;
  const mode15 = mode === 15;
  // An omitted compatibilityMode is neither mode 15 nor an authored earlier
  // mode: each rule above states how it treats the omitted value.
  const authoredPreMode15 = mode !== undefined && mode < 15;
  return Object.freeze({
    pageWideFootnoteArea: mode15,
    pageWideEndnoteArea: mode15,
    perColumnFootnoteBands: producer.footnoteColumnBands === true || authoredPreMode15,
    continuousSectionKeepsFootnotePage: mode15,
    footnoteLayoutLikeWW8: settings?.footnoteLayoutLikeWW8 === true,
    pageBreakBeforeSuppressesTopSpacing: mode15,
  });
}

/** Profile of a document without settings (every mode-gated fact off). Used
 * when a layout input carries no profile, as for synthetic inputs. */
export const DEFAULT_DOCX_COMPAT_PROFILE: DocxCompatProfile = buildDocxCompatProfile(undefined);
