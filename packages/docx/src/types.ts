// ===== Output JSON model (mirrors Rust types) =====

import type { MathNode, ChartModel, Duotone } from '@silurus/ooxml-core';

export interface DocxDocumentModel {
  section: SectionProps;
  body: BodyElement[];
  headers: HeadersFooters;
  footers: HeadersFooters;
  /** Theme `<a:fontScheme><a:majorFont><a:latin@typeface>` (heading face). */
  majorFont?: string;
  /** Theme `<a:fontScheme><a:minorFont><a:latin@typeface>` (body face). */
  minorFont?: string;
  /**
   * ECMA-376 §17.8.3.10 — font family classification from `word/fontTable.xml`.
   * Maps font name to `<w:family @w:val>`: "roman" | "swiss" | "modern" |
   * "script" | "decorative" | "auto". The renderer uses this as the primary
   * source for serif/sans-serif decisions (roman→serif, swiss→sans-serif,
   * modern→monospace), falling back to name-pattern matching only when the
   * entry is absent or classified as "auto".
   */
  fontFamilyClasses?: Record<string, string>;
  /**
   * ECMA-376 §17.8.3.29 — per-font pitch from `word/fontTable.xml`
   * (`<w:pitch>`, ST_Pitch §17.18.66): font name → "fixed" | "variable" |
   * "default". Present only for fonts that declare `<w:pitch>`. The renderer
   * pairs this with {@link fontFamilyClasses}: a `family="modern"` face is
   * treated as monospace ONLY when its pitch is "fixed"; "variable" /
   * "default" / absent fall through to name-pattern / CJK-sans classification
   * (§17.8.3.10 `family` classifies the design, not the pitch — issue #855).
   */
  fontFamilyPitches?: Record<string, string>;
  /** ECMA-376 §17.8.3.3-.6 — embedded fonts from `word/fontTable.xml`, resolved
   *  to their `.odttf` part paths + fontKey. The viewer de-obfuscates (§17.8.1)
   *  and registers each as a FontFace before pagination so text measures/draws
   *  with the authored typeface. */
  embeddedFonts?: EmbeddedFontRef[];
  /** ECMA-376 §17.13.5 — flat list of `<w:ins>` / `<w:del>` events in the
   *  body. Each entry carries author / date / text. The renderer marks
   *  runs inline via {@link DocxTextRun.revision}; this array is primarily for
   *  tooling (MCP, agents, change-summary panels). */
  revisions?: DocRevision[];
  /** ECMA-376 §17.13.4 — `word/comments.xml`. Each comment carries id,
   *  author, initials, date, and plain-text body. */
  comments?: DocComment[];
  /** ECMA-376 §17.11.10 — `word/footnotes.xml` (id + text). Excludes the
   *  spec-defined separator / continuation-separator entries. */
  footnotes?: DocNote[];
  /** ECMA-376 §17.11.4 — `word/endnotes.xml` (id + text). Same shape as
   *  `footnotes`. */
  endnotes?: DocNote[];
  /** ECMA-376 §17.15.1.* — document-wide compatibility / typography settings
   *  from `word/settings.xml`. Currently carries the Japanese line-breaking
   *  (kinsoku) configuration. Absent when settings.xml has no relevant
   *  elements (the renderer then uses spec defaults: kinsoku ON). */
  settings?: DocSettings;
  /** RB7 partial degradation: set when `word/document.xml` (the body part) could
   *  not be read or parsed. The document still "opens" — `body` is empty and this
   *  part-tagged error (e.g. `"word/document.xml: <detail>"`) is carried — so the
   *  viewer shows a visible placeholder page instead of throwing. Absent
   *  (`undefined`) for every healthy document. */
  parseError?: string;
}

/** ECMA-376 §17.8.3.3-.6 — one embedded font-style slot from
 *  `word/fontTable.xml`, resolved to its obfuscated part path + fontKey. */
export interface EmbeddedFontRef {
  fontName: string;
  style: 'regular' | 'bold' | 'italic' | 'boldItalic';
  partPath: string;
  fontKey: string;
}

export interface DocSettings {
  /** §17.15.1.58 `w:kinsoku` — East-Asian line-breaking toggle. `undefined`
   *  means the element is absent; the spec default is ON (treated as `true`). */
  kinsoku?: boolean;
  /** §17.15.1.60 `w:noLineBreaksBefore@w:val` — custom set of characters that
   *  cannot begin a line (行頭禁則). When present it REPLACES the application
   *  default set. Word's per-`w:lang` sets are merged into one string. */
  noLineBreaksBefore?: string;
  /** §17.15.1.59 `w:noLineBreaksAfter@w:val` — custom set of characters that
   *  cannot end a line (行末禁則). Replaces the default when present. */
  noLineBreaksAfter?: string;
  /** §22.1.2.30 `m:mathPr/m:defJc@m:val` — document-wide default math
   *  justification (ST_Jc math: left|right|center|centerGroup). `undefined`
   *  ⇒ the renderer uses the spec default `centerGroup`. */
  mathDefJc?: string;
  /** §17.15.1.25 `w:defaultTabStop@w:val` — interval (points) between automatic
   *  tab stops generated after all custom stops. `undefined` ⇒ the renderer
   *  uses the spec default of 720 twips (36pt). */
  defaultTabStop?: number;
  /** §17.15.1.18 `w:characterSpacingControl@w:val` — East Asian punctuation /
   *  character-spacing control. */
  characterSpacingControl?: string;
  /** §17.15.3.1 `w:compat/w:useFELayout` — Far East layout compatibility. */
  useFeLayout?: boolean;
  /** §17.15.3.1 `w:compat/w:balanceSingleByteDoubleByteWidth` — balance
   *  single-byte and double-byte widths for East Asian layout. */
  balanceSingleByteDoubleByteWidth?: boolean;
  /** §17.15.3.1 `w:compat/w:adjustLineHeightInTable` — apply the section
   *  document-grid line pitch to text in table cells. */
  adjustLineHeightInTable?: boolean;
}

export interface DocRevision {
  /** "insertion" | "deletion" */
  kind: 'insertion' | 'deletion' | string;
  author?: string;
  /** ISO-8601 timestamp */
  date?: string;
  text: string;
}

export interface DocComment {
  id: string;
  author?: string;
  initials?: string;
  date?: string;
  text: string;
}

export interface DocNote {
  id: string;
  /** ECMA-376 §17.11.2 / §17.11.10 — the note's block-level content
   *  (paragraphs / nested tables), parsed with the document's styles +
   *  numbering. The leading run is the `<w:footnoteRef/>` auto-number marker
   *  (carries a {@link DocxTextRun.noteRef}). Use {@link noteText} to extract
   *  the plain-text body without the marker. */
  content: BodyElement[];
}

/** Flatten a footnote/endnote's content to its plain-text body, excluding the
 *  auto-number reference marker. Convenience for data-only consumers
 *  (the renderer draws {@link DocNote.content} directly). */
export function noteText(note: DocNote): string {
  const parts: string[] = [];
  for (const el of note.content) {
    if (el.type !== 'paragraph') continue;
    let s = '';
    for (const run of (el as DocParagraph).runs) {
      if (run.type === 'text' && !(run as DocxTextRun).noteRef) {
        s += (run as DocxTextRun).text;
      }
    }
    s = s.trim();
    if (s) parts.push(s);
  }
  return parts.join(' ');
}

export interface HeadersFooters {
  default: HeaderFooter | null;
  first: HeaderFooter | null;
  even: HeaderFooter | null;
}

export interface HeaderFooter {
  body: BodyElement[];
}

/** ECMA-376 §17.6.12 `<w:pgNumType>` — a section's page-numbering settings.
 *  Mirrors the Rust `PageNumType`. Only the two attributes that change the
 *  DISPLAYED page number are carried:
 *  - `start` — the number shown on the FIRST page of the section (§17.6.12);
 *    absent ⇒ numbering continues from the previous section's highest number.
 *    Kept as a possibly-zero / possibly-negative integer; zero is valid source
 *    markup.
 *  - `fmt` — the ST_NumberFormat (§17.18.59) for the section's page numbers
 *    (decimal / upperRoman / lowerLetter / …); absent ⇒ decimal.
 *  `chapStyle`/`chapSep` (chapter-prefixed numbering) are out of scope for this
 *  pass and never surfaced. Field names match the Rust `PageNumType` serialization
 *  (`start`, `fmt`). */
export interface PageNumType {
  start?: number;
  fmt?: string;
}

/** ECMA-376 §17.6.10 `<w:pgBorders>` — page borders drawn around each page of a
 *  section. Mirrors the Rust `PageBorders`. Each edge is a CT_Border (§17.18.4);
 *  the container carries the placement globals. Absent on {@link SectionProps}
 *  (`pageBorders` undefined) ⇒ no page border (the common case). Art borders
 *  (§17.18.2 decorative-image styles) are unsupported — retained Canvas paint
 *  draws only the standard line styles (single/double/dashed/dotted/thick/…). */
export interface PageBorders {
  /** `@w:offsetFrom` (§17.18.63): "page" ⇒ each edge's `space` is from the PAGE
   *  edge; "text" (the default) ⇒ from the text margin. */
  offsetFrom: string;
  /** `@w:display` (§17.18.62): "allPages" (default) | "firstPage" |
   *  "notFirstPage" — which physical pages of the section show the border. */
  display: string;
  /** `@w:zOrder` (§17.18.64): "front" (default; over text) | "back" (under). */
  zOrder: string;
  top?: PageBorderEdge;
  bottom?: PageBorderEdge;
  left?: PageBorderEdge;
  right?: PageBorderEdge;
}

/** ECMA-376 §17.18.4 CT_Border for one edge of `<w:pgBorders>`. Mirrors the Rust
 *  `PageBorderEdge`. Same shape as a paragraph border edge. */
export interface PageBorderEdge {
  /** `@w:val` — ST_Border line style ("single" | "double" | "dashed" | …). */
  style: string;
  /** `@w:color` hex 6, or absent for "auto" (renderer defaults to black). */
  color?: string;
  /** `@w:sz` in pt (eighths of a point ÷ 8). */
  width: number;
  /** `@w:space` in pt — a POINT measure (§17.18.68, 0–31) for page borders, NOT
   *  twips — the inset from the `offsetFrom` reference. */
  space: number;
}

/** ECMA-376 §17.6.8 `<w:lnNumType>` — line numbering for a section. Mirrors the
 *  Rust `LineNumbering`. A number is drawn in the left margin of each body line
 *  whose count is a multiple of `countBy`. Absent on {@link SectionProps}
 *  (`lineNumbering` undefined) ⇒ line numbering off. */
export interface LineNumbering {
  /** `@w:countBy` — only lines whose number is a multiple of this display a
   *  number. Required (absent ⇒ the whole struct is absent per §17.6.8). */
  countBy: number;
  /** `@w:start` — the starting number after each restart. Default 1. */
  start: number;
  /** `@w:distance` in pt (twips ÷ 20) — gap between the text margin and the
   *  number glyphs. Absent ⇒ implementation-defined (renderer uses a default). */
  distance?: number;
  /** `@w:restart` (§17.18.47): "newPage" (default) | "newSection" |
   *  "continuous" — when the counter resets to `start`. */
  restart: string;
}

/** ECMA-376 §17.6.13 `<w:pgSz>` + §17.6.11 `<w:pgMar>` — a section's page
 *  geometry: page size + margins + header/footer distances (pt). Mirrors the Rust
 *  `SectionGeom`. Carried on a {@link BodyElement} `sectionBreak` arm (`geom`) so a
 *  mid-body section keeps its own page size; the FINAL section's geometry lives on
 *  {@link DocxDocumentModel.section}. Canonical layout retains the resolved page
 *  box on each physical page. `orient` is omitted because the parsed landscape
 *  geometry already carries the resolved width and height.
 *
 *  ⚠ Spread over the body-level {@link SectionProps} in `renderDocumentToCanvas`
 *  (`{ ...doc.section, ...pageGeom }`): only add per-section PAGE-BOX fields that
 *  exist on `SectionProps` with the same name and semantics — an optional field
 *  colliding with a non-geometry `SectionProps` name would silently override the
 *  body-level value the renderer promises to preserve. */
export interface SectionGeom {
  pageWidth: number;   // pt
  pageHeight: number;  // pt
  /** §17.6.11 — top/bottom MAY be negative (ST_SignedTwipsMeasure); keep the sign. */
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  headerDistance: number;
  footerDistance: number;
}

export interface SectionProps {
  pageWidth: number;   // pt
  pageHeight: number;  // pt
  // ECMA-376 §17.6.11 — top/bottom are ST_SignedTwipsMeasure (§17.18.81) and MAY be
  // negative (the body is then measured |margin| from the page edge and overlaps the
  // header/footer). Keep the SIGN here: header/footerOverflowPt need it to decide
  // overlap-vs-reserve. The renderer's bodyMarginInsetPt derives the body inset (|margin|).
  // Do NOT Math.abs at the parser or overflow sites. left/right are ST_TwipsMeasure (unsigned).
  marginTop: number;   // pt — signed (§17.6.11)
  marginRight: number;
  marginBottom: number; // pt — signed (§17.6.11)
  marginLeft: number;
  headerDistance: number;   // pt — top of page to header
  footerDistance: number;   // pt — bottom of page to footer
  titlePage: boolean;
  evenAndOddHeaders: boolean;
  /** ECMA-376 §17.6.22 ST_SectionMark — the body (final) section's `<w:type>`
   *  start type ("continuous" | "nextPage" | "oddPage" | "evenPage"). Governs
   *  how the last section begins relative to the previous one; consumed by the
   *  paginator at the boundary INTO the final section. Absent ⇒ "nextPage" (the
   *  spec default). Non-final sections carry their start type on their own
   *  SectionBreak marker. */
  sectionStart?: string | null;
  /** ECMA-376 §17.6.20 `<w:textDirection w:val>` — the section's flow direction,
   *  using the Transitional ST_TextDirection enum (Part 4 §14.11.7:
   *  `lrTb`|`tbRl`|`btLr`|`lrTbV`|`tbLrV`|`tbRlV`), NOT the Part 1 §17.18.93
   *  Strict set. Absent / `null` ⇒ "lrTb" (horizontal, left→right / top→bottom,
   *  the default). `"tbRl"` = vertical Japanese (glyphs stack top→bottom, lines
   *  advance right→left); the renderer (see `isVerticalSection`) lays the page out
   *  horizontally and rotates it +90° at paint for the vertical values
   *  (`tbRl`/`tbRlV`/`tbLrV`), keeping CJK glyphs upright and Latin sideways. Only
   *  a non-default value is emitted by the parser, so horizontal documents keep
   *  byte-identical rendering. */
  textDirection?: string | null;
  /** ECMA-376 §17.6.5 w:docGrid/@w:type — "default" | "lines" | "linesAndChars" | "snapToChars". */
  docGridType?: string | null;
  /** ECMA-376 §17.6.5 w:docGrid/@w:linePitch in pt. When docGridType is "lines" or
   *  "linesAndChars", auto line spacing multiplies against this pitch instead of
   *  the font's natural line height. */
  docGridLinePitch?: number | null;
  /** ECMA-376 §17.6.5 w:docGrid/@w:charSpace (ST_DecimalNumber, signed). The
   *  raw character-grid spacing in 1/4096ths of an em (NOT twips). When
   *  docGridType is "linesAndChars" or "snapToChars", every full-width East-
   *  Asian glyph occupies a fixed cell of width `fontSizePt + charSpace/4096` pt
   *  (negative = tighter). Absent ⇒ East-Asian glyphs keep their natural em
   *  advance. */
  docGridCharSpace?: number | null;
  /** ECMA-376 §17.6.4 `<w:cols>` — newspaper-style multi-column layout. `null`
   *  (or absent) ⇒ single full-width column (unchanged behavior). When present,
   *  body text flows top-to-bottom through `count` columns (newspaper fill);
   *  see {@link computeColumns}. */
  columns?: ColumnsSpec | null;
  /** ECMA-376 §17.6.12 `<w:pgNumType>` — the body (final) section's page-numbering
   *  settings (start / fmt). `null`/absent ⇒ numbering continues; decimal. The
   *  renderer resolves the displayed page number per physical page from this plus
   *  the per-section `SectionBreak.pageNumType` markers. */
  pageNumType?: PageNumType | null;
  /** ECMA-376 §17.6.10 `<w:pgBorders>` — page borders for this section.
   *  `null`/absent ⇒ no page border (the common case). */
  pageBorders?: PageBorders | null;
  /** ECMA-376 §17.6.8 `<w:lnNumType>` — line numbering for this section.
   *  `null`/absent ⇒ line numbering off. */
  lineNumbering?: LineNumbering | null;
  /** ECMA-376 §17.6.23 `<w:vAlign w:val>` — body vertical alignment between the
   *  top/bottom margins ("top" | "center" | "both" | "bottom"). `null`/absent ⇒
   *  "top" (body flows from the top margin unchanged). "both" (vertical
   *  justification) is parsed but rendered as "top" until distribution is
   *  implemented (see renderer note). */
  vAlign?: string | null;
}

/** ECMA-376 §17.6.4 `<w:cols>` — the section's multi-column configuration. */
export interface ColumnsSpec {
  /** `@w:num` — number of columns (>= 2 when emitted). */
  count: number;
  /** `@w:space` in pt — inter-column gap for equal-width columns (default 36pt
   *  = 720 twips per the spec). */
  spacePt: number;
  /** `@w:equalWidth` (default true) — all columns share one width + `spacePt`.
   *  When false, `cols` carries explicit per-column geometry. */
  equalWidth: boolean;
  /** `@w:sep` — draw vertical separator rules between columns. */
  sep: boolean;
  /** Per-column `<w:col>` entries (width + trailing space, pt). Empty when
   *  `equalWidth` is true. */
  cols: ColSpec[];
}

/** ECMA-376 §17.6.3 `<w:col>` — one column's width and trailing space (pt). */
export interface ColSpec {
  widthPt: number;
  spacePt: number;
}

/** ECMA-376 §17.6.4 — a single newspaper column resolved to its page-absolute
 *  left edge (`xPt`) and text width (`wPt`), in pt. Produced by `computeColumns`
 *  (in the renderer) from a section's {@link ColumnsSpec} + the page geometry. */
export interface ColumnGeom {
  xPt: number;
  wPt: number;
}

export type BodyElement =
  | { type: 'paragraph' } & DocParagraph
  | { type: 'table' } & DocTable
  | {
      type: 'pageBreak';
      parity?: 'odd' | 'even';
      /** The hard break followed visible content in the same source paragraph. */
      sameParagraphAsPrevious?: boolean;
    }
  /** ECMA-376 §17.3.1.20 `<w:br w:type="column"/>` — force the following content
   *  into the next newspaper column (or the next page's first column when
   *  already in the last column). Hoisted to the body level by the parser. */
  | { type: 'columnBreak' }
  /** ECMA-376 §17.6.x — a section boundary in the body. A `<w:sectPr>` carried in
   *  a paragraph's `pPr` (or a loose mid-body one) defines the section that ENDS
   *  here; the FINAL section's settings live on {@link DocxDocumentModel.section}.
   *  `columns` is the ENDING section's `<w:cols>` (§17.6.4; absent ⇒ single
   *  full-width column — the spec default for `@w:num` 1). `kind` is the
   *  ST_SectionMark (§17.18.79) controlling how the NEXT section starts:
   *  "continuous" (same page), "nextPage" (default; page break), "oddPage" /
   *  "evenPage" (page break + parity padding). The paginator switches its active
   *  newspaper-column geometry per section at each marker — fixing the regression
   *  where every section inherited the body-level section's columns.
   *
   *  `headers`/`footers` carry the ENDING section's resolved (§17.10.1-inherited)
   *  header/footer set, and `titlePage` its own `<w:titlePg>` flag, so the renderer
   *  can pick the active section's header/footer per page (mirroring how `columns`
   *  drives per-section column geometry). The body-level (final) section's sets
   *  live on {@link DocxDocumentModel.section}/`.headers`/`.footers` instead. */
  | {
      type: 'sectionBreak';
      kind: 'continuous' | 'nextPage' | 'oddPage' | 'evenPage' | string;
      columns?: ColumnsSpec | null;
      headers?: HeadersFooters;
      footers?: HeadersFooters;
      titlePage?: boolean;
      /** ECMA-376 §17.6.13 / §17.6.11 — this ENDING section's page geometry
       *  (size + margins). Absent when the sectPr inherits both pgSz and pgMar
       *  (the renderer then falls back to the body-level section geometry). */
      geom?: SectionGeom;
      /** ECMA-376 §17.6.12 `<w:pgNumType>` — this ENDING section's page-numbering
       *  settings (start / fmt). Absent ⇒ numbering continues; decimal. Carried
       *  separately from `geom` because a section may inherit its geometry yet
       *  still restart / re-format its page numbers. */
      pageNumType?: PageNumType | null;
      /** ECMA-376 §17.6.20 `<w:textDirection w:val>` — this ENDING section's
       *  flow direction (TRANSITIONAL ST_TextDirection, same enum and semantics
       *  as {@link SectionProps.textDirection}), so a vertical (tbRl/btLr)
       *  non-final section can coexist with a horizontal final section (issue
       *  #1000). Absent ⇒ horizontal ("lrTb" is collapsed by the parser).
       *  Carried separately from `geom` (like `pageNumType`) because a section
       *  may inherit its page geometry yet still set its own flow direction. */
      textDirection?: string | null;
    };


export interface DocParagraph {
  /**
   * Authored `w14:paraId` for the source paragraph, preserved verbatim and
   * absent when the source paragraph does not carry that identifier.
   */
  paragraphId?: string;
  /**
   * ECMA-376 §17.18.44 ST_Jc. Renderer honors left, start, center, right, end,
   * both, distribute. Other values (kashida variants, numTab, thaiDistribute)
   * are treated as start-aligned.
   */
  alignment:
    | 'left' | 'start'
    | 'center'
    | 'right' | 'end'
    | 'justify' | 'both'
    | 'distribute'
    // §17.18.44 Arabic kashida + Thai justification variants — mapped onto the
    // existing justify/distribute slack kernel (see bidi-line `resolveAlignEdge`).
    | 'lowKashida' | 'mediumKashida' | 'highKashida'
    | 'thaiDistribute'
    | string;
  indentLeft: number;   // pt
  indentRight: number;  // pt
  indentFirst: number;  // pt
  spaceBefore: number;  // pt
  spaceAfter: number;   // pt
  lineSpacing: LineSpacing | null;
  numbering: NumberingInfo | null;
  tabStops: TabStop[];
  runs: DocRun[];
  /**
   * ECMA-376 §17.13.6.2 `<w:bookmarkStart w:name>` — names of the bookmarks that
   * start within (or at the head of) this paragraph, in document order. A
   * `<w:hyperlink w:anchor="X">` internal link (§17.16.23) targets the paragraph
   * whose `bookmarks` contains `"X"`; {@link buildBookmarkPageMap} turns these
   * into a `bookmarkName → pageIndex` map after pagination. Absent (`undefined`)
   * for the common paragraph that anchors nothing.
   */
  bookmarks?: string[];
  /** Paragraph background hex color (w:shd fill) */
  shading?: string | null;
  /** Force a page break before this paragraph (w:pageBreakBefore) */
  pageBreakBefore?: boolean;
  /** ECMA-376 §17.3.1.9 `<w:contextualSpacing>` — between adjacent same-style
   *  paragraphs, `word-contextual-spacing-per-side` drops the toggling
   *  paragraph's own contribution to the collapsed inter-paragraph gap. */
  contextualSpacing?: boolean;
  /** Keep paragraph on same page as the next paragraph (w:keepNext) */
  keepNext?: boolean;
  /** Keep all lines of this paragraph on the same page (w:keepLines) */
  keepLines?: boolean;
  /** ECMA-376 §17.3.1.29 + §17.3.2.41 — the paragraph MARK's resolved `w:vanish`
   *  (hidden text). An inkless paragraph whose mark is vanished collapses to zero
   *  height in the normal/print view (hidden-text off), the same way the parser
   *  strips hidden runs; the paginator drops it whole. Absent = mark is visible. */
  markVanish?: boolean;
  /** Widow/orphan control (w:widowControl). ECMA-376 default is true. */
  widowControl?: boolean;
  /** ECMA-376 §17.3.1.21 `<w:overflowPunct>` — permit one trailing
   *  punctuation character beyond paragraph indents/margins. Omission is true. */
  overflowPunct?: boolean;
  /** Paragraph borders (w:pBdr) */
  borders?: ParagraphBorders | null;
  /** Style ID of the applied paragraph style */
  styleId?: string | null;
  /** Default font size (pt) inherited from style + direct pPr/rPr. Falls back to 10pt. */
  defaultFontSize?: number;
  /** Default font family resolved from the style chain. Used to size empty
   *  paragraphs (no runs) with the intended font's line metrics. */
  defaultFontFamily?: string | null;
  /** Default East Asian font family resolved from the style chain. Empty /
   *  anchor-only paragraph marks in East Asian documents use this axis for line
   *  metrics instead of the ASCII fallback. */
  defaultFontFamilyEastAsia?: string | null;
  /** ECMA-376 §17.3.1.29 — the paragraph mark run's resolved `w:color` (direct
   *  pPr/rPr → pStyle chain → docDefaults; hex 6 without `#`, lowercased; an
   *  explicit `auto` surfaces as absent, §17.3.2.6). The numbering level rPr
   *  (§17.9.24) layers over the mark's run properties, so the renderer uses
   *  this as the marker-color fallback when {@link NumberingInfo.color} is
   *  absent. */
  paragraphMarkColor?: string | null;
  /**
   * ECMA-376 §17.3.1.6 `<w:bidi>` — right-to-left paragraph. `true` = RTL,
   * `false` = explicitly LTR, absent = unspecified (inherit). The renderer uses
   * this as the paragraph base direction: it seeds the UAX#9 reordering pass
   * (`computeLineVisualOrder`), swaps the left/right indents, resolves the
   * `w:jc` start/end edges (`resolveAlignEdge`), and lays out lines from the
   * right.
   */
  bidi?: boolean;
  /**
   * ECMA-376 §17.3.1.32 `<w:snapToGrid>` — when `false`, this paragraph opts out
   * of the section's document grid (`w:docGrid`): its lines use natural font
   * metrics / the line-spacing multiplier directly instead of snapping to the
   * grid pitch. `undefined` = inherit (default on). Set on Word's "Footnote
   * Text" style, so footnote bodies use compact natural line height.
   */
  snapToGrid?: boolean;
  /**
   * ECMA-376 §17.3.1.11 `<w:framePr>` — text-frame / drop-cap properties.
   * Present ⇒ this paragraph is part of a text frame; the renderer positions it
   * as a frame (drop cap or generic frame) and registers a wrap exclusion so
   * following body text flows around it. Absent ⇒ ordinary in-flow paragraph.
   */
  framePr?: FramePr;
}

/**
 * ECMA-376 §17.3.1.11 `<w:framePr>` — text-frame / drop-cap properties.
 *
 * Lengths are pt (parser converts from twips). Per the spec, `x`/`y` are
 * ignored when `xAlign`/`yAlign` are set, and for a drop cap `y`/`yAlign` are
 * ignored entirely while `lines` drives the height. `w`/`h`/`x`/`y` are
 * `undefined` when the attribute was absent (distinct from an explicit 0).
 */
export interface FramePr {
  /** ST_DropCap (§17.18.20): 'none' | 'drop' | 'margin'. Default 'none'. */
  dropCap: 'none' | 'drop' | 'margin' | string;
  /** §17.3.1.11 `lines` — drop-cap vertical height in anchor lines. Default 1. */
  lines: number;
  /** ST_Wrap (§17.18.104): 'around'|'auto'|'none'|'notBeside'|'through'|'tight'. Default 'around'. */
  wrap: 'around' | 'auto' | 'none' | 'notBeside' | 'through' | 'tight' | string;
  /** ST_HAnchor (§17.18.35): 'text'(=column) | 'margin' | 'page'. Default 'page'. */
  hAnchor: 'text' | 'margin' | 'page' | string;
  /** ST_VAnchor (§17.18.100): 'text' | 'margin' | 'page'. Default 'page'. */
  vAnchor: 'text' | 'margin' | 'page' | string;
  /** ST_HeightRule (§17.18.37): 'auto' | 'atLeast' | 'exact'. Default 'auto'. */
  hRule: 'auto' | 'atLeast' | 'exact' | string;
  /** hSpace — min wrap padding L/R when wrap='around' (pt). Default 0. */
  hSpace: number;
  /** vSpace — min wrap padding top/bottom (pt). Default 0. */
  vSpace: number;
  /** w — exact frame width (pt). Absent ⇒ auto (max content line width). */
  w?: number;
  /** h — frame height (pt). Meaning gated by hRule. Absent ⇒ auto. */
  h?: number;
  /** x — absolute horizontal offset from hAnchor (pt). Ignored when xAlign set. */
  x?: number;
  /** y — absolute vertical offset from vAnchor (pt). Ignored when yAlign set / drop cap. */
  y?: number;
  /** ST_XAlign (§22.9.2.18): 'left'|'center'|'right'|'inside'|'outside'. Supersedes x. */
  xAlign?: 'left' | 'center' | 'right' | 'inside' | 'outside' | string;
  /** ST_YAlign (§22.9.2.20): 'inline'|'top'|'center'|'bottom'|'inside'|'outside'. Supersedes y. */
  yAlign?: 'inline' | 'top' | 'center' | 'bottom' | 'inside' | 'outside' | string;
}

export interface ParagraphBorders {
  top: ParaBorderEdge | null;
  bottom: ParaBorderEdge | null;
  left: ParaBorderEdge | null;
  right: ParaBorderEdge | null;
  between: ParaBorderEdge | null;
}

export interface ParaBorderEdge {
  style: string;
  color: string | null;
  /** pt (sz / 8) */
  width: number;
  /** pt spacing between border and text */
  space: number;
}

/** ECMA-376 §17.3.2.4 `<w:bdr>` — a run-level border drawn as a box around the
 *  run's text. Parallel to {@link ParaBorderEdge} but applies per run. */
export interface DocxRunBorder {
  /** "single" | "double" | "dashed" | ... (w:bdr/@w:val) */
  style: string;
  /** hex 6-char, or null for automatic (renderer falls back to text color) */
  color?: string | null;
  /** pt (sz / 8) */
  width: number;
  /** pt spacing between the border and the run text (w:space) */
  space: number;
}

export interface TabStop {
  /** tab stop position in pt (from the left of paragraph content area) */
  pos: number;
  /** ECMA-376 ST_TabJc. `num` is the list tab between a numbering marker and
   * paragraph contents; start/end are logical-direction aliases. */
  alignment: 'left' | 'start' | 'center' | 'right' | 'end' | 'decimal' | 'bar' | 'clear' | 'num';
  leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

export interface LineSpacing {
  value: number;   // multiplier (auto) or pt (exact/atLeast)
  rule: 'auto' | 'exact' | 'atLeast';
  /** True when `w:spacing/@w:line` was set on the paragraph's own pPr or on a
   *  named style (not inherited solely from docDefault). Per ECMA-376 §17.6.5,
   *  an inherited-only paragraph in a docGrid section snaps to one grid pitch
   *  per line, ignoring the multiplier. Defaults to false on JSON parse. */
  explicit?: boolean;
}

export interface NumberingInfo {
  numId: number;
  level: number;
  format: string;
  text: string;       // resolved bullet text, e.g. "1." or "•"
  indentLeft: number; // pt
  tab: number;        // pt
  /** ECMA-376 §17.9.28 `<w:suff>` — "tab" (default) | "space" | "nothing".
   *  Where body text starts after the marker on the first line. */
  suff: string;
  /** ECMA-376 §17.9.8 `<w:lvlJc>` — marker justification: "left" (default) |
   *  "right" (period-aligned numerals: marker RIGHT edge at the hanging-indent
   *  position) | "center". The renderer offsets the marker draw accordingly.
   *  Always emitted by the parser; optional here so hand-built fixtures may omit
   *  it (the renderer treats absent as "left"). */
  jc?: string;
  /** ECMA-376 §17.3.2.26 ascii axis for the marker glyph, resolved through the
   *  level's `rPr` (§17.9.6) merged over the paragraph's run formatting. The
   *  renderer draws Latin marker chars (e.g. a decimal "1") with this family, so
   *  a heading whose ascii=Times renders its auto-number in Times (serif) even
   *  when eastAsia=Gothic. Absent ⇒ the renderer falls back to its default. */
  fontFamily?: string | null;
  /** ECMA-376 §17.3.2.26 eastAsia axis for the marker glyph (same resolution as
   *  {@link NumberingInfo.fontFamily}). The renderer draws CJK marker chars with
   *  this family. Absent ⇒ the renderer falls back to
   *  {@link NumberingInfo.fontFamily}. */
  fontFamilyEastAsia?: string | null;
  /** ECMA-376 §17.9.24 — the numbering level rPr's `w:color` (hex 6 without
   *  `#`, lowercased). Colors the marker glyph only, never the paragraph's
   *  runs. Absent ⇒ the renderer falls back to
   *  {@link DocParagraph.paragraphMarkColor} (§17.3.1.29 — the level rPr layers
   *  over the paragraph mark's run properties) and finally to its
   *  default ink. An explicit `w:val="auto"` is absent here + {@link colorAuto}. */
  color?: string | null;
  /** ECMA-376 §17.3.2.6 / ST_HexColorAuto (§17.18.39) — true when the level
   *  rPr carries an EXPLICIT `w:color w:val="auto"`. Auto names no concrete
   *  color but is not "unset": it breaks the paragraph-mark fallback, so the
   *  marker draws the automatic (default) ink instead of
   *  {@link DocParagraph.paragraphMarkColor}. */
  colorAuto?: boolean;
  /** ECMA-376 §17.9.9/§17.9.20 — when the level uses a `<w:lvlPicBulletId>`,
   *  the marker is this image (zip path, e.g. `word/media/image1.gif`), drawn in
   *  place of {@link NumberingInfo.text}. Absent ⇒ ordinary text/glyph marker. */
  picBulletImagePath?: string;
  /** MIME type of {@link NumberingInfo.picBulletImagePath} (e.g. `image/gif`). */
  picBulletMimeType?: string;
  /** Picture-bullet marker width in pt (from the `<v:shape style="width">`). */
  picBulletWidthPt?: number;
  /** Picture-bullet marker height in pt (from the `<v:shape style="height">`). */
  picBulletHeightPt?: number;
}

export type DocRun =
  | { type: 'text' } & DocxTextRun
  | { type: 'anchorHost' } & AnchorHostMetrics
  | { type: 'image' } & ImageRun
  | { type: 'chart' } & ChartRun
  | { type: 'break'; breakType: 'line' | 'page' | 'column' }
  | { type: 'field' } & FieldRun
  | { type: 'shape' } & ShapeRun
  | { type: 'math'; nodes: MathNode[]; display: boolean; fontSize: number; jc?: string }
  | { type: 'ptab' } & PTabRun;

/** ECMA-376 §21.2 — a DrawingML chart embedded in the run flow via
 *  `<w:drawing><wp:inline|wp:anchor>…<a:graphicData uri=".../chart"><c:chart r:id>`.
 *  Mirrors the Rust `ChartRun`. `chart` is the shared {@link ChartModel} the
 *  core `renderChart` consumes (identical to what pptx/xlsx pass), so a docx
 *  chart draws at the same quality through the same code path. `widthPt`/
 *  `heightPt` are the `<wp:extent>` natural size. An inline chart flows as an
 *  inline box of that size; anchor acquisition retains an anchored chart
 *  (§20.4.2.3) at its resolved page box while contributing wrap exclusions
 *  when required — all paint paths use `renderChart`. */
export interface ChartRun {
  chart: ChartModel;
  widthPt: number;
  heightPt: number;
  /** true = `<wp:anchor>` (absolute page position, drawn by the anchor path);
   *  false = `<wp:inline>` (flows with text). */
  anchor: boolean;
  anchorXPt?: number;
  anchorYPt?: number;
  anchorXFromMargin?: boolean;
  anchorYFromPara?: boolean;
  /**
   * Wrap mode for anchored charts (ECMA-376 §20.4.2.16/.17):
   *   "square" | "topAndBottom" | "none" | "tight" | "through"
   * Inline charts and undetermined cases leave this undefined. The renderer
   * treats "tight" and "through" as "square", matching anchored images.
   */
  wrapMode?: string;
  /** Padding top (pt). Anchor-only (ECMA-376 §20.4.2.16/.17). */
  distTop?: number;
  /** Padding bottom (pt). Anchor-only (ECMA-376 §20.4.2.16/.17). */
  distBottom?: number;
  /** Padding left (pt). Anchor-only (ECMA-376 §20.4.2.16/.17). */
  distLeft?: number;
  /** Padding right (pt). Anchor-only (ECMA-376 §20.4.2.16/.17). */
  distRight?: number;
  /** wrapText attribute: "bothSides" | "left" | "right" | "largest". */
  wrapSide?: string;
  /**
   * ECMA-376 §20.4.2.3 `wp:anchor/@allowOverlap`. The parser omits this
   * field when true, so renderers must read it as `allowOverlap ?? true`.
   */
  allowOverlap?: boolean;
  /** ECMA-376 §20.4.3.1 wp:align horizontal: "left" | "center" | "right" |
   *  "inside" | "outside". */
  anchorXAlign?: string | null;
  /** Vertical equivalent of anchorXAlign: "top" | "center" | "bottom". */
  anchorYAlign?: string | null;
  /**
   * ECMA-376 §20.4.3.2 `<wp:positionH/@relativeFrom>` / §20.4.3.5
   * `<wp:positionV/@relativeFrom>` — the raw anchor placement containers.
   */
  anchorXRelativeFrom?: string | null;
  anchorYRelativeFrom?: string | null;
}

/** ECMA-376 §17.3.3.23 `<w:ptab>` — an absolute-position tab. Advances to a
 *  position derived from {@link PTabRun.alignment} and {@link PTabRun.relativeTo},
 *  independent of the paragraph's custom tab stops / default-tab interval. */
export interface PTabRun {
  /** ST_PTabAlignment (§17.18.71): where on the line the tab lands, and how the
   *  following text aligns to it. */
  alignment: 'left' | 'center' | 'right';
  /** ST_PTabRelativeTo (§17.18.73): the base the position is measured from —
   *  the text margins or the paragraph indents. */
  relativeTo: 'margin' | 'indent';
  /** ST_PTabLeader (§17.18.72): the character repeated to fill the tab gap. */
  leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'middleDot';
  /** Resolved run font size (pt) — matches the surrounding text's leader/gap. */
  fontSize: number;
}

export type PathCmd =
  | { cmd: 'moveTo'; x: number; y: number }
  | { cmd: 'lineTo'; x: number; y: number }
  | { cmd: 'cubicBezTo'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { cmd: 'arcTo'; wr: number; hr: number; stAng: number; swAng: number }
  | { cmd: 'close' };

/** Resolved formatting of the WordprocessingML anchor character that hosts a
 * floating drawing. The drawing has zero inline advance, but these metrics still
 * participate in the containing line's height and document-grid allocation. */
export interface AnchorHostMetrics {
  /** Effective `<w:sz>` in points. */
  fontSize: number;
  /** Resolved ascii/hAnsi font face. */
  fontFamily?: string | null;
  /** Resolved East Asian font face, retained for Far East line metrics. */
  fontFamilyEastAsia?: string | null;
  bold?: boolean;
  italic?: boolean;
}

export interface ShapeRun {
  widthPt: number;
  heightPt: number;
  /** X offset in pt */
  anchorXPt: number;
  /** Y offset in pt */
  anchorYPt: number;
  anchorXFromMargin: boolean;
  anchorYFromPara: boolean;
  /** ECMA-376 §20.4.3.1 wp:align horizontal: "left" | "center" | "right" |
   *  "inside" | "outside". When set the renderer aligns the shape inside the
   *  container indicated by `anchorXFromMargin` and ignores `anchorXPt`. */
  anchorXAlign?: string | null;
  /** Vertical equivalent of anchorXAlign: "top" | "center" | "bottom". */
  anchorYAlign?: string | null;
  /** ECMA-376 §20.4.2.7 wp14:pctPosHOffset / pctPosVOffset normalised to a
   *  fraction in `[0, 1]`. When set the renderer multiplies it by the
   *  relativeFrom container's width / height and uses that as the
   *  shape's offset within the container, ignoring anchorXPt / anchorYPt. */
  pctPosH?: number | null;
  pctPosV?: number | null;
  /** Raw `relativeFrom` value from `<wp:positionH>` / `<wp:positionV>` —
   *  e.g. "page", "margin", "topMargin", "rightMargin",
   *  "insideMargin", "paragraph", "line". Drives container selection
   *  for both pctPos* and anchor*Align positioning. */
  anchorXRelativeFrom?: string | null;
  anchorYRelativeFrom?: string | null;
  /** ECMA-376 §20.4.2.18 wp14:sizeRelH/sizeRelV — width/height as a
   *  fraction of the relativeFrom container. When set, the renderer uses
   *  this in place of `widthPt` / `heightPt` for layout. `pct == 0` from
   *  the source is dropped at parse time (treated as "use extent"). */
  widthPct?: number | null;
  heightPct?: number | null;
  widthRelativeFrom?: string | null;
  heightRelativeFrom?: string | null;
  /** Parent wgp group dimensions (pt) — set only when this shape is a child
   *  of a `<wpg:wgp>`. Used by `resolveAnchor*` so align/pctPos resolve the
   *  GROUP's origin, then `anchor[XY]Pt` adds the within-group offset. */
  groupWidthPt?: number | null;
  groupHeightPt?: number | null;
  /** Draw behind text when true (wp:anchor behindDoc="1"). */
  behindDoc?: boolean;
  /** ECMA-376 §20.4.2.3 `wp:anchor/@relativeHeight`: lower values render first. */
  zOrder: number;
  /** Normalized [0,1] custom-geometry sub-paths. Empty when `presetGeometry`
   *  is set; the renderer chooses between buildCustomPath and buildShapePath. */
  subpaths: PathCmd[][];
  /** OOXML <a:prstGeom prst> name (e.g. "rect", "ellipse", "rtTriangle").
   *  When set the renderer calls core's buildShapePath with `adjValues`. */
  presetGeometry?: string | null;
  /** <a:gd name="adj{n}"> values from prstGeom/avLst in adj1..adj8 order.
   *  `null` preserves omitted named guides so the preset engine can use the
   *  geometry's default for that index. */
  adjValues?: Array<number | null>;
  fill: ShapeFill | null;
  stroke: string | null;
  strokeWidth?: number;
  /** `<a:ln><a:prstDash val>` — ECMA-376 §20.1.8.48. Absent = solid. */
  strokeDash?: string | null;
  /** Normalized line cap: `butt` | `round` | `square`. */
  strokeCap?: CanvasLineCap | null;
  /** `<a:ln><a:headEnd>` line-start decoration (ECMA-376 §20.1.8.3). */
  headEnd?: LineEnd | null;
  /** `<a:ln><a:tailEnd>` line-end decoration (ECMA-376 §20.1.8.3). */
  tailEnd?: LineEnd | null;
  rotation?: number;
  /** `<a:xfrm flipH>` (§20.1.7.6) — mirror about the vertical centre line. */
  flipH?: boolean;
  /** `<a:xfrm flipV>` (§20.1.7.6) — mirror about the horizontal centre line. */
  flipV?: boolean;
  wrapMode?: string | null;
  /** Padding top (pt). Anchor-only. Mirrors {@link ImageRun.distTop}; an anchored
   *  wrap-shape uses these to size its float-exclusion band (ECMA-376 §20.4.2.x). */
  distTop?: number;
  /** Padding bottom (pt). Anchor-only. */
  distBottom?: number;
  /** Padding left (pt). Anchor-only. */
  distLeft?: number;
  /** Padding right (pt). Anchor-only. */
  distRight?: number;
  /** wrapText attribute: "bothSides" | "left" | "right" | "largest". */
  wrapSide?: string | null;
  /** Text rendered INSIDE the shape's bounding box (`<wps:txbx><w:txbxContent>`). */
  textBlocks?: ShapeText[];
  /** ECMA-376 §20.1.4.1.17 `<wps:style><a:fontRef>` — the shape's DEFAULT text
   *  color (hex, no `#`). A text-box run ({@link ShapeTextRun}) with no explicit
   *  {@link ShapeTextRun.color} inherits this before falling back to the
   *  document/theme default (black); an explicit run color still wins. This is
   *  the color axis of the fontRef only — the `@idx` (major/minor/none) font-face
   *  selection is out of scope (fonts resolve via rFonts/docDefaults). Mirrors
   *  pptx's per-shape default text color from the placeholder fontRef. Absent ⇒
   *  no shape default (the run color or black applies). */
  defaultTextColor?: string | null;
  /** "t" | "ctr" | "b" — vertical anchor for the shape's text body (`<wps:bodyPr @anchor>`). */
  textAnchor?: string | null;
  /** ECMA-376 §21.1.2.1.1 auto-fit mode from `<wps:bodyPr>`, normalized to the
   *  shared core `autoFit` vocabulary (core `src/types/common.ts`): "none"
   *  (`<a:noAutofit/>`, fixed box — overflowing text is CLIPPED to the box),
   *  "sp" (`<a:spAutoFit/>`, box grows to text), or "norm" (`<a:normAutofit/>`,
   *  text shrinks). Absent ⇒ overflow visible. */
  textAutofit?: string | null;
  textInsetL?: number;  // pt
  textInsetT?: number;  // pt
  textInsetR?: number;  // pt
  textInsetB?: number;  // pt
  /** ECMA-376 §20.1.10.83 ST_TextVerticalType — the text-body flow direction from
   *  `<wps:bodyPr vert>` / `<a:bodyPr vert>`. Recognised vertical values:
   *  "vert" (all glyphs 90° CW, chars T→B, lines R→L), "vert270" (all glyphs 270°
   *  CW = 90° CCW, chars B→T, lines L→R), and "eaVert" (East-Asian upright: CJK
   *  stands upright, non-EA rotated 90°, chars T→B, lines R→L). "horz"/absent ⇒
   *  horizontal (unchanged). Unrecognised values ("mongolianVert", "wordArtVert",
   *  …) fall back to horizontal until implemented. */
  textVert?: string | null;
  /** ECMA-376 Part 4 §19.1.2.23 `<v:textpath>` — WordArt text laid on the
   *  shape path (a text watermark). When set the renderer draws this string,
   *  scaled to fill the box (`fitshape`), rotated by {@link ShapeRun.rotation},
   *  filled with {@link ShapeRun.fill} at {@link ShapeRun.fillOpacity} alpha —
   *  INSTEAD of a fill/stroke panel + body text. */
  textPath?: TextPath | null;
  /** ECMA-376 Part 4 §19.1.2.5 `<v:fill opacity>` — fill alpha in `[0, 1]`
   *  (default 1 = opaque). Used with {@link ShapeRun.textPath} to draw the
   *  watermark semi-transparently. Absent ⇒ opaque. */
  fillOpacity?: number | null;
}

/** ECMA-376 Part 4 §19.1.2.23 `<v:textpath>` — a WordArt vector text path,
 *  emitted by Word for text watermarks (the `PowerPlusWaterMarkObject` shape).
 *  The text is stretched to fit the shape box (`fitshape`, the WordArt
 *  `#_x0000_t136` shapetype default), so its drawn size derives from the shape
 *  geometry rather than the nominal `font-size` in the textpath style. */
export interface TextPath {
  /** The `string` attribute — the watermark text (e.g. "DRAFT"). */
  string: string;
  /** `font-family` from the textpath style (quotes stripped). */
  fontFamily?: string | null;
  bold?: boolean;
  italic?: boolean;
}

/** DrawingML line-end (arrow head). ECMA-376 §20.1.8.3 CT_LineEndProperties.
 *  Maps 1:1 to core's `ArrowEnd`. */
export interface LineEnd {
  /** "triangle" | "stealth" | "diamond" | "oval" | "arrow" (never "none"). */
  type: string;
  /** Width step: "sm" | "med" | "lg". */
  w: string;
  /** Length step: "sm" | "med" | "lg". */
  len: string;
}

/** One formatting run (`<w:r>`) inside a shape-text paragraph. Mirrors the
 *  character-formatting fields of {@link ShapeText}; the renderer lays a
 *  paragraph's {@link ShapeText.runs} out as rich text so mixed bold/non-bold
 *  runs each keep their own font. */
export interface ShapeTextRun {
  text: string;
  fontSizePt: number;
  color?: string | null;
  /** ECMA-376 §17.3.2.26 ascii axis (`<w:rFonts w:ascii>`), resolved through
   *  docDefaults. Latin letters/digits in this run render with this family. */
  fontFamily?: string | null;
  /** ECMA-376 §17.3.2.26 eastAsia axis (`<w:rFonts w:eastAsia>`), resolved
   *  through docDefaults. CJK characters in this run render with this family;
   *  the renderer falls back to {@link ShapeTextRun.fontFamily} when absent. */
  fontFamilyEastAsia?: string | null;
  bold?: boolean;
  italic?: boolean;
  /** ECMA-376 §17.3.3.25 ruby annotation (furigana) for text-box runs. */
  ruby?: RubyAnnotation | null;
}

export interface ShapeText {
  text: string;
  fontSizePt: number;
  color?: string | null;
  fontFamily?: string | null;
  bold?: boolean;
  italic?: boolean;
  /** Per-run formatting for this paragraph (one entry per `<w:r>` with text).
   *  When non-empty the renderer draws the block as rich text (each run's
   *  font); otherwise it uses the single block-level format fields above
   *  (image blocks / legacy single-format paragraphs). Absent for image-only
   *  paragraphs. */
  runs?: ShapeTextRun[];
  /** ECMA-376 §17.9 paragraph numbering for text-box paragraphs. */
  numbering?: NumberingInfo | null;
  alignment: string;
  /** ECMA-376 §17.3.1.33 `<w:spacing w:before>` of this text-box paragraph, in
   *  pt — reserved ABOVE the paragraph inside the box. Absent/0 ⇒ no offset. */
  spaceBefore?: number;
  /** ECMA-376 §17.3.1.33 `<w:spacing w:after>` of this text-box paragraph, in
   *  pt — reserved BELOW the paragraph. Absent/0 ⇒ no offset. */
  spaceAfter?: number;
  /** ECMA-376 §17.3.1.33 line spacing value (style-chain resolved). Encoded per
   *  {@link lineSpacingRule}: "auto" ⇒ a MULTIPLIER on the natural line box
   *  (1.15 = 276/240), "exact"/"atLeast" ⇒ pt. Absent ⇒ single (natural). */
  lineSpacingVal?: number;
  /** "auto" | "exact" | "atLeast" — see {@link lineSpacingVal}. */
  lineSpacingRule?: string;
  /** ECMA-376 §17.3.1.12 `<w:ind w:left/@start>` — paragraph left indent (pt).
   *  Absent/0 ⇒ flush to the box's inner left edge. */
  indentLeft?: number;
  /** ECMA-376 §17.3.1.12 `<w:ind w:right/@end>` — paragraph right indent (pt).
   *  Absent/0 ⇒ flush to the box's inner right edge. */
  indentRight?: number;
  /** `<w:ind>` first-line indent (pt, SIGNED: `w:firstLine` positive,
   *  `w:hanging` negative). A negative value hangs the first line further LEFT
   *  than the continuation lines (the body renderer honors the sign too — Word
   *  applies a signed hanging first-line list-independently). Absent/0 ⇒ the
   *  first line aligns with the continuation lines. */
  indentFirst?: number;
  /** ECMA-376 §17.3.1.37 `<w:tabs>` — explicit tab stops of this text-box
   *  paragraph, resolved through the style chain like {@link DocParagraph.tabStops}.
   *  Absent/empty ⇒ only the automatic default-tab grid applies. The renderer
   *  feeds these to the SAME line engine the body uses so a `\t` inside a text box
   *  advances to its stop (the old shape wrapper dropped tabs entirely). */
  tabStops?: TabStop[];
  /** ECMA-376 §17.3.1.6 `<w:bidi>` — right-to-left text-box paragraph, resolved
   *  through the style chain like {@link DocParagraph.bidi}. `true` = RTL,
   *  `false` = explicitly LTR, absent = unspecified. Consumed as the paragraph
   *  base direction for the UAX#9 reordering pass (the body renderer reads the
   *  identical field). */
  bidi?: boolean;
  /** ECMA-376 §17.3.1.9 `<w:contextualSpacing>` — resolved through the style
   *  chain in the parser. When set,
   *  `word-contextual-spacing-per-side` drops this text-box paragraph's own
   *  contribution against an adjacent paragraph with the same
   *  {@link ShapeText.styleId}. Absent ⇒ no suppression. */
  contextualSpacing?: boolean;
  /** Resolved paragraph style id of this text-box paragraph — the explicit
   *  `<w:pStyle>`, else the document default paragraph style, else "Normal" (the
   *  same stable id {@link DocParagraph.styleId} carries). Paired with
   *  {@link ShapeText.contextualSpacing} to group adjacent same-style paragraphs
   *  for §17.3.1.9. */
  styleId?: string | null;
  /** Zip path of an inline image inside this text-box paragraph
   *  (`<w:drawing><wp:inline><a:blip r:embed>`), e.g. `word/media/image1.emf`.
   *  Absent for a text-only paragraph. */
  imagePath?: string;
  /** MIME type of the blip at {@link ShapeText.imagePath}. */
  mimeType?: string;
  /** Zip path of the vector original (`asvg:svgBlip` extension), preferred over
   *  `imagePath` when present. */
  svgImagePath?: string;
  /** Inline image natural width in pt (from `<wp:extent cx>`). */
  imageWidthPt?: number;
  /** Inline image natural height in pt (from `<wp:extent cy>`). */
  imageHeightPt?: number;
}

export type ShapeFill =
  | { fillType: 'solid'; color: string }
  | { fillType: 'gradient'; stops: GradientStop[]; angle: number; gradType: string };

export interface GradientStop {
  /** 0.0–1.0 */
  position: number;
  /** hex 6-char */
  color: string;
}

export interface FieldRun {
  /** "page" | "numPages" | "other" */
  fieldType: string;
  instruction: string;
  fallbackText: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  fontSize: number;  // pt
  color: string | null;
  fontFamily: string | null;
  background: string | null;
  vertAlign: 'super' | 'sub' | null;
  allCaps?: boolean;
  smallCaps?: boolean;
  doubleStrikethrough?: boolean;
  highlight?: string | null;
  /** ECMA-376 §17.3.2.12 `<w:em w:val>` — emphasis (boten / 圏点) mark, mirrors
   *  {@link DocxTextRun.emphasisMark} (§17.18.24 ST_Em). Absent (or the
   *  authored `val="none"`) ⇒ no mark. */
  emphasisMark?: EmphasisMark;
}

export interface DocxTextRun {
  text: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  /** ECMA-376 §17.3.2.40 `<w:u w:val>` — the raw ST_Underline (§17.18.99) style
   *  value (`double` / `thick` / `dotted` / `wave` / `dashLong` / …). Absent for
   *  the plain single rule (or no underline). The renderer normalizes this
   *  WordprocessingML vocabulary to the shared DrawingML ST_TextUnderlineType
   *  (§20.1.10.82) that `core.drawUnderline` dispatches on. */
  underlineStyle?: string;
  /** ECMA-376 §17.3.2.40 `<w:u w:color>` — underline-only colour (hex 6, or the
   *  literal `auto`). Absent ⇒ the underline follows the glyph colour. */
  underlineColor?: string;
  strikethrough: boolean;
  fontSize: number;  // pt
  color: string | null;
  fontFamily: string | null;
  /** ECMA-376 §17.3.2.26 eastAsia axis (`<w:rFonts w:eastAsia>`), resolved
   *  through the style chain + docDefaults. CJK code points in this run render
   *  with this family; {@link DocxTextRun.fontFamily} keeps the conflated single-
   *  font fallback (ascii → eastAsia) for paths that do not split per character.
   *  The renderer routes consecutive CJK code points to this axis (the same per-
   *  script rule {@link ShapeTextRun.fontFamilyEastAsia} uses), so a Gothic
   *  eastAsia title sits beside a serif ascii number with no name heuristics.
   *  Absent ⇒ the renderer falls back to {@link DocxTextRun.fontFamily}. */
  fontFamilyEastAsia?: string | null;
  isLink: boolean;
  background: string | null;
  /** ECMA-376 §17.3.2.6 — `<w:color w:val="auto"/>` was set on this run. When
   *  true and {@link DocxTextRun.color} is absent, the renderer resolves the
   *  glyph color from the effective background (an implementation-defined
   *  black/white contrast pick; ECMA-376 gives no normative algorithm) instead
   *  of the default text color. */
  colorAuto?: boolean | null;
  /** ECMA-376 §17.3.2.4 `<w:bdr>` — a run-level border (box) drawn around the
   *  run text. Absent when the run has no border. */
  border?: DocxRunBorder | null;
  vertAlign: 'super' | 'sub' | null;
  /** Target URL for hyperlinks (resolved from relationships.xml) */
  hyperlink: string | null;
  /** ECMA-376 §17.16.23 `<w:hyperlink w:anchor>` — internal bookmark name this
   *  link jumps to (a `<w:bookmarkStart w:name>` in the same document). Set for an
   *  internal cross-reference / TOC entry. When a link carries both `r:id` and
   *  `w:anchor`, {@link DocxTextRun.hyperlink} (external) wins and this still
   *  records the anchor. Absent when the link has no anchor. */
  hyperlinkAnchor?: string | null;
  allCaps?: boolean;
  smallCaps?: boolean;
  doubleStrikethrough?: boolean;
  highlight?: string | null;
  /** ECMA-376 §17.3.2.12 `<w:em w:val>` — emphasis (boten / 圏点) mark drawn on
   *  every non-space character of the run (§17.18.24 ST_Em). `'dot'` = filled
   *  dot above, `'comma'` = sesame/comma above, `'circle'` = hollow circle
   *  above, `'underDot'` = filled dot below (horizontal writing). Absent (or the
   *  authored `val="none"`) ⇒ no mark. The renderer stamps the mark per glyph
   *  after the text and does NOT change the glyph advance. */
  emphasisMark?: EmphasisMark;
  /** ECMA-376 §17.3.3.25 ruby annotation (furigana). Renders above the
   *  base text in a smaller font; line height is expanded to fit it. */
  ruby?: RubyAnnotation;
  /** ECMA-376 §17.13.5 — set when this run sits inside `<w:ins>` or
   *  `<w:del>`. The renderer paints insertions with an author-coloured
   *  underline and deletions with an author-coloured strikethrough so
   *  tracked changes appear inline. */
  revision?: RunRevision;
  /** ECMA-376 §17.3.2.30 `<w:rtl>` — complex-script / right-to-left run.
   *  `true` = RTL, `false` = explicitly LTR, absent = unspecified. The renderer
   *  treats a `true` run as RTL for the UAX#9 pass (it forces complex-script
   *  shaping and marks the segment so `computeLineVisualOrder` reorders it), and
   *  draws the slice with `ctx.direction = 'rtl'` so Canvas mirrors the glyphs. */
  rtl?: boolean;
  /** ECMA-376 §17.3.2.7 `<w:cs/>` — complex-script run toggle: cs formatting
   *  applies to ALL characters of the run (§17.3.2.26). Distinct from
   *  `rFonts@cs` (`fontFamilyCs`), which is only a font slot. */
  cs?: boolean;
  /** ECMA-376 §17.3.2.26 `<w:rFonts w:cs>` — complex-script typeface
   *  (theme references resolved to a literal family). */
  fontFamilyCs?: string;
  /** ECMA-376 §17.3.2.39 `<w:szCs>` — complex-script font size in pt
   *  (same units as `fontSize`). */
  fontSizeCs?: number;
  /** ECMA-376 §17.3.2.3 `<w:bCs>` — complex-script bold toggle. */
  boldCs?: boolean;
  /** ECMA-376 §17.3.2.17 `<w:iCs>` — complex-script italic toggle. */
  italicCs?: boolean;
  /** ECMA-376 §17.3.2.20 `<w:lang w:bidi>` — complex-script (RTL) language tag,
   *  lower-cased (e.g. "ar-sa", "ae-ar"). Drives Word's AN digit ordering. */
  langBidi?: string;
  /** ECMA-376 §17.3.2.34 `<w:snapToGrid>` — false opts this run out of the
   *  section character grid; absent inherits participation. */
  snapToGrid?: boolean;
  /** ECMA-376 §17.3.2.35 `<w:spacing w:val>` — character-spacing adjustment in
   *  POINTS (signed): the extra pitch added after each character before the next
   *  is rendered. The renderer feeds it to `ctx.letterSpacing` on BOTH the
   *  measure and paint passes so line breaking / pagination stay consistent.
   *  Absent ⇒ no extra pitch. */
  charSpacing?: number;
  /** ECMA-376 §17.3.2.14 `<w:fitText>` — manual run-width target in TWIPS
   *  (`w:val`, 1/20 pt) plus the optional `w:id` that links consecutive runs
   *  into one region. The arbitrary-precision XSD integer id is serialized as a
   *  string; numeric synthetic inputs remain supported for layout tests and
   *  direct model construction. An id-less run is always standalone. */
  fitTextVal?: number;
  fitTextId?: string | number;
  /** ECMA-376 §17.3.2.43 `<w:w w:val>` — horizontal text scale as a FRACTION of
   *  normal character width (0.67 = 67%, 2.0 = 200%). Stretches each glyph's
   *  width, not the gap between glyphs. Absent ⇒ 100%. */
  charScale?: number;
  /** ECMA-376 §17.3.2.24 `<w:position w:val>` — baseline raise (positive) /
   *  lower (negative) in POINTS, without changing the font size or line box.
   *  Absent ⇒ no shift. */
  position?: number;
  /** ECMA-376 §17.3.2.19 `<w:kern w:val>` — font-kerning threshold in POINTS
   *  (the smallest font size that is kerned). Presence enables kerning subject
   *  to the threshold; absent ⇒ kerning off (the hierarchy default). `0` = kern
   *  at all sizes. */
  kerning?: number;
  /** ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:vert>` — horizontal-in-vertical
   *  (縦中横 / tate-chū-yoko). `true` means that in a VERTICAL (tbRl) page this
   *  run's characters are laid out horizontally side by side within ONE cell of
   *  the vertical line (rotated 90° relative to the vertical flow). Absent ⇒
   *  normal vertical stacking. Inert in a horizontal page. */
  eastAsianVert?: boolean;
  /** ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:vertCompress>` — compress the
   *  縦中横 run to fit the existing line height without growing the line. Ignored
   *  unless {@link eastAsianVert} is set. Absent ⇒ not compressed. */
  eastAsianVertCompress?: boolean;
  /** ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:combine>` — two-lines-in-one.
   *  PARSED for completeness; not yet rendered (no fixture). */
  eastAsianCombine?: boolean;
  /** ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:combineBrackets>` (§17.18.8) —
   *  bracket style around two-lines-in-one text. PARSED for completeness; the
   *  two-lines-in-one draw is a follow-up. */
  eastAsianCombineBrackets?: string;
  /** ECMA-376 §17.11.6/.7/.16/.17 — set when this run is a footnote/endnote
   *  reference marker (`<w:footnoteReference>` in the body, `<w:footnoteRef>` at
   *  the start of the note's content, and the endnote equivalents). `text` holds
   *  the raw `@w:id`; the renderer overrides the displayed glyph with the note's
   *  sequential number. */
  noteRef?: NoteRef;
}

/** A footnote / endnote reference marker (ECMA-376 §17.11). */
export interface NoteRef {
  /** "footnote" | "endnote" */
  kind: 'footnote' | 'endnote' | string;
  /** `@w:id` linking the marker to its note. Empty for the in-note
   *  `<w:footnoteRef/>` placeholder (the renderer uses the enclosing note). */
  id: string;
}

export interface RunRevision {
  /** "insertion" or "deletion" */
  kind: 'insertion' | 'deletion' | string;
  /** `<w:ins w:author>` / `<w:del w:author>`. Used to colour the markup. */
  author?: string;
  /** ISO-8601 timestamp. */
  date?: string;
}

export interface RubyAnnotation {
  text: string;
  /** Annotation font size in pt; `<w:hps>` stores half-points. */
  fontSizePt: number;
  /** Distance “between the phonetic guide base text and the phonetic guide
   *  text” in pt; `<w:hpsRaise>` stores half-points
   *  (ECMA-376 §17.3.3.12). */
  hpsRaisePt?: number;
}

/** ECMA-376 §17.18.24 ST_Em — the emphasis-mark styles a run may carry via
 *  `<w:em w:val>` (§17.3.2.12). `'none'` is filtered out by the parser, so the
 *  model only ever carries one of these four positive marks (or `undefined`). */
export type EmphasisMark = 'dot' | 'comma' | 'circle' | 'underDot';

export interface ImageRun {
  /**
   * Embedded zip path of the raster blip (e.g. `word/media/image1.png`) — the
   * raster fallback (PNG/JPEG), or the SVG part itself when no raster blip is
   * embedded. The renderer fetches the bytes lazily by path (see {@link
   * DocxDocument.getImage}) instead of inlining base64.
   */
  imagePath: string;
  /** MIME type of the blip at {@link ImageRun.imagePath} (e.g. `image/png`, or
   *  `image/svg+xml` for an svg-only picture). */
  mimeType: string;
  /**
   * Vector original from the Microsoft `asvg:svgBlip` extension (MS-ODRAWXML) —
   * the zip path of the `.svg` part. When present the renderer prefers it over
   * {@link ImageRun.imagePath} (the raster fallback). Absent for a plain raster
   * image. Its MIME is always `image/svg+xml` and is owned by the SVG decoder.
   */
  svgImagePath?: string;
  /**
   * ECMA-376 §20.1.8.55 `<a:srcRect>` — the source-rectangle crop applied to
   * the decoded bitmap before it is drawn into the display box. The four values
   * are signed fractions of the source bitmap measured from each edge
   * (`l`/`t` from left/top, `r`/`b` from right/bottom); the visible source
   * region is `[l, t, 1−r, 1−b]`. The parser converts the raw ST_Percentage
   * (1000ths of a percent) to fractions, so the renderer crops in bitmap pixels
   * (`sx = l*w`, `sy = t*h`, `sw = (1−l−r)*w`, `sh = (1−t−b)*h`) without unit
   * knowledge. Absent / null when there is no crop (the full bitmap is drawn).
   */
  srcRect?: { l: number; t: number; r: number; b: number } | null;
  widthPt: number;
  heightPt: number;
  /** Effective DrawingML transform for grouped pictures. */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  /** true = wp:anchor (absolute positioned), false/undefined = wp:inline (flows with text) */
  anchor?: boolean;
  /** X offset in pt (anchor only) */
  anchorXPt?: number;
  /** Y offset in pt (anchor only) */
  anchorYPt?: number;
  /**
   * If true, anchorXPt is relative to the left margin — add section.marginLeft to get page X.
   * If false/absent, anchorXPt is already page-absolute.
   */
  anchorXFromMargin?: boolean;
  /**
   * If true, anchorYPt is relative to the paragraph's top Y in the renderer.
   * If false/absent, anchorYPt is already page-absolute.
   */
  anchorYFromPara?: boolean;
  /**
   * When set, the renderer replaces all pixels of this hex color (e.g. "FFFFFF") with full
   * transparency. Implements a:clrChange (make-background-transparent).
   */
  colorReplaceFrom?: string;
  /**
   * ECMA-376 §20.1.8.23 `<a:duotone>` recolour, resolved to its two endpoint
   * colours (through the document theme). Absent ⇒ no duotone. When present the
   * renderer decodes the raster once, remaps it along the `clr1`→`clr2`
   * luminance ramp, and caches the recoloured bitmap under a colour-suffixed key.
   */
  duotone?: Duotone;
  /**
   * ECMA-376 §20.1.8.6 `<a:alphaModFix@amt>` opacity as 0..1. Absent ⇒ fully
   * opaque. When present the renderer multiplies the picture's `globalAlpha` by
   * this fraction.
   */
  alpha?: number;
  /**
   * Wrap mode for anchor images:
   *   "square" | "topAndBottom" | "none" | "tight" | "through"
   * Inline images and undetermined cases leave this undefined.
   * MVP renders "tight" and "through" as "square".
   */
  wrapMode?: string;
  /** Padding top (pt). Anchor-only. */
  distTop?: number;
  /** Padding bottom (pt). Anchor-only. */
  distBottom?: number;
  /** Padding left (pt). Anchor-only. */
  distLeft?: number;
  /** Padding right (pt). Anchor-only. */
  distRight?: number;
  /** wrapText attribute: "bothSides" | "left" | "right" | "largest". */
  wrapSide?: string;
  /**
   * ECMA-376 §20.4.2.3 `wp:anchor/@allowOverlap` — whether this floating object
   * may overlap other floats. Spec default is true (the attribute is optional);
   * absent/undefined is treated as true. `false` mandates the renderer
   * reposition the object to prevent any overlap.
   */
  allowOverlap?: boolean;
  /** ECMA-376 §20.4.3.1 wp:align horizontal: "left" | "center" | "right" |
   *  "inside" | "outside". When set the renderer aligns the image inside the
   *  container indicated by `anchorXFromMargin` and ignores `anchorXPt`.
   *  Mirrors {@link ShapeRun.anchorXAlign}. Absent for inline images and
   *  offset-based anchors. */
  anchorXAlign?: string | null;
  /** Vertical equivalent of anchorXAlign: "top" | "center" | "bottom". */
  anchorYAlign?: string | null;
  /**
   * ECMA-376 §20.4.3.2 `<wp:positionH/@relativeFrom>` / §20.4.3.5
   * `<wp:positionV/@relativeFrom>` — names the container the offset / align /
   * pctPos is measured against. Raw spec string: `"page"`, `"margin"`,
   * `"paragraph"`, `"line"`, `"leftMargin"`, `"rightMargin"`, `"topMargin"`,
   * `"bottomMargin"`, `"insideMargin"`, `"outsideMargin"`, `"column"`,
   * `"character"`. Mirrors {@link ShapeRun.anchorXRelativeFrom} /
   * {@link ShapeRun.anchorYRelativeFrom}. When present, supersedes the legacy
   * coarse boolean hints (`anchorXFromMargin` / `anchorYFromPara`) for the
   * align and pctPos paths so e.g. `relativeFrom="margin"` + `align="top"`
   * pins the image to the top content margin rather than the page top. Absent
   * for inline images and for anchors that omitted `<wp:positionH/V>`.
   */
  anchorXRelativeFrom?: string | null;
  anchorYRelativeFrom?: string | null;
}

// ===== Table =====

/**
 * ECMA-376 §17.4.57 `<w:tblpPr>` — authored floating-table positioning.
 * Under `word-effective-floating-table-positioning`, lexical presence alone
 * does not determine whether the table leaves ordinary flow.
 * All fields are optional in the source.
 */
export interface TblpPr {
  /** §17.4.57 minimum distance to wrapping text (dist padding), pt. Default 0. */
  leftFromText: number;
  rightFromText: number;
  topFromText: number;
  bottomFromText: number;
  /** §17.4.57 ST_HAnchor {text,margin,page}. Default 'page'. */
  horzAnchor: 'text' | 'margin' | 'page' | string;
  /** True iff the source `<w:tblpPr>` carried ANY horizontal positioning hint
   *  (horzAnchor, tblpX, or tblpXSpec). When false, no horizontal position was
   *  given. The retained floating-table resolver uses the text/column frame for
   *  this compatibility default rather than claiming an application-specific
   *  normative placement. */
  horzSpecified: boolean;
  /** §17.4.57 ST_VAnchor {text,margin,page}. Default 'page'. */
  vertAnchor: 'text' | 'margin' | 'page' | string;
  /** §17.4.57 absolute signed offset from the horz/vert anchor edge, pt.
   *  Default 0. Ignored when the matching `*Spec` is present. */
  tblpX: number;
  tblpY: number;
  /** §17.4.57 ST_XAlign {left,center,right,inside,outside}. Supersedes tblpX. */
  tblpXSpec?: 'left' | 'center' | 'right' | 'inside' | 'outside' | string;
  /** §17.4.57 ST_YAlign {inline,top,center,bottom,inside,outside}. Supersedes
   *  tblpY, UNLESS vertAnchor='text' (relative vertical positioning is not
   *  allowed there ⇒ tblpYSpec is ignored, fall back to tblpY). */
  tblpYSpec?: 'inline' | 'top' | 'center' | 'bottom' | 'inside' | 'outside' | string;
}

export interface DocTable {
  colWidths: number[];  // pt
  rows: DocTableRow[];
  borders: TableBorders;
  cellMarginTop: number;
  cellMarginBottom: number;
  cellMarginLeft: number;
  cellMarginRight: number;
  /** table horizontal alignment on the page: 'left' | 'center' | 'right'. */
  jc: string;
  /** ECMA-376 §17.4.50 `<w:tblInd>` — indentation added before the table's
   *  LEADING edge (left in an LTR table, right in an RTL/`bidiVisual` table), in
   *  pt. SIGNED: a negative value pulls the table outward past the leading margin
   *  toward the page edge. `type="dxa"` only; `pct`/`auto` are dropped by the
   *  parser per §17.4.50. Absent ⇒ no direct indent. The renderer applies it only
   *  when the resolved `jc` is left/leading (§17.4.50). */
  tblInd?: number;
  /** ECMA-376 §17.4.52 `<w:tblLayout w:type>` — 'fixed' | 'autofit'. Absent
   *  (undefined) means the specification default, 'autofit'. The table grid is
   *  the initial shared grid for §17.18.87; preferred table/cell widths remain
   *  constraints on both layout modes. */
  layout?: string;
  /** ECMA-376 §17.4.63 `<w:tblW>` preferred table width (type="dxa"), pt. */
  widthPt?: number;
  /** `<w:tblW>` type="pct": 50ths of a percent of available content width. */
  widthPct?: number;
  /**
   * ECMA-376 §17.4.1 `<w:bidiVisual>` — render columns in right-to-left
   * (visual) order. `true` = RTL columns, `false` = explicitly LTR, absent =
   * unspecified. When `true` the renderer mirrors the grid so logical column 0
   * is placed rightmost, and flips per-cell left/right borders accordingly.
   */
  bidiVisual?: boolean;
  /** ECMA-376 §17.4.57 `<w:tblpPr>` authored positioning payload. The lexical
   *  payload remains available under `word-effective-floating-table-positioning`;
   *  presence alone is not an effective-floating test. */
  tblpPr?: TblpPr;
  /** ECMA-376 §17.4.56 `<w:tblOverlap w:val>` — 'never' | 'overlap'. 'never' ⇒
   *  the floating table must be repositioned to avoid overlapping other floats.
   *  Default 'overlap' (omitted ⇒ overlap allowed). Ignored when not floating. */
  overlap?: string;
}

export interface TableBorders {
  top: BorderSpec | null;
  bottom: BorderSpec | null;
  left: BorderSpec | null;
  right: BorderSpec | null;
  insideH: BorderSpec | null;
  insideV: BorderSpec | null;
}

export interface BorderSpec {
  width: number;   // pt
  color: string | null;
  style: string;
}

export interface DocTableRow {
  cells: DocTableCell[];
  /** ECMA-376 §17.4.15 `<w:gridBefore>` — shared table-grid columns skipped
   *  before placing this row's first real cell. Omitted/zero means none. */
  gridBefore?: number;
  /** ECMA-376 §17.4.14 `<w:gridAfter>` — shared table-grid columns skipped
   *  after this row's last real cell. Omitted/zero means none. */
  gridAfter?: number;
  rowHeight: number | null;  // pt
  /** ECMA-376 §17.4.80 hRule. "auto" (default) = informational; "atLeast" =
   *  lower bound; "exact" = fixed clip. */
  rowHeightRule: 'auto' | 'atLeast' | 'exact' | string;
  isHeader: boolean;
  /** ECMA-376 §17.4.6 `<w:cantSplit>` — when true, the row must not be split
   *  across page boundaries. Omitted/false rows may split at page boundaries. */
  cantSplit?: boolean;
}

/** ECMA-376 §17.4.7: a table cell may contain paragraphs AND nested tables. */
export type CellElement =
  | { type: 'paragraph' } & DocParagraph
  | { type: 'table' } & DocTable;

export interface DocTableCell {
  content: CellElement[];
  colSpan: number;
  vMerge: boolean | null;  // null=no merge, true=start, false=continuation
  borders: CellBorders;
  background: string | null;
  vAlign: 'top' | 'center' | 'bottom';
  /** ECMA-376 §17.4.71 `<w:tcW>` preferred cell width (type="dxa"), pt. This is
   *  a preferred constraint, not an exact cell box width. */
  widthPt: number | null;
  /** `<w:tcW>` type="pct": 50ths of a percent of the final table width. */
  widthPct?: number;
  /** Per-cell margins (pt) from `<w:tcPr><w:tcMar>` (ECMA-376 §17.4.42). Each
   *  edge overrides the table-level `cellMargin*` default when set; null/absent
   *  = inherit the table default. */
  marginTop?: number | null;
  marginBottom?: number | null;
  marginLeft?: number | null;
  marginRight?: number | null;
}

export interface CellBorders {
  top: BorderSpec | null;
  bottom: BorderSpec | null;
  left: BorderSpec | null;
  right: BorderSpec | null;
  /** ECMA-376 §17.4.34 (tcBorders w:insideH/w:insideV): the interior
   *  horizontal/vertical border this cell contributes. Folded from the cell's
   *  inline tcBorders OVER the resolved conditional table-style borders (§17.7.6)
   *  at parse time. `null` = unset (the renderer falls back to the table-level
   *  insideH/insideV); a spec with style "nil"/"none" = an explicit "no interior
   *  border" (e.g. banded data rows in Medium List 2 / Medium Shading 2). */
  insideH: BorderSpec | null;
  insideV: BorderSpec | null;
}

// ===== Worker message protocol =====

export type WorkerRequest =
  | { type: 'init'; wasmUrl: string }
  | { type: 'parse'; id: number; data: ArrayBuffer; maxZipEntryBytes?: number; maxZipTotalBytes?: number; maxZipEntries?: number }
  | { type: 'extractImage'; id: number; path: string }
  // Project the retained archive to GitHub-flavoured markdown (`DocxArchive.to_markdown`,
  // the handle already opened at `parse` — no re-copy of the file). Twin of
  // `extractImage`: the archive stays in the worker, only the string crosses back.
  | { type: 'toMarkdown'; id: number };

export type WorkerResponse =
  // The model crosses the worker boundary as raw UTF-8 JSON bytes (transferred,
  // not cloned); the main thread does the single `TextDecoder.decode` +
  // `JSON.parse` into a `DocxDocumentModel`. See `parse_docx` (Rust) for why.
  | { type: 'parsed'; id: number; documentJson: ArrayBuffer }
  | { type: 'imageExtracted'; id: number; bytes: ArrayBuffer }
  | { type: 'markdownRendered'; id: number; markdown: string }
  | {
      type: 'error';
      id: number;
      message: string;
      errorName?: string;
      code?: string;
      reason?: string;
      outgoingColumnIndex?: number;
      outgoingColumnCount?: number;
      incomingColumnCount?: number;
    };

// ===== Public API types =====

export interface RenderPageOptions {
  /** Canvas CSS width in px; height is auto-computed from page aspect ratio.
   *  Applies per CALL — pages of different physical widths (per-section pgSz,
   *  §17.6.13) rendered at the same `width` get different px-per-pt scales.
   *  For a uniform document scale, derive a per-page width from
   *  `DocxDocument.pageSize(i)` instead of passing a constant. */
  width?: number;
  dpr?: number;
  defaultTextColor?: string;
  /** Called for each rendered text segment. Used to build a transparent text
   *  selection overlay. On a vertical (§17.6.20 tbRl) page `x`/`y` are the
   *  PHYSICAL top-left and `transform` is the CSS rotation the overlay span
   *  applies about its top-left; absent for horizontal pages. */
  onTextRun?: (run: { text: string; x: number; y: number; w: number; h: number; fontSize: number; font: string; transform?: string }) => void;
  /** Default `true`. When false, ECMA-376 §17.13.5 track-changes runs render
   *  in their normal style (no author colour, no underline / strikethrough)
   *  — equivalent to Word's "Final / No Markup" view. */
  showTrackChanges?: boolean;
  /** ECMA-376 §17.16.5.16 DATE / §17.16.5.72 TIME — the "current" instant a
   *  DATE/TIME field formats through its `\@` date picture (§17.16.4.1). A `Date`
   *  or epoch-ms number. Default = the real current time at render. Set a fixed
   *  value for deterministic / reproducible DATE/TIME field output. */
  currentDate?: Date | number;
}
