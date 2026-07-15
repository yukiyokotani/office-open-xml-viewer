use serde::Serialize;
use std::collections::BTreeMap;

#[derive(Serialize, Debug, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct Document {
    pub section: SectionProps,
    pub body: Vec<BodyElement>,
    pub headers: HeadersFooters,
    pub footers: HeadersFooters,
    /// ECMA-376 §17.7.4 theme `<a:fontScheme><a:majorFont><a:latin@typeface>`
    /// — the heading typeface declared by the document theme. `None` when
    /// the document has no theme part or no Latin major typeface is
    /// declared.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub major_font: Option<String>,
    /// ECMA-376 §17.7.4 theme `<a:fontScheme><a:minorFont><a:latin@typeface>`
    /// — the body typeface declared by the document theme. `None` when the
    /// document has no theme part or no Latin minor typeface is declared.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub minor_font: Option<String>,
    /// ECMA-376 §17.8.3.10 — font family classification from `word/fontTable.xml`.
    /// Maps font name to `<w:family @w:val>` (one of: "roman", "swiss", "modern",
    /// "script", "decorative", "auto"). Used by the renderer to select the
    /// correct CSS generic family (roman→serif, swiss→sans-serif, modern→monospace)
    /// without relying on name-pattern heuristics. Empty when fontTable.xml
    /// is absent or malformed.
    /// Serialized as a `BTreeMap` so JSON key order is deterministic (font names
    /// sorted), making the parser output byte-stable for identical input.
    #[serde(skip_serializing_if = "BTreeMap::is_empty")]
    pub font_family_classes: BTreeMap<String, String>,
    /// ECMA-376 §17.8.3.29 — per-font pitch from `word/fontTable.xml`
    /// (`<w:pitch w:val="…"/>`, ST_Pitch §17.18.66: "fixed" | "variable" |
    /// "default"). Maps font name → pitch value; a font is present only when it
    /// declares `<w:pitch>` (an omitted element is assumed "default" per
    /// §17.8.3.29, which the renderer treats as non-fixed). The renderer uses this
    /// to decide whether a `family="modern"` (§17.8.3.10) face is genuinely
    /// monospace: only "fixed" is. Empty when fontTable.xml is absent or declares
    /// no pitches. BTreeMap for deterministic (byte-stable) JSON key order.
    #[serde(skip_serializing_if = "BTreeMap::is_empty")]
    pub font_family_pitches: BTreeMap<String, String>,
    /// ECMA-376 §17.8.3.1 font name → w:charset hexadecimal byte.
    #[serde(skip_serializing_if = "BTreeMap::is_empty")]
    pub font_family_charsets: BTreeMap<String, String>,
    /// ECMA-376 §17.8.3.3-.6 — embedded fonts declared in `word/fontTable.xml`
    /// (`<w:embedRegular>` / `embedBold` / `embedItalic` / `embedBoldItalic`),
    /// resolved through `word/_rels/fontTable.xml.rels` to their obfuscated
    /// `.odttf` part paths. The renderer de-obfuscates (§17.8.1, via the
    /// `fontKey`) and registers each as a FontFace so text draws with the
    /// authored typeface. Empty when the document embeds no fonts.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub embedded_fonts: Vec<EmbeddedFont>,
    /// ECMA-376 §17.13.5 — track-changes events found in the body. Each entry
    /// is one `<w:ins>` or `<w:del>` block, with the change author / date /
    /// text content. Empty when the document has no tracked changes.
    /// Renderer ignores this; surfaced for tools (MCP, agents).
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub revisions: Vec<DocxRevision>,
    /// ECMA-376 §17.13.4 — `word/comments.xml` flat list. Each comment carries
    /// its id (matches `<w:commentReference w:id>` in the body), author, date,
    /// and plain-text body. Empty when the document has no comments part.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub comments: Vec<DocxComment>,
    /// ECMA-376 §17.11.10 — `word/footnotes.xml` flat list (id + text).
    /// Excludes the spec-defined separator and continuation-separator entries
    /// (id="-1" / "0"). Empty when the document has no footnotes part.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub footnotes: Vec<DocxNote>,
    /// ECMA-376 §17.11.4 — `word/endnotes.xml` flat list. Same shape as
    /// `footnotes`.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub endnotes: Vec<DocxNote>,
    /// ECMA-376 §17.15.1.* — document-wide compatibility / typography settings
    /// from `word/settings.xml`. Currently the Japanese line-breaking (kinsoku)
    /// configuration. `None` when settings.xml carries none of these elements
    /// (the renderer then uses spec defaults: kinsoku ON).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub settings: Option<DocumentSettings>,
    /// RB7 partial degradation: set when `word/document.xml` — the body part —
    /// could not be read or parsed. The document still "opens" (so the viewer
    /// shows a placeholder page instead of throwing an opaque error) with an
    /// empty `body` and this part-tagged error (e.g.
    /// `"word/document.xml: <detail>"`). `None` (and omitted from JSON) for every
    /// healthy document, so existing snapshots are byte-for-byte unchanged. The
    /// renderer paints a visible error placeholder.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parse_error: Option<String>,
}

/// One embedded font-style slot from `word/fontTable.xml`. `style` is one of
/// "regular" | "bold" | "italic" | "boldItalic" (the `<w:embed*>` element).
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct EmbeddedFont {
    /// The `<w:font w:name>` this style belongs to (FontFace family name).
    pub font_name: String,
    /// Style slot: "regular" | "bold" | "italic" | "boldItalic".
    pub style: String,
    /// Resolved zip part path, e.g. "word/fonts/font1.odttf".
    pub part_path: String,
    /// `<w:embed* w:fontKey>` GUID for §17.8.1 de-obfuscation.
    pub font_key: String,
}

/// Document-wide settings surfaced from `word/settings.xml`. Only the
/// typography settings the renderer needs are extracted.
#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct DocumentSettings {
    /// §17.15.1.58 `w:kinsoku` — East-Asian line-breaking toggle. `None` means
    /// the element is absent; the spec default is ON, so the renderer treats
    /// `None` and `Some(true)` identically. `Some(false)` disables kinsoku.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub kinsoku: Option<bool>,
    /// §17.15.1.60 `w:noLineBreaksBefore@w:val` — custom set of characters that
    /// cannot begin a line (行頭禁則). When present it REPLACES the application
    /// default set. Multiple per-`w:lang` elements are concatenated.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub no_line_breaks_before: Option<String>,
    /// §17.15.1.59 `w:noLineBreaksAfter@w:val` — custom set of characters that
    /// cannot end a line (行末禁則). Replaces the default when present.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub no_line_breaks_after: Option<String>,
    /// §22.1.2.30 `m:mathPr/m:defJc@m:val` — document-wide default math
    /// justification (ST_Jc math). `None` when absent; the renderer then falls
    /// back to the spec default `centerGroup`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub math_def_jc: Option<String>,
    /// §17.15.1.25 `w:defaultTabStop@w:val` — the interval (in points, converted
    /// from twips) at which automatic tab stops are generated after all custom
    /// stops. `None` when the element is absent; the spec default is then 720
    /// twips (0.5" = 36pt), which the renderer applies.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_tab_stop: Option<f64>,
    /// §17.15.1.18 `w:characterSpacingControl@w:val` — document-wide East Asian
    /// punctuation compression / spacing control.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub character_spacing_control: Option<String>,
    /// §17.15.3.1 `w:compat` / `w:useFELayout` — enable Far East layout
    /// compatibility behavior.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub use_fe_layout: Option<bool>,
    /// §17.15.3.1 `w:compat` / `w:balanceSingleByteDoubleByteWidth` — balance
    /// single-byte and double-byte character widths in East Asian layout.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub balance_single_byte_double_byte_width: Option<bool>,
    /// §17.15.3.1 `w:compat` / `w:adjustLineHeightInTable` — apply the section
    /// document-grid line pitch to text in table cells.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub adjust_line_height_in_table: Option<bool>,
}

/// Single track-changes event extracted from a body `<w:ins>` / `<w:del>`.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DocxRevision {
    /// "insertion" | "deletion"
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    /// `<w:ins w:date>` / `<w:del w:date>` — ISO-8601 timestamp.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    /// Concatenated plain text. For deletions, this is the deleted text
    /// captured from `<w:delText>` runs.
    pub text: String,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DocxComment {
    pub id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    /// `<w:initials>` from `word/comments.xml`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub initials: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    pub text: String,
}

/// ECMA-376 §17.11.2 (endnote) / §17.11.10 (footnote) — one note's block-level
/// content. Retains the full paragraph/run structure (FootnoteText style, the
/// `<w:footnoteRef/>` auto-number placeholder, inline formatting) so the
/// renderer can lay it out faithfully at page bottom / document end, and so the
/// data-only API can surface the plain text.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DocxNote {
    pub id: String,
    /// Block-level content of the note (paragraphs / nested tables), parsed with
    /// the document's styles + numbering. Empty for an empty note.
    pub content: Vec<BodyElement>,
}

#[derive(Serialize, Debug, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct HeadersFooters {
    pub default: Option<HeaderFooter>,
    pub first: Option<HeaderFooter>,
    pub even: Option<HeaderFooter>,
}

#[derive(Serialize, Debug, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooter {
    pub body: Vec<BodyElement>,
}

/// ECMA-376 §17.6.12 `<w:pgNumType>` — a section's page-numbering settings. Only
/// the two attributes that affect the DISPLAYED page number are carried:
///
/// * `start` (`@w:start`, ST_DecimalNumber) — the page number shown on the FIRST
///   page of the section. When absent (`None`) numbering continues from the
///   previous section's highest page number (§17.6.12). Kept as a signed integer
///   because Word writes `start="0"` (and, in principle, negatives).
/// * `fmt` (`@w:fmt`, ST_NumberFormat §17.18.59) — the number format for every
///   page number in the section (decimal / upperRoman / lowerLetter / …). `None`
///   ⇒ the spec default `decimal`. Carried verbatim; the TS renderer maps it via
///   the shared `formatOrdinalNumber` kernel.
///
/// `chapStyle`/`chapSep` (chapter-prefixed numbering) are intentionally NOT
/// modeled — out of scope for this pass (they require resolving heading numbering
/// state); a `<w:pgNumType>` that carries only those is treated as "no start / no
/// fmt" and numbering continues normally.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct PageNumType {
    /// `@w:start` — the first page number of the section. `None` ⇒ continue from
    /// the previous section (§17.6.12).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start: Option<i64>,
    /// `@w:fmt` — ST_NumberFormat (§17.18.59) for this section's page numbers.
    /// `None` ⇒ decimal (the spec default).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fmt: Option<String>,
}

/// ECMA-376 §17.6.10 `<w:pgBorders>` — the page borders drawn around each page of
/// a section. Each edge is a `CT_Border` (§17.18.4: val / sz / space / color);
/// the container carries the placement globals `offsetFrom`, `display`, `zOrder`.
/// `None` on `SectionProps` when the sectPr declares no `<w:pgBorders>` (the
/// common case — no page border), so existing snapshots are byte-identical.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct PageBorders {
    /// `@w:offsetFrom` (ST_PageBorderOffset §17.18.63): "page" ⇒ each edge's
    /// `space` is measured from the PAGE edge; "text" (the spec DEFAULT) ⇒ from
    /// the text margin. Carried verbatim; the renderer positions the rectangle
    /// accordingly.
    pub offset_from: String,
    /// `@w:display` (ST_PageBorderDisplay §17.18.62): "allPages" (default) |
    /// "firstPage" | "notFirstPage". Governs which physical pages of the section
    /// show the border.
    pub display: String,
    /// `@w:zOrder` (ST_PageBorderZOrder §17.18.64): "front" (default) ⇒ the border
    /// is painted OVER intersecting text/objects; "back" ⇒ UNDER them.
    pub z_order: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<PageBorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<PageBorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<PageBorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<PageBorderEdge>,
}

/// ECMA-376 §17.18.4 `CT_Border` as used by one edge of `<w:pgBorders>`. Same
/// shape as `ParaBorderEdge`: a line style + color + width (pt) + spacing (pt).
/// Art borders (`@w:val` naming a decorative image, §17.18.2 ST_Border art
/// values) are NOT modeled — an art `val` is carried verbatim and the renderer
/// falls back to no ink (documented as unsupported).
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct PageBorderEdge {
    /// `@w:val` — ST_Border line style ("single" | "double" | "dashed" | "dotted"
    /// | "thick" | …). Reused by the renderer's shared border-line drawing.
    pub style: String,
    /// `@w:color` — hex 6, or `None` for "auto" (renderer defaults to black).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// `@w:sz` in pt (eighths of a point ÷ 8), like every other CT_Border width.
    pub width: f64,
    /// `@w:space` in pt. For page borders `space` is a POINT measure (ST_PointMeasure
    /// §17.18.68, 0–31), NOT twips — the offset the border is inset by from the
    /// `offset_from` reference (page edge or text margin).
    pub space: f64,
}

/// ECMA-376 §17.6.8 `<w:lnNumType>` — line numbering for a section: a number is
/// drawn in the left margin of each body line that is a multiple of `count_by`.
/// `None` on `SectionProps` when the sectPr declares no `<w:lnNumType>` (line
/// numbering is off — the common case; snapshots stay byte-identical).
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct LineNumbering {
    /// `@w:countBy` (§17.6.8): only lines whose number is an even multiple of this
    /// value display a number. Required for the struct to exist — an absent
    /// `countBy` means NO line numbering, so the whole struct is `None` then.
    pub count_by: i64,
    /// `@w:start` — the starting line number after each restart. Default 1.
    pub start: i64,
    /// `@w:distance` in pt (twips ÷ 20) — the gap between the text margin and the
    /// line-number glyphs. `None` ⇒ implementation-defined positioning (§17.6.8);
    /// the renderer uses a default gap then.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub distance: Option<f64>,
    /// `@w:restart` (ST_LineNumberRestart §17.18.47): "newPage" (default) |
    /// "newSection" | "continuous". When the counter is reset to `start`.
    pub restart: String,
}

/// ECMA-376 §17.6.13 `<w:pgSz>` + §17.6.11 `<w:pgMar>` — a section's page
/// geometry: page size + margins + header/footer distances (all pt, converted
/// from twips). Carried on a `BodyElement::SectionBreak` (`geom`) so mid-body
/// sections keep their own page size (the FINAL section's geometry lives on the
/// `SectionProps` fields of `Document.section`). Mirrored on the TS side as
/// `SectionGeom`; the renderer resolves per-page geometry from it.
///
/// `orient` is intentionally omitted: Word swaps `w`/`h` for a landscape page,
/// so the verbatim `w`/`h` already give the correct dims — no orientation flag
/// is needed to size the page.
///
/// No `Default` derive on purpose: zeros ≠ spec defaults; a derived default is an
/// all-zeros geometry (0×0 page, zero margins), which is NOT the ECMA-376 spec
/// default (US Letter portrait, 1" margins, 0.5" header/footer). Use
/// `spec_default_geom()` (parser.rs) when a spec-default geometry is needed.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct SectionGeom {
    /// page width in pt (converted from twips)
    pub page_width: f64,
    /// page height in pt
    pub page_height: f64,
    /// §17.6.11 top/bottom are ST_SignedTwipsMeasure and MAY be negative — keep
    /// the sign (identical to `SectionProps.margin_top/bottom`).
    pub margin_top: f64,
    pub margin_right: f64,
    pub margin_bottom: f64,
    pub margin_left: f64,
    /// distance from top of page to header (pt)
    pub header_distance: f64,
    /// distance from bottom of page to footer (pt)
    pub footer_distance: f64,
}

#[derive(Serialize, Debug, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct SectionProps {
    /// page width in pt (converted from twips)
    pub page_width: f64,
    /// page height in pt
    pub page_height: f64,
    pub margin_top: f64,
    pub margin_right: f64,
    pub margin_bottom: f64,
    pub margin_left: f64,
    /// distance from top of page to header (pt)
    pub header_distance: f64,
    /// distance from bottom of page to footer (pt)
    pub footer_distance: f64,
    /// whether first page has its own header/footer
    pub title_page: bool,
    /// whether even pages have distinct header/footer
    pub even_and_odd_headers: bool,
    /// ECMA-376 §17.6.22 ST_SectionMark — the body (final) section's `<w:type>`
    /// start type ("continuous" | "nextPage" | "oddPage" | "evenPage"). Governs
    /// how the last section begins relative to the previous one; the paginator
    /// consumes it at the boundary INTO the final section (non-final sections
    /// carry their start type on their own `SectionBreak` marker). `None` ⇒
    /// "nextPage" (the spec default).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub section_start: Option<String>,
    /// ECMA-376 §17.6.20 `<w:textDirection w:val>` — the section's flow
    /// direction, using the TRANSITIONAL ST_TextDirection enum Word writes
    /// (Part 4 §14.11.7: `lrTb`|`tbRl`|`btLr`|`lrTbV`|`tbLrV`|`tbRlV`), NOT the
    /// Part 1 §17.18.93 Strict set. "lrTb" (the default; horizontal, left→right /
    /// top→bottom) is treated as `None` here so horizontal documents keep
    /// byte-identical rendering; "tbRl" (vertical Japanese: glyphs stack downward,
    /// lines advance right→left) and the other non-default values are carried so
    /// the renderer can decide which flow vertically. `None` ⇒ horizontal (lrTb).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_direction: Option<String>,
    /// ECMA-376 §17.6.5 w:docGrid/@w:type ("default" | "lines" |
    /// "linesAndChars" | "snapToChars"). None = default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_grid_type: Option<String>,
    /// ECMA-376 §17.6.5 w:docGrid/@w:linePitch in pt (converted from twentieths
    /// of a point). When docGridType is "lines" or "linesAndChars", this is
    /// the vertical pitch per grid line — auto line spacing multiplies against
    /// this instead of the font's natural line height.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_grid_line_pitch: Option<f64>,
    /// ECMA-376 §17.6.5 `<w:docGrid w:charSpace>` (ST_DecimalNumber, signed).
    /// The per-character-grid spacing in 1/4096ths of an em (NOT twips). When
    /// `doc_grid_type` is "linesAndChars" or "snapToChars", every full-width
    /// East-Asian glyph occupies a fixed cell of width `fontSizePt +
    /// charSpace/4096` pt; a positive value loosens the cell, a negative value
    /// (the common case) tightens it. `None` when the attribute is absent (the
    /// renderer then leaves East-Asian glyphs at their natural em advance).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_grid_char_space: Option<f64>,
    /// ECMA-376 §17.6.4 `<w:cols>` — newspaper-style multi-column layout for the
    /// section. `None` when the section is single-column (`<w:cols>` absent or
    /// `@w:num` <= 1), in which case the renderer keeps its single full-width
    /// content column (unchanged behavior).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub columns: Option<ColumnsSpec>,
    /// ECMA-376 §17.6.12 `<w:pgNumType>` — the body (final) section's
    /// page-numbering settings (start / fmt). `None` when the body-level sectPr
    /// omits `<w:pgNumType>`. The renderer resolves the DISPLAYED page number per
    /// physical page from this + the per-section `geom.pageNumType`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_num_type: Option<PageNumType>,
    /// ECMA-376 §17.6.10 `<w:pgBorders>` — page borders for this section. `None`
    /// when the sectPr declares none (no page border — the common case).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_borders: Option<PageBorders>,
    /// ECMA-376 §17.6.8 `<w:lnNumType>` — line numbering for this section. `None`
    /// when line numbering is off (no `<w:lnNumType countBy>`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_numbering: Option<LineNumbering>,
    /// ECMA-376 §17.6.23 `<w:vAlign w:val>` — vertical alignment of the body text
    /// between the top/bottom margins ("top" | "center" | "both" | "bottom").
    /// `None` ⇒ "top" (the default; body flows from the top margin unchanged).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub v_align: Option<String>,
}

/// ECMA-376 §17.6.4 `<w:cols>` — the section's multi-column configuration.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ColumnsSpec {
    /// `@w:num` — number of columns (>= 2 when this struct is emitted).
    pub count: usize,
    /// `@w:space` in pt (converted from twips) — the inter-column gap used when
    /// `equal_width` is true. Default 720 twips (36 pt) per the spec.
    pub space_pt: f64,
    /// `@w:equalWidth` — when true (the default), every column has the same
    /// width and `space_pt` is the uniform gap; `cols` is empty. When false the
    /// per-column `cols` entries define the geometry verbatim.
    pub equal_width: bool,
    /// `@w:sep` — draw vertical separator rules between columns.
    pub sep: bool,
    /// Per-column `<w:col>` entries (width + trailing space, pt). Empty when
    /// `equal_width` is true.
    pub cols: Vec<ColSpec>,
}

/// ECMA-376 §17.6.3 `<w:col>` — one column's width and trailing space.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ColSpec {
    /// `@w:w` — the column's width in pt (converted from twips).
    pub width_pt: f64,
    /// `@w:space` — space after this column in pt (converted from twips).
    pub space_pt: f64,
}

/** Parser-private section placement facts. This wire is deliberately emitted
 * under `__sectionPlacement` on SectionBreak and projected into an internal TS
 * sidecar; it is not part of the stable public BodyElement declaration. */
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct SectionPlacementWire {
    pub section_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub v_align: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_numbering: Option<LineNumbering>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum BodyElement {
    // Keep the block wrapper compact as paragraph metadata grows. `Box` is
    // transparent to Serde, so the parser wire shape remains unchanged while
    // body vectors avoid reserving the largest paragraph payload per element.
    Paragraph(Box<DocParagraph>),
    // Tables are similarly recursive and substantially larger than marker
    // variants; indirection keeps every body-vector slot compact.
    Table(Box<DocTable>),
    /// Page break. `parity` carries section-break parity intent:
    /// `Some("odd")` = oddPage break (next content must start on an odd
    /// 1-based page), `Some("even")` = evenPage. `None` = a plain `nextPage`
    /// or `<w:br w:type="page"/>` break.
    PageBreak {
        #[serde(skip_serializing_if = "Option::is_none")]
        parity: Option<String>,
    },
    /// ECMA-376 §17.3.1.20 `<w:br w:type="column"/>` — force the following
    /// content into the next newspaper column of the current section (or the
    /// first column of the next page when already in the last column). Hoisted
    /// to the body level (mirroring `PageBreak`) so the paginator can act on it
    /// without inspecting runs mid-paragraph. In a single-column section it
    /// behaves like a page break.
    ColumnBreak,
    /// ECMA-376 §17.6.x — a section boundary in the body. A `<w:sectPr>` carried
    /// in a paragraph's `pPr` (or a loose mid-body `<w:sectPr>`) defines the
    /// section that ENDS at that point; the FINAL section is the body-level
    /// `<w:sectPr>` and is surfaced on `Document.section` instead of here.
    ///
    /// `columns` is THAT ending section's `<w:cols>` (§17.6.4), parsed via
    /// `parse_columns` (so `None` ⇒ a single full-width column — the spec default
    /// when `@w:num` is absent or 1). `kind` is the section's ST_SectionMark
    /// (`<w:type w:val>`, §17.18.79) which controls how the NEXT section starts:
    /// "continuous" ⇒ no page break (the next section's columns begin on the same
    /// page); "nextPage" (default) ⇒ a plain page break; "oddPage"/"evenPage" ⇒ a
    /// page break + parity padding.
    ///
    /// The renderer switches its active column geometry to the *next* section's
    /// columns at each marker (each marker carries the columns of the section it
    /// terminates, so the renderer peeks forward to the next marker to size the
    /// current section).
    #[serde(rename_all = "camelCase")]
    SectionBreak {
        /// ST_SectionMark token: "continuous" | "nextPage" | "oddPage" |
        /// "evenPage". Always one of these (the parser normalizes the default to
        /// "nextPage").
        kind: String,
        /// The terminating section's `<w:cols>` (§17.6.4). `None` ⇒ single
        /// full-width column.
        #[serde(skip_serializing_if = "Option::is_none")]
        columns: Option<ColumnsSpec>,
        /// ECMA-376 §17.10.1 — the resolved header set for the section that ENDS
        /// at this marker (its own `<w:headerReference>`s layered onto the
        /// inherited running state, then loaded from the package). Lets the
        /// renderer pick the active section's header per page, exactly as
        /// `columns` lets it pick the section's column geometry. Empty when no
        /// section in the inheritance chain declared a header.
        #[serde(default)]
        headers: Box<HeadersFooters>,
        /// ECMA-376 §17.10.1 — the resolved footer set for the section that ENDS
        /// at this marker (see `headers`). sample-13's first section declares a
        /// `first` footer (the DOI line) here; the renderer renders it on that
        /// section's first page.
        #[serde(default)]
        footers: Box<HeadersFooters>,
        /// ECMA-376 §17.10.1 `<w:titlePg>` — whether THIS ending section has a
        /// distinct first-page header/footer. NOT inherited (each sectPr's flag
        /// stands alone, like the body-level `Document.section.title_page`).
        title_page: bool,
        /// ECMA-376 §17.6.13 / §17.6.11 — this ENDING section's page geometry
        /// (size + margins). `None` when the sectPr declares no `<w:pgSz>` /
        /// `<w:pgMar>` (the renderer then falls back to the body-level section).
        /// The final (body-level) section's geometry stays on `Document.section`.
        #[serde(skip_serializing_if = "Option::is_none")]
        geom: Option<Box<SectionGeom>>,
        /// ECMA-376 §17.6.12 `<w:pgNumType>` — this ENDING section's page-numbering
        /// settings (start / fmt). `None` when the sectPr omits `<w:pgNumType>` (or
        /// carries only chapter attributes) — numbering continues; decimal. Carried
        /// SEPARATELY from `geom` (not bundled) because a section may inherit its
        /// page geometry yet still restart / re-format its page numbers; the
        /// renderer resolves the displayed number per physical page from this + the
        /// body-level `Document.section.page_num_type`. Mirrors how `columns` /
        /// `headers` are carried per-terminating-section.
        #[serde(skip_serializing_if = "Option::is_none")]
        page_num_type: Option<PageNumType>,
        /// ECMA-376 §17.6.20 `<w:textDirection w:val>` — this ENDING section's
        /// flow direction (TRANSITIONAL ST_TextDirection, Part 4 §14.11.7), so a
        /// vertical (tbRl/btLr) non-final section can coexist with a horizontal
        /// final section (issue #1000 per-section mixing). Same handling as
        /// `SectionProps.text_direction`: the default "lrTb" (and an absent
        /// element) collapse to `None`; other tokens are carried verbatim.
        /// Carried SEPARATELY from `geom` (like `page_num_type`) because a
        /// section may inherit its page geometry yet still set its own flow
        /// direction. `None` ⇒ horizontal (lrTb).
        #[serde(skip_serializing_if = "Option::is_none")]
        text_direction: Option<String>,
        /// Internal A3 placement projection. Kept off the stable TS public
        /// BodyElement contract and consumed only by parser-model normalization.
        /// Boxed because this private diagnostic payload would otherwise set the
        /// allocation size of every body element; Serde keeps the wire unchanged.
        #[serde(rename = "__sectionPlacement")]
        section_placement: Box<SectionPlacementWire>,
    },
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct DocParagraph {
    /// "left" | "center" | "right" | "both"
    pub alignment: String,
    /// pt
    pub indent_left: f64,
    /// pt
    pub indent_right: f64,
    /// pt (negative = hanging)
    pub indent_first: f64,
    /// pt
    pub space_before: f64,
    /// pt
    pub space_after: f64,
    /// None = single (1.0), Some(LineSpacing)
    pub line_spacing: Option<LineSpacing>,
    /// Boxed: `NumberingInfo` carries several resolved strings (text + marker
    /// font axes); boxing keeps `DocParagraph` small enough that the
    /// `BodyElement`/`CellElement` enums stay balanced (clippy::large_enum_variant).
    /// Serde flattens the Box, so the JSON shape is unchanged.
    pub numbering: Option<Box<NumberingInfo>>,
    /// Explicit tab stops from w:tabs. Empty means use default tab interval.
    pub tab_stops: Vec<TabStop>,
    pub runs: Vec<DocRun>,
    /// ECMA-376 §17.13.6.2 `<w:bookmarkStart w:name>` — the names of every
    /// bookmark that STARTS within (or at the head of) this paragraph, in
    /// document order. A `<w:hyperlink w:anchor="X">` (§17.16.23) jumps to the
    /// paragraph whose `bookmarks` contains `"X"`; the TS side turns these into a
    /// `bookmarkName → pageIndex` map after pagination so an internal-link click
    /// can scroll/render the destination page. Empty (and omitted from JSON) for
    /// the common paragraph that anchors nothing, so existing snapshots are
    /// unchanged. The reserved `_GoBack` bookmark Word inserts is kept — callers
    /// can ignore it; dropping it here would be a policy decision the map builder
    /// is free to make instead.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub bookmarks: Vec<String>,
    /// Paragraph background hex color (w:shd fill on paragraph)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub shading: Option<String>,
    /// Force a page break before this paragraph (w:pageBreakBefore)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub page_break_before: bool,
    /// Suppress spacing between adjacent same-style paragraphs (w:contextualSpacing)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub contextual_spacing: bool,
    /// Keep paragraph on the same page as the next paragraph (w:keepNext)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub keep_next: bool,
    /// Keep all lines of this paragraph on the same page (w:keepLines)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub keep_lines: bool,
    /// ECMA-376 §17.3.1.29 + §17.3.2.41 — the paragraph MARK's resolved `w:vanish`
    /// (hidden text). An inkless paragraph whose mark is vanished collapses to zero
    /// height in the normal/print view (hidden-text off), the same way a hidden run
    /// is stripped; the renderer's paginator drops it whole.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub mark_vanish: bool,
    /// Widow/orphan control (w:widowControl). Default per spec: true.
    pub widow_control: bool,
    /// Paragraph borders (w:pBdr)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub borders: Option<ParagraphBorders>,
    /// Style ID of the applied paragraph style (for contextual spacing resolution)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
    /// Default font size in pt inherited from style + direct pPr/rPr. Used for
    /// sizing empty paragraphs (lines with no runs) correctly.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_font_size: Option<f64>,
    /// Default font family resolved from the style chain + direct pPr/rPr.
    /// Used to size empty paragraphs (no runs) when the intended font's line
    /// height differs from the fallback (e.g. an empty Meiryo cell that forms a
    /// résumé "bar" must reserve Meiryo's tall line box). None when unresolved.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_font_family: Option<String>,
    /// Default East Asian font family resolved from the style chain + direct
    /// pPr/rPr. The paragraph mark has run properties (§17.3.1.29), and in an
    /// East Asian document its mark glyph must use the eastAsia axis rather than
    /// the ASCII fallback.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_font_family_east_asia: Option<String>,
    /// Internal effective paragraph-mark run facts used by the shared text
    /// resolver for empty and anchor-only line metrics. Not part of the stable
    /// public TypeScript document model.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paragraph_mark_font_facts: Option<RunFontFacts>,
    /// ECMA-376 §17.3.1.29 — the paragraph MARK run's resolved `<w:color>`
    /// (direct `pPr/rPr` → pStyle chain → docDefaults, the same `mark_run`
    /// resolution that feeds `default_font_size`; hex 6 lowercased; an
    /// explicit `auto` breaks the chain and surfaces as `None`, §17.3.2.6).
    /// Word formats a numbering marker with the level rPr (§17.9.24) layered
    /// over the mark's run properties, so a mark-colored list item tints its
    /// bullet/number even when the level rPr carries no color — the renderer
    /// reads this as the marker-color fallback after `NumberingInfo::color`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paragraph_mark_color: Option<String>,
    /// ECMA-376 §17.3.1.20 `<w:outlineLvl w:val="N">` (0–8). Resolved through
    /// the style chain: explicit pPr → linked paragraph style → docDefaults.
    /// `None` for body paragraphs that don't appear in the document outline.
    /// Surfaced so MCP / agents can build a heading hierarchy without relying
    /// on the styleId string ("Heading1", "見出し1", etc.).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u32>,
    /// ECMA-376 §17.3.1.6 `<w:bidi>` — right-to-left paragraph. `Some(true)` =
    /// RTL, `Some(false)` = explicitly LTR, `None` = unspecified (inherit).
    /// The renderer consumes this as the paragraph base direction: it seeds the
    /// UAX#9 reordering pass, swaps the left/right indents, and resolves the
    /// `w:jc` start/end alignment edges against it.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi: Option<bool>,
    /// ECMA-376 §17.3.1.32 `<w:snapToGrid>` — `Some(false)` opts the paragraph
    /// OUT of the section's document grid, so its lines use natural font
    /// metrics / the line-spacing multiplier directly instead of snapping to the
    /// grid pitch. `None` = inherit (default on). Set on Word's "Footnote Text"
    /// style.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub snap_to_grid: Option<bool>,
    /// ECMA-376 §17.3.1.11 `<w:framePr>` — text-frame / drop-cap properties.
    /// `Some` ⇒ this paragraph is part of a text frame; the renderer positions
    /// it as a frame (drop cap or generic frame) and registers a wrap exclusion
    /// for following body text. `None` ⇒ ordinary in-flow paragraph.
    /// Boxed (like `numbering`) so `DocParagraph` stays small enough that the
    /// `BodyElement` / `CellElement` enum variants remain balanced
    /// (clippy::large_enum_variant); serde flattens the Box, so JSON is unchanged.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub frame_pr: Option<Box<FramePr>>,
    #[serde(
        rename = "__paragraphTypographyAcquisition",
        skip_serializing_if = "Option::is_none"
    )]
    pub(crate) paragraph_typography_acquisition: Option<ParagraphTypographyWire>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphBorders {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<ParaBorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<ParaBorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<ParaBorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<ParaBorderEdge>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub between: Option<ParaBorderEdge>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ParaBorderEdge {
    /// "single" | "double" | "dashed" | ...
    pub style: String,
    pub color: Option<String>,
    /// pt (sz / 8)
    pub width: f64,
    /// pt spacing between border and text
    pub space: f64,
}

/// ECMA-376 §17.3.2.4 `<w:bdr>` — a run-level border ("box" around the run).
/// Serialized shape matches `DocxRunBorder` on the TS side.
#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct RunBorder {
    /// "single" | "double" | "dashed" | ... (w:bdr/@w:val)
    pub style: String,
    /// hex 6, or None for automatic (renderer falls back to text color)
    pub color: Option<String>,
    /// pt (w:sz / 8)
    pub width: f64,
    /// pt spacing between the border and the run text (w:space)
    pub space: f64,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct TabStop {
    /// tab stop position in pt (from the left of the paragraph content area)
    pub pos: f64,
    /// "left" | "center" | "right" | "decimal" | "bar" | "clear"
    pub alignment: String,
    /// "none" | "dot" | "hyphen" | "underscore" | "heavy" | "middleDot"
    pub leader: String,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct LineSpacing {
    /// multiplier (e.g. 1.15) or exact pt
    pub value: f64,
    /// "auto" | "exact" | "atLeast"
    pub rule: String,
    /// True when `w:spacing/@w:line` is set on the paragraph's own pPr or on
    /// a named style; false when only docDefault provides it. Controls the
    /// docGrid interaction per ECMA-376 §17.6.5 — inherited-only paragraphs
    /// snap to one grid pitch per line regardless of the multiplier.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub explicit: bool,
}

/// ECMA-376 §17.3.1.11 `<w:framePr>` — text-frame / drop-cap properties for a
/// paragraph. The presence of `framePr` makes the paragraph part of a text
/// frame; adjacent paragraphs whose attribute sets are identical belong to the
/// same frame. Lengths are normalized to pt (twip / 20); the raw enum strings
/// are carried verbatim so the renderer maps them with no re-derivation.
///
/// Attribute defaults (per §17.3.1.11):
///   wrap     → "around"   hRule → "auto"   vRule (implied) → none here
///   hAnchor  → "page"     vAnchor → "page"
///   lines    → 1          h/w/x/y → 0
/// `x`/`y` are ignored when `xAlign`/`yAlign` are set; for a drop cap, `y`/
/// `yAlign` are ignored entirely and `lines` drives the height.
#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct FramePr {
    /// ECMA-376 §17.3.1.11 `anchorLock`. This is retained only for effective
    /// CT_FramePr identity during layout acquisition: exposing it as a normal
    /// camelCase member would widen the stable public TypeScript FramePr API.
    #[serde(rename = "__anchorLock", skip_serializing_if = "std::ops::Not::not")]
    pub(crate) anchor_lock: bool,
    /// ST_DropCap (§17.18.20): "none" | "drop" | "margin". Absent ⇒ "none".
    pub drop_cap: String,
    /// §17.3.1.11 `lines` — drop-cap vertical height in anchor lines. Default 1.
    pub lines: u32,
    /// ST_Wrap (§17.18.104): "around" | "auto" | "none" | "notBeside" |
    /// "through" | "tight". Absent ⇒ "around".
    pub wrap: String,
    /// ST_HAnchor (§17.18.35): "text"(=column) | "margin" | "page". Default "page".
    pub h_anchor: String,
    /// ST_VAnchor (§17.18.100): "text" | "margin" | "page". Default "page".
    pub v_anchor: String,
    /// ST_HeightRule (§17.18.37): "auto" | "atLeast" | "exact". Default "auto".
    pub h_rule: String,
    /// hSpace — min wrap padding L/R when wrap="around" (pt). Default 0.
    pub h_space: f64,
    /// vSpace — min wrap padding top/bottom (pt). Default 0.
    pub v_space: f64,
    /// w — exact frame width (pt). 0/absent ⇒ auto (max content line width).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub w: Option<f64>,
    /// h — frame height (pt). Meaning gated by h_rule. 0/absent ⇒ auto.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub h: Option<f64>,
    /// x — absolute horizontal offset from h_anchor (pt). Ignored when x_align set.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub x: Option<f64>,
    /// y — absolute vertical offset from v_anchor (pt). Ignored when y_align set
    /// or when this is a drop cap.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub y: Option<f64>,
    /// ST_XAlign (§22.9.2.18): "left" | "center" | "right" | "inside" |
    /// "outside". Supersedes `x` when present.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub x_align: Option<String>,
    /// ST_YAlign (§22.9.2.20): "inline" | "top" | "center" | "bottom" |
    /// "inside" | "outside". Supersedes `y` when present (ignored if v_anchor=text).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub y_align: Option<String>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct NumberingInfo {
    pub num_id: u32,
    pub level: u32,
    /// "decimal" | "bullet" | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman"
    pub format: String,
    /// resolved text, e.g. "1." or "•"
    pub text: String,
    /// indent for the entire numbered paragraph (pt)
    pub indent_left: f64,
    /// tab stop after bullet/number (pt)
    pub tab: f64,
    /// ECMA-376 §17.9.28 `<w:suff>` — "tab" (default) | "space" | "nothing".
    /// Determines where body text starts after the marker on the first line.
    pub suff: String,
    /// ECMA-376 §17.9.8 `<w:lvlJc>` — marker justification: "left" (default) |
    /// "right" (period-aligned numerals — marker right edge at the hanging-indent
    /// position) | "center". The renderer offsets the marker draw accordingly.
    pub jc: String,
    /// ECMA-376 §17.3.2.26 ascii axis for the marker glyph, resolved through the
    /// level's `rPr` (§17.9.6) merged over the paragraph's run formatting. The
    /// renderer draws Latin marker chars (e.g. a decimal "1") with this family,
    /// so a heading whose ascii=Times renders its auto-number in Times (serif)
    /// even when eastAsia=Gothic. `None` ⇒ renderer falls back to its default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    /// ECMA-376 §17.3.2.26 eastAsia axis for the marker glyph (same resolution as
    /// `font_family`). The renderer draws CJK marker chars (e.g. an ideographic
    /// bullet) with this family. `None` ⇒ renderer falls back to `font_family`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family_east_asia: Option<String>,
    /// Internal effective numbering-level run facts. Kept outside the stable TS
    /// public model, but serialized so layout can apply the normative four-slot
    /// §17.3.2.26 selection to every scalar in mixed marker text.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_facts: Option<RunFontFacts>,
    /// ECMA-376 §17.9.24 — the numbering level rPr's `<w:color w:val>` (hex 6,
    /// lowercased like run colors). Colors the marker glyph only, never the
    /// paragraph's runs. `None` (absent `<w:color>`) lets the renderer fall
    /// back to the paragraph MARK's resolved color
    /// (`DocParagraph::paragraph_mark_color`, §17.3.1.29 — Word layers the
    /// level rPr over the mark's run properties), and finally to its default
    /// ink. An explicit `val="auto"` is `None` + `color_auto` below.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// ECMA-376 §17.3.2.6 / ST_HexColorAuto (§17.18.39) — true when the level
    /// rPr carries an EXPLICIT `<w:color w:val="auto"/>`. Auto names no
    /// concrete color but is NOT "unset": layered over the paragraph mark
    /// (§17.9.24 over §17.3.1.29) it breaks an inherited concrete mark color,
    /// so the renderer must NOT fall back to `paragraph_mark_color` and draws
    /// the automatic (default) ink instead.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub color_auto: bool,
    /// ECMA-376 §17.9.9/§17.9.20 — when the level uses a `<w:lvlPicBulletId>`,
    /// the marker is this image (zip path, e.g. `word/media/image1.gif`) drawn in
    /// place of `text`. `None` ⇒ ordinary text/glyph marker.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pic_bullet_image_path: Option<String>,
    /// MIME type of {@link pic_bullet_image_path} (e.g. `image/gif`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pic_bullet_mime_type: Option<String>,
    /// Picture-bullet marker width in pt (from the `<v:shape style="width">`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pic_bullet_width_pt: Option<f64>,
    /// Picture-bullet marker height in pt (from the `<v:shape style="height">`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pic_bullet_height_pt: Option<f64>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum DocRun {
    // Boxed: TextRun is the largest non-Shape Run variant (it carries the full
    // run-format axis incl. the eastAsia font + complex-script fields); boxing
    // keeps the enum compact (clippy::large_enum_variant). Serde flattens the
    // Box, so the JSON tag/shape is unchanged.
    Text(Box<TextRun>),
    /// Zero-advance formatting contribution of the WordprocessingML run that
    /// hosts a floating DrawingML object. Kept independent of the drawing kind
    /// because the same `<wp:anchor>` can contain a picture, chart, shape, or
    /// group, while the host character participates in line sizing exactly once.
    AnchorHost(AnchorHostMetrics),
    // Boxed with Field because these rich payloads otherwise make every run
    // reserve image/field-sized storage; Serde keeps the wire unchanged.
    Image(Box<ImageRun>),
    /// ECMA-376 §21.2 DrawingML chart embedded in a `<w:drawing>` whose
    /// `<a:graphicData uri=".../chart">` carries a `<c:chart r:id>`. The chart
    /// model is the shared [`ooxml_common::chart::ChartModel`] (the same
    /// superset pptx/xlsx emit), pre-resolved by `parse()` from the referenced
    /// `word/charts/chartN.xml` part. Boxed because `ChartModel` is large.
    Chart(Box<ChartRun>),
    /// `rename_all` on the enum only renames variant tags; the field
    /// `break_type` would otherwise serialize as snake_case while the TS
    /// side reads `breakType`. Re-apply camelCase at the variant level so
    /// the JSON tag matches.
    #[serde(rename_all = "camelCase")]
    Break {
        break_type: BreakType,
    },
    Field(Box<FieldRun>),
    // Boxed: ShapeRun is by far the largest Run variant; boxing keeps the
    // enum compact (clippy::large_enum_variant). Serde flattens the Box, so
    // the JSON tag/shape is unchanged.
    Shape(Box<ShapeRun>),
    /// An OMML equation (`m:oMath`). `display` = block (`m:oMathPara`).
    #[serde(rename_all = "camelCase")]
    Math {
        nodes: Vec<crate::math::MathNode>,
        display: bool,
        /// Resolved paragraph font size (pt) so the equation matches surrounding text.
        font_size: f64,
        /// ECMA-376 §22.1.2.88 `m:oMathPara/m:oMathParaPr/m:jc` — per-instance
        /// justification of a display equation (ST_Jc math: left|right|center|
        /// centerGroup). Only set for display math; inline `m:oMath` is `None`.
        /// Document-default (`m:defJc`) resolution is left to the renderer.
        #[serde(skip_serializing_if = "Option::is_none")]
        jc: Option<String>,
    },
    /// ECMA-376 §17.3.3.23 `<w:ptab>` — an absolute-position tab. Unlike a plain
    /// `<w:tab>` (§17.3.3.22) it ignores the paragraph's custom tab stops and the
    /// default-tab interval, advancing instead to a position derived from its own
    /// `alignment` (§17.18.71) and `relativeTo` (§17.18.73) attributes, filling
    /// the gap with the `leader` (§17.18.72) character. A separate variant (not a
    /// `"\t"` Text run) so the layout can resolve the jump geometrically.
    ///
    /// `#[serde(rename = "ptab")]`: the enum-level `rename_all = "camelCase"`
    /// treats a leading run of capitals as a single word, so `PTab` would
    /// otherwise serialize its tag as `"pTab"` (only the leading `P`
    /// lowercased). The TS discriminant union (types.ts) uses the fully
    /// lowercase `'ptab'`, matching the other single/lowercase-joined
    /// variant tags (`text`/`image`/`break`/`field`/`shape`/`math`) — so
    /// without this override every `<w:ptab>` silently failed to match any
    /// arm in the TS render switch and was dropped from the page.
    #[serde(rename = "ptab", rename_all = "camelCase")]
    PTab {
        /// ST_PTabAlignment (§17.18.71): "left" | "center" | "right".
        alignment: String,
        /// ST_PTabRelativeTo (§17.18.73): "margin" | "indent".
        relative_to: String,
        /// ST_PTabLeader (§17.18.72): "none" | "dot" | "hyphen" | "underscore" |
        /// "middleDot". Fills the space created by the tab.
        leader: String,
        /// Resolved run font size (pt) so the leader glyphs / gap height match the
        /// surrounding text — mirrors how `<w:tab>` carries the run's font.
        font_size: f64,
    },
}

/// A DrawingML line-end decoration (arrow head). ECMA-376 §20.1.8.3
/// (CT_LineEndProperties): `type` ∈ none/triangle/stealth/diamond/oval/arrow,
/// `w`/`len` ∈ sm/med/lg. Maps 1:1 to core's `ArrowEnd`.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct LineEnd {
    /// ST_LineEndType (§20.1.10.33).
    pub r#type: String,
    /// ST_LineEndWidth (§20.1.10.32). Absent in source ⇒ "med".
    pub w: String,
    /// ST_LineEndLength (§20.1.10.31). Absent in source ⇒ "med".
    pub len: String,
}

/// Resolved character formatting of the WordprocessingML run that hosts a
/// floating drawing. The drawing is out of inline flow, but Word still uses its
/// anchor character's run metrics when sizing the containing text line.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct AnchorHostMetrics {
    /// Effective run size in points (§17.3.2.38 `<w:sz>`).
    pub font_size: f64,
    /// Resolved ascii/hAnsi font face (§17.3.2.26 `<w:rFonts>`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    /// Resolved East Asian font face, kept separately so document-grid line
    /// allocation can use the Far East design metrics.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family_east_asia: Option<String>,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,
    /// Parser-private identity of the `wp:anchor` occurrence hosted by this
    /// zero-advance character. Kept off the stable TypeScript model.
    #[serde(
        rename = "__anchorOccurrenceId",
        skip_serializing_if = "Option::is_none"
    )]
    pub(crate) anchor_occurrence_id: Option<String>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorEdgesWire {
    pub(crate) top_pt: Option<f64>,
    pub(crate) top_status: AnchorValueStatusWire,
    pub(crate) right_pt: Option<f64>,
    pub(crate) right_status: AnchorValueStatusWire,
    pub(crate) bottom_pt: Option<f64>,
    pub(crate) bottom_status: AnchorValueStatusWire,
    pub(crate) left_pt: Option<f64>,
    pub(crate) left_status: AnchorValueStatusWire,
}

#[derive(Serialize, Debug, Clone, Copy, PartialEq, Eq, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) enum AnchorValueStatusWire {
    #[default]
    Missing,
    Invalid,
    Valid,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub(crate) enum AnchorAxisChoiceWire {
    #[default]
    Missing,
    Invalid,
    Align {
        value: String,
    },
    Offset {
        #[serde(rename = "valuePt")]
        value_pt: f64,
    },
    Percent {
        fraction: f64,
    },
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorAxisWire {
    pub(crate) relative_from: Option<String>,
    pub(crate) relative_from_status: AnchorValueStatusWire,
    pub(crate) choice: AnchorAxisChoiceWire,
}

impl Default for AnchorAxisWire {
    fn default() -> Self {
        Self {
            relative_from: None,
            relative_from_status: AnchorValueStatusWire::Missing,
            choice: AnchorAxisChoiceWire::Missing,
        }
    }
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorSimplePositionWire {
    pub(crate) enabled: Option<bool>,
    pub(crate) status: AnchorValueStatusWire,
    pub(crate) x_pt: Option<f64>,
    pub(crate) x_status: AnchorValueStatusWire,
    pub(crate) y_pt: Option<f64>,
    pub(crate) y_status: AnchorValueStatusWire,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorExtentWire {
    pub(crate) width_pt: Option<f64>,
    pub(crate) height_pt: Option<f64>,
    pub(crate) width_status: AnchorValueStatusWire,
    pub(crate) height_status: AnchorValueStatusWire,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorRelativeSizeAxisWire {
    pub(crate) relative_from: Option<String>,
    pub(crate) relative_from_status: AnchorValueStatusWire,
    pub(crate) fraction: Option<f64>,
    pub(crate) fraction_status: AnchorValueStatusWire,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorRelativeSizeWire {
    pub(crate) horizontal: Option<AnchorRelativeSizeAxisWire>,
    pub(crate) vertical: Option<AnchorRelativeSizeAxisWire>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorPointWire {
    pub(crate) x: Option<i64>,
    pub(crate) y: Option<i64>,
    pub(crate) raw_x: Option<String>,
    pub(crate) raw_y: Option<String>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorPolygonWire {
    pub(crate) edited: bool,
    pub(crate) coordinate_space: AnchorPolygonSpaceWire,
    pub(crate) points: Vec<AnchorPointWire>,
    pub(crate) invalid_point_count: usize,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorPolygonSpaceWire {
    pub(crate) width: u32,
    pub(crate) height: u32,
}

#[derive(Serialize, Debug, Clone, Copy, PartialEq, Eq, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) enum AnchorWrapKindWire {
    #[default]
    Missing,
    Invalid,
    None,
    Square,
    Tight,
    Through,
    TopAndBottom,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorWrapWire {
    pub(crate) kind: AnchorWrapKindWire,
    pub(crate) authored_kinds: Vec<String>,
    pub(crate) side: Option<String>,
    pub(crate) distances: AnchorEdgesWire,
    pub(crate) effect_extent: Option<AnchorEdgesWire>,
    pub(crate) polygon: Option<AnchorPolygonWire>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorBehaviorWire {
    pub(crate) behind_doc: Option<bool>,
    pub(crate) behind_doc_status: AnchorValueStatusWire,
    pub(crate) relative_height: Option<u32>,
    pub(crate) relative_height_status: AnchorValueStatusWire,
    pub(crate) locked: Option<bool>,
    pub(crate) locked_status: AnchorValueStatusWire,
    pub(crate) allow_overlap: Option<bool>,
    pub(crate) allow_overlap_status: AnchorValueStatusWire,
    pub(crate) layout_in_cell: Option<bool>,
    pub(crate) layout_in_cell_status: AnchorValueStatusWire,
}

impl Default for AnchorBehaviorWire {
    fn default() -> Self {
        Self {
            behind_doc: None,
            behind_doc_status: AnchorValueStatusWire::Missing,
            relative_height: None,
            relative_height_status: AnchorValueStatusWire::Missing,
            locked: None,
            locked_status: AnchorValueStatusWire::Missing,
            allow_overlap: None,
            allow_overlap_status: AnchorValueStatusWire::Missing,
            layout_in_cell: None,
            layout_in_cell_status: AnchorValueStatusWire::Missing,
        }
    }
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorRawTransformWire {
    pub(crate) offset_x_emu: Option<f64>,
    pub(crate) offset_y_emu: Option<f64>,
    pub(crate) extent_width_emu: Option<f64>,
    pub(crate) extent_height_emu: Option<f64>,
    pub(crate) child_offset_x_emu: Option<f64>,
    pub(crate) child_offset_y_emu: Option<f64>,
    pub(crate) child_extent_width_emu: Option<f64>,
    pub(crate) child_extent_height_emu: Option<f64>,
    pub(crate) rotation_units: Option<f64>,
    pub(crate) flip_h: Option<bool>,
    pub(crate) flip_v: Option<bool>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorResolvedChildFrameWire {
    pub(crate) offset_x_pt: f64,
    pub(crate) offset_y_pt: f64,
    pub(crate) width_pt: f64,
    pub(crate) height_pt: f64,
    pub(crate) rotation_deg: f64,
    pub(crate) flip_h: bool,
    pub(crate) flip_v: bool,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorGroupWire {
    pub(crate) child_source_id: String,
    pub(crate) source_index: usize,
    pub(crate) source_count: usize,
    pub(crate) transform_chain: Vec<AnchorRawTransformWire>,
    pub(crate) child_transform: Option<AnchorRawTransformWire>,
    pub(crate) resolved_child_frame: AnchorResolvedChildFrameWire,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct AnchorAcquisitionWire {
    pub(crate) occurrence_id: String,
    pub(crate) simple_position: AnchorSimplePositionWire,
    pub(crate) horizontal: AnchorAxisWire,
    pub(crate) vertical: AnchorAxisWire,
    pub(crate) extent: AnchorExtentWire,
    pub(crate) parent_effect_extent: AnchorEdgesWire,
    pub(crate) anchor_distances: AnchorEdgesWire,
    pub(crate) relative_size: AnchorRelativeSizeWire,
    pub(crate) wrap: AnchorWrapWire,
    pub(crate) behavior: AnchorBehaviorWire,
    pub(crate) group: Option<AnchorGroupWire>,
}

/// A drawn shape (wps:wsp inside wp:anchor). Positioned like an anchor image
/// and rendered via core's buildCustomPath + paint primitives.
#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct ShapeRun {
    /// pt
    pub width_pt: f64,
    /// pt
    pub height_pt: f64,
    /// anchor X (pt). Ignored if `anchor_x_align` is set.
    pub anchor_x_pt: f64,
    /// anchor Y (pt). Ignored if `anchor_y_align` is set.
    pub anchor_y_pt: f64,
    pub anchor_x_from_margin: bool,
    pub anchor_y_from_para: bool,
    /// ECMA-376 §20.4.3.1 wp:align (positionH/wp:align). When present the
    /// renderer centers / left-aligns / right-aligns the shape within the
    /// container indicated by `anchor_x_from_margin`. Values: "left",
    /// "center", "right" (others fall back to "left").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_x_align: Option<String>,
    /// Vertical equivalent of anchor_x_align (positionV/wp:align).
    /// Values: "top", "center", "bottom".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_y_align: Option<String>,
    /// ECMA-376 §20.4.2.7 wp14:pctPosHOffset / pctPosVOffset — fraction
    /// of the relativeFrom container's width / height in 1/100,000ths of
    /// a percent. When set, renderer ignores anchor_x_pt / anchor_y_pt
    /// and computes the offset as `container_size * pct / 100000`. The
    /// container is `anchor_*_relative_from` (page / margin / etc.).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pct_pos_h: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pct_pos_v: Option<f64>,
    /// `<wp:positionH/V relativeFrom="…">` — overrides the looser
    /// `anchor_x_from_margin`/`anchor_y_from_para` booleans for the
    /// pct-pos and align paths. Common values: "page", "margin",
    /// "leftMargin"/"rightMargin", "insideMargin"/"outsideMargin",
    /// "topMargin"/"bottomMargin", "paragraph", "line".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_x_relative_from: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_y_relative_from: Option<String>,
    /// ECMA-376 §20.4.2.18 wp14:sizeRelH/sizeRelV — width/height as a
    /// fraction of the relativeFrom container, in 1/100,000ths of a percent.
    /// When set, the renderer uses this in place of `width_pt` / `height_pt`
    /// for layout (text frame, align centering, shape draw rect). The
    /// container is resolved by `width_relative_from` / `height_relative_from`
    /// using the same rules as anchor positioning. `pct == 0` is treated as
    /// "fall back to extent" (matches Word's empirical behavior).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_pct: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height_pct: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_relative_from: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height_relative_from: Option<String>,
    /// Parent wgp group dimensions (pt). Set only when this shape is a child
    /// of a `<wpg:wgp>` group; `None` for standalone wsp anchors. The renderer
    /// uses these for align/pctPos math so the GROUP is positioned within
    /// the relativeFrom container, and per-shape offsets (`anchor_x_pt` /
    /// `anchor_y_pt`) carry the child's offset within the group.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub group_width_pt: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub group_height_pt: Option<f64>,
    /// If true, draw the shape behind text (wp:anchor behindDoc="1"). Renderer
    /// should draw background shapes BEFORE body text.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub behind_doc: bool,
    /// ECMA-376 §20.4.2.3 `wp:anchor/@relativeHeight`: stacking order among
    /// anchors sharing the same behindDoc value. Lower value = drawn first.
    /// Group children add their document-order index to the anchor value.
    pub z_order: u32,
    /// normalized [0,1] custom path commands (one or more sub-paths). Empty
    /// when `preset_geometry` is set; the renderer chooses between
    /// buildCustomPath (custGeom) and buildShapePath (prstGeom).
    pub subpaths: Vec<Vec<PathCmd>>,
    /// OOXML <a:prstGeom prst="..."> name (e.g. "rect", "ellipse",
    /// "roundRect", "rtTriangle"). Empty when the shape is custGeom.
    /// `adj_values` carries <a:gd name="adj{n}"> values in adj1..adj8 order
    /// (0–100000 scale), preserving omitted named guides as `None` so the shared
    /// preset engine can fall back to the preset's declared default per index.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub preset_geometry: Option<String>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub adj_values: Vec<Option<f64>>,
    /// Fill (solid or gradient). None = no fill.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<ShapeFill>,
    /// stroke hex. None = no stroke.
    pub stroke: Option<String>,
    /// stroke width in pt.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub stroke_width: f64,
    /// `<a:ln><a:prstDash val>` (ECMA-376 §20.1.8.48). None ⇒ solid.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_dash: Option<String>,
    /// VML `<v:stroke endcap>` / DrawingML line cap, normalized for Canvas.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_cap: Option<String>,
    /// `<a:ln><a:headEnd>` decoration (line start). None ⇒ no head.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub head_end: Option<LineEnd>,
    /// `<a:ln><a:tailEnd>` decoration (line end). None ⇒ no tail.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tail_end: Option<LineEnd>,
    /// rotation in degrees (clockwise).
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub rotation: f64,
    /// Horizontal flip (`<a:xfrm flipH="1">`, §20.1.7.6). Mirrors the shape
    /// about its vertical centre line — also swaps a connector's start/end so
    /// arrow heads land on the correct tip.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub flip_h: bool,
    /// Vertical flip (`<a:xfrm flipV="1">`, §20.1.7.6).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub flip_v: bool,
    /// Text blocks from <wps:txbx><w:txbxContent> — text rendered INSIDE the
    /// shape's bounding box (ECMA-376 §17.3.4.7). Each block is one paragraph
    /// reduced to plain text + the first run's formatting; advanced layout
    /// (numbering, paragraph styles, mixed-format runs within a single
    /// paragraph) isn't supported in shape bodies yet.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub text_blocks: Vec<ShapeText>,
    /// ECMA-376 §20.1.4.1.17 `<wps:style><a:fontRef>` — the shape's DEFAULT text
    /// color (hex, no `#`). A text-box run that sets no explicit `<w:color>`
    /// takes this before falling back to the document/theme default (black). The
    /// same DrawingML `<a:fontRef>` PowerPoint resolves for placeholder text
    /// (pptx `shape.rs`); Word applies it to `<wps:txbx>` runs. This resolves the
    /// COLOR axis only — the `fontRef @idx` (major/minor/none) font-face effect is
    /// out of scope here (fonts already resolve through `rFonts`/docDefaults).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_text_color: Option<String>,
    /// Vertical anchor for the shape text box: "t" (top), "ctr" (center),
    /// "b" (bottom). Read from <wps:bodyPr @anchor>. Default = "t".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_anchor: Option<String>,
    /// Text auto-fit mode from the `<wps:bodyPr>` child (ECMA-376 §21.1.2.1.1),
    /// normalized to the shared core vocabulary (core `src/types/common.ts`
    /// `autoFit`): "none" (`<a:noAutofit/>`, fixed box — overflowing text is
    /// CLIPPED to the box), "sp" (`<a:spAutoFit/>`, box grows to text), or
    /// "norm" (`<a:normAutofit/>`, text shrinks to fit). Absent ⇒ None (renderer
    /// treats as the spec default: overflow visible / no clip).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_autofit: Option<String>,
    /// Text-body flow direction from `<wps:bodyPr vert>` (ECMA-376 §20.1.10.83
    /// ST_TextVerticalType): "vert" (all glyphs 90° CW, chars T→B, lines R→L),
    /// "vert270" (all glyphs 270° CW = 90° CCW, chars B→T, lines L→R), "eaVert"
    /// (East-Asian upright: CJK stands upright, non-EA rotated 90°). "horz"/absent
    /// ⇒ None (horizontal). Other values ("mongolianVert", "wordArtVert", …) are
    /// carried verbatim; the renderer falls them back to horizontal until handled.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_vert: Option<String>,
    /// Body-pr text insets in pt (left/top/right/bottom). Default 0 each.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub text_inset_l: f64,
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub text_inset_t: f64,
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub text_inset_r: f64,
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub text_inset_b: f64,
    /// Wrap mode matching ImageRun.wrap_mode semantics.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_mode: Option<String>,
    /// distT (top padding, pt). Anchor-only. Mirrors ImageRun.dist_top so an
    /// anchored wrap-shape reserves the same float-exclusion band as an image
    /// (ECMA-376 §20.4.2.x — distT/B/L/R are the min text↔object distance).
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_top: f64,
    /// distB (bottom padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_bottom: f64,
    /// distL (left padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_left: f64,
    /// distR (right padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_right: f64,
    /// wrapSquare/wrapTight "wrapText" attribute: "bothSides" | "left" | "right"
    /// | "largest". Defaults to "bothSides" when absent. Mirrors ImageRun.wrap_side.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_side: Option<String>,
    /// ECMA-376 Part 4 §19.1.2.23 `<v:textpath>` — WordArt text plus its
    /// resolved CT_Path / CT_TextPath controls. `textpathok` and `on` govern
    /// display; `fitshape` / `fitpath` govern fitting. Rotation and paint remain
    /// shape facts. `None` for an ordinary VML shape.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_path: Option<TextPath>,
    /// ECMA-376 Part 4 §19.1.2.5 `<v:fill opacity>` — the fill's alpha in
    /// [0.0, 1.0] (default 1.0 ⇒ opaque). Used with `text_path` to draw the
    /// watermark semi-transparently. `None` ⇒ fully opaque.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_opacity: Option<f64>,
    #[serde(
        rename = "__anchorAcquisition",
        skip_serializing_if = "Option::is_none"
    )]
    pub(crate) anchor_acquisition: Option<AnchorAcquisitionWire>,
}

/// ECMA-376 Part 4 §19.1.2.23 `<v:textpath>` — a WordArt vector text path.
/// The authored font properties are preserved together with resolved controls;
/// downstream acquisition applies display and fitting only when those controls
/// request it.
#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TextPath {
    /// The `string` attribute — the watermark text (e.g. "DRAFT").
    pub string: String,
    /// `font-family` from the textpath `style` CSS `font`/`font-family` (quotes
    /// stripped). `None` ⇒ the renderer's default family.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    /// `font-weight:bold` (or the `bold` keyword in the `font` shorthand).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    /// `font-style:italic` (or the `italic` keyword in the `font` shorthand).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,
    /// Resolved transitional VML control facts. These are parser/layout
    /// implementation detail: serde flattens them into the wire object so the
    /// internal TypeScript parser-model can acquire WordArt deterministically,
    /// while the stable public `TextPath` declaration remains unchanged.
    #[serde(flatten)]
    pub(crate) vml: VmlTextPathFacts,
}

/// Private resolved facts from CT_Path / CT_TextPath. Parser construction fills
/// every boolean with `Some`, including the specified false defaults. `None` is
/// reserved for a non-parser/default Rust value and therefore remains distinct
/// from explicit parser provenance on the serialized wire object.
#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct VmlTextPathFacts {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) text_path_ok: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) on: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) fit_shape: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) fit_path: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) trim: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) x_scale: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) font_size_pt: Option<f64>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(tag = "fillType", rename_all = "camelCase")]
pub enum ShapeFill {
    Solid {
        color: String,
    },
    #[serde(rename_all = "camelCase")]
    Gradient {
        stops: Vec<GradientStop>,
        /// degrees: 0 = left→right, 90 = top→bottom
        angle: f64,
        /// "linear" | "radial"
        grad_type: String,
    },
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct GradientStop {
    /// 0.0–1.0
    pub position: f64,
    /// hex 6-char
    pub color: String,
}

/// Custom geometry path command (shape rendering). Mirrors the pptx
/// PathCmd type to keep JSON output compatible with core's buildCustomPath.
#[derive(Serialize, Debug, Clone)]
#[serde(tag = "cmd", rename_all = "camelCase")]
pub enum PathCmd {
    MoveTo {
        x: f64,
        y: f64,
    },
    LineTo {
        x: f64,
        y: f64,
    },
    CubicBezTo {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        x: f64,
        y: f64,
    },
    /// Elliptical arc (all angles in degrees).
    ///
    /// The enum-level `#[serde(tag = ..., rename_all = "camelCase")]` renames the
    /// variant *tag* (`ArcTo` → `arcTo`) but NOT the variant's struct fields, so
    /// a per-variant `rename_all` is required for the multi-word fields. Without
    /// it the JSON carried `st_ang`/`sw_ang`, while the TS `PathCmd`
    /// (packages/docx/src/types.ts) and core's `buildCustomPath`
    /// (packages/core/src/shape/custGeom.ts) read `stAng`/`swAng` — the mismatch
    /// left the angles `undefined`, producing `NaN` coordinates and a missing
    /// arc. (The degenerate-arc guard `wr<=0||hr<=0` short-circuits before the
    /// angles are read, so only non-degenerate arcs surface this.) Mirrors the
    /// pptx fix (#489) and the per-variant `rename_all` used by sibling enums.
    #[serde(rename_all = "camelCase")]
    ArcTo {
        wr: f64,
        hr: f64,
        st_ang: f64,
        sw_ang: f64,
    },
    Close,
}

#[derive(Serialize, Debug, Clone, Default, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct RunFontAxisValues {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ascii: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub high_ansi: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub east_asia: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub complex_script: Option<String>,
}

#[derive(Serialize, Debug, Clone, Default, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct RunFontAxisPresence {
    pub ascii: bool,
    pub high_ansi: bool,
    pub east_asia: bool,
    pub complex_script: bool,
}

#[derive(Serialize, Debug, Clone, Default, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct RunFontSlots {
    pub direct: RunFontAxisValues,
    pub theme: RunFontAxisValues,
    pub theme_present: RunFontAxisPresence,
}

/// Internal projection of the effective run properties needed by the DOCX text
/// service. This is shared by numbering-level text (§17.9.6) and paragraph-mark
/// line probes (§17.3.1.29); the stable public TypeScript model intentionally
/// remains unchanged.
#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct RunFontFacts {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family_high_ansi: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_slots: Option<RunFontSlots>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family_east_asia: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_hint: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtl: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cs: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family_cs: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_cs: Option<f64>,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold_cs: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic_cs: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub lang_bidi: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub lang_east_asia: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub kerning: Option<f64>,
}

#[derive(Serialize, Debug, Clone, Copy, Default, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub(crate) enum TypographyValueStatusWire {
    #[default]
    Missing,
    Invalid,
    Valid,
}

/// Parser-private lexical + normalized value. Keeping the lexical form is the
/// diagnostic seam: required/enum values that are absent or malformed remain
/// distinguishable from valid zero/false values instead of being guessed.
#[derive(Serialize, Debug, Clone, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TypographyValueWire<T> {
    pub status: TypographyValueStatusWire,
    pub raw: Option<String>,
    pub value: Option<T>,
}

#[derive(Serialize, Debug, Clone, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct CtBorderTypographyWire {
    pub val: TypographyValueWire<String>,
    pub color: TypographyValueWire<String>,
    pub theme_color: TypographyValueWire<String>,
    pub theme_tint: TypographyValueWire<String>,
    pub theme_shade: TypographyValueWire<String>,
    pub size_pt: TypographyValueWire<f64>,
    pub space_pt: TypographyValueWire<f64>,
    pub shadow: TypographyValueWire<bool>,
    pub frame: TypographyValueWire<bool>,
}

#[derive(Serialize, Debug, Clone, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct UnderlineTypographyWire {
    pub val: TypographyValueWire<String>,
    pub color: TypographyValueWire<String>,
    pub theme_color: TypographyValueWire<String>,
    pub theme_tint: TypographyValueWire<String>,
    pub theme_shade: TypographyValueWire<String>,
}

#[derive(Serialize, Debug, Clone, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TypographyLanguagesWire {
    pub east_asia: Option<String>,
    pub bidi: Option<String>,
}

#[derive(Serialize, Debug, Clone, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct EastAsianLayoutTypographyWire {
    pub vert: Option<bool>,
    pub vert_compress: Option<bool>,
    pub combine: Option<bool>,
    pub combine_brackets: TypographyValueWire<String>,
}

#[derive(Serialize, Debug, Clone, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct RubyGuideRunTypographyWire {
    pub text: String,
    pub font_family: Option<String>,
    pub font_size_pt: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    pub color: Option<String>,
    pub language: Option<String>,
}

#[derive(Serialize, Debug, Clone, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct RubyTypographyWire {
    pub align: TypographyValueWire<String>,
    pub base_font_size_pt: TypographyValueWire<f64>,
    pub raise_pt: TypographyValueWire<f64>,
    pub language: TypographyValueWire<String>,
    pub guide_runs: Vec<RubyGuideRunTypographyWire>,
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct RevisionTypographyWire {
    pub kind: String,
    pub id: TypographyValueWire<String>,
    pub author: Option<String>,
    pub date: Option<String>,
}

#[derive(Serialize, Debug, Clone, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct RunTypographyWire {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<UnderlineTypographyWire>,
    pub strike: bool,
    pub double_strike: bool,
    pub caps: bool,
    pub small_caps: bool,
    pub color_auto: bool,
    pub vertical_align: TypographyValueWire<String>,
    pub position_pt: TypographyValueWire<f64>,
    pub snap_to_grid: Option<bool>,
    pub character_spacing_pt: Option<f64>,
    pub character_scale: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fit_text: Option<FitTextSpecWire>,
    pub kerning_threshold_pt: Option<f64>,
    pub emphasis: TypographyValueWire<String>,
    pub languages: TypographyLanguagesWire,
    pub east_asian_layout: EastAsianLayoutTypographyWire,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border: Option<CtBorderTypographyWire>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ruby: Option<RubyTypographyWire>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub revision: Option<RevisionTypographyWire>,
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FitTextSpecWire {
    pub val_twips: f64,
    pub id: Option<String>,
}

#[derive(Serialize, Debug, Clone, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ParagraphTypographyBordersWire {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<CtBorderTypographyWire>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<CtBorderTypographyWire>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<CtBorderTypographyWire>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<CtBorderTypographyWire>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub between: Option<CtBorderTypographyWire>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bar: Option<CtBorderTypographyWire>,
}

#[derive(Serialize, Debug, Clone, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ParagraphTypographyWire {
    pub borders: ParagraphTypographyBordersWire,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct FieldRun {
    /// "page" | "numPages" | "other"
    pub field_type: String,
    /// original instruction text (e.g. "PAGE \\* MERGEFORMAT")
    pub instruction: String,
    /// fallback text captured between fldChar separate and end (shown if field_type is "other")
    pub fallback_text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    /// pt
    pub font_size: f64,
    pub color: Option<String>,
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family_high_ansi: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_slots: Option<RunFontSlots>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family_east_asia: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_hint: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtl: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cs: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family_cs: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_cs: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold_cs: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic_cs: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub lang_bidi: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub lang_east_asia: Option<String>,
    pub background: Option<String>,
    /// "super" | "sub" | None
    pub vert_align: Option<String>,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub all_caps: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub small_caps: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub double_strikethrough: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub highlight: Option<String>,
    /// ECMA-376 §17.3.2.12 `<w:em w:val>` — emphasis (boten / 圏点) mark, mirrors
    /// `TextRun::emphasis_mark` (§17.18.24 ST_Em). `None` when unset or
    /// `val="none"`. A field's displayed text (its resolved PAGE/NUMPAGES value
    /// or fallback text) draws through the same per-glyph stamp as a plain run.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub emphasis_mark: Option<String>,
    #[serde(
        rename = "__typographyAcquisition",
        skip_serializing_if = "Option::is_none"
    )]
    pub(crate) typography_acquisition: Option<RunTypographyWire>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TextRun {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    /// ECMA-376 §17.3.2.40 `<w:u w:val>` — the raw ST_Underline (§17.18.99) style
    /// value (`double` / `thick` / `dotted` / `wave` / `dashLong` / …). `None`
    /// for the plain single rule (or no underline). The renderer normalizes this
    /// WordprocessingML vocabulary to the shared DrawingML ST_TextUnderlineType
    /// (§20.1.10.82) that `core::drawUnderline` dispatches on.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline_style: Option<String>,
    /// ECMA-376 §17.3.2.40 `<w:u w:color>` — underline-only colour (hex 6, or the
    /// literal `auto`). `None` ⇒ the underline follows the glyph colour.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline_color: Option<String>,
    pub strikethrough: bool,
    /// pt
    pub font_size: f64,
    pub color: Option<String>,
    pub font_family: Option<String>,
    /// ECMA-376 §17.3.2.26 hAnsi axis, preserved independently from ascii.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family_high_ansi: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_slots: Option<RunFontSlots>,
    /// ECMA-376 §17.3.2.26 eastAsia axis (`<w:rFonts w:eastAsia>`), resolved
    /// through the style chain + docDefaults. CJK characters in this run render
    /// with this family; `font_family` keeps the conflated single-font fallback
    /// (ascii → eastAsia) for any path that does not split per character. The
    /// renderer routes consecutive CJK code points to this axis (the same per-
    /// script rule `ShapeTextRun` already uses), so a Gothic eastAsia title sits
    /// alongside a serif ascii number with no name heuristics. `None` ⇒ renderer
    /// falls back to `font_family`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family_east_asia: Option<String>,
    /// Resolved ECMA-376 §17.3.2.26 rFonts@hint.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_hint: Option<String>,
    pub is_link: bool,
    pub background: Option<String>,
    /// ECMA-376 §17.3.2.6 — `<w:color w:val="auto"/>` was set on this run. The
    /// renderer resolves the glyph color from the effective background
    /// (implementation-defined black/white pick; no normative algorithm) when
    /// this is true and no concrete `color` is present.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub color_auto: bool,
    /// ECMA-376 §17.3.2.4 `<w:bdr>` — a run-level border (box). `None` when the
    /// run has no border.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border: Option<RunBorder>,
    /// "super" | "sub" | None
    pub vert_align: Option<String>,
    /// Target URL for hyperlinks (from relationships.xml), None if not a link or no URL
    pub hyperlink: Option<String>,
    /// ECMA-376 §17.16.23 `<w:hyperlink w:anchor>` — internal bookmark name this
    /// link jumps to (a `<w:bookmarkStart w:name>` in the same document). Set for
    /// an internal cross-reference / TOC entry. When a `<w:hyperlink>` carries
    /// BOTH `r:id` and `w:anchor`, the external URL wins for `hyperlink` while the
    /// anchor is still recorded here. `None` when the link has no anchor.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hyperlink_anchor: Option<String>,
    /// Transform all characters to uppercase (w:caps)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub all_caps: bool,
    /// Render as small capitals (uppercase at ~80% size, w:smallCaps)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub small_caps: bool,
    /// Double strikethrough (w:dstrike)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub double_strikethrough: bool,
    /// OOXML highlight color name: "yellow" | "cyan" | "green" | ... (w:highlight)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub highlight: Option<String>,
    /// ECMA-376 §17.3.2.12 `<w:em w:val>` — emphasis (boten / 圏点) mark drawn on
    /// every non-space character of the run (§17.18.24 ST_Em). One of "dot"
    /// (filled dot above), "comma" (sesame/comma above), "circle" (hollow circle
    /// above), or "underDot" (filled dot below, in horizontal writing). `None`
    /// when unset or `val="none"`. The renderer paints the mark per glyph after
    /// the text; it does not alter the glyph advance.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub emphasis_mark: Option<String>,
    /// ECMA-376 §17.3.3.25 ruby annotation (furigana). Set on the rubyBase
    /// text run; renders as a small inline annotation above the base glyphs.
    /// `text` is the annotation string (e.g. "すわ" above "坐"). `font_size_pt`
    /// is the annotation's font size in pt — Word stores this as `<w:hps>`
    /// (half-points) inside the rubyPr.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ruby: Option<RubyAnnotation>,
    /// ECMA-376 §17.13.5 — set when this run sits inside a `<w:ins>` or
    /// `<w:del>` block. The renderer paints insertions with a per-author
    /// underline and deletions with a per-author strikethrough so reviewers
    /// can see tracked edits inline.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub revision: Option<RunRevision>,
    /// ECMA-376 §17.3.2.30 `<w:rtl>` — complex-script / right-to-left run.
    /// `Some(true)` = RTL, `Some(false)` = explicitly LTR, `None` = unspecified.
    /// The renderer marks a `Some(true)` run as RTL for the UAX#9 pass (forcing
    /// complex-script shaping) and draws it with `ctx.direction = 'rtl'` so the
    /// glyph order is mirrored.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtl: Option<bool>,
    /// ECMA-376 §17.3.2.7 `<w:cs/>` — complex-script run toggle: cs
    /// formatting applies to ALL characters of the run (§17.3.2.26).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cs: Option<bool>,
    /// ECMA-376 §17.3.2.26 `<w:rFonts w:cs>` — complex-script typeface (theme
    /// references resolved to a literal family). `None` when unspecified.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family_cs: Option<String>,
    /// ECMA-376 §17.3.2.39 `<w:szCs>` — complex-script font size in pt (same
    /// units as `font_size`). `None` when unspecified.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_cs: Option<f64>,
    /// ECMA-376 §17.3.2.3 `<w:bCs>` — complex-script bold toggle. `None` when
    /// unspecified (renderer falls back to `bold`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold_cs: Option<bool>,
    /// ECMA-376 §17.3.2.17 `<w:iCs>` — complex-script italic toggle. `None`
    /// when unspecified (renderer falls back to `italic`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic_cs: Option<bool>,
    /// ECMA-376 §17.3.2.20 `<w:lang w:bidi>` — the complex-script (RTL) language
    /// tag, lower-cased (e.g. "ar-sa", "ae-ar", "he-il"). Used to decide whether
    /// European digits in a complex-script run are classified as AN (Word's
    /// Arabic/Hebrew digit ordering). `None` when unspecified.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub lang_bidi: Option<String>,
    /// Resolved ECMA-376 §17.3.2.20 w:lang/@w:eastAsia.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub lang_east_asia: Option<String>,
    /// ECMA-376 §17.3.2.34 `<w:snapToGrid>` — run participation in the section
    /// character grid. `Some(false)` opts this run out; `None` inherits.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub snap_to_grid: Option<bool>,
    /// ECMA-376 §17.3.2.35 `<w:spacing w:val>` — character-spacing adjustment in
    /// POINTS (signed): the extra pitch added after each character before the
    /// next. The renderer feeds it to `ctx.letterSpacing` on BOTH the measure and
    /// paint passes so wrapping/pagination stay consistent. `None` = no extra
    /// pitch.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub char_spacing: Option<f64>,
    /// ECMA-376 §17.3.2.14 `<w:fitText w:val>` — manual run-width target in
    /// TWIPS (1/20 pt). Consecutive runs with the same `fit_text_id` form one
    /// region; an id-less run is standalone.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fit_text_val: Option<f64>,
    /// ECMA-376 §17.3.2.14 `<w:fitText w:id>` — arbitrary-precision signed
    /// XSD integer serialized as a string for exact equality linking.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fit_text_id: Option<String>,
    /// ECMA-376 §17.3.2.43 `<w:w w:val>` — horizontal text scale as a FRACTION of
    /// normal character width (e.g. 0.67 = 67%, 2.0 = 200%). Stretches each
    /// glyph's width (not the gap). `None` = 100%.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub char_scale: Option<f64>,
    /// ECMA-376 §17.3.2.24 `<w:position w:val>` — baseline raise (positive) /
    /// lower (negative) in POINTS, without changing the font size or line box.
    /// `None` = no shift.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<f64>,
    /// ECMA-376 §17.3.2.19 `<w:kern w:val>` — font-kerning threshold in POINTS
    /// (the smallest font size that is kerned). Presence enables kerning; `None`
    /// = kerning off (the hierarchy default). `Some(0.0)` = kern at all sizes.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub kerning: Option<f64>,
    /// ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:vert>` — horizontal-in-vertical
    /// (縦中横 / tate-chū-yoko). `Some(true)` means that in a VERTICAL (tbRl) page
    /// this run's characters are laid out horizontally within one cell of the
    /// vertical line. `None`/absent = normal vertical stacking. Inert in a
    /// horizontal page (the property is only meaningful in vertical text).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub east_asian_vert: Option<bool>,
    /// ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:vertCompress>` — compress the
    /// horizontally-laid-out (縦中横) run to fit the existing line height. Ignored
    /// unless `east_asian_vert` is set. `None`/absent = not compressed.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub east_asian_vert_compress: Option<bool>,
    /// ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:combine>` — two-lines-in-one.
    /// PARSED for completeness; not yet rendered (no fixture). `None`/absent = off.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub east_asian_combine: Option<bool>,
    /// ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:combineBrackets>` (§17.18.8) —
    /// bracket style around two-lines-in-one text. PARSED for completeness; the
    /// two-lines-in-one draw is a follow-up. `None` = no brackets / `none`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub east_asian_combine_brackets: Option<String>,
    /// ECMA-376 §17.11.6 / §17.11.7 / §17.11.16 / §17.11.17 — set when this run
    /// is a footnote/endnote reference mark (`<w:footnoteReference>` in the body,
    /// `<w:footnoteRef>` inside the note content, and the endnote equivalents).
    /// `text` holds the raw `@w:id` as a fallback; the renderer overrides the
    /// displayed glyph with the note's sequential number (the displayed number is
    /// the note's 1-based position, not the raw id). `None` for ordinary runs.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub note_ref: Option<NoteRef>,
    #[serde(
        rename = "__typographyAcquisition",
        skip_serializing_if = "Option::is_none"
    )]
    pub(crate) typography_acquisition: Option<RunTypographyWire>,
}

/// A footnote / endnote reference marker. The displayed number is resolved by
/// the renderer from the note's sequential position (ECMA-376 numbering), so
/// only the kind + id need to travel through the model.
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct NoteRef {
    /// "footnote" | "endnote".
    pub kind: String,
    /// The `@w:id` linking the marker to its note in footnotes.xml / endnotes.xml.
    pub id: String,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct RunRevision {
    /// "insertion" or "deletion".
    pub kind: String,
    /// `<w:ins w:author>` / `<w:del w:author>`. Used by the renderer to pick
    /// a stable per-author colour (modulo a fixed palette).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    /// `<w:ins w:date>` / `<w:del w:date>` ISO-8601 timestamp.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    /// ECMA-376 §17.13.5 requires `w:id`. It stays off the stable public
    /// revision shape and is projected through the private typography wire.
    #[serde(skip)]
    pub(crate) typography_id: TypographyValueWire<String>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct RubyAnnotation {
    pub text: String,
    /// pt
    pub font_size_pt: f64,
    /// Distance “between the phonetic guide base text and the phonetic guide text”
    /// in pt (ECMA-376 §17.3.3.12).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hps_raise_pt: Option<f64>,
    /// Rich rubyPr/rt facts are parser acquisition data, not a public model
    /// expansion. The owning TextRun projects them through its private wire.
    #[serde(skip)]
    pub(crate) typography: Option<RubyTypographyWire>,
}

/// One formatting run (`<w:r>`) inside a shape-text paragraph. Mirrors the
/// fields of {@link ShapeText} that carry character formatting, resolved
/// through the SAME chain (`parse_run_fmt` + docDefaults font fallback). The
/// renderer lays a paragraph's `runs` out as rich text (per-run font), so a
/// bold label followed by non-bold body text keeps each run's formatting
/// instead of collapsing to the first run's.
#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct ShapeTextRun {
    pub text: String,
    pub font_size_pt: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// ECMA-376 §17.3.2.26 ascii axis (`<w:rFonts w:ascii>`), resolved through
    /// docDefaults. The renderer draws Latin letters/digits in this run with this
    /// family.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    /// ECMA-376 §17.3.2.26 eastAsia axis (`<w:rFonts w:eastAsia>`), resolved
    /// through docDefaults. The renderer draws CJK characters in this run with
    /// this family (falling back to `font_family` when absent). Splitting the two
    /// axes lets a serif ascii face and a gothic eastAsia face coexist in one run
    /// (e.g. serif digits inside a gothic Japanese title).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family_east_asia: Option<String>,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,
    /// ECMA-376 §17.3.3.25 w:ruby inside VML/text-box content. Carried through
    /// the same shape rich-text path as normal run formatting so the renderer can
    /// center the annotation above the base glyphs instead of flattening it into
    /// the paragraph text.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ruby: Option<RubyAnnotation>,
}

/// One paragraph of text rendered inside a shape (`<wps:txbx><w:txbxContent>`).
/// The single `text`/format fields carry the concatenated paragraph string and
/// the FIRST run's effective formatting (kept for backward compatibility with
/// existing consumers and the image-block path); `runs` additionally preserves
/// PER-RUN formatting so the renderer can draw mixed bold/non-bold runs as rich
/// text (ECMA-376 §17.3.2 — each `<w:r>` resolves its own rPr).
#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct ShapeText {
    pub text: String,
    pub font_size_pt: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,
    /// Per-run formatting for this paragraph (one entry per `<w:r>` that carries
    /// text). Empty for an image-only paragraph (the image path is unchanged).
    /// The renderer prefers `runs` (rich layout) over the single format fields
    /// when non-empty.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub runs: Vec<ShapeTextRun>,
    /// ECMA-376 §17.9 numbering for text-box paragraphs. Word applies the same
    /// paragraph numbering model inside `<w:txbxContent>` as in the body, so list
    /// markers (for example sample-33's `※`) must travel separately from the run
    /// text and render in the hanging margin.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub numbering: Option<Box<NumberingInfo>>,
    /// Paragraph alignment ("left" | "center" | "right" | "both").
    pub alignment: String,
    /// ECMA-376 §17.3.1.33 `<w:spacing w:before>` of this text-box paragraph, in
    /// pt. Word reserves it ABOVE the paragraph inside the text box (e.g.
    /// sample-13's "Journal homepage" line sits 50 pt below the box top because
    /// its txbxContent paragraph carries `w:before="1000"`). 0 when absent.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub space_before: f64,
    /// ECMA-376 §17.3.1.33 `<w:spacing w:after>` of this text-box paragraph, in
    /// pt — reserved below the paragraph. 0 when absent.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub space_after: f64,
    /// ECMA-376 §17.3.1.33 `<w:spacing w:line>` line-spacing value, resolved
    /// through the style chain (incl. docDefaults). Encoded per `line_spacing_rule`:
    /// "auto" ⇒ a MULTIPLIER on the natural line box (276/240 = 1.15); "exact" /
    /// "atLeast" ⇒ pt. Absent ⇒ single spacing (natural line box).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_spacing_val: Option<f64>,
    /// Line-spacing rule: "auto" | "exact" | "atLeast" (see `line_spacing_val`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_spacing_rule: Option<String>,
    /// ECMA-376 §17.3.1.12 `<w:ind w:left/@start>` — paragraph left indent (pt).
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub indent_left: f64,
    /// ECMA-376 §17.3.1.12 `<w:ind w:right/@end>` — paragraph right indent (pt).
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub indent_right: f64,
    /// `<w:ind>` first-line indent (pt, SIGNED: w:firstLine positive, w:hanging
    /// negative). Mirrors the body renderer, which honors a signed hanging
    /// first-line indent list-independently (Word's behaviour).
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub indent_first: f64,
    /// ECMA-376 §17.3.1.37 `<w:tabs>` — explicit tab stops of this text-box
    /// paragraph, resolved through the style chain like the body paragraph's
    /// `tab_stops`. Empty ⇒ only the automatic default-tab grid applies. Carried
    /// so the shape text is laid out by the SAME line engine the body uses (tabs
    /// were previously dropped on this path, so a `\t` in a text box collapsed to
    /// nothing).
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub tab_stops: Vec<TabStop>,
    /// ECMA-376 §17.3.1.6 `<w:bidi>` — right-to-left text-box paragraph, resolved
    /// through the style chain like the body paragraph's `bidi`. `Some(true)` =
    /// RTL, `Some(false)` = explicitly LTR, `None` = unspecified (inherit). The
    /// renderer consumes it as the paragraph base direction for the UAX#9
    /// reordering pass (the body renderer reads the identical field).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi: Option<bool>,
    /// ECMA-376 §17.3.1.9 `<w:contextualSpacing>` — resolved through the style
    /// chain (direct `<w:pPr>` wins, else the paragraph style incl. docDefaults).
    /// When set, Word suppresses `space_before`/`space_after` between this
    /// text-box paragraph and an ADJACENT paragraph that shares its `style_id`
    /// and also sets the toggle — identical to the body paragraph's
    /// `contextual_spacing` (parser.rs `resolve_para` / renderer `contextualSuppressed`).
    /// Without it a `<w:contextualSpacing/>` ListParagraph list inside a fixed
    /// text box kept inheriting the docDefault `after=160` (8 pt) gap, which
    /// inflated its line pitch and clipped the trailing line (sample-32).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub contextual_spacing: bool,
    /// Resolved paragraph style id of this text-box paragraph: the explicit
    /// `<w:pStyle w:val>` when present, else the document default paragraph style
    /// (`w:default="1"`), else "Normal" — the SAME stable id the body path stamps
    /// (parser.rs ~2578) so §17.3.1.9 `contextual_spacing` can group adjacent
    /// same-style paragraphs even under locale-specific default style ids.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
    /// Embedded zip path of an inline image living inside this text-box
    /// paragraph (`<w:drawing><wp:inline>…<a:blip r:embed>`), e.g.
    /// `word/media/image1.emf`. `None` for a text-only paragraph. Resolved
    /// the same way body images are (`resolve_blip_urls`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub image_path: Option<String>,
    /// MIME type of the blip at `image_path` (e.g. `image/x-wmf`,
    /// `image/png`). `None` when there is no image.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mime_type: Option<String>,
    /// Vector original from the Microsoft `asvg:svgBlip` extension (the zip
    /// path of the `.svg` part), preferred over `image_path` when present.
    /// `None` for a plain raster/metafile blip or a text-only paragraph.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub svg_image_path: Option<String>,
    /// Inline image natural width in pt (from `<wp:extent cx=>` EMU→pt). 0
    /// when there is no image.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub image_width_pt: f64,
    /// Inline image natural height in pt (from `<wp:extent cy=>` EMU→pt). 0
    /// when there is no image.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub image_height_pt: f64,
}

/// ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop, shared across the docx,
/// pptx and xlsx parsers (see `ooxml_common::blip`).
pub use ooxml_common::blip::{Duotone, SrcRect};

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ImageRun {
    /// Embedded zip path of the raster blip (e.g. `word/media/image1.png`), the
    /// raster fallback (PNG/JPEG), or the SVG part itself when the blip embeds no
    /// raster `r:embed`. The renderer fetches the bytes lazily by path rather
    /// than inlining base64.
    pub image_path: String,
    /// MIME type of the blip at `image_path` (e.g. `image/png`, or
    /// `image/svg+xml` for an svg-only picture).
    pub mime_type: String,
    /// Vector original from the Microsoft `asvg:svgBlip` extension (MS-ODRAWXML)
    /// — the zip path of the `.svg` part. When present the renderer prefers it
    /// over `image_path` (the raster fallback). `None` for a plain raster blip.
    /// Its MIME is always `image/svg+xml` and is owned by the SVG decoder.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub svg_image_path: Option<String>,
    /// ECMA-376 §20.1.8.55 `<a:srcRect>` source-rectangle crop, as fractions
    /// 0..1 of the decoded bitmap. `None` when absent or all-zero (no crop).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub src_rect: Option<SrcRect>,
    /// pt
    pub width_pt: f64,
    /// pt
    pub height_pt: f64,
    /// Effective DrawingML rotation after composing a picture's parent groups.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub rotation: f64,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub flip_h: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub flip_v: bool,
    /// true = wp:anchor (absolute page position), false = wp:inline (flows with text)
    pub anchor: bool,
    /// X offset in pt (anchor only).  Interpretation depends on anchor_x_from_margin.
    pub anchor_x_pt: f64,
    /// Y offset in pt (anchor only).  Interpretation depends on anchor_y_from_para.
    pub anchor_y_pt: f64,
    /// If true anchorXPt is relative to the left margin; add section.marginLeft to get page-abs X.
    /// If false anchorXPt is already page-absolute.
    pub anchor_x_from_margin: bool,
    /// If true anchorYPt is relative to the paragraph's top Y in the renderer (add paragraphTopPx).
    /// If false anchorYPt is already page-absolute.
    pub anchor_y_from_para: bool,
    /// When set, the renderer should replace all pixels of this hex color (e.g. "FFFFFF") with
    /// full transparency. Used to implement a:clrChange (make-background-transparent).
    pub color_replace_from: Option<String>,
    /// ECMA-376 §20.1.8.23 `<a:duotone>` recolour effect, resolved to its two
    /// endpoint colours through the document theme. `None` = no duotone (the
    /// common case). When set, the renderer remaps the decoded raster along the
    /// `clr1`→`clr2` luminance ramp at decode time (shared core `applyDuotone`),
    /// matching Word's picture recolour.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub duotone: Option<Duotone>,
    /// ECMA-376 §20.1.8.6 `<a:blip><a:alphaModFix@amt>` opacity as a fraction
    /// (0.0–1.0). `None` = fully opaque (the common case). When set, the renderer
    /// multiplies the picture's `globalAlpha` by this fraction, exactly as the
    /// pptx/xlsx picture paths do.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alpha: Option<f64>,
    /// Wrap mode for anchor images. One of:
    ///   "square" | "topAndBottom" | "none" | "tight" | "through"
    /// Inline images and anchors without an explicit wrap element use "none".
    /// "tight" and "through" fall back to "square" rendering in the MVP.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_mode: Option<String>,
    /// distT (top padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_top: f64,
    /// distB (bottom padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_bottom: f64,
    /// distL (left padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_left: f64,
    /// distR (right padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_right: f64,
    /// wrapSquare/wrapTight "wrapText" attribute: "bothSides" | "left" | "right" | "largest".
    /// Defaults to "bothSides" (equivalent).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_side: Option<String>,
    /// ECMA-376 §20.4.2.3 `wp:anchor/@allowOverlap` — whether this floating
    /// object may overlap other floats. CT_Anchor requires the attribute;
    /// parser-originated missing/invalid status is retained on the private wire.
    /// This public boolean keeps the historical hand-built-model fallback.
    /// Inline images carry true (the no-constraint value).
    pub allow_overlap: bool,
    /// ECMA-376 §20.4.3.1 wp:align (positionH/wp:align). When present the
    /// renderer centers / left-aligns / right-aligns the image within the
    /// container indicated by `anchor_x_from_margin`. Values: "left",
    /// "center", "right" (others fall back to "left"). Mirrors
    /// `ShapeRun::anchor_x_align`. `None` for inline images and offset-based
    /// anchors.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_x_align: Option<String>,
    /// Vertical equivalent of anchor_x_align (positionV/wp:align).
    /// Values: "top", "center", "bottom".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_y_align: Option<String>,
    /// ECMA-376 §20.4.3.2 `<wp:positionH>` / §20.4.3.5 `<wp:positionV>`
    /// `@relativeFrom` — names the container the offset / align / pctPos
    /// is measured against (raw spec string: "page", "margin", "paragraph",
    /// "line", "leftMargin", "rightMargin", "topMargin", "bottomMargin",
    /// "insideMargin", "outsideMargin", "column", "character"). Mirrors
    /// `ShapeRun::anchor_x_relative_from` / `anchor_y_relative_from`. The
    /// renderer routes this to `xContainer` / `yContainer` so e.g.
    /// `relativeFrom="margin"` + `align="top"` pins the image to the top
    /// content margin instead of the page top. `None` for inline images and
    /// for anchors that didn't carry a positionH/V (preserve the legacy
    /// boolean hints `anchor_x_from_margin` / `anchor_y_from_para`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_x_relative_from: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_y_relative_from: Option<String>,
    #[serde(
        rename = "__anchorAcquisition",
        skip_serializing_if = "Option::is_none"
    )]
    pub(crate) anchor_acquisition: Option<AnchorAcquisitionWire>,
}

fn is_zero_f64(v: &f64) -> bool {
    *v == 0.0
}

fn is_true(v: &bool) -> bool {
    *v
}

/// A DrawingML chart embedded in the run flow (ECMA-376 §21.2). Positioned like
/// a picture: the `<wp:extent cx/cy>` natural size in points governs the box
/// the chart is drawn into. The `chart` payload is the shared
/// [`ooxml_common::chart::ChartModel`] — identical to what pptx/xlsx pass to the
/// core chart renderer — so the docx renderer draws charts at pptx/xlsx quality
/// through the same `renderChart` entry point.
///
/// Both placements are drawn: an inline chart (`anchor == false`) flows with the
/// text like an inline picture, and an anchored (floating) chart
/// (`anchor == true`, §20.4.2.3 `<wp:anchor>`) carries the same anchor and
/// text-wrap fields an `ImageRun` does — align/relativeFrom placement
/// (§20.4.3.x) plus the wrap mode/side, dist* padding, and allowOverlap the
/// float-exclusion machinery consumes (§20.4.2.16/.17) — so the renderer
/// positions it and wraps body text around it with picture parity
/// (issues #752, #788).
#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ChartRun {
    /// The shared chart model (barChart / lineChart / pie / …). Emitted 1:1 as
    /// the `chart` object the core `renderChart` consumes.
    pub chart: ooxml_common::chart::ChartModel,
    /// Natural width from `<wp:extent cx>` (EMU → pt).
    pub width_pt: f64,
    /// Natural height from `<wp:extent cy>` (EMU → pt).
    pub height_pt: f64,
    /// true = `<wp:anchor>` (absolute page position, drawn by the renderer's
    /// anchor path), false = `<wp:inline>` (flows with text).
    pub anchor: bool,
    /// Anchor X offset (pt). Anchor-only; interpretation mirrors `ImageRun`.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub anchor_x_pt: f64,
    /// Anchor Y offset (pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub anchor_y_pt: f64,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub anchor_x_from_margin: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub anchor_y_from_para: bool,
    /// ECMA-376 §20.4.2.16/.17 text-wrap mode for anchor charts. One of:
    ///   "square" | "topAndBottom" | "none" | "tight" | "through"
    /// Inline charts and anchors without an explicit wrap element use "none".
    /// "tight" and "through" fall back to "square" rendering in the MVP.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_mode: Option<String>,
    /// ECMA-376 §20.4.2.3 distT (top padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_top: f64,
    /// ECMA-376 §20.4.2.3 distB (bottom padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_bottom: f64,
    /// ECMA-376 §20.4.2.3 distL (left padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_left: f64,
    /// ECMA-376 §20.4.2.3 distR (right padding, pt). Anchor-only.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub dist_right: f64,
    /// ECMA-376 §20.4.2.16/.17 wrapSquare/wrapTight `wrapText` attribute:
    /// "bothSides" | "left" | "right" | "largest". Defaults to
    /// "bothSides" (equivalent).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_side: Option<String>,
    /// ECMA-376 §20.4.2.3 `wp:anchor/@allowOverlap` — whether this floating
    /// object may overlap other floats. CT_Anchor requires the attribute;
    /// parser-originated missing/invalid status is retained on the private wire.
    /// This public boolean keeps the historical hand-built-model fallback.
    /// Inline charts carry true (the no-constraint value), which is omitted from
    /// JSON; the TypeScript side reads an absent value with `?? true`.
    #[serde(skip_serializing_if = "is_true")]
    pub allow_overlap: bool,
    /// ECMA-376 §20.4.3.1 wp:align (positionH/wp:align). When present the
    /// renderer centers / left-aligns / right-aligns the chart within the
    /// container indicated by `anchor_x_from_margin`. Values: "left",
    /// "center", "right" (others fall back to "left"). Mirrors
    /// `ImageRun::anchor_x_align`. `None` for inline charts and offset-based
    /// anchors.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_x_align: Option<String>,
    /// Vertical equivalent of anchor_x_align (positionV/wp:align).
    /// Values: "top", "center", "bottom".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_y_align: Option<String>,
    /// ECMA-376 §20.4.3.4 ST_RelFromH / §20.4.3.5 ST_RelFromV raw
    /// `@relativeFrom` string — names the container the offset or align is
    /// measured against ("page", "margin", "paragraph", "line",
    /// "leftMargin", "rightMargin", "topMargin", "bottomMargin",
    /// "insideMargin", "outsideMargin", "column", "character"). Mirrors
    /// `ImageRun::anchor_x_relative_from` / `anchor_y_relative_from`. `None`
    /// for inline charts and for anchors that didn't carry a positionH/V
    /// (preserve the legacy boolean hints `anchor_x_from_margin` /
    /// `anchor_y_from_para`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_x_relative_from: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor_y_relative_from: Option<String>,
    #[serde(
        rename = "__anchorAcquisition",
        skip_serializing_if = "Option::is_none"
    )]
    pub(crate) anchor_acquisition: Option<AnchorAcquisitionWire>,
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub enum BreakType {
    Line,
    Page,
    Column,
    /// ECMA-376 §17.3.1.20 — Word's saved hint about where the previous
    /// render placed a page break. A layout cache, not authoritative: we
    /// paginate ourselves and ignore it uniformly (the parser strips these
    /// runs). Kept distinct from `Page` only so the stripping stays explicit.
    /// Serialized as `renderedPage` for the TS side.
    RenderedPage,
}

// ===== Table =====

/// ECMA-376 §17.4.57 `<w:tblpPr>` — floating-table positioning. Its mere
/// presence in `<w:tblPr>` makes the table FLOAT (out of the main text flow,
/// absolutely positioned by its top-left corner). All attributes are optional.
#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TblpPr {
    /// §17.4.57 leftFromText/rightFromText/topFromText/bottomFromText
    /// (ST_TwipsMeasure): minimum distance to wrapping text (dist padding), pt.
    /// Default 0.
    pub left_from_text: f64,
    pub right_from_text: f64,
    pub top_from_text: f64,
    pub bottom_from_text: f64,
    /// §17.4.57 horzAnchor (ST_HAnchor {text,margin,page}). Default "page".
    pub horz_anchor: String,
    /// True iff the source `<w:tblpPr>` carried ANY horizontal positioning hint
    /// (horzAnchor, tblpX, or tblpXSpec). When false, NO horizontal position was
    /// given. ECMA-376's literal default would then be horzAnchor="page" +
    /// tblpX=0 (the page edge), but Word anchors such an auto-converted floating
    /// table at the anchor paragraph's text/column left instead. The renderer
    /// uses this flag to apply that Word-runtime placement (documented there).
    pub horz_specified: bool,
    /// §17.4.57 vertAnchor (ST_VAnchor {text,margin,page}). Default "page".
    pub vert_anchor: String,
    /// §17.4.57 tblpX/tblpY (ST_SignedTwipsMeasure): absolute signed offset from
    /// the horz/vert anchor edge, pt. Default 0. Ignored when the corresponding
    /// `*Spec` is present.
    pub tblp_x: f64,
    pub tblp_y: f64,
    /// §17.4.57 tblpXSpec (ST_XAlign {left,center,right,inside,outside}).
    /// Supersedes tblpX when present. None ⇒ use the absolute offset tblpX.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tblp_x_spec: Option<String>,
    /// §17.4.57 tblpYSpec (ST_YAlign {inline,top,center,bottom,inside,outside}).
    /// Supersedes tblpY when present, UNLESS vertAnchor="text" (relative vertical
    /// positioning is not allowed there ⇒ tblpYSpec is ignored, fall back to
    /// tblpY). None ⇒ use the absolute offset tblpY.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tblp_y_spec: Option<String>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct DocTable {
    /// column widths in pt
    pub col_widths: Vec<f64>,
    pub rows: Vec<DocTableRow>,
    /// table-level borders
    pub borders: TableBorders,
    /// cell margin defaults pt
    pub cell_margin_top: f64,
    pub cell_margin_bottom: f64,
    pub cell_margin_left: f64,
    pub cell_margin_right: f64,
    /// table horizontal alignment on the page: "left" | "center" | "right" (w:tblPr/w:jc).
    pub jc: String,
    /// ECMA-376 §17.4.50 `<w:tblInd>` — indentation added before the table's
    /// LEADING edge (left in LTR, right in RTL/`bidi_visual`), in pt (signed —
    /// a negative value pulls the table outward past the leading margin toward
    /// the page edge). `type="dxa"` only; `pct`/`auto` are ignored per §17.4.50.
    /// `None` ⇒ no direct indent (the renderer adds nothing). Applied by the
    /// renderer only when the resolved `jc` is left/leading.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tbl_ind: Option<f64>,
    /// ECMA-376 §17.4.52 `<w:tblLayout w:type>`. "fixed" or "autofit".
    /// Absent in the source ⇒ None, which the renderer treats as the spec
    /// default "autofit" (size columns by preferred widths). When "fixed" the
    /// renderer uses the tblGrid widths verbatim (scaled to fit), ignoring tcW.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub layout: Option<String>,
    /// ECMA-376 §17.4.63 `<w:tblW>` preferred table width. `width_pt` is set
    /// only for type="dxa"; `width_pct` carries type="pct" (value in 50ths of a
    /// percent of the available content width). type="auto"/"nil"/0 ⇒ both None
    /// (table width is dictated by its columns).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_pt: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_pct: Option<f64>,
    /// ECMA-376 §17.4.1 `<w:bidiVisual>` — render columns in right-to-left
    /// (visual) order. `Some(true)` = RTL columns, `Some(false)` = explicitly
    /// LTR, `None` = unspecified. When `Some(true)` the renderer mirrors the
    /// grid so logical column 0 is placed rightmost and flips the per-cell
    /// left/right borders accordingly.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi_visual: Option<bool>,
    /// ECMA-376 §17.4.57 `<w:tblpPr>` — when present the table is FLOATING
    /// (absolutely positioned, out of the main text flow). None ⇒ an ordinary
    /// block table.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tblp_pr: Option<TblpPr>,
    /// ECMA-376 §17.4.56 `<w:tblOverlap w:val>` (ST_TblOverlap {never,overlap}).
    /// "never" ⇒ the floating table must be repositioned to avoid overlapping
    /// other floats. Default "overlap" (omitted ⇒ overlap allowed). Ignored when
    /// the table is not floating (no `tblp_pr`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub overlap: Option<String>,
    /// Parser/layout acquisition facts which must retain authored presence and
    /// lexical OOXML values. Kept off the stable TypeScript `DocTable` surface;
    /// `parser-model.ts` snapshots this serde-only wire before layout.
    #[serde(rename = "__tableLayout")]
    pub table_layout: TableLayoutAcquisitionWire,
}

/// ECMA-376 CT_TblWidth as authored. `None` on either attribute is distinct
/// from the element being absent (represented by the owning `Option`). Layout
/// resolves the lexical kind only after its containing-block width is known.
#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableWidthAcquisitionWire {
    pub kind: Option<String>,
    pub value: Option<String>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableLayoutKindAcquisitionWire {
    pub kind: Option<String>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableGridColumnAcquisitionWire {
    /// Raw `gridCol/@w` twips token. Missing is not an inferred width: per
    /// §17.4.16 it contributes zero until the table-width algorithm runs.
    pub width: Option<String>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableGridAcquisitionWire {
    pub authored: bool,
    pub columns: Vec<TableGridColumnAcquisitionWire>,
    /// The grid may need augmentation when rows address more columns than
    /// `tblGrid` authored (§17.4.17/§17.4.48). Preserve that lower bound rather
    /// than manufacturing widths in the parser.
    pub required_column_count: u32,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableMarginAcquisitionWire {
    pub top: Option<TableWidthAcquisitionWire>,
    pub bottom: Option<TableWidthAcquisitionWire>,
    pub start: Option<TableWidthAcquisitionWire>,
    pub end: Option<TableWidthAcquisitionWire>,
    pub left: Option<TableWidthAcquisitionWire>,
    pub right: Option<TableWidthAcquisitionWire>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableLayoutAcquisitionWire {
    pub effective_style_id: Option<String>,
    /// Whether Word treats this table as part of ordinary body flow after its
    /// tblpPr compatibility exceptions have been applied. This is retained
    /// separately from `DocTable::tblp_pr`: that public field preserves the
    /// authored positioning payload, whose mere presence is not an effective
    /// floating-status test ([MS-OI29500] 2.1.162(b-c)).
    pub ordinary_flow: bool,
    pub grid: TableGridAcquisitionWire,
    pub preferred_width: Option<TableWidthAcquisitionWire>,
    pub layout: Option<TableLayoutKindAcquisitionWire>,
    pub cell_spacing: Option<TableWidthAcquisitionWire>,
    pub cell_margins: Option<TableMarginAcquisitionWire>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableBorders {
    pub top: Option<BorderSpec>,
    pub bottom: Option<BorderSpec>,
    pub left: Option<BorderSpec>,
    pub right: Option<BorderSpec>,
    pub inside_h: Option<BorderSpec>,
    pub inside_v: Option<BorderSpec>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct BorderSpec {
    /// pt
    pub width: f64,
    pub color: Option<String>,
    /// "single" | "double" | "none" | ...
    pub style: String,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct DocTableRow {
    pub cells: Vec<DocTableCell>,
    /// ECMA-376 §17.4.15: number of table-grid columns skipped before the first
    /// real cell in this row. Omitted means zero.
    pub grid_before: u32,
    /// ECMA-376 §17.4.14: number of table-grid columns skipped after the last
    /// real cell in this row. Omitted means zero.
    pub grid_after: u32,
    /// pt, None = auto
    pub row_height: Option<f64>,
    /// ECMA-376 §17.4.80 w:trHeight/@hRule. "auto" (default) = treat row_height
    /// as informational and size to content; "atLeast" = use as a lower bound;
    /// "exact" = honor the value exactly. Stored as the raw OOXML token so the
    /// renderer can branch without re-parsing.
    pub row_height_rule: String,
    pub is_header: bool,
    /// ECMA-376 §17.4.6 w:cantSplit. When true, this row must not be split
    /// across page boundaries.
    pub cant_split: bool,
    #[serde(rename = "__tableRowLayout")]
    pub table_row_layout: TableRowLayoutAcquisitionWire,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableRowHeightAcquisitionWire {
    pub value: Option<String>,
    pub rule: String,
    /// `[MS-OI29500]` 2.1.180 gives Word a different omitted-hRule behavior;
    /// retaining presence lets a later compatibility policy distinguish it
    /// from an explicitly authored `auto` without guessing from the value.
    pub rule_authored: bool,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TablePropertyExceptionAcquisitionWire {
    pub preferred_width: Option<TableWidthAcquisitionWire>,
    pub layout: Option<TableLayoutKindAcquisitionWire>,
    pub justification: Option<String>,
    pub indent: Option<TableWidthAcquisitionWire>,
    pub borders: Option<TableBorders>,
    pub cell_margins: Option<TableMarginAcquisitionWire>,
    pub cell_spacing: Option<TableWidthAcquisitionWire>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableRowLayoutAcquisitionWire {
    pub height: Option<TableRowHeightAcquisitionWire>,
    /// ECMA-376 §17.4.27 direct alignment for this row. It precedes the
    /// table-property exception and parent table alignment during acquisition.
    pub justification: Option<String>,
    pub before_width: Option<TableWidthAcquisitionWire>,
    pub after_width: Option<TableWidthAcquisitionWire>,
    pub cell_spacing: Option<TableWidthAcquisitionWire>,
    /// Effective whole-table/conditional table-style spacing for this row.
    /// Direct row, exception, and table properties remain separate higher layers.
    pub style_cell_spacing: Option<TableWidthAcquisitionWire>,
    /// Effective whole-table/conditional table-style margins for this row.
    pub style_cell_margins: Option<TableMarginAcquisitionWire>,
    pub exception: Option<TablePropertyExceptionAcquisitionWire>,
}

/// One block-level entry inside a table cell. ECMA-376 §17.4.7 (w:tc) allows
/// any BlockLevelElts inside a cell — paragraphs and nested tables.
#[derive(Serialize, Debug, Clone)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum CellElement {
    // Match `BodyElement`: nested cell block vectors must not inherit the full
    // inline size of the paragraph model, and Serde still emits the same wire.
    Paragraph(Box<DocParagraph>),
    Table(Box<DocTable>),
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct DocTableCell {
    pub content: Vec<CellElement>,
    pub col_span: u32,
    /// VMerge: None = no merge, Some(true) = start of vertical merge, Some(false) = continuation
    pub v_merge: Option<bool>,
    pub borders: CellBorders,
    /// hex color background
    pub background: Option<String>,
    /// "top" | "center" | "bottom"
    pub v_align: String,
    /// ECMA-376 §17.4.71 `<w:tcW>` preferred cell width, type="dxa", in pt.
    /// Drives autofit column sizing (the per-column preferred width is the max
    /// over the cells anchored in it).
    pub width_pt: Option<f64>,
    /// `<w:tcW>` type="pct": 50ths of a percent of the available content width.
    /// Resolved against the available width in the renderer (parse time has no
    /// table width). None unless the cell uses type="pct".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width_pct: Option<f64>,
    /// Per-cell margins from `<w:tcPr><w:tcMar>` (ECMA-376 §17.4.42), in pt.
    /// Each edge overrides the table-level `<w:tblCellMar>` default (§17.4.41)
    /// when present; None = inherit the table default. Used e.g. by résumé
    /// templates that add a top margin to a single cell to space its content.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margin_top: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margin_bottom: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margin_left: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margin_right: Option<f64>,
    #[serde(rename = "__tableCellLayout")]
    pub table_cell_layout: TableCellLayoutAcquisitionWire,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableCellLayoutAcquisitionWire {
    pub preferred_width: Option<TableWidthAcquisitionWire>,
    pub margins: Option<TableMarginAcquisitionWire>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct CellBorders {
    pub top: Option<BorderSpec>,
    pub bottom: Option<BorderSpec>,
    pub left: Option<BorderSpec>,
    pub right: Option<BorderSpec>,
    /// ECMA-376 §17.4.34 (tcBorders w:insideH/w:insideV): the interior
    /// horizontal/vertical borders a cell contributes. A conditional table-style
    /// block (`w:tblStylePr`) commonly sets these to `nil` to suppress the
    /// gridlines on banded data rows (e.g. Medium List 2 / Medium Shading 2). We
    /// keep them as part of the cell so the conditional formatting (§17.7.6),
    /// folded in at parse time, can override the table-level insideH/insideV on a
    /// per-cell basis. `None` = unset (renderer falls back to the table inside
    /// spec); a `Some` with style "nil"/"none" = an explicit "no interior border".
    pub inside_h: Option<BorderSpec>,
    pub inside_v: Option<BorderSpec>,
}
