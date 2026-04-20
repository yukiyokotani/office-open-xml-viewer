use serde::Serialize;

#[derive(Serialize, Debug, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct Document {
    pub section: SectionProps,
    pub body: Vec<BodyElement>,
    pub headers: HeadersFooters,
    pub footers: HeadersFooters,
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
}

#[derive(Serialize, Debug, Clone)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum BodyElement {
    Paragraph(DocParagraph),
    Table(DocTable),
    PageBreak,
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
    pub numbering: Option<NumberingInfo>,
    /// Explicit tab stops from w:tabs. Empty means use default tab interval.
    pub tab_stops: Vec<TabStop>,
    pub runs: Vec<DocRun>,
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
}

#[derive(Serialize, Debug, Clone)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum DocRun {
    Text(TextRun),
    Image(ImageRun),
    Break { break_type: BreakType },
    Field(FieldRun),
    Shape(ShapeRun),
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
    /// anchor X (pt)
    pub anchor_x_pt: f64,
    /// anchor Y (pt)
    pub anchor_y_pt: f64,
    pub anchor_x_from_margin: bool,
    pub anchor_y_from_para: bool,
    /// If true, draw the shape behind text (wp:anchor behindDoc="1"). Renderer
    /// should draw background shapes BEFORE body text.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub behind_doc: bool,
    /// Document order within the wp:anchor (for correct z-ordering among shapes
    /// sharing the same behindDoc value). Lower value = drawn first.
    pub z_order: u32,
    /// normalized [0,1] custom path commands (one or more sub-paths)
    pub subpaths: Vec<Vec<PathCmd>>,
    /// Fill (solid or gradient). None = no fill.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<ShapeFill>,
    /// stroke hex. None = no stroke.
    pub stroke: Option<String>,
    /// stroke width in pt.
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub stroke_width: f64,
    /// rotation in degrees (clockwise).
    #[serde(skip_serializing_if = "is_zero_f64")]
    pub rotation: f64,
    /// Wrap mode matching ImageRun.wrap_mode semantics.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub wrap_mode: Option<String>,
}

#[derive(Serialize, Debug, Clone)]
#[serde(tag = "fillType", rename_all = "camelCase")]
pub enum ShapeFill {
    Solid { color: String },
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
    MoveTo { x: f64, y: f64 },
    LineTo { x: f64, y: f64 },
    CubicBezTo { x1: f64, y1: f64, x2: f64, y2: f64, x: f64, y: f64 },
    ArcTo { wr: f64, hr: f64, st_ang: f64, sw_ang: f64 },
    Close,
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
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct TextRun {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    /// pt
    pub font_size: f64,
    pub color: Option<String>,
    pub font_family: Option<String>,
    pub is_link: bool,
    pub background: Option<String>,
    /// "super" | "sub" | None
    pub vert_align: Option<String>,
    /// Target URL for hyperlinks (from relationships.xml), None if not a link or no URL
    pub hyperlink: Option<String>,
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
}

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ImageRun {
    /// data:<mime>;base64,...
    pub data_url: String,
    /// pt
    pub width_pt: f64,
    /// pt
    pub height_pt: f64,
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
}

fn is_zero_f64(v: &f64) -> bool { *v == 0.0 }

#[derive(Serialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub enum BreakType {
    Line,
    Page,
    Column,
}

// ===== Table =====

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
    /// pt, None = auto
    pub row_height: Option<f64>,
    pub is_header: bool,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct DocTableCell {
    pub content: Vec<DocParagraph>,
    pub col_span: u32,
    /// VMerge: None = no merge, Some(true) = start of vertical merge, Some(false) = continuation
    pub v_merge: Option<bool>,
    pub borders: CellBorders,
    /// hex color background
    pub background: Option<String>,
    /// "top" | "center" | "bottom"
    pub v_align: String,
    /// pt
    pub width_pt: Option<f64>,
}

#[derive(Serialize, Debug, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct CellBorders {
    pub top: Option<BorderSpec>,
    pub bottom: Option<BorderSpec>,
    pub left: Option<BorderSpec>,
    pub right: Option<BorderSpec>,
}
