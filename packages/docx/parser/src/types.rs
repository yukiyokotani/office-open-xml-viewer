use serde::Serialize;

#[derive(Serialize, Debug, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct Document {
    pub section: SectionProps,
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
    pub runs: Vec<DocRun>,
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
}

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
