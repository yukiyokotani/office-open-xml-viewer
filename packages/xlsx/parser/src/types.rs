use serde::Serialize;
use std::collections::BTreeMap;

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Workbook {
    pub sheets: Vec<SheetMeta>,
    /// Workbook date system (`<workbookPr date1904>`, ECMA-376 §18.2.28).
    /// `true` selects the 1904 date system (Mac-authored workbooks). Serial
    /// dates in cells are resolved against this base (§18.17.4.1). Omitted from
    /// JSON when false (the default 1900 system) for wire parity.
    #[serde(skip_serializing_if = "std::ops::Not::not", default)]
    pub date1904: bool,
    /// #773 partial degradation: a WORKBOOK-LEVEL degradation that leaves every
    /// sheet openable. Set when a shared workbook part was PRESENT but corrupt —
    /// most commonly `xl/sharedStrings.xml` (§18.4.9): a broken shared-string
    /// table silently blanks every string cell across ALL sheets, so unlike a
    /// per-sheet break it can't be attributed to one `Worksheet::placeholder`.
    /// Surfacing it here (tagged with the offending part, e.g.
    /// `"xl/sharedStrings.xml: <detail>"`) makes the loss visible instead of
    /// silent, while every sheet still renders its non-string content. `None`
    /// (and omitted from JSON) when every shared part read cleanly, so existing
    /// workbook snapshots are byte-for-byte unchanged.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parse_error: Option<String>,
}

/// Sheet visibility (`<sheet state>`, ECMA-376 §18.2.19 `ST_SheetState`).
/// `Hidden` = hidden but user-unhideable via the UI; `VeryHidden` = revealable
/// only programmatically. Default is `Visible`.
#[derive(Debug, Clone, Serialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum SheetVisibility {
    Visible,
    Hidden,
    VeryHidden,
}

impl SheetVisibility {
    /// For `skip_serializing_if`: omit the default (`Visible`) so existing
    /// workbook JSON snapshots are unchanged.
    pub fn is_visible(&self) -> bool {
        matches!(self, SheetVisibility::Visible)
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetMeta {
    pub name: String,
    pub sheet_id: u32,
    pub r_id: String,
    /// Sheet tab color (`<sheetPr><tabColor>`, ECMA-376 §18.3.1.93) resolved
    /// to `#RRGGBB`. Surfaced at workbook-list time so the viewer can paint
    /// every tab without eagerly parsing each worksheet. `None` when the
    /// sheet declares no tab color.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<String>,
    /// Sheet visibility (`<sheet state>`, ECMA-376 §18.2.19). Omitted from JSON
    /// when `Visible` (the default) so existing snapshots are unchanged.
    #[serde(skip_serializing_if = "SheetVisibility::is_visible")]
    pub visibility: SheetVisibility,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MergeCell {
    pub top: u32,
    pub left: u32,
    pub bottom: u32,
    pub right: u32,
}

/// Compact representation of a `<col min max width>` declaration. SpreadsheetML
/// permits one declaration to cover all 16,384 columns, so expanding the range
/// into `col_widths` would unnecessarily inflate the parser wire payload.
#[derive(Debug, Clone, Serialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ColumnWidthRange {
    pub min: u32,
    pub max: u32,
    pub width: f64,
}

/// Compact `<col min max style>` default formatting. SpreadsheetML applies
/// this style to cells in the column that do not carry their own `c/@s`.
#[derive(Debug, Clone, Serialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ColumnStyleRange {
    pub min: u32,
    pub max: u32,
    pub style_index: u32,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Worksheet {
    pub name: String,
    /// `true` for an `xl/chartsheets/*.xml` part. A chart sheet has no
    /// SpreadsheetML cell grid or `sheetData`; its drawing is the content.
    /// Omitted for ordinary worksheets to preserve the existing wire shape.
    #[serde(skip_serializing_if = "std::ops::Not::not", default)]
    pub is_chart_sheet: bool,
    pub rows: Vec<Row>,
    /// Serialized as `BTreeMap`s so JSON key order is deterministic (columns /
    /// rows in ascending index order), making the parser output byte-stable for
    /// identical input.
    pub col_widths: BTreeMap<u32, f64>,
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub col_width_ranges: Vec<ColumnWidthRange>,
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub col_style_ranges: Vec<ColumnStyleRange>,
    pub row_heights: BTreeMap<u32, f64>,
    /// Per-column outline (grouping) depth, 0-7 (ECMA-376 §18.3.1.13
    /// `<col outlineLevel>`). Only columns whose level is non-zero are recorded,
    /// so a sheet without column outlining serializes an empty map (omitted from
    /// JSON). Keyed by 1-based column index like {@link col_widths}.
    #[serde(skip_serializing_if = "BTreeMap::is_empty", default)]
    pub col_outline_levels: BTreeMap<u32, u8>,
    /// Per-column `<col collapsed>` (§18.3.1.13): set on the *summary* column
    /// when the columns one level deeper are collapsed. Only `true` entries are
    /// recorded. Keyed by 1-based column index.
    #[serde(skip_serializing_if = "BTreeMap::is_empty", default)]
    pub col_collapsed: BTreeMap<u32, bool>,
    /// Per-column `<col hidden>` (§18.3.1.13): `true` when the column is hidden
    /// (e.g. a collapsed outline hides its detail columns). Recorded explicitly —
    /// distinct from `width == 0` — so the outline gutter can tell collapsed
    /// detail columns apart. Only `true` entries are recorded.
    #[serde(skip_serializing_if = "BTreeMap::is_empty", default)]
    pub col_hidden: BTreeMap<u32, bool>,
    pub default_col_width: f64,
    pub default_row_height: f64,
    /// `<sheetFormatPr customHeight>` (§18.3.1.81): the sheet-wide default row
    /// height was manually set. Omitted when false, the schema default.
    #[serde(skip_serializing_if = "std::ops::Not::not", default)]
    pub default_row_height_custom: bool,
    pub merge_cells: Vec<MergeCell>,
    pub freeze_rows: u32,
    pub freeze_cols: u32,
    pub conditional_formats: Vec<ConditionalFormat>,
    pub images: Vec<ImageAnchor>,
    pub charts: Vec<ChartAnchor>,
    /// Grouped shapes from `<xdr:grpSp>` inside a twoCellAnchor (ECMA-376
    /// §20.5.2.17). Each anchor flattens its shape tree to a list of leaf
    /// shapes with normalized geometry for the renderer.
    pub shape_groups: Vec<ShapeAnchor>,
    /// Whether to display zero values in cells (ECMA-376 §18.3.1.94)
    pub show_zeros: bool,
    /// Whether to draw default grid lines on this sheet. Mirrors the "View →
    /// Gridlines" checkbox in Excel; parsed from `<sheetView showGridLines>`
    /// (ECMA-376 §18.3.1.83). Defaults to true.
    pub show_gridlines: bool,
    /// Whether the sheet grid is laid out right-to-left, mirroring the entire
    /// grid so column A sits on the right. Parsed from `<sheetView rightToLeft>`
    /// (ECMA-376 §18.3.1.87). Defaults to false (left-to-right).
    pub right_to_left: bool,
    /// Outline display properties from `<sheetPr><outlinePr>` (ECMA-376
    /// §18.3.1.61): `summaryBelow` / `summaryRight` decide which side of a group
    /// the summary row/column (and its +/- toggle) sits on. `None` (omitted from
    /// JSON) when the sheet declares no `<outlinePr>`; the viewer then applies the
    /// schema defaults (both `true`). Only surfaced when the sheet actually has
    /// outlining so outline-free sheets keep byte-identical JSON.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline_pr: Option<OutlinePr>,
    /// Tab color for the sheet tab (ECMA-376 §18.3.1.79)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<String>,
    /// AutoFilter range (ECMA-376 §18.3.1.2)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<CellRange>,
    /// Hyperlinks in this worksheet (ECMA-376 §18.3.1.47)
    pub hyperlinks: Vec<Hyperlink>,
    /// Cell refs (A1-style) that have an associated <comment> in xl/commentsN.xml.
    /// Excel shows a small red triangle in the top-right corner of each.
    pub comment_refs: Vec<String>,
    /// Full-fidelity comment bodies (text + author) for each `<comment>` in
    /// xl/commentsN.xml — agents that want to read the comment content (not
    /// just the cell that has one) consume this. Empty when the sheet has no
    /// comments file. Keep in sync with `comment_refs` (one entry per ref).
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub comments: Vec<XlsxComment>,
    /// `<dataValidations>` rules (ECMA-376 §18.3.1.32). Each rule covers one
    /// or more cell ranges and constrains permitted input ("list", "decimal",
    /// "date", "textLength", "custom", …). Empty when the sheet declares
    /// none. Renderer ignores this for now — exposed for tooling.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub data_validations: Vec<DataValidation>,
    /// Defined names in scope for this sheet. Includes workbook-global names and
    /// any names whose `localSheetId` matches this sheet's position in the
    /// workbook. Used by conditional-formatting `expression` rules that call
    /// named ranges like `task_start`, `today`, etc. (ECMA-376 §18.2.5).
    pub defined_names: Vec<DefinedName>,
    /// Excel Tables defined for this sheet (ECMA-376 §18.5). Rendered with a
    /// built-in table style (bold header, banded rows, etc.) on top of the
    /// cells' own styles.
    pub tables: Vec<TableInfo>,
    /// Slicers anchored to the sheet's drawing (Office 2010+ extension —
    /// `http://schemas.microsoft.com/office/drawing/2010/slicer` inside
    /// `<mc:AlternateContent>/<mc:Choice>`). Each slicer resolves its cache
    /// and referenced pivotCacheDefinition so the renderer can draw a static
    /// button list with the saved selection state.
    pub slicers: Vec<SlicerAnchor>,
    /// Metadata-only descriptions of saved pivot tables (ECMA-376 §18.10).
    /// These facts never alter saved worksheet cells or styles.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub pivot_tables: Vec<PivotTableMetadata>,
    /// Per-pivot failures that do not prevent the worksheet itself from loading.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub pivot_diagnostics: Vec<PivotDiagnostic>,
    /// Sparkline groups defined in the worksheet's `<extLst>` (Office 2010
    /// extension `http://schemas.microsoft.com/office/spreadsheetml/2009/9/main`,
    /// element `<x14:sparklineGroup>`, ECMA-376 §18.2 / Part 4).
    pub sparkline_groups: Vec<SparklineGroup>,
    /// Family name of the workbook's Normal-style font, resolved from
    /// `<cellStyleXfs>[0].fontId` → `<fonts>[fontId].name.val`. Used by the
    /// renderer to compute the Max Digit Width (ECMA-376 §18.3.1.13) for the
    /// active sheet's column widths. Denormalized onto every worksheet for
    /// renderer convenience; the value is workbook-wide.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_font_family: Option<String>,
    /// Point size of the workbook's Normal-style font (`<fonts>[N].sz.val`).
    /// Used together with `default_font_family` to compute Max Digit Width.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_font_size: Option<f64>,
    /// Workbook date system (`<workbookPr date1904>`, ECMA-376 §18.2.28),
    /// denormalized onto every worksheet so the cell formatter can resolve
    /// serial dates (§18.17.4.1) without a workbook back-reference. `true` =
    /// 1904 date system. Omitted from JSON when false (default 1900 system).
    #[serde(skip_serializing_if = "std::ops::Not::not", default)]
    pub date1904: bool,
    /// RB7 partial degradation: when THIS sheet's part could not be read/parsed,
    /// the workbook still opens with the OTHER sheets intact and this one becomes
    /// an empty placeholder carrying the part-tagged error (e.g.
    /// `"xl/worksheets/sheet3.xml: <detail>"`). `None` (and omitted from JSON) for
    /// every healthy sheet, so existing snapshots are byte-for-byte unchanged. The
    /// renderer paints a visible error overlay on the sheet grid.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parse_error: Option<String>,
}

impl Worksheet {
    fn empty(name: &str) -> Self {
        Worksheet {
            name: name.to_string(),
            is_chart_sheet: false,
            rows: Vec::new(),
            col_widths: BTreeMap::new(),
            col_width_ranges: Vec::new(),
            col_style_ranges: Vec::new(),
            row_heights: BTreeMap::new(),
            col_outline_levels: BTreeMap::new(),
            col_collapsed: BTreeMap::new(),
            col_hidden: BTreeMap::new(),
            default_col_width: 0.0,
            default_row_height: 0.0,
            default_row_height_custom: false,
            merge_cells: Vec::new(),
            freeze_rows: 0,
            freeze_cols: 0,
            conditional_formats: Vec::new(),
            images: Vec::new(),
            charts: Vec::new(),
            shape_groups: Vec::new(),
            show_zeros: true,
            show_gridlines: true,
            right_to_left: false,
            outline_pr: None,
            tab_color: None,
            auto_filter: None,
            hyperlinks: Vec::new(),
            comment_refs: Vec::new(),
            comments: Vec::new(),
            data_validations: Vec::new(),
            defined_names: Vec::new(),
            tables: Vec::new(),
            slicers: Vec::new(),
            pivot_tables: Vec::new(),
            pivot_diagnostics: Vec::new(),
            sparkline_groups: Vec::new(),
            default_font_family: None,
            default_font_size: None,
            date1904: false,
            parse_error: None,
        }
    }

    /// A healthy row-free chart sheet. Its drawing relationships are attached
    /// by the same finalization path as an ordinary worksheet.
    pub fn chart_sheet(name: &str) -> Self {
        let mut worksheet = Self::empty(name);
        worksheet.is_chart_sheet = true;
        worksheet.show_gridlines = false;
        worksheet
    }

    /// A minimal empty sheet carrying a part-tagged parse error (RB7). Used when
    /// one sheet's XML can't be read or parsed so the workbook still opens.
    pub fn placeholder(name: &str, parse_error: String) -> Self {
        let mut worksheet = Self::empty(name);
        worksheet.parse_error = Some(parse_error);
        worksheet
    }
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotTableMetadata {
    pub name: String,
    /// `cacheId` is retained as an identity fact only. Part resolution follows
    /// the pivot table part's own relationship, never this number.
    pub cache_id: u32,
    pub location: PivotLocation,
    /// ECMA-376 §18.10.1.92/§18.10.1.15 `CT_Field@x` is signed; `-2` is the
    /// Values pseudo-field sentinel and must remain distinguishable.
    pub row_fields: Vec<i32>,
    pub column_fields: Vec<i32>,
    pub page_fields: Vec<PivotPageField>,
    pub data_fields: Vec<PivotDataField>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub refresh_on_load: Option<bool>,
    /// `pivotCacheDefinition@invalid`, known only after the cache part parses.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cache_invalid: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cache_definition_part: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cache_source: Option<PivotCacheSource>,
    pub status: PivotMetadataStatus,
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub extension_uris: Vec<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotLocation {
    #[serde(flatten)]
    pub range: CellRange,
    pub first_header_row: u32,
    pub first_data_row: u32,
    pub first_data_col: u32,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotPageField {
    /// ECMA-376 §18.10.1.66 `CT_PageField@fld` is signed.
    pub field: i32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub item: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotDataField {
    /// ECMA-376 §18.10.1.16 `CT_DataField@fld` is unsigned.
    pub field: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subtotal: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub raw_subtotal: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum PivotCacheSource {
    Worksheet {
        #[serde(skip_serializing_if = "Option::is_none")]
        sheet: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        reference: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        name: Option<String>,
        #[serde(rename = "relationshipId", skip_serializing_if = "Option::is_none")]
        relationship_id: Option<String>,
    },
    External,
    Consolidation,
    Scenario,
}

#[derive(Debug, Serialize)]
#[serde(tag = "state", rename_all = "camelCase")]
pub enum PivotMetadataStatus {
    Complete,
    Partial { reasons: Vec<PivotPartialReason> },
}

#[derive(Debug, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum PivotPartialReason {
    MissingCacheRelationship,
    MalformedCacheRelationships,
    UnreadableCacheRelationships,
    ExternalCacheRelationship,
    AmbiguousCacheRelationship,
    UnreadableCacheDefinition,
    MalformedCacheDefinition,
    MalformedField {
        field: String,
    },
    UnsupportedCacheSource {
        #[serde(rename = "sourceType")]
        source_type: String,
    },
    UnresolvedWorksheetSourceRelationship,
    UnsupportedSemanticFeature {
        feature: String,
    },
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotDiagnostic {
    pub part: String,
    pub reason: PivotDiagnosticReason,
}

#[derive(Debug, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum PivotDiagnosticReason {
    UnreadableWorksheetRelationships,
    MalformedWorksheetRelationships,
    MalformedPivotRelationship,
    ExternalPivotRelationship,
    UnreadablePart,
    MalformedXml,
    MissingIdentity,
    InvalidLocation,
}

/// Single sparkline group (`<x14:sparklineGroup>`). Holds the shared formatting
/// for every individual sparkline cell that belongs to the group.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SparklineGroup {
    /// `line` (default) | `column` | `stem` (win/loss).
    pub kind: SparklineType,
    /// Show a marker dot at every data point (line type only).
    pub markers: bool,
    /// Highlight high / low / first / last / negative points.
    pub high: bool,
    pub low: bool,
    pub first: bool,
    pub last: bool,
    pub negative: bool,
    /// Show the horizontal axis line when data crosses zero.
    pub display_x_axis: bool,
    /// `gap` (default) | `zero` | `span` — how empty cells in the source
    /// range are treated. We only honor `gap` (default) at render time.
    pub display_empty_cells_as: String,
    /// Per-axis-bound type: `individual` (default) / `group` / `custom`.
    pub min_axis_type: String,
    pub max_axis_type: String,
    /// Used when *AxisType=`custom`. f64::NAN otherwise.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub manual_min: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub manual_max: Option<f64>,
    /// Stroke weight in pt (line type). ECMA-376 default 0.75.
    pub line_weight: f64,
    /// Resolved RGB hex strings (e.g. `#4472C4`) — theme + tint flattened
    /// at parse time so the renderer never sees a theme index. `None` means
    /// the property was not specified and the renderer should fall back.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_series: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_negative: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_axis: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_markers: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_first: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_last: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_high: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_low: Option<String>,
    /// Individual sparklines (one per destination cell).
    pub sparklines: Vec<Sparkline>,
}

#[derive(Debug, Serialize, Clone, Copy)]
#[serde(rename_all = "camelCase")]
pub enum SparklineType {
    Line,
    Column,
    Stem,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Sparkline {
    /// 1-based row of the cell that displays this sparkline (`<xm:sqref>`).
    pub row: u32,
    /// 1-based column of the cell.
    pub col: u32,
    /// Numeric values resolved from the `<xm:f>` data range. `None` for
    /// empty / non-numeric / out-of-bounds cells.
    pub values: Vec<Option<f64>>,
}

/// Excel Table metadata (ECMA-376 §18.5 `<table>`). The renderer overlays a
/// built-in style on top of the cell styles inside `range`.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableInfo {
    /// Inclusive table area including the header row.
    pub range: CellRange,
    /// Built-in style name like "TableStyleLight18" (ECMA-376 §18.5.1.4).
    pub style_name: String,
    /// Number of header rows (default 1).
    pub header_row_count: u32,
    /// Number of totals rows at the bottom (default 0).
    pub totals_row_count: u32,
    /// `<tableStyleInfo showRowStripes>` — banded rows in the data region.
    pub show_row_stripes: bool,
    /// `<tableStyleInfo showColumnStripes>`.
    pub show_column_stripes: bool,
    /// `<tableStyleInfo showFirstColumn>`.
    pub show_first_column: bool,
    /// `<tableStyleInfo showLastColumn>`.
    pub show_last_column: bool,
    /// Accent color resolved from the built-in style name against this file's
    /// theme accents (e.g. `TableStyleLight18` → accent3 of theme1.xml). Used
    /// by the renderer to draw banding, header background, and rules.
    pub accent_color: String,
    /// `true` when `style_name` is defined in the file's `<tableStyles>` block,
    /// i.e. a *custom* style (ECMA-376 §18.5.1.2). The renderer must draw such
    /// tables strictly from their declared element dxfs and must NOT apply the
    /// accent-based approximation (banding, synthesized rules/header) that is
    /// reserved for built-in style names whose definitions are absent.
    #[serde(default)]
    pub is_custom: bool,
    /// Dxf index for the `wholeTable` element of a custom `<tableStyle>`
    /// (ECMA-376 §18.8.83). When set, its border/fill apply to every cell
    /// of the table as a base layer. Built-in style names use the renderer's
    /// accent-based fallback, not this field.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub whole_table_dxf: Option<u32>,
    /// Dxf index for the `headerRow` element of a custom `<tableStyle>`.
    /// Provides the header background fill, font color/weight, and any
    /// vertical separator borders shown between header cells.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_row_dxf: Option<u32>,
    /// Dxf index for the `totalRow` element (ECMA-376 §18.18.93).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub total_row_dxf: Option<u32>,
    /// Dxf index for the `firstColumn` element.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_column_dxf: Option<u32>,
    /// Dxf index for the `lastColumn` element.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_column_dxf: Option<u32>,
    /// Dxf index for the `firstRowStripe` (band1 horizontal) element — the odd
    /// banded-row stripe applied when `show_row_stripes` is set.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub band1_horizontal_dxf: Option<u32>,
    /// Dxf index for the `secondRowStripe` (band2 horizontal) element — the
    /// even banded-row stripe.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub band2_horizontal_dxf: Option<u32>,
    /// Per-column DXF references (ECMA-376 §18.5.1.3 `tableColumn`). Length
    /// matches the number of `<tableColumn>` children in the table XML, so
    /// `columns[c - range.left]` gives the DXFs for the cell column. The
    /// renderer can use these to apply column-level overlays for named-style
    /// tables that don't pre-bake column DXFs into the cell `xf`. For files
    /// where Excel pre-bakes the result into `xf` (the common case), reading
    /// the cell `xf` already reflects the column DXF and these fields are
    /// purely informational.
    pub columns: Vec<TableColumnInfo>,
}

/// Per-column DXF references inside a `<table>` element
/// (ECMA-376 §18.5.1.3 `tableColumn`).
#[derive(Debug, Serialize, Default, Clone)]
#[serde(rename_all = "camelCase")]
pub struct TableColumnInfo {
    /// `<tableColumn dataDxfId>` — applied to data-region cells in this column.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_dxf_id: Option<u32>,
    /// `<tableColumn headerRowDxfId>` — applied to the header cell of this column.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_row_dxf_id: Option<u32>,
    /// `<tableColumn totalsRowDxfId>` — applied to the totals cell of this column.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub totals_row_dxf_id: Option<u32>,
}

/// One cell comment. Sourced from the classic notes file `xl/commentsN.xml`
/// (ECMA-376 §18.7) when present, otherwise from the Office-365 threaded
/// comments part `xl/threadedComments/threadedCommentN.xml` (MS-XLSX schema
/// `…/office/spreadsheetml/2018/threadedcomments`, `personId` resolved against
/// `xl/persons/person*.xml`). `text` is the joined plain text — every `<r><t>`
/// run for classic notes, every reply in the thread (newline-joined) for
/// threaded comments; rich-text formatting is dropped.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct XlsxComment {
    /// A1-style cell reference (`@ref` on the comment element).
    pub cell_ref: String,
    /// Resolved author name. For classic notes this is the `<authors>` entry
    /// indexed by `@authorId`; for threaded comments it is the `<person>`
    /// `displayName` matching `@personId`. `None` when unresolved / absent.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    /// Concatenated plain text (every run for classic; every threaded reply,
    /// newline-joined, for threaded).
    pub text: String,
}

/// One `<dataValidation>` rule (ECMA-376 §18.3.1.32). `type` covers the
/// constraint class ("list", "whole", "decimal", "date", "time", "textLength",
/// "custom"). `operator` qualifies it ("between", "notBetween", "equal",
/// "notEqual", "lessThan", …). `formula1` / `formula2` are the rule operands.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DataValidation {
    /// Affected cell ranges, written verbatim from `@sqref` (space-separated).
    pub sqref: String,
    /// Constraint class. None means the validator is the spec's default
    /// (`"none"`, treated as no constraint).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub validation_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub operator: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula1: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula2: Option<String>,
    #[serde(skip_serializing_if = "std::ops::Not::not", default)]
    pub allow_blank: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prompt_title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prompt: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_message: Option<String>,
}

/// Workbook- or sheet-scoped defined name (ECMA-376 §18.2.5 `definedName`).
/// `formula` is the raw formula text (typically a cell/range reference, e.g.
/// `ProjectSchedule!$E1`). Relative references inside are shifted relative to
/// A1 when substituted into a formula.
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DefinedName {
    pub name: String,
    pub formula: String,
}

/// A chart anchored to a rectangular range of cells (ECMA-376 §20.5 twoCellAnchor).
/// Offsets are EMU (914400 EMU = 1 inch, 9525 EMU = 1 px @ 96 DPI).
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartAnchor {
    /// Document-order byte position of the `<xdr:graphicFrame>` in its drawing
    /// part. DrawingML uses element order as z-order (§20.5.2.35).
    pub z_order: u64,
    pub from_col: u32,
    pub from_col_off: i64,
    pub from_row: u32,
    pub from_row_off: i64,
    pub to_col: u32,
    pub to_col_off: i64,
    pub to_row: u32,
    pub to_row_off: i64,
    /// The emitted chart payload — the canonical shared `ChartModel`, produced
    /// directly by `ooxml_common::chart::parse_chart_part` (the single superset
    /// parser for pptx + xlsx).
    pub chart: ooxml_common::chart::ChartModel,
}

/// A grouped-shape anchor (ECMA-376 §20.5.2.17, `<xdr:grpSp>` inside a
/// `<xdr:twoCellAnchor>`). Leaf shape elements (`<xdr:sp>`) from any nesting
/// level are flattened into `shapes` with normalized coordinates so the
/// renderer only needs to scale to the anchor rect.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeAnchor {
    pub from_col: u32,
    pub from_col_off: i64,
    pub from_row: u32,
    pub from_row_off: i64,
    pub to_col: u32,
    pub to_col_off: i64,
    pub to_row: u32,
    pub to_row_off: i64,
    /// `twoCellAnchor@editAs` (ECMA-376 §20.5.2.33). With `"oneCell"` the
    /// renderer uses `native_ext_cx/cy` for the on-sheet size instead of the
    /// from/to-derived rect (Excel's "Move but don't size with cells").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub edit_as: Option<String>,
    /// Group's `<xdr:grpSpPr><a:xfrm><a:ext cx cy>` (or `<xdr:spPr><a:xfrm>`
    /// for stand-alone sp/pic) in EMU. The saved on-sheet size, used as the
    /// authoritative extent when `editAs == "oneCell"`. 0 = unavailable.
    pub native_ext_cx: i64,
    pub native_ext_cy: i64,
    pub shapes: Vec<ShapeInfo>,
}

/// A leaf shape extracted from a grpSp/sp tree. Position/size are normalized
/// to [0,1] relative to the top-level grpSp extent (which itself maps to the
/// anchor rect).
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeInfo {
    /// Document-order byte position of this leaf in the drawing part.
    pub z_order: u64,
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
    /// Rotation in degrees (clockwise). DrawingML `a:xfrm/@rot` is in 60000ths
    /// of a degree; the parser converts to degrees here.
    pub rot: f64,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub flip_h: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub flip_v: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<ShapeFill>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_color: Option<String>,
    /// Stroke width in EMU (914400 = 1 inch). 0 = no stroke.
    pub stroke_width: i64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_fill: Option<ShapeStrokeFill>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_dash_style: Option<String>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub stroke_custom_dash: Vec<ShapeLineDashSegment>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_line_cap: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_line_join: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_miter_limit: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_alignment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_cmpd: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_head_end: Option<ShapeLineEnd>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_tail_end: Option<ShapeLineEnd>,
    pub geom: ShapeGeom,
    /// Text content from `<xdr:txBody>` (ECMA-376 §20.5.2.34). Present on
    /// shapes that carry visible text — typically text boxes (`txBox="1"`)
    /// but also any `<xdr:sp>` whose `<a:p>` runs render visibly. `None`
    /// when the shape has no text body or only contains an empty paragraph.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<ShapeText>,
}

#[derive(Debug, Serialize)]
#[serde(tag = "fillType", rename_all = "camelCase")]
pub enum ShapeStrokeFill {
    Gradient {
        stops: Vec<ooxml_common::fill::GradStop>,
        angle: f64,
        grad_type: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        scaled: Option<bool>,
        #[serde(skip_serializing_if = "Option::is_none")]
        path: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        fill_to_rect: Option<ooxml_common::fill::FillRect>,
        #[serde(skip_serializing_if = "Option::is_none")]
        tile_rect: Option<ooxml_common::fill::FillRect>,
        #[serde(skip_serializing_if = "Option::is_none")]
        flip: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        rot_with_shape: Option<bool>,
    },
    Pattern {
        fg: String,
        bg: String,
        preset: String,
    },
}

#[derive(Debug, Serialize)]
#[serde(tag = "fillType", rename_all = "camelCase")]
pub enum ShapeFill {
    Solid {
        color: String,
    },
    Gradient {
        stops: Vec<ooxml_common::fill::GradStop>,
        angle: f64,
        grad_type: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        scaled: Option<bool>,
        #[serde(skip_serializing_if = "Option::is_none")]
        path: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        fill_to_rect: Option<ooxml_common::fill::FillRect>,
        #[serde(skip_serializing_if = "Option::is_none")]
        tile_rect: Option<ooxml_common::fill::FillRect>,
        #[serde(skip_serializing_if = "Option::is_none")]
        flip: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        rot_with_shape: Option<bool>,
    },
    Pattern {
        fg: String,
        bg: String,
        preset: String,
    },
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeLineDashSegment {
    pub dash: f64,
    pub space: f64,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeLineEnd {
    pub r#type: String,
    pub w: String,
    pub len: String,
}

/// Paragraph line spacing (`<a:pPr>/<a:lnSpc>`, ECMA-376 §21.1.2.2.5). Shared
/// with the pptx parser (and mirroring core's TS `SpaceLine`): a percentage of
/// the natural single line, or an absolute per-line height in points.
pub use ooxml_common::text::SpaceLine;

/// Text body inside a shape (`<xdr:txBody>`, ECMA-376 §20.1.2.2). Holds
/// the paragraphs plus body-level formatting (`<a:bodyPr>`).
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeText {
    /// `<a:bodyPr@anchor>` — vertical alignment of the text block within the
    /// shape rect. `t` (top, default), `ctr` (middle), `b` (bottom),
    /// `just`/`dist` (treated as top).
    pub anchor: String,
    /// `<a:bodyPr@wrap>` — `square` (default = wrap to shape width),
    /// `none` (no wrap).
    pub wrap: String,
    /// `<a:bodyPr>` autofit child (ECMA-376 §21.1.2.1.1-.3): `sp`
    /// (`spAutoFit`), `norm` (`normAutofit`), or `none` (`noAutofit`/absent).
    /// Always emitted (default `none`), mirroring the pptx `TextBody.autoFit`.
    pub auto_fit: String,
    /// `<a:normAutofit@fontScale>` — PowerPoint/Excel's stored font-shrink
    /// fraction (e.g. 0.625 for `fontScale="62500"`). `None` when unset.
    /// Modeled for parity; the xlsx renderer does not currently apply it.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_scale: Option<f64>,
    /// `<a:normAutofit@lnSpcReduction>` — stored line-spacing reduction fraction
    /// (e.g. 0.20 for `lnSpcReduction="20000"`). `None` when unset.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ln_spc_reduction: Option<f64>,
    /// `<a:bodyPr@lIns>` — left text inset in EMU (ECMA-376 §21.1.2.1.1
    /// `CT_TextBodyProperties`). Emitted even at the default so the renderer uses
    /// the spec inset (91440 EMU = 7.2 pt) instead of an empirical constant.
    /// Same EMU convention as `ShapeParagraph.marL`.
    pub l_ins: i64,
    /// `<a:bodyPr@tIns>` — top text inset in EMU. Default 45720 (3.6 pt).
    pub t_ins: i64,
    /// `<a:bodyPr@rIns>` — right text inset in EMU. Default 91440 (7.2 pt).
    pub r_ins: i64,
    /// `<a:bodyPr@bIns>` — bottom text inset in EMU. Default 45720 (3.6 pt).
    pub b_ins: i64,
    pub paragraphs: Vec<ShapeParagraph>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeParagraph {
    /// `<a:pPr@algn>` — `l` (default), `ctr`, `r`, `just`, `dist`.
    pub align: String,
    /// `<a:pPr@rtl>` — whether the paragraph reads right-to-left (ECMA-376
    /// §21.1.2.2.7). Omitted from JSON when false so existing output stays
    /// byte-identical (additive, like the other Phase 0 direction flags).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub rtl: bool,
    /// `<a:pPr@marL>` — left margin (EMU). ECMA-376 §21.1.2.2.7
    /// (`CT_TextParagraphProperties`). Direct attribute only (xlsx text boxes
    /// have no lstStyle/level cascade). `None` = unset. Omitted from JSON when
    /// `None` so existing output stays byte-identical (additive).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mar_l: Option<i64>,
    /// `<a:pPr@marR>` — right margin (EMU). ECMA-376 §21.1.2.2.7. `None` = unset.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mar_r: Option<i64>,
    /// `<a:pPr@indent>` — first-line indent (EMU; negative = hanging). ECMA-376
    /// §21.1.2.2.7. `None` = unset.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<i64>,
    /// `<a:pPr>/<a:lnSpc>` line spacing (ECMA-376 §21.1.2.2.5). Direct-only.
    /// `None` = unset. Omitted from JSON when `None` so existing output stays
    /// byte-identical (additive).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub space_line: Option<SpaceLine>,
    pub runs: Vec<ShapeTextRun>,
}

/// A run within a shape paragraph. Tagged union (mirrors the pptx `TextRun`
/// shape) so a run is either styled text, a soft line break, or an OMML
/// equation. Excel stores "Insert > Equation" as OMML inside the shared
/// DrawingML `<xdr:txBody>` grammar (ECMA-376 §22.1), exactly like PowerPoint.
#[derive(Debug, Serialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum ShapeTextRun {
    /// The enum-level `rename_all = "camelCase"` (tag = "type") renames only the
    /// variant tags, not the fields, so the multi-word `font_face` needs a
    /// per-variant `rename_all` to serialize as the camelCase key the TS
    /// renderer reads (`renderer.ts`: `run.fontFace`). Without it the key lands
    /// as snake_case, the renderer reads `undefined`, and the shape text's font
    /// face is silently ignored (falls back to the default stack). Same root
    /// cause as the pptx fix in PR #489 / the xlsx ArcTo fix in PR #491.
    #[serde(rename_all = "camelCase")]
    Text {
        text: String,
        bold: bool,
        italic: bool,
        /// Font size in points (`<a:rPr@sz>` is in 100ths of a point; this
        /// field is already converted). 0 means "inherit from default" →
        /// renderer falls back to its own default.
        size: f64,
        #[serde(skip_serializing_if = "Option::is_none")]
        color: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        font_face: Option<String>,
        /// East-Asian typeface (`<a:ea@typeface>`, ECMA-376 §21.1.2.3.1). The
        /// common encoding for Japanese shape text sets Meiryo here while
        /// leaving `<a:latin>` default. The renderer floors the line box by
        /// this face's design line too. Additive — omitted from JSON when None
        /// so shapes without `<a:ea>` stay byte-identical.
        #[serde(skip_serializing_if = "Option::is_none")]
        font_face_ea: Option<String>,
        /// Complex-script typeface (`<a:cs@typeface>`, ECMA-376 §21.1.2.3.1).
        /// Additive — omitted from JSON when None.
        #[serde(skip_serializing_if = "Option::is_none")]
        font_face_cs: Option<String>,
    },
    /// Soft line break (`<a:br>`).
    Break,
    /// Inline (`display:false`) or block (`display:true`) OMML equation.
    ///
    /// Like `Text`, the multi-word `font_size` needs a per-variant
    /// `rename_all = "camelCase"` so it serializes as `fontSize` (the key the
    /// renderer reads at `renderer.ts`: `run.fontSize`). Without it the
    /// snake_case key is read as `undefined` and the equation always falls back
    /// to the inherited/default size instead of its explicit `rPr@sz`.
    #[serde(rename_all = "camelCase")]
    Math {
        nodes: Vec<ooxml_common::math::MathNode>,
        display: bool,
        /// Point size for the equation, when the run carries an explicit
        /// `rPr@sz`; otherwise the renderer inherits the surrounding size.
        #[serde(skip_serializing_if = "Option::is_none")]
        font_size: Option<f64>,
        #[serde(skip_serializing_if = "Option::is_none")]
        color: Option<String>,
    },
}

#[derive(Debug, Serialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum ShapeGeom {
    /// Preset geometry (rect, ellipse, roundRect, triangle, etc.).
    /// ECMA-376 §20.1.9.18 `a:prstGeom/@prst`.
    ///
    /// `adj` carries the shape's adjust handles from `<a:avLst><a:gd>` in
    /// declaration order (index 0 = adj/adj1, 1 = adj2, …); `None` entries mean
    /// "use the preset's declared default". ECMA-376 §19.5.31.3 / §20.1.9.5.
    Preset {
        name: String,
        #[serde(skip_serializing_if = "Vec::is_empty")]
        adj: Vec<Option<f64>>,
    },
    /// Freeform path geometry (ECMA-376 §20.1.9.2 `a:custGeom`).
    Custom { paths: Vec<PathInfo> },
    /// Bitmap (or vector) image leaf inside a `<xdr:grpSp>` tree (ECMA-376
    /// §20.5.2.17). `image_path` is the zip path of the drawing's relationship
    /// target (png/jpg/gif/svg/…) — the blip's raster `r:embed` fallback, or the
    /// SVG itself when no raster is embedded — and `mime_type` its MIME via the
    /// shared `mime_from_ext`. `svg_image_path` carries the Microsoft svgBlip
    /// extension's vector original when present, so the renderer can prefer it
    /// and fall back to `image_path` on a decode failure. `None` otherwise. Its
    /// MIME is always `image/svg+xml` and is owned by the SVG decoder. The
    /// renderer fetches bytes lazily via `extract_image`.
    #[serde(rename_all = "camelCase")]
    Image {
        image_path: String,
        mime_type: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        svg_image_path: Option<String>,
        /// ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop on the leaf pic,
        /// present only when cropped. Honored by the renderer like the top-level
        /// `ImageAnchor.src_rect` so a `oneCellAnchor` / `grpSp` leaf pic crops
        /// the same as a `twoCellAnchor` picture.
        #[serde(skip_serializing_if = "Option::is_none")]
        src_rect: Option<SrcRect>,
        /// ECMA-376 §20.1.8.6 `<a:alphaModFix@amt>` on the leaf pic (0.0–1.0).
        /// `None` = opaque. Applied via `globalAlpha`, like the top-level anchor.
        #[serde(skip_serializing_if = "Option::is_none")]
        alpha: Option<f64>,
        /// ECMA-376 §20.1.8.23 `<a:duotone>` recolour on the leaf pic. `None` =
        /// no effect. Remaps the leaf image along the `clr1`→`clr2` ramp.
        #[serde(skip_serializing_if = "Option::is_none")]
        duotone: Option<Duotone>,
    },
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PathInfo {
    /// Path's own coordinate system width.
    pub w: f64,
    /// Path's own coordinate system height.
    pub h: f64,
    pub commands: Vec<PathCmd>,
}

#[derive(Debug, Serialize)]
#[serde(tag = "op", rename_all = "camelCase")]
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
        x3: f64,
        y3: f64,
    },
    QuadBezTo {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
    },
    /// ECMA-376 §20.1.9.3 `a:arcTo`. `stAng`/`swAng` are in 60000ths of a
    /// degree. The start point is the current pen position; the ellipse
    /// center is derived so the pen lies on the ellipse at `stAng`.
    ///
    /// The enum-level `rename_all = "camelCase"` (tag = "op") renames only the
    /// variant tags, not the fields, so `st_ang`/`sw_ang` need a per-variant
    /// `rename_all` to serialize as the camelCase keys the TS renderer reads
    /// (`renderer.ts`: `cmd.stAng`/`cmd.swAng`, then `/ 60000`). Without it the
    /// keys land as snake_case, the renderer reads `undefined` → `NaN`, and the
    /// arc fails to draw. Same root cause as the pptx fix in PR #489. The raw
    /// 60000ths convention is unchanged (the renderer does the division).
    #[serde(rename_all = "camelCase")]
    ArcTo {
        wr: f64,
        hr: f64,
        st_ang: f64,
        sw_ang: f64,
    },
    Close,
}

/// Slicer anchor — a button bank that filters a connected pivot table or
/// Excel Table. Office stores slicers in a 2010 extension namespace
/// (`sle:slicer`) wrapped in `<mc:AlternateContent>`, with the cache data in
/// `xl/slicerCaches/*.xml` and the underlying item list in the linked
/// pivotCache's `sharedItems`.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SlicerAnchor {
    pub from_col: u32,
    pub from_col_off: i64,
    pub from_row: u32,
    pub from_row_off: i64,
    pub to_col: u32,
    pub to_col_off: i64,
    pub to_row: u32,
    pub to_row_off: i64,
    /// Slicer header text. Typically the `caption` of the slicer definition
    /// (`xl/slicers/slicerN.xml`), falling back to the drawing `cNvPr` name.
    pub caption: String,
    /// One row per cache item in display order. Items flagged "selected" are
    /// the ones currently active in the filter; non-selected items are drawn
    /// with the ghost style. When the cache selection state is unavailable,
    /// all items are emitted as selected (Excel's default).
    pub items: Vec<SlicerItem>,
    /// Resolved custom slicer style from SpreadsheetML `<tableStyles>` plus
    /// the Office 2010 `x14:slicerStyles` extension. Absent for built-in or
    /// unresolved styles, in which case the renderer uses its default skin.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style: Option<SlicerStyle>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SlicerItem {
    pub name: String,
    pub selected: bool,
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct SlicerStyle {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub whole: Option<SlicerElementStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header: Option<SlicerElementStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub selected_item_with_data: Option<SlicerElementStyle>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub unselected_item_with_data: Option<SlicerElementStyle>,
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct SlicerElementStyle {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub border_color: Option<String>,
}

/// An image anchored to a rectangular range of cells
/// (ECMA-376 §20.5, `<xdr:twoCellAnchor>`). Offsets are EMU (English
/// Metric Unit): 914400 EMU = 1 inch, and 9525 EMU = 1 pixel at 96 DPI.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageAnchor {
    /// Document-order byte position of the `<xdr:pic>` in its drawing part.
    pub z_order: u64,
    pub from_col: u32,
    pub from_col_off: i64,
    pub from_row: u32,
    pub from_row_off: i64,
    pub to_col: u32,
    pub to_col_off: i64,
    pub to_row: u32,
    pub to_row_off: i64,
    /// `twoCellAnchor@editAs` (ECMA-376 §20.5.2.33). Possible values: `"twoCell"`
    /// (default), `"oneCell"`, `"absolute"`. With `"oneCell"`, Excel preserves
    /// the picture's native EMU size (below) when cells are resized; with
    /// `"twoCell"`, the from/to anchor rect IS the rendered size.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub edit_as: Option<String>,
    /// `<xdr:pic><xdr:spPr><a:xfrm><a:ext cx cy>` in EMU. The picture's saved
    /// size at insert/edit time. Used as the authoritative size when
    /// `editAs == "oneCell"`. 0 = absent / use from/to rect.
    pub native_ext_cx: i64,
    pub native_ext_cy: i64,
    /// Zip path of the blip inside the package (e.g. `xl/media/image1.png`).
    /// The blip's own `r:embed` raster fallback when an svgBlip extension is
    /// present; otherwise the only source. Falls back to the SVG part itself
    /// when the picture has no raster `r:embed` (an icon inserted as a pure
    /// SVG), so the element is always drawable. The renderer fetches the bytes
    /// lazily via `extract_image` rather than receiving an inlined base64 URL.
    pub image_path: String,
    /// MIME of the blip at `image_path`, derived from its extension via the
    /// shared `mime_from_ext` (e.g. `image/png`, or `image/svg+xml` for the
    /// SVG-only fallback).
    pub mime_type: String,
    /// Microsoft 2016 SVG extension (`<a:blip><a:extLst><a:ext
    /// uri="{96DAC541-…}"><asvg:svgBlip r:embed>`, MS-ODRAWXML): the zip path of
    /// the vector *original*, so the renderer can prefer it and fall back to
    /// `image_path` on a decode failure. `None` when the picture carries no
    /// svgBlip extension (the common case). Its MIME is always `image/svg+xml`
    /// and is owned by the SVG decoder.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub svg_image_path: Option<String>,
    /// ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop, present only when the
    /// picture is cropped. `None` (the common case) ⇒ the whole blip fills the
    /// anchor rect. When set, the renderer draws only the visible sub-rectangle.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub src_rect: Option<SrcRect>,
    /// ECMA-376 §20.1.8.6 `<a:alphaModFix@amt>` — the blip's overall opacity as a
    /// fraction (0.0–1.0). `None` = fully opaque. The renderer applies it via
    /// `globalAlpha` so the picture composites over the cells beneath it.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alpha: Option<f64>,
    /// ECMA-376 §20.1.8.23 `<a:duotone>` recolour effect, resolved to its two
    /// endpoint colours. `None` = no duotone (the common case). When set, the
    /// renderer remaps the image along the `clr1`→`clr2` luminance ramp.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub duotone: Option<Duotone>,
}

/// ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop, shared across the docx,
/// pptx and xlsx parsers (see `ooxml_common::blip`).
pub use ooxml_common::blip::SrcRect;

/// ECMA-376 §20.1.8.23 `<a:duotone>` image effect, shared across the docx, pptx
/// and xlsx parsers (see `ooxml_common::blip`).
pub use ooxml_common::blip::Duotone;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellRange {
    pub top: u32,
    pub left: u32,
    pub bottom: u32,
    pub right: u32,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ConditionalFormat {
    pub sqref: Vec<CellRange>,
    pub rules: Vec<CfRule>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase", tag = "type")]
pub enum CfRule {
    #[serde(rename_all = "camelCase")]
    CellIs {
        operator: String,
        formulas: Vec<String>,
        dxf_id: Option<u32>,
        priority: i32,
    },
    #[serde(rename_all = "camelCase")]
    Expression {
        formula: String,
        dxf_id: Option<u32>,
        priority: i32,
        stop_if_true: bool,
    },
    #[serde(rename_all = "camelCase")]
    ColorScale { stops: Vec<CfStop>, priority: i32 },
    #[serde(rename_all = "camelCase")]
    DataBar {
        color: String,
        min: CfValue,
        max: CfValue,
        priority: i32,
        gradient: bool,
    },
    #[serde(rename_all = "camelCase")]
    Top10 {
        top: bool,
        percent: bool,
        rank: u32,
        dxf_id: Option<u32>,
        priority: i32,
    },
    #[serde(rename_all = "camelCase")]
    AboveAverage {
        above_average: bool,
        /// ECMA-376 §18.3.1.10 `equalAverage`: include cells equal to the
        /// average in the highlighted set. Default false.
        #[serde(skip_serializing_if = "std::ops::Not::not")]
        equal_average: bool,
        /// ECMA-376 §18.3.1.10 `stdDev`: when present, the threshold becomes
        /// `mean ± stdDev · σ` (population standard deviation) instead of the
        /// plain mean. Absent (None) means a simple above/below-average rule.
        #[serde(skip_serializing_if = "Option::is_none")]
        std_dev: Option<u32>,
        dxf_id: Option<u32>,
        priority: i32,
    },
    #[serde(rename_all = "camelCase")]
    IconSet {
        icon_set: String,
        cfvos: Vec<CfValue>,
        reverse: bool,
        priority: i32,
        #[serde(skip_serializing_if = "Option::is_none")]
        custom_icons: Option<Vec<CfIcon>>,
    },
    #[serde(rename_all = "camelCase")]
    Other { kind: String, priority: i32 },
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CfIcon {
    pub icon_set: String,
    pub icon_id: u32,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CfStop {
    pub kind: String,
    pub value: Option<String>,
    pub color: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CfValue {
    pub kind: String,
    pub value: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Hyperlink {
    pub col: u32,
    pub row: u32,
    /// External target (ECMA-376 §18.3.1.47 `r:id` → resolved via worksheet
    /// rels). `None` for a purely internal hyperlink.
    pub url: Option<String>,
    /// Internal target (§18.3.1.47 `location` attribute): a defined name or a
    /// cell reference such as `Sheet1!A1`. Inline — no rels resolution needed.
    /// A `<hyperlink>` may carry `r:id` (external), `location` (internal), or
    /// both.
    pub location: Option<String>,
    /// Optional display text (§18.3.1.47 `display`). Not used for rendering
    /// (the cell value drives that) but preserved for callers.
    pub display: Option<String>,
}

/// `<sheetPr><outlinePr>` display flags (ECMA-376 §18.3.1.61). When
/// `summary_below` is `true` (the default) a group's summary row sits *below*
/// its detail rows; when `false`, above. `summary_right` is the column analogue
/// (summary to the right of detail by default). These decide where the outline
/// gutter draws each group's collapse toggle.
#[derive(Debug, Serialize, Clone, Copy)]
#[serde(rename_all = "camelCase")]
pub struct OutlinePr {
    pub summary_below: bool,
    pub summary_right: bool,
}

impl Default for OutlinePr {
    fn default() -> Self {
        // ECMA-376 §18.3.1.61 defaults: both flags true.
        OutlinePr {
            summary_below: true,
            summary_right: true,
        }
    }
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    pub index: u32,
    pub height: Option<f64>,
    /// `<row customHeight>` (§18.3.1.73): the row height was manually set.
    /// Retained even when producers omit `ht`, because Office then
    /// keeps the sheet default rather than applying display-time auto-fit.
    #[serde(skip_serializing_if = "std::ops::Not::not", default)]
    pub custom_height: bool,
    pub cells: Vec<Cell>,
    /// Outline (grouping) depth of this row, 0-7 (ECMA-376 §18.3.1.73
    /// `<row outlineLevel>`). `0` (the ungrouped default) is omitted from JSON
    /// so a sheet without outlining serializes byte-for-byte as before.
    #[serde(skip_serializing_if = "is_zero_u8", default)]
    pub outline_level: u8,
    /// `<row collapsed>` (§18.3.1.73): set on the *summary* row when the rows one
    /// outline level deeper are in the collapsed state. Drives which +/- toggle
    /// the gutter draws. Omitted from JSON when false (the schema default).
    #[serde(skip_serializing_if = "std::ops::Not::not", default)]
    pub collapsed: bool,
    /// `<row hidden>` (§18.3.1.73): `1` when the row is hidden — most commonly
    /// because a collapsed outline hides its detail rows, but also plain
    /// user-hidden rows. Kept as an explicit flag (distinct from `height == 0`)
    /// so the outline gutter can tell collapsed-detail rows apart from
    /// zero-height rows. Omitted from JSON when false.
    #[serde(skip_serializing_if = "std::ops::Not::not", default)]
    pub hidden: bool,
}

/// serde `skip_serializing_if` predicate for the common ungrouped `outlineLevel`
/// of `0`, so a sheet without outlining keeps its byte-identical JSON.
fn is_zero_u8(v: &u8) -> bool {
    *v == 0
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Cell {
    pub col: u32,
    pub row: u32,
    pub value: CellValue,
    /// Explicit `c/@s`. `None` inherits the owning column's `<col style>`;
    /// `Some(0)` is intentionally distinct and resets to the Normal XF.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_index: Option<u32>,
    /// Raw `<f>` formula text (ECMA-376 §18.3.1.40), when present. The
    /// renderer uses this to recompute volatile functions like TODAY() /
    /// NOW() at display time so the cached `<v>` (frozen when the file was
    /// last saved) doesn't show a stale date.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    /// ECMA-376 §18.3.1.4 `<c ph="1">` — whether this cell should display its
    /// phonetic hint (furigana). Defaults to `false` (the schema default), so a
    /// cell whose String Item carries `<rPh>` runs still shows NO furigana
    /// unless the cell opts in with `ph="1"`. This mirrors Excel: the shared
    /// string keeps the reading, but display is a per-cell toggle.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub show_phonetic: bool,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase", tag = "type")]
pub enum CellValue {
    #[default]
    Empty,
    Text {
        text: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        runs: Option<Vec<Run>>,
        /// ECMA-376 §18.4.6 phonetic runs (furigana) for this text, carried
        /// over from the resolved String Item. Only populated for inline
        /// strings here; shared-string cells ship a `Shared { si }` reference
        /// and the consumer pulls the phonetic data off the `SharedString`.
        #[serde(skip_serializing_if = "Vec::is_empty")]
        phonetic_runs: Vec<PhoneticRun>,
        /// ECMA-376 §18.4.3 phonetic display properties for the furigana above.
        #[serde(skip_serializing_if = "Option::is_none")]
        phonetic_pr: Option<PhoneticProperties>,
    },
    Number {
        number: f64,
    },
    Bool {
        bool: bool,
    },
    Error {
        error: String,
    },
    /// Shared-string reference into the workbook `sharedStrings` table
    /// (ECMA-376 §18.4.8 `<si>` / §18.3.1.4 cell `t="s"`), resolved on the
    /// consumer side so the full string + runs are shipped once in the
    /// workbook, not cloned per cell.
    Shared {
        si: usize,
    },
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Run {
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font: Option<RunFont>,
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct RunFont {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// ECMA-376 §18.4.13 ST_UnderlineValues — see Font.underline_style.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline_style: Option<String>,
    /// ECMA-376 §18.4.6 ST_VerticalAlignRun — "superscript" | "subscript" |
    /// "baseline". Absent leaves the run on the baseline.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<String>,
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct SharedString {
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub runs: Option<Vec<Run>>,
    /// ECMA-376 §18.4.6 `<rPh>` phonetic runs (furigana) for this string item.
    /// Each run's `sb`/`eb` are zero-based character offsets into `text`; the
    /// hint applies over `text[sb..eb)`. Empty when the `<si>` carries no
    /// furigana. Kept separate from `text`/`runs` so the base string a
    /// consumer searches / measures never includes the phonetic hint (matching
    /// Excel, whose Find searches base text, not furigana).
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub phonetic_runs: Vec<PhoneticRun>,
    /// ECMA-376 §18.4.3 `<phoneticPr>` — how the furigana is displayed (font
    /// record index, character set, alignment). Absent when the `<si>` has no
    /// `<phoneticPr>`; the renderer then applies the schema defaults
    /// (fullwidthKatakana / left) and falls back to font 0 for the size.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub phonetic_pr: Option<PhoneticProperties>,
}

/// ECMA-376 §18.4.6 `<rPh sb=".." eb=".."><t>..</t></rPh>` — one furigana run.
/// `sb`/`eb` are zero-based **character** (Unicode scalar) offsets into the
/// String Item's base text; the phonetic hint `text` is shown over the base
/// characters `[sb, eb)`. The spec requires `sb < eb` and both within the base
/// length; the renderer positions the hint over that base span.
#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct PhoneticRun {
    /// Zero-based start character offset into the base text (inclusive).
    pub sb: u32,
    /// Zero-based end character offset into the base text (exclusive).
    pub eb: u32,
    /// The phonetic hint text (e.g. the katakana reading).
    pub text: String,
}

/// ECMA-376 §18.4.3 `<phoneticPr>` — phonetic display properties for a String
/// Item. `fontId` is a required zero-based index into `styles.xml`'s font
/// records; out of bounds falls back to font 0 (§18.4.3). `type` and
/// `alignment` default to `fullwidthKatakana` and `left` respectively when the
/// attribute is absent (CT_PhoneticPr schema defaults).
#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct PhoneticProperties {
    /// Zero-based index into the style sheet's font records (§18.18.32 ST_FontId).
    pub font_id: u32,
    /// ECMA-376 §18.18.57 ST_PhoneticType — "fullwidthKatakana" (default) |
    /// "halfwidthKatakana" | "Hiragana" | "noConversion".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub r#type: Option<String>,
    /// ECMA-376 §18.18.56 ST_PhoneticAlignment — "left" (default) | "center" |
    /// "distributed" | "noControl".
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<String>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Styles {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
    pub cell_xfs: Vec<CellXf>,
    pub num_fmts: Vec<NumFmt>,
    pub dxfs: Vec<Dxf>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Dxf {
    pub font: Option<Font>,
    pub fill: Option<Fill>,
    pub border: Option<Border>,
    /// Number format override applied when the conditional-formatting rule
    /// matches. ECMA-376 §18.8.17 allows `<dxf>` to carry a `<numFmt>` that
    /// replaces the cell's own style numFmt (e.g. switching a calendar cell
    /// from `d` to `m"月"d"日"` on the first of each month).
    pub num_fmt: Option<NumFmt>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Font {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub size: f64,
    pub color: Option<String>,
    pub name: Option<String>,
    /// ECMA-376 §18.4.13 ST_UnderlineValues. Only emitted when not the default
    /// "single" — values: "double", "singleAccounting", "doubleAccounting".
    /// "none" sets `underline = false` and leaves this field absent.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline_style: Option<String>,
    /// ECMA-376 §18.4.6 ST_VerticalAlignRun on a cell-level <font> —
    /// "superscript" | "subscript". Absent leaves the run on the baseline.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<String>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Fill {
    pub pattern_type: String,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
    /// When the fill element is a <gradientFill>, this carries the gradient
    /// stops + type + rotation. patternType stays "none" because xlsx does
    /// not mix gradient + pattern in the same fill.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gradient: Option<GradientFillSpec>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct GradientFillSpec {
    /// "linear" (default) or "path". Linear uses `degree`; path uses top/bottom/left/right.
    pub gradient_type: String,
    /// Linear-gradient rotation in degrees (0 = left→right).
    pub degree: f64,
    /// Path-gradient bounding box (0..1 within the cell). Unused for linear.
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub stops: Vec<GradientStopSpec>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct GradientStopSpec {
    pub position: f64,
    pub color: String,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Border {
    pub left: Option<BorderEdge>,
    pub right: Option<BorderEdge>,
    pub top: Option<BorderEdge>,
    pub bottom: Option<BorderEdge>,
    /// Diagonal line from bottom-left to top-right (ECMA-376 §18.8.4 diagonalUp)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagonal_up: Option<BorderEdge>,
    /// Diagonal line from top-left to bottom-right (ECMA-376 §18.8.4 diagonalDown)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagonal_down: Option<BorderEdge>,
    /// Inner horizontal rule between rows inside a region (ECMA-376 §18.8.40
    /// `tableStyleElement/border/horizontal`). Only set on table-style dxfs;
    /// ignored on cell-level borders.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<BorderEdge>,
    /// Inner vertical rule between columns inside a region (same ECMA-376
    /// section). Only set on table-style dxfs.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical: Option<BorderEdge>,
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct BorderEdge {
    pub style: String,
    pub color: Option<String>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct CellXf {
    pub font_id: u32,
    pub fill_id: u32,
    pub border_id: u32,
    pub num_fmt_id: u32,
    pub align_h: Option<String>,
    pub align_v: Option<String>,
    pub wrap_text: bool,
    /// Text indentation level (each level ≈ 3 characters wide, ECMA-376 §18.8.44)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<u32>,
    /// Text rotation in degrees: 0–90 = counter-clockwise, 91–180 = (value−90)° clockwise, 255 = stacked (ECMA-376 §18.8.44)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_rotation: Option<u32>,
    /// Shrink text to fit the cell width (ECMA-376 §18.8.44)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub shrink_to_fit: bool,
    /// `<alignment readingOrder>` (ECMA-376 §18.8.1) — 0 = context (default),
    /// 1 = left-to-right, 2 = right-to-left. Drives canvas `direction`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reading_order: Option<u32>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct NumFmt {
    pub num_fmt_id: u32,
    pub format_code: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParsedWorkbook {
    pub workbook: Workbook,
    pub styles: Styles,
    pub shared_strings: Vec<SharedString>,
}
