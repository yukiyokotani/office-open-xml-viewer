use crate::xml_util::*;
use ooxml_common::depth::parse_guarded;
use std::collections::{HashMap, HashSet};

/// ECMA-376 §17.3.2.14 `<w:fitText>` as one cascaded run property.
#[derive(Clone, PartialEq, Debug)]
pub struct FitTextSpec {
    /// Target manual run width in TWIPS (`ST_TwipsMeasure`, 1/20 pt).
    pub val: f64,
    /// Raw, trimmed `ST_DecimalNumber` identifier. Its arbitrary-precision XSD
    /// integer domain is preserved because it is used only for equality linking.
    pub id: Option<String>,
}

/// Resolved run (character) formatting.
#[derive(Debug, Clone, Default)]
pub struct RunFmt {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    /// ECMA-376 §17.3.2.40 `<w:u w:val>` — the raw ST_Underline (§17.18.99)
    /// style value: `single` / `double` / `thick` / `dotted` / `dottedHeavy` /
    /// `dash` / `dashedHeavy` / `dashLong` / `dashLongHeavy` / `dotDash` /
    /// `dashDotHeavy` / `dotDotDash` / `dashDotDotHeavy` / `wave` / `wavyHeavy` /
    /// `wavyDouble` / `words`. `None` for absent or `none`/`single` (the default
    /// single rule needs no style hint; the renderer draws single from
    /// `underline` alone). The renderer normalizes this WordprocessingML
    /// vocabulary to the shared DrawingML ST_TextUnderlineType (§20.1.10.82).
    pub underline_style: Option<String>,
    /// ECMA-376 §17.3.2.40 `<w:u w:color>` — an underline-only colour (hex 6, or
    /// the literal `auto`). `None` means the underline follows the glyph colour.
    pub underline_color: Option<String>,
    /// Complete private `<w:u>` facts. The public bool/style/color projection is
    /// retained for API compatibility; layout acquisition consumes this lexical
    /// value/status representation so invalid enums are not normalized away.
    pub(crate) underline_typography: Option<crate::types::UnderlineTypographyWire>,
    pub strikethrough: Option<bool>,
    pub font_size: Option<f64>, // pt
    pub color: Option<String>,  // hex 6
    /// ECMA-376 §17.3.2.6 `<w:color w:val="auto"/>` — an explicit *automatic*
    /// color. Unlike a concrete `color`, auto does not name a value; the final
    /// color is resolved from the effective background at render time. The
    /// black/white contrast resolution is implementation-defined — ECMA-376
    /// gives no normative algorithm; `auto` itself is defined by §17.3.2.6 /
    /// ST_HexColorAuto §17.18.39. `color_auto` also breaks inheritance in
    /// `apply_run`, dropping any
    /// inherited concrete color so e.g. PlaceholderText gray does not survive.
    pub color_auto: bool,
    pub font_family_ascii: Option<String>,
    pub font_family_high_ansi: Option<String>,
    pub font_family_east_asia: Option<String>,
    pub font_family_ascii_direct: Option<String>,
    pub font_family_ascii_theme: Option<String>,
    pub font_family_high_ansi_direct: Option<String>,
    pub font_family_high_ansi_theme: Option<String>,
    pub font_family_east_asia_direct: Option<String>,
    pub font_family_east_asia_theme: Option<String>,
    /// ECMA-376 §17.3.2.26 rFonts@hint, resolved through the run style chain.
    pub font_hint: Option<String>,
    pub background: Option<String>, // hex 6
    /// "super" | "sub" — mapped from w:vertAlign val="superscript|subscript"
    pub vert_align: Option<String>,
    /// All caps (w:caps)
    pub all_caps: Option<bool>,
    /// Small capitals (w:smallCaps)
    pub small_caps: Option<bool>,
    /// Double strikethrough (w:dstrike)
    pub dstrike: Option<bool>,
    /// ECMA-376 §17.3.2.41 w:vanish — "Hidden Text". Hidden in the normal
    /// (print/page) view this renderer produces, so the parser skips the run.
    /// This is NOT `webHidden`: vanish hides in all non-web views.
    pub vanish: Option<bool>,
    /// ECMA-376 §17.3.2.44 w:webHidden — text hidden ONLY when the document is
    /// shown in *web page view* (§17.18.102). In the normal/print layout we
    /// render, webHidden text is visible, so it must NOT set `vanish` (doing so
    /// dropped TOC dot-leader tabs and PAGEREF page numbers, which Word marks
    /// `<w:webHidden/>`). Preserved here so a future web-view mode can honor it.
    pub web_hidden: Option<bool>,
    /// Highlight color name: "yellow" | "cyan" | "green" | ... (w:highlight)
    pub highlight: Option<String>,
    /// ECMA-376 §17.3.2.12 `<w:em w:val>` — emphasis (boten) mark applied to
    /// each non-space character of the run (§17.18.24 ST_Em). One of "dot" |
    /// "comma" | "circle" | "underDot"; `val="none"` filters to `None` (no mark).
    /// Resolved through the same style chain as `highlight`: a level that sets a
    /// concrete value wins. (Like `highlight`, `none` collapses to `None` at
    /// parse time, so it reads as "unset" rather than an inherit-clearing
    /// override — emphasis marks are effectively never inherited-then-cleared in
    /// practice, and this keeps the value-property resolution identical to
    /// `highlight`.)
    pub emphasis_mark: Option<String>,
    /// ECMA-376 §17.3.2.4 `<w:bdr>` — a run-level border drawn as a box around
    /// the run's text. Reuses `EdgeBorder` (width pt, color, style, space pt).
    pub border: Option<EdgeBorder>,
    /// Full CT_Border (§17.3.4) attributes, including theme/shadow/frame facts
    /// omitted by the stable public EdgeBorder projection.
    pub(crate) border_typography: Option<crate::types::CtBorderTypographyWire>,
    /// ECMA-376 §17.3.2.30 w:rtl — this run contains complex-script (RTL)
    /// content. Resolved through the style chain onto the run model, where the
    /// renderer uses it to force complex-script shaping and feed the UAX#9
    /// reordering / `ctx.direction = 'rtl'` mirroring pass.
    pub rtl: Option<bool>,
    /// ECMA-376 §17.3.2.34 `<w:snapToGrid>` — whether this run participates in
    /// the section character grid. `None` inherits; explicit false opts out.
    pub snap_to_grid: Option<bool>,
    /// ECMA-376 §17.3.2.7 `<w:cs/>` — treat the run as complex-script,
    /// applying the cs formatting axis to ALL its characters (§17.3.2.26).
    pub cs_toggle: Option<bool>,
    /// ECMA-376 §17.3.2.26 w:rFonts/@w:cs — complex-script typeface. Parallel
    /// to `font_family_ascii` / `font_family_east_asia`; carries the same
    /// "@theme:<ref>" marker convention for @cstheme references.
    pub font_family_cs: Option<String>,
    pub font_family_cs_direct: Option<String>,
    pub font_family_cs_theme: Option<String>,
    /// ECMA-376 §17.3.2.39 w:szCs/@w:val — complex-script font size in pt
    /// (converted from half-points, same units as `font_size`).
    pub font_size_cs: Option<f64>,
    /// True when `w:sz` was set at THIS level (direct rPr or a named style),
    /// regardless of whether `w:szCs` was also set. Word resolves the
    /// complex-script size by mirroring a directly-applied `w:sz` when no
    /// `w:szCs` accompanies it at the same level (a directly-set Latin size
    /// shadows an inherited complex-script size); §17.3.2.38 / §17.3.2.18.
    pub font_size_set_here: bool,
    /// True when `w:szCs` was set at THIS level. Lets the merge mirror an
    /// unaccompanied `w:sz` into `font_size_cs` (see `apply_run`).
    pub font_size_cs_set_here: bool,
    /// ECMA-376 §17.3.2.3 w:bCs — complex-script bold toggle.
    pub bold_cs: Option<bool>,
    /// ECMA-376 §17.3.2.17 w:iCs — complex-script italic toggle.
    pub italic_cs: Option<bool>,
    /// ECMA-376 §17.3.2.20 w:lang/@w:bidi — complex-script (RTL) language tag,
    /// lower-cased (e.g. "ar-sa", "ae-ar"). Drives Word's AN digit ordering.
    pub lang_bidi: Option<String>,
    /// ECMA-376 §17.3.2.20 w:lang/@w:eastAsia, lower-cased.
    pub lang_east_asia: Option<String>,
    /// ECMA-376 §17.3.2.35 `<w:spacing w:val>` — character-spacing adjustment,
    /// the pitch added AFTER each character before the next is rendered
    /// ("equivalent to the additional character pitch added by a document
    /// grid"). Stored in POINTS, signed (source is ST_SignedTwipsMeasure =
    /// twips = 1/20 pt); the renderer feeds it to `ctx.letterSpacing` for both
    /// measure and paint so wrapping/pagination stay measure==paint. `None` =
    /// inherit (no additional pitch when never set in the style hierarchy).
    pub char_spacing: Option<f64>,
    /// ECMA-376 §17.3.2.14 `<w:fitText>` — manual run width and optional link
    /// id. The element cascades as one composite property, so a direct element
    /// without `w:id` clears an inherited id instead of retaining it.
    pub fit_text: Option<FitTextSpec>,
    /// ECMA-376 §17.3.2.43 `<w:w w:val>` — horizontal Expanded/Compressed text
    /// scale. ST_TextScale is a percentage of the normal (100%) character width
    /// (1%–600%), stored here as a FRACTION (e.g. `w:val="67"` or `"67%"` →
    /// 0.67). Unlike `char_spacing` this stretches each glyph's WIDTH, not the
    /// gap between glyphs. `None` = inherit (100% when never set).
    pub char_scale: Option<f64>,
    /// ECMA-376 §17.3.2.24 `<w:position w:val>` — baseline raise (positive) or
    /// lower (negative) WITHOUT changing the font size. Stored in POINTS, signed
    /// (source is ST_SignedHpsMeasure = half-points = 1/144 in). `None` =
    /// inherit (no shift when never set). Word does not grow the line box for a
    /// positioned run; the shift is a pure baseline offset (§17.3.2.24).
    pub position: Option<f64>,
    /// ECMA-376 §17.3.2.19 `<w:kern w:val>` — the SMALLEST font size (threshold)
    /// that has automatic font kerning applied; a run whose `sz` is below this
    /// value is not kerned. Stored in POINTS (source is ST_HpsMeasure =
    /// half-points). Presence itself enables kerning (subject to the threshold);
    /// `None` = inherit, and "never set in the hierarchy" ⇒ no kerning at all
    /// (Word's default is OFF, unlike Canvas's default `fontKerning='auto'`).
    /// `Some(0.0)` = kern at every size.
    pub kerning: Option<f64>,
    /// ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:vert>` — "Horizontal in Vertical
    /// (Rotate Text)" (縦中横 / tate-chū-yoko). When `Some(true)`, in a VERTICAL
    /// (tbRl) document the run's characters are laid out HORIZONTALLY within a
    /// single cell of the vertical line ("keeping the text on the same line"),
    /// i.e. rotated 90° relative to the surrounding vertical flow. `None` =
    /// inherit; the property is inert in a horizontal document (§17.3.2.10 is only
    /// meaningful in vertical text).
    pub east_asian_vert: Option<bool>,
    /// ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:vertCompress>` — "Compress
    /// Rotated Text to Line Height". Ignored unless `east_asian_vert` is set. When
    /// `Some(true)`, the horizontally-laid-out (rotated) run is compressed so it
    /// fits into the existing line height without growing the line. `None` =
    /// inherit (default: not compressed).
    pub east_asian_vert_compress: Option<bool>,
    /// ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:combine>` — "Two Lines in One":
    /// the run's characters are written on two sub-lines within one logical line.
    /// PARSED for completeness (§17.3.2.10) but not yet rendered (no fixture);
    /// carried so a future two-lines-in-one implementation has the flag. `None` =
    /// inherit / off.
    pub east_asian_combine: Option<bool>,
    /// ECMA-376 §17.3.2.10 `<w:eastAsianLayout w:combineBrackets>` — the bracket
    /// style (§17.18.8 ST_CombineBrackets: `none` | `round` | `square` | `angle` |
    /// `curly`) enclosing two-lines-in-one text. Ignored unless `combine` is set.
    /// PARSED for completeness; the two-lines-in-one draw is a follow-up. `None` =
    /// no brackets.
    pub east_asian_combine_brackets: Option<String>,
    pub(crate) vertical_align_typography: Option<crate::types::TypographyValueWire<String>>,
    pub(crate) position_typography: Option<crate::types::TypographyValueWire<f64>>,
    pub(crate) emphasis_typography: Option<crate::types::TypographyValueWire<String>>,
    pub(crate) combine_brackets_typography: Option<crate::types::TypographyValueWire<String>>,
}

/// Resolved paragraph formatting.
#[derive(Debug, Clone, Default)]
pub struct ParaFmt {
    pub alignment: Option<String>,
    pub indent_left: Option<f64>,  // pt
    pub indent_right: Option<f64>, // pt
    pub indent_first: Option<f64>, // pt
    pub space_before: Option<f64>, // pt
    pub space_after: Option<f64>,  // pt
    pub line_spacing_val: Option<f64>,
    pub line_spacing_rule: Option<String>,
    /// True when `w:spacing/@w:line` was declared on the paragraph's own pPr
    /// or on one of its named styles (i.e. NOT inherited from docDefaults).
    /// Per ECMA-376 §17.6.5, when docGrid is active a paragraph that only
    /// inherits line from docDefault uses the grid pitch (1 grid line per
    /// text line, ignoring the multiplier), while an explicitly-set line
    /// multiplies against the pitch as usual.
    pub line_spacing_explicit: Option<bool>,
    pub num_id: Option<u32>,
    pub num_level: Option<u32>,
    /// Explicit tab stops (pos_pt, alignment, leader). None = inherit from parent style chain.
    pub tab_stops: Option<Vec<(f64, String, String)>>,
    /// merged run defaults from pPr/rPr
    pub run: RunFmt,
    /// Paragraph background hex color (w:shd fill on paragraph)
    pub shading: Option<String>,
    /// Force page break before paragraph (w:pageBreakBefore)
    pub page_break_before: Option<bool>,
    /// Suppress spacing between adjacent same-style paragraphs (w:contextualSpacing)
    pub contextual_spacing: Option<bool>,
    /// Keep paragraph on same page as the next paragraph (w:keepNext)
    pub keep_next: Option<bool>,
    /// Keep all lines of this paragraph on the same page (w:keepLines)
    pub keep_lines: Option<bool>,
    /// Widow/orphan control (w:widowControl). Default per spec: true.
    pub widow_control: Option<bool>,
    /// ECMA-376 §17.3.1.21 w:overflowPunct — allow one trailing punctuation
    /// character beyond the paragraph extents. The final default (`true`) is
    /// applied when the resolved paragraph is built so style/direct `false`
    /// remains distinguishable from omission during the cascade.
    pub overflow_punct: Option<bool>,
    /// ECMA-376 §17.3.1.1 w:adjustRightInd — allow the consumer to adjust the
    /// effective right indent when a document grid is active. Retained as an
    /// Option through the style cascade because omission inherits and the final
    /// specification default is true.
    pub adjust_right_ind: Option<bool>,
    /// Paragraph border edges (w:pBdr)
    pub para_borders: Option<crate::types::ParagraphBorders>,
    pub(crate) paragraph_typography: Option<crate::types::ParagraphTypographyWire>,
    /// Heading outline level (w:outlineLvl, 0–8) when set. Word's built-in
    /// heading styles (Heading 1–9) are rendered with an implicit
    /// `w:keepNext` even when not spelled out in styles.xml; downstream
    /// code uses this to infer that behavior.
    pub outline_level: Option<u32>,
    /// ECMA-376 §17.3.1.6 w:bidi — right-to-left paragraph (text and column
    /// order flow RTL). Resolved through the style chain onto the paragraph
    /// model, where the renderer uses it as the base direction for UAX#9
    /// reordering, indent swapping, and `w:jc` start/end edge resolution.
    pub bidi: Option<bool>,
    /// ECMA-376 §17.3.1.32 w:snapToGrid — when `Some(false)` this paragraph
    /// opts OUT of the document grid (`w:docGrid`), so its lines use natural
    /// font metrics / the line-spacing multiplier directly instead of snapping
    /// to the grid pitch. `None` = inherit (default true). Word's built-in
    /// "Footnote Text" style sets this off, which is why footnote bodies use
    /// compact natural line height rather than the 18 pt grid pitch.
    pub snap_to_grid: Option<bool>,
    /// ECMA-376 §17.3.1.11 w:framePr — text-frame / drop-cap properties.
    /// Resolved through the style chain like other pPr; `Some` ⇒ this paragraph
    /// is part of a text frame. Boxed to match `DocParagraph::frame_pr`.
    pub frame_pr: Option<Box<crate::types::FramePr>>,
}

#[derive(Debug, Default)]
pub struct StyleDef {
    pub para: ParaFmt,
    pub run: RunFmt,
    pub based_on: Option<String>,
}

/// One border edge from a table style (val/sz/color), pt-converted.
#[derive(Debug, Default, Clone)]
pub struct EdgeBorder {
    pub width: f64,
    pub color: Option<String>,
    pub style: String,
    /// ECMA-376 CT_Border `@w:space` — spacing between the border and the
    /// content, in points (not eighths). Defaults to 0; harmless for table
    /// borders, which never set it.
    pub space: f64,
}

#[derive(Debug, Default, Clone)]
pub struct RawTblBorders {
    pub top: Option<EdgeBorder>,
    pub bottom: Option<EdgeBorder>,
    pub left: Option<EdgeBorder>,
    pub right: Option<EdgeBorder>,
    pub inside_h: Option<EdgeBorder>,
    pub inside_v: Option<EdgeBorder>,
}

/// Conditional formatting block (`w:tblStylePr`) — the subset we resolve.
#[derive(Debug, Default, Clone)]
pub struct CondFmt {
    pub shd: Option<String>,
    pub borders: RawTblBorders,
    /// ECMA-376 §17.7.6: the conditional block's `<w:rPr>` — run defaults that
    /// apply to runs in cells covered by this condition (e.g. Calendar 3's
    /// firstRow sets `<w:color w:val="365F91"/>` for the "Sun/Mon/…" header).
    /// Layered as a BASE below paragraph/character styles and direct rPr
    /// (§17.7.2), so a directly-colored run still wins.
    pub run: Option<RunFmt>,
    /// ECMA-376 §17.7.6: the conditional block's `<w:pPr>` — paragraph defaults
    /// for cells covered by this condition (e.g. firstRow `<w:jc w:val="right"/>`).
    pub para: Option<ParaFmt>,
    /// Row-wide table-style spacing remains lexical until the layout boundary
    /// applies CT_TblWidth defaults and the documented Word deviations.
    pub cell_spacing: Option<crate::types::TableWidthAcquisitionWire>,
    /// Full-row conditional table margin layer (§17.7.6.3).
    pub cell_margins: Option<crate::types::TableMarginAcquisitionWire>,
    /// Row pagination properties participate in the table-style cascade; an
    /// explicit `false` must remain distinguishable from an omitted property.
    pub tbl_header: Option<bool>,
    pub cant_split: Option<bool>,
}

/// Fold an ordered list of conditional-format layers into one effective
/// [`CondFmt`]. `layers` must be supplied **lowest precedence first**; later
/// layers override earlier ones, matching ECMA-376 §17.7.6 ("these conditional
/// formats shall be applied in the following order […] therefore subsequent
/// formats override properties on previous formats"). The §17.7.6 order is
/// wholeTable < band*Vert < band*Horz < firstRow/lastRow < firstCol/lastCol <
/// corner cells; the caller is responsible for assembling `layers` in that
/// order. Within each layer the same set-overrides-unset merge semantics used
/// elsewhere apply (shd, borders, rPr, pPr), so a cell covered by both a band
/// and firstCol picks up firstCol's run color while keeping band borders the
/// firstCol layer left unset.
pub fn merge_cond_layers(layers: &[&CondFmt]) -> CondFmt {
    let mut out = CondFmt::default();
    for layer in layers {
        if layer.shd.is_some() {
            out.shd = layer.shd.clone();
        }
        merge_raw_borders(&mut out.borders, &layer.borders);
        if let Some(r) = &layer.run {
            apply_run(out.run.get_or_insert_with(RunFmt::default), r);
        }
        if let Some(p) = &layer.para {
            apply_para(out.para.get_or_insert_with(ParaFmt::default), p);
        }
        if layer.cell_spacing.is_some() {
            out.cell_spacing = layer.cell_spacing.clone();
        }
        merge_table_margin_layer(&mut out.cell_margins, &layer.cell_margins);
        if layer.tbl_header.is_some() {
            out.tbl_header = layer.tbl_header;
        }
        if layer.cant_split.is_some() {
            out.cant_split = layer.cant_split;
        }
    }
    out
}

/// Table style (`w:style w:type="table"`) cell/border formatting.
#[derive(Debug, Default, Clone)]
pub struct TableStyleDef {
    pub based_on: Option<String>,
    pub borders: RawTblBorders,
    pub cell_shd: Option<String>,
    pub cell_valign: Option<String>,
    /// ECMA-376 §17.7.6: the table style's whole-table `<w:rPr>` — run defaults
    /// applied to every cell (e.g. Calendar 3's `<w:color w:val="7F7F7F"/>`
    /// makes day numbers gray). Layered below the conditional `run` but above
    /// docDefaults (§17.7.2).
    pub run: Option<RunFmt>,
    /// ECMA-376 §17.7.6: the table style's whole-table `<w:pPr>` — paragraph
    /// defaults for every cell (e.g. Calendar 3's `<w:jc w:val="right"/>`).
    /// "Table Grid"'s line/after spacing is already threaded through
    /// `StyleMap::styles`; this field carries the table style's pPr for the
    /// conditional-formatting resolution path used by `resolve_table_cond`.
    pub para: Option<ParaFmt>,
    /// ECMA-376 §17.7.6.7 `<w:tblStyleRowBandSize>`: number of rows per
    /// horizontal band (band1Horz/band2Horz alternate every N rows). `None`
    /// means the element was omitted; callers treat that as the spec default 1.
    /// Kept as `Option` so a derived style that omits it inherits the base
    /// style's value through `resolve_table_style` instead of resetting to 1.
    pub row_band_size: Option<usize>,
    /// ECMA-376 §17.7.6.5 `<w:tblStyleColBandSize>`: number of columns per
    /// vertical band (band1Vert/band2Vert alternate every N columns). `None` =
    /// omitted; callers treat that as the spec default 1. See `row_band_size`
    /// for why this is an `Option`.
    pub col_band_size: Option<usize>,
    /// keyed by w:tblStylePr w:type (firstRow, band1Horz, band2Horz, …).
    pub cond: HashMap<String, CondFmt>,
    /// ECMA-376 §17.4.42 `<w:tblPr><w:tblCellMar>`: per-edge default cell
    /// margins (in points). `None` means the style omitted that edge, in
    /// which case the value inherits from the basedOn chain via
    /// `resolve_table_style`. The default table style (`<w:style
    /// w:type="table" w:default="1">`, typically "TableNormal") carries the
    /// values Word/Office apply when a table omits `<w:tblCellMar>`
    /// entirely; the parser falls back to these values through
    /// `default_table_style_id`. If a margin is never specified in the
    /// style hierarchy, §17.4.34 / §17.4.11 / §17.4.5 / §17.4.75 define the
    /// spec defaults (115 twips for left/right, 0 for top/bottom) which the
    /// caller applies.
    pub cell_margin_top: Option<f64>,
    pub cell_margin_bottom: Option<f64>,
    pub cell_margin_left: Option<f64>,
    pub cell_margin_right: Option<f64>,
    /// Lexical margin layer retained independently from the compatibility
    /// numeric projection so start/end can be mapped after bidiVisual is known.
    pub cell_margins: Option<crate::types::TableMarginAcquisitionWire>,
    /// Whole-table style spacing, inherited through basedOn.
    pub cell_spacing: Option<crate::types::TableWidthAcquisitionWire>,
    /// ECMA-376 §17.7.6.10/.11: row properties supplied by the table style.
    /// `Option` preserves inheritance and explicit CT_OnOff false values.
    pub tbl_header: Option<bool>,
    pub cant_split: Option<bool>,
}

#[derive(Default)]
pub struct StyleMap {
    styles: HashMap<String, StyleDef>,
    /// Style ID of the paragraph style whose primary name is `Normal`.
    /// `w:styleId` is an arbitrary reference; Word's built-in-style identity is
    /// defined by the reserved primary style name ([MS-OE376] §2.1.236).
    normal_para_style_id: Option<String>,
    table_styles: HashMap<String, TableStyleDef>,
    defaults_para: ParaFmt,
    defaults_run: RunFmt,
    /// styleId of the style with w:default="1" and w:type="paragraph".
    /// Applied to paragraphs that have no explicit pStyle.
    default_para_style_id: Option<String>,
    /// styleId of the style with `w:default="1"` and `w:type="table"`
    /// (typically "TableNormal" / "Normal Table"). ECMA-376 §17.7.4 makes
    /// this the implicit table style for tables that omit `<w:tblStyle>`,
    /// so its `<w:tblCellMar>` (and any other tblPr defaults) apply when a
    /// table inherits silently. Stored separately from
    /// `default_para_style_id` since the paragraph default and table default
    /// are independent.
    default_table_style_id: Option<String>,
}

impl StyleMap {
    /// Style ID of the paragraph style marked `w:default="1"` in styles.xml.
    /// International templates may use non-English IDs (e.g. "a", "標準").
    pub fn default_para_style_id(&self) -> Option<&str> {
        self.default_para_style_id.as_deref()
    }

    /// Style ID of the `w:type="table" w:default="1"` style (typically
    /// "TableNormal" / "Normal Table"). ECMA-376 §17.4.42 + §17.7.4: a
    /// table that omits `<w:tblCellMar>` inherits from its associated table
    /// style, and a table that omits `<w:tblStyle>` is associated with the
    /// default table style. Callers in `parser.rs` use this to resolve cell
    /// margin defaults when the document never names a table style on the
    /// table itself.
    pub fn default_table_style_id(&self) -> Option<&str> {
        self.default_table_style_id.as_deref()
    }

    /// Whether `style_id` names an existing `<w:style w:type="table">`. Only
    /// table-typed styles are indexed in `table_styles`, so this rejects both a
    /// dangling reference and a reference to a paragraph/character style. Used
    /// by ECMA-376 Part 1 §17.4.37 logical-table membership: an explicit but
    /// unresolved or non-table style is not a valid grouping identity. This
    /// does not alter the omitted-style / default-table-style policy, which
    /// remains owned by `default_table_style_id`.
    pub fn contains_table_style(&self, style_id: &str) -> bool {
        self.table_styles.contains_key(style_id)
    }

    pub fn parse(xml: &str) -> Self {
        let doc = match parse_guarded(xml) {
            Ok(d) => d,
            Err(_) => return Self::empty(),
        };
        let root = doc.root_element();
        let mut styles: HashMap<String, StyleDef> = HashMap::new();
        let mut defaults_para = ParaFmt::default();
        let mut defaults_run = RunFmt::default();

        // Parse docDefaults
        if let Some(dd) = child_w(root, "docDefaults") {
            if let Some(rpr_def) = child_w(dd, "rPrDefault").and_then(|n| child_w(n, "rPr")) {
                defaults_run = parse_run_fmt(rpr_def);
            }
            if let Some(ppr_def) = child_w(dd, "pPrDefault").and_then(|n| child_w(n, "pPr")) {
                defaults_para = parse_para_fmt(ppr_def);
                // docDefaults is the implicit fallback; ECMA-376 §17.6.5 +
                // §17.3.1.33 imply that a paragraph whose line spacing is
                // only satisfied by docDefault (not declared on pPr or a
                // named style) should be treated as "no explicit line" —
                // in a docGrid section that yields 1 grid line per text
                // line rather than pitch × M.
                defaults_para.line_spacing_explicit = None;
            }
        }

        // Parse each style (paragraph, character, or table).
        // ECMA-376 §17.7.6 ST_StyleType: table styles may carry pPr that
        // applies to cell paragraphs (e.g. "Table Grid" sets
        // `w:spacing w:line="240" w:lineRule="auto" w:after="0"`). We
        // index them in the same StyleMap so cell resolution can look
        // them up by ID.
        let mut default_para_style_id: Option<String> = None;
        let mut normal_name_para_style_id: Option<String> = None;
        let mut normal_id_para_style_id: Option<String> = None;
        let mut default_table_style_id: Option<String> = None;
        let mut table_styles: HashMap<String, TableStyleDef> = HashMap::new();
        for style_node in children_w(root, "style") {
            let Some(style_id) = attr_w(style_node, "styleId") else {
                continue;
            };
            let style_type = attr_w(style_node, "type").unwrap_or_default();
            if style_type != "paragraph" && style_type != "character" && style_type != "table" {
                continue;
            }

            if style_type == "paragraph" && attr_w(style_node, "default").as_deref() == Some("1") {
                default_para_style_id = Some(style_id.clone());
            }
            if style_type == "paragraph" {
                let primary_name = child_w(style_node, "name")
                    .and_then(|node| attr_w(node, "val"))
                    .map(|name| name.trim().to_string());
                if primary_name.as_deref() == Some("Normal") && normal_name_para_style_id.is_none()
                {
                    normal_name_para_style_id = Some(style_id.clone());
                }
                if style_id == "Normal" && normal_id_para_style_id.is_none() {
                    normal_id_para_style_id = Some(style_id.clone());
                }
            }
            // §17.7.4: track `<w:style w:type="table" w:default="1">` so
            // tables that omit `<w:tblStyle>` can inherit its tblCellMar etc.
            if style_type == "table" && attr_w(style_node, "default").as_deref() == Some("1") {
                default_table_style_id = Some(style_id.clone());
            }

            let based_on = child_w(style_node, "basedOn").and_then(|n| attr_w(n, "val"));

            if style_type == "table" {
                table_styles.insert(
                    style_id.clone(),
                    parse_tbl_style_def(style_node, based_on.clone()),
                );
            }

            let para = if let Some(ppr) = child_w(style_node, "pPr") {
                parse_para_fmt(ppr)
            } else {
                ParaFmt::default()
            };

            let run = if let Some(rpr) = child_w(style_node, "rPr") {
                parse_run_fmt(rpr)
            } else {
                RunFmt::default()
            };

            styles.insert(
                style_id,
                StyleDef {
                    para,
                    run,
                    based_on,
                },
            );
        }

        StyleMap {
            styles,
            normal_para_style_id: normal_name_para_style_id.or(normal_id_para_style_id),
            table_styles,
            defaults_para,
            defaults_run,
            default_para_style_id,
            default_table_style_id,
        }
    }

    fn empty() -> Self {
        StyleMap {
            styles: HashMap::new(),
            normal_para_style_id: None,
            table_styles: HashMap::new(),
            defaults_para: ParaFmt::default(),
            defaults_run: RunFmt::default(),
            default_para_style_id: None,
            default_table_style_id: None,
        }
    }

    /// Resolve a table style by ID, flattening its basedOn chain (base first, then
    /// the derived style overrides). Returns defaults if the ID is unknown.
    pub fn resolve_table_style(&self, style_id: &str) -> TableStyleDef {
        let mut chain: Vec<&TableStyleDef> = Vec::new();
        let mut current = Some(style_id);
        let mut visited = HashSet::new();
        while let Some(current_id) = current {
            if !visited.insert(current_id) {
                break;
            }
            let Some(def) = self.table_styles.get(current_id) else {
                break;
            };
            chain.push(def);
            current = def.based_on.as_deref();
        }
        // Merge from base (end of chain) to derived (front).
        let mut out = TableStyleDef::default();
        for def in chain.into_iter().rev() {
            merge_raw_borders(&mut out.borders, &def.borders);
            if def.cell_shd.is_some() {
                out.cell_shd = def.cell_shd.clone();
            }
            if def.cell_valign.is_some() {
                out.cell_valign = def.cell_valign.clone();
            }
            // §17.7.6.5 / §17.7.6.7: band widths inherit through basedOn — a
            // derived style that omits them (None) keeps the base value.
            if def.row_band_size.is_some() {
                out.row_band_size = def.row_band_size;
            }
            if def.col_band_size.is_some() {
                out.col_band_size = def.col_band_size;
            }
            // §17.4.42: per-edge tblCellMar inherits through basedOn — a
            // derived style that omits an edge keeps the base value, and
            // an explicit edge value overrides. Each edge is independent
            // (Word writes `<w:left w:w="108"/>` without writing `<w:top>`
            // in TableNormal, and a derived style may override only
            // `<w:bottom>` without resetting the others).
            if def.cell_margin_top.is_some() {
                out.cell_margin_top = def.cell_margin_top;
            }
            if def.cell_margin_bottom.is_some() {
                out.cell_margin_bottom = def.cell_margin_bottom;
            }
            if def.cell_margin_left.is_some() {
                out.cell_margin_left = def.cell_margin_left;
            }
            if def.cell_margin_right.is_some() {
                out.cell_margin_right = def.cell_margin_right;
            }
            merge_table_margin_layer(&mut out.cell_margins, &def.cell_margins);
            if def.cell_spacing.is_some() {
                out.cell_spacing = def.cell_spacing.clone();
            }
            if def.tbl_header.is_some() {
                out.tbl_header = def.tbl_header;
            }
            if def.cant_split.is_some() {
                out.cant_split = def.cant_split;
            }
            // §17.7.6: a derived table style's whole-table rPr/pPr layers ON TOP
            // of the base style's. We fold each level into a single accumulated
            // RunFmt/ParaFmt with the standard merge semantics (later wins).
            if let Some(r) = &def.run {
                apply_run(out.run.get_or_insert_with(RunFmt::default), r);
            }
            if let Some(p) = &def.para {
                apply_para(out.para.get_or_insert_with(ParaFmt::default), p);
            }
            for (k, v) in &def.cond {
                let slot = out.cond.entry(k.clone()).or_default();
                if v.shd.is_some() {
                    slot.shd = v.shd.clone();
                }
                merge_raw_borders(&mut slot.borders, &v.borders);
                if let Some(r) = &v.run {
                    apply_run(slot.run.get_or_insert_with(RunFmt::default), r);
                }
                if let Some(p) = &v.para {
                    apply_para(slot.para.get_or_insert_with(ParaFmt::default), p);
                }
                if v.cell_spacing.is_some() {
                    slot.cell_spacing = v.cell_spacing.clone();
                }
                merge_table_margin_layer(&mut slot.cell_margins, &v.cell_margins);
                if v.tbl_header.is_some() {
                    slot.tbl_header = v.tbl_header;
                }
                if v.cant_split.is_some() {
                    slot.cant_split = v.cant_split;
                }
            }
        }
        out
    }

    /// Resolve all formatting for a paragraph style ID, merging inherited chain.
    /// Priority (lowest to highest): docDefaults → table style pPr (if inside a
    /// table) → basedOn chain of the paragraph style → paragraph style itself.
    /// Within each level: style rPr then pPr/rPr (both are paragraph-level run defaults).
    pub fn resolve_para(
        &self,
        style_id: Option<&str>,
        table_style_id: Option<&str>,
    ) -> (ParaFmt, RunFmt) {
        self.resolve_para_cond(style_id, table_style_id, None)
    }

    /// Resolve the paragraph style whose primary name is `Normal`.
    ///
    /// ECMA-376 §17.6.5 defines document-grid character pitch relative to the
    /// font size of the Normal style, which is not necessarily the paragraph
    /// style marked `w:default="1"`. For non-conforming documents that omit a
    /// Normal paragraph style, preserve the ordinary default-style fallback.
    /// A paragraph style whose ID is literally `Normal` is accepted only as a
    /// compatibility fallback for producers that omit its required name.
    pub fn resolve_normal_paragraph_style(&self) -> (ParaFmt, RunFmt) {
        if let Some(style_id) = self.normal_para_style_id.as_deref() {
            self.resolve_para(Some(style_id), None)
        } else {
            self.resolve_para(None, None)
        }
    }

    /// Resolve only the named/default paragraph-style chain, without document
    /// defaults or table-style layers. Callers that combine numbering paragraph
    /// properties need this provenance because §17.7.2 places style-origin
    /// numbering below the paragraph style, while direct `numPr` and its
    /// associated properties are applied in the final direct-formatting layer.
    pub(crate) fn resolve_paragraph_style_layer(&self, style_id: Option<&str>) -> ParaFmt {
        let mut paragraph = ParaFmt::default();
        let mut run = RunFmt::default();
        let effective_id = style_id
            .map(str::to_string)
            .or_else(|| self.default_para_style_id.clone());
        if let Some(id) = effective_id.as_deref() {
            self.apply_style_chain(id, &mut paragraph, &mut run);
        }
        paragraph
    }

    /// Like [`resolve_para`], but additionally layers a table style's resolved
    /// conditional formatting (`w:tblStylePr`'s `<w:rPr>`/`<w:pPr>`, §17.7.6)
    /// onto the cell's base formatting.
    ///
    /// ECMA-376 §17.7.2 style-application order (low→high): docDefaults <
    /// table-style (whole-table) < **table conditional** < numbering <
    /// paragraph style < character style < direct. So the conditional layer is
    /// applied AFTER the whole-table table-style chain but BEFORE the paragraph
    /// style — it is a BASE the paragraph/character styles and direct rPr (which
    /// the caller applies via `apply_direct_*` on top of the returned values)
    /// override, never the other way around. This is why a cell whose run
    /// carries a direct color still wins over a conditional row color.
    pub fn resolve_para_cond(
        &self,
        style_id: Option<&str>,
        table_style_id: Option<&str>,
        cond: Option<&CondFmt>,
    ) -> (ParaFmt, RunFmt) {
        let mut merged_para = ParaFmt::default();
        let mut merged_run = RunFmt::default();

        apply_para(&mut merged_para, &self.defaults_para);
        apply_run(&mut merged_run, &self.defaults_run);

        // Table style pPr applies to every paragraph inside the table, below
        // the paragraph style (§17.7.6). "Table Grid" sets line=240 after=0;
        // without this, cell paragraphs inherit docDefault's M=1.15 spacing
        // and render ~3pt taller per line than Word. This step also folds in the
        // table style's whole-table rPr (e.g. Calendar 3's gray day-number
        // color) since table styles are indexed in `self.styles`.
        if let Some(tid) = table_style_id {
            self.apply_style_chain(tid, &mut merged_para, &mut merged_run);
        }

        // §17.7.2: the row/band conditional formatting sits one layer above the
        // whole-table table style and below the paragraph style. Apply its pPr
        // then rPr with the standard merge semantics (set values override, unset
        // inherit) so e.g. firstRow's `<w:color w:val="365F91"/>` colors the
        // header runs unless a more specific layer overrides it.
        if let Some(c) = cond {
            if let Some(p) = &c.para {
                apply_para(&mut merged_para, p);
            }
            if let Some(r) = &c.run {
                apply_run(&mut merged_run, r);
            }
        }

        // Use explicit pStyle if present, otherwise fall back to the
        // paragraph style marked w:default="1" (typically "Normal").
        let effective_id = style_id
            .map(str::to_string)
            .or_else(|| self.default_para_style_id.clone());
        if let Some(id) = effective_id.as_deref() {
            self.apply_style_chain(id, &mut merged_para, &mut merged_run);
        }

        (merged_para, merged_run)
    }

    fn apply_style_chain(&self, id: &str, merged_para: &mut ParaFmt, merged_run: &mut RunFmt) {
        let mut chain = Vec::new();
        let mut visited = HashSet::new();
        let mut current = Some(id);
        while let Some(style_id) = current {
            if !visited.insert(style_id) {
                break;
            }
            let Some(def) = self.styles.get(style_id) else {
                break;
            };
            chain.push(def);
            current = def.based_on.as_deref();
        }

        // ECMA-376 Part 1 §17.7.4.3 defines the parent chain. MS-OI29500
        // §2.1.233 requires Word-authored chains not to contain circular
        // references. A malformed cycle is therefore terminated at the edge
        // that revisits a style; every unique style already found still
        // participates once, from the base-most entry to the requested style.
        for def in chain.into_iter().rev() {
            apply_para(merged_para, &def.para);
            apply_run(merged_run, &def.run);
            // pPr/rPr (paragraph mark run properties) also apply to runs
            apply_run(merged_run, &def.para.run);
        }
    }

    /// ECMA-376 §17.9.23 (`<w:lvl><w:pStyle>`) — resolve each style's
    /// paragraph-style ↔ numbering-level ASSOCIATION into that style's own
    /// `num_level`.
    ///
    /// Word's "Define Multilevel List → Link level to style" authoring stores
    /// the level on the NUMBERING side: the style's `numPr` carries only a
    /// `<w:numId>` (no `<w:ilvl>`), and the abstractNum's `<w:lvl>` names the
    /// style back via `<w:pStyle>` (sample-28's KPMGHeading1/2/3 →
    /// abstractNum 67 levels 0/1/2). §17.9.23: paragraphs of that style "shall
    /// automatically utilize this numbering level", and "any numbering level
    /// defined by the numPr element [of the style] shall be ignored" — so the
    /// association also overrides an explicit (redundant or stale) ilvl written
    /// in the style's own numPr.
    ///
    /// Rewriting the level AT THE STYLE LAYER (rather than per paragraph) keeps
    /// every downstream rule intact for free: `basedOn` children inherit the
    /// corrected level through the normal cascade (`apply_para`), and a
    /// paragraph's DIRECT `<w:ilvl>` (Word's Tab-demotion) still outranks it as
    /// direct-over-style formatting (§17.7.2).
    ///
    /// The lookup uses the style's EFFECTIVE numId (its own `numPr` numId or
    /// the nearest `basedOn` ancestor's, §17.7.4.17) so a derived style that
    /// only inherits its list still finds its own association. A style whose
    /// chain carries no numbering never fires — which also discharges
    /// §17.9.23's "not a paragraph style → can be ignored" clause, since only
    /// styles with (inherited) paragraph numbering are affected.
    ///
    /// Call once after BOTH `styles.xml` and `numbering.xml` are parsed.
    pub fn resolve_numbering_level_backlinks(&mut self, num_map: &crate::numbering::NumberingMap) {
        let ids: Vec<String> = self.styles.keys().cloned().collect();
        for id in ids {
            let Some(num_id) = self.effective_num_id(&id) else {
                continue;
            };
            if let Some(level) = num_map.level_for_style(num_id, &id) {
                if let Some(def) = self.styles.get_mut(&id) {
                    def.para.num_level = Some(level);
                }
            }
        }
    }

    /// The style's EFFECTIVE numId: its own `pPr/numPr/numId` or the nearest
    /// `basedOn` ancestor's (§17.7.4.17 style inheritance). Hop-capped so a
    /// malformed `basedOn` cycle terminates.
    fn effective_num_id(&self, id: &str) -> Option<u32> {
        let mut cur = id;
        let mut visited = HashSet::new();
        while visited.insert(cur) {
            let def = self.styles.get(cur)?;
            if def.para.num_id.is_some() {
                return def.para.num_id;
            }
            cur = def.based_on.as_deref()?;
        }
        None
    }

    /// Resolve a character style (rStyle, §17.3.2.29) chain WITHOUT prepending
    /// docDefaults or the default paragraph style. ECMA-376 §17.7.2 (Style
    /// Hierarchy) layers character styles ON TOP of the paragraph's
    /// already-resolved run formatting — pulling docDefaults in here would
    /// overwrite values the paragraph style legitimately set
    /// (e.g. Normal's Meiryo UI 9pt being clobbered by docDefault Calibri 11pt
    /// for a run that only says `<w:rStyle w:val="PlaceholderText"/>`).
    pub fn resolve_run_style(&self, style_id: &str) -> RunFmt {
        let mut merged_run = RunFmt::default();
        // Walk only the rStyle's basedOn chain. No docDefaults, no table
        // style, no default paragraph style — those are baseline contributions
        // already folded into the caller's `base_run`.
        let mut merged_para = ParaFmt::default();
        self.apply_style_chain(style_id, &mut merged_para, &mut merged_run);
        merged_run
    }
}

/// ECMA-376 §17.3.1.37 tab-stop inheritance — MERGE a more-specific set (`src`,
/// e.g. a paragraph's direct `<w:tabs>`) over an inherited one (`base`, e.g. the
/// resolved style-chain stops) by POSITION, not wholesale replace. Word treats
/// each `<w:tab>` as keyed on its `pos`: a more-specific stop at the same
/// position OVERRIDES the inherited one (its alignment/leader win), a
/// `val="clear"` stop REMOVES the inherited stop at that position, and stops at
/// positions the inherited set lacks are ADDED. This is what lets a TOC entry's
/// direct left tab coexist with the TOC style's right dot-/underscore-leader tab
/// (the leader + page-number column). Positions compare within a 1/20-pt epsilon
/// (twips round-trip). The result is sorted; `clear` markers are stripped (they
/// only delete). A wholesale replace here silently dropped the style's leader
/// tab whenever a paragraph added ANY direct tab (issue #820 TOC leaders).
pub(crate) fn merge_tab_stops(
    base: &[(f64, String, String)],
    src: &[(f64, String, String)],
) -> Vec<(f64, String, String)> {
    const EPS: f64 = 0.05; // 1/20 pt — the twips grid
    let mut out: Vec<(f64, String, String)> = base.to_vec();
    for s in src {
        // Remove any inherited stop at the same position (override or clear).
        out.retain(|b| (b.0 - s.0).abs() > EPS);
        if s.1 != "clear" {
            out.push(s.clone());
        }
    }
    out.retain(|t| t.1 != "clear");
    out.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal));
    out
}

/// Layer the paragraph format `src` OVER `dst`: every property `src` explicitly
/// sets replaces `dst`'s. Used both for the style cascade (basedOn parent →
/// child) and — via parser.rs — for a paragraph's DIRECT pPr over its resolved
/// style. Keep this the single source of truth for "which paragraph properties
/// override on merge"; the direct-formatting path reuses it so the two can never
/// drift (a drift previously dropped direct pBdr / shd / pageBreakBefore — see
/// `direct_paragraph_ppr_properties_survive_merge`).
pub(crate) fn apply_para(dst: &mut ParaFmt, src: &ParaFmt) {
    if src.alignment.is_some() {
        dst.alignment = src.alignment.clone();
    }
    if src.indent_left.is_some() {
        dst.indent_left = src.indent_left;
    }
    if src.indent_right.is_some() {
        dst.indent_right = src.indent_right;
    }
    if src.indent_first.is_some() {
        dst.indent_first = src.indent_first;
    }
    if src.space_before.is_some() {
        dst.space_before = src.space_before;
    }
    if src.space_after.is_some() {
        dst.space_after = src.space_after;
    }
    if src.line_spacing_val.is_some() {
        dst.line_spacing_val = src.line_spacing_val;
    }
    if src.line_spacing_rule.is_some() {
        dst.line_spacing_rule = src.line_spacing_rule.clone();
    }
    if src.line_spacing_explicit.is_some() {
        dst.line_spacing_explicit = src.line_spacing_explicit;
    }
    if src.num_id.is_some() {
        dst.num_id = src.num_id;
    }
    if src.num_level.is_some() {
        dst.num_level = src.num_level;
    }
    if let Some(src_tabs) = &src.tab_stops {
        // §17.3.1.37 — MERGE by position (see `merge_tab_stops`), not replace, so
        // a direct/child tab set keeps the inherited style's other stops (e.g. a
        // TOC leader tab). `dst` empty ⇒ merge is just the sorted src.
        let base = dst.tab_stops.take().unwrap_or_default();
        dst.tab_stops = Some(merge_tab_stops(&base, src_tabs));
    }
    if src.shading.is_some() {
        dst.shading = src.shading.clone();
    }
    if src.page_break_before.is_some() {
        dst.page_break_before = src.page_break_before;
    }
    if src.contextual_spacing.is_some() {
        dst.contextual_spacing = src.contextual_spacing;
    }
    if src.keep_next.is_some() {
        dst.keep_next = src.keep_next;
    }
    if src.outline_level.is_some() {
        dst.outline_level = src.outline_level;
    }
    if src.keep_lines.is_some() {
        dst.keep_lines = src.keep_lines;
    }
    if src.widow_control.is_some() {
        dst.widow_control = src.widow_control;
    }
    if src.overflow_punct.is_some() {
        dst.overflow_punct = src.overflow_punct;
    }
    if src.adjust_right_ind.is_some() {
        dst.adjust_right_ind = src.adjust_right_ind;
    }
    if let Some(src_b) = &src.para_borders {
        // Each pBdr EDGE inherits INDEPENDENTLY across the style hierarchy — bottom
        // §17.3.1.7, left §17.3.1.17, right §17.3.1.28, top §17.3.1.42, between
        // §17.3.1.5 each carry the same clause: "if this element is
        // omitted ... its value is determined by the setting previously set at any
        // level of the style hierarchy". So merge edge-by-edge over the inherited
        // box rather than replacing it wholesale: a level that names only
        // `<w:bottom>` keeps the inherited top/left/right/between. (Contrast
        // `frame_pr` below, §17.3.1.11 — one grouped element, replaced wholesale.)
        // An EXPLICIT `<w:bottom w:val="nil"/>` is a present "cleared" edge (style
        // "none", emitted by parse_edge — distinct from an OMITTED edge which is
        // `None` and inherits), so it overrides here and removes the inherited edge.
        let dst_b = dst.para_borders.get_or_insert_with(Default::default);
        if src_b.top.is_some() {
            dst_b.top = src_b.top.clone();
        }
        if src_b.bottom.is_some() {
            dst_b.bottom = src_b.bottom.clone();
        }
        if src_b.left.is_some() {
            dst_b.left = src_b.left.clone();
        }
        if src_b.right.is_some() {
            dst_b.right = src_b.right.clone();
        }
        if src_b.between.is_some() {
            dst_b.between = src_b.between.clone();
        }
    }
    if let Some(src_typography) = &src.paragraph_typography {
        let dst_typography = dst
            .paragraph_typography
            .get_or_insert_with(Default::default);
        let src_borders = &src_typography.borders;
        let dst_borders = &mut dst_typography.borders;
        if src_borders.top.is_some() {
            dst_borders.top = src_borders.top.clone();
        }
        if src_borders.right.is_some() {
            dst_borders.right = src_borders.right.clone();
        }
        if src_borders.bottom.is_some() {
            dst_borders.bottom = src_borders.bottom.clone();
        }
        if src_borders.left.is_some() {
            dst_borders.left = src_borders.left.clone();
        }
        if src_borders.between.is_some() {
            dst_borders.between = src_borders.between.clone();
        }
        if src_borders.bar.is_some() {
            dst_borders.bar = src_borders.bar.clone();
        }
    }
    if src.bidi.is_some() {
        dst.bidi = src.bidi;
    }
    if src.snap_to_grid.is_some() {
        dst.snap_to_grid = src.snap_to_grid;
    }
    // §17.3.1.11: framePr is a single grouped element — a later level that
    // specifies it replaces the whole frame definition (Word does not merge
    // individual frame attributes across the style chain).
    if src.frame_pr.is_some() {
        dst.frame_pr = src.frame_pr.clone();
    }
}

/// Layer the run format `src` OVER `dst`: every property `src` explicitly sets
/// replaces `dst`'s. The single source of truth for "which run properties
/// override on merge" — used by the style cascade (docDefaults → style chain →
/// rStyle) AND, via `parser::apply_direct_run`, for a run's DIRECT rPr. Keep new
/// `RunFmt` fields here only; the direct path reuses this so the two can never
/// drift. The trailing `*_set_here` markers are OR-propagated (see below) so a
/// folded rStyle sub-chain still mirrors `w:sz`→szCs when re-applied; the direct
/// path resets them afterwards because it is the terminal merge.
pub(crate) fn apply_run(dst: &mut RunFmt, src: &RunFmt) {
    if src.bold.is_some() {
        dst.bold = src.bold;
    }
    if src.italic.is_some() {
        dst.italic = src.italic;
    }
    if src.underline.is_some() {
        dst.underline = src.underline;
    }
    // §17.3.2.40 underline style / colour merge with the same set-value-wins rule
    // as `highlight` (an Option<String>): a level that names a style/colour wins,
    // absence inherits. A `<w:u w:val="none"/>` clears `underline` (drawn off) but
    // leaves a stale inherited style/colour — harmless because the renderer never
    // draws an underline whose `underline` bool is false.
    if src.underline_style.is_some() {
        dst.underline_style = src.underline_style.clone();
    }
    if src.underline_color.is_some() {
        dst.underline_color = src.underline_color.clone();
    }
    if let Some(src_u) = &src.underline_typography {
        let dst_u = dst
            .underline_typography
            .get_or_insert_with(Default::default);
        // [MS-OI29500] §2.1.100: omitted underline attributes continue through
        // the style hierarchy. Invalid authored values still override so the
        // retained layer can emit a diagnostic instead of silently inheriting.
        let merge = |dst: &mut crate::types::TypographyValueWire<String>,
                     src: &crate::types::TypographyValueWire<String>| {
            if src.status != crate::types::TypographyValueStatusWire::Missing {
                *dst = src.clone();
            }
        };
        merge(&mut dst_u.val, &src_u.val);
        merge(&mut dst_u.color, &src_u.color);
        merge(&mut dst_u.theme_color, &src_u.theme_color);
        merge(&mut dst_u.theme_tint, &src_u.theme_tint);
        merge(&mut dst_u.theme_shade, &src_u.theme_shade);
    }
    if src.strikethrough.is_some() {
        dst.strikethrough = src.strikethrough;
    }
    if src.font_size.is_some() {
        dst.font_size = src.font_size;
    }
    if src.color_auto {
        // §17.3.2.6: explicit auto breaks inheritance (an inherited style color
        // such as PlaceholderText gray must not survive) and defers the final
        // color to background-contrast resolution at render time (an
        // implementation-defined black/white pick; no normative algorithm).
        dst.color = None;
        dst.color_auto = true;
    } else if src.color.is_some() {
        dst.color = src.color.clone();
        dst.color_auto = false;
    }
    if src.font_family_ascii.is_some() {
        dst.font_family_ascii = src.font_family_ascii.clone();
    }
    if src.font_family_ascii_direct.is_some() || src.font_family_ascii_theme.is_some() {
        dst.font_family_ascii_direct = src.font_family_ascii_direct.clone();
        dst.font_family_ascii_theme = src.font_family_ascii_theme.clone();
    }
    if src.font_family_high_ansi.is_some() {
        dst.font_family_high_ansi = src.font_family_high_ansi.clone();
    }
    if src.font_family_high_ansi_direct.is_some() || src.font_family_high_ansi_theme.is_some() {
        dst.font_family_high_ansi_direct = src.font_family_high_ansi_direct.clone();
        dst.font_family_high_ansi_theme = src.font_family_high_ansi_theme.clone();
    }
    if src.font_family_east_asia.is_some() {
        dst.font_family_east_asia = src.font_family_east_asia.clone();
    }
    if src.font_family_east_asia_direct.is_some() || src.font_family_east_asia_theme.is_some() {
        dst.font_family_east_asia_direct = src.font_family_east_asia_direct.clone();
        dst.font_family_east_asia_theme = src.font_family_east_asia_theme.clone();
    }
    if src.font_hint.is_some() {
        dst.font_hint = src.font_hint.clone();
    }
    if src.background.is_some() {
        dst.background = src.background.clone();
    }
    if src.vert_align.is_some() {
        dst.vert_align = src.vert_align.clone();
    }
    if src.vertical_align_typography.is_some() {
        dst.vertical_align_typography = src.vertical_align_typography.clone();
    }
    if src.all_caps.is_some() {
        dst.all_caps = src.all_caps;
    }
    if src.small_caps.is_some() {
        dst.small_caps = src.small_caps;
    }
    if src.dstrike.is_some() {
        dst.dstrike = src.dstrike;
    }
    if src.vanish.is_some() {
        dst.vanish = src.vanish;
    }
    if src.web_hidden.is_some() {
        dst.web_hidden = src.web_hidden;
    }
    if src.highlight.is_some() {
        dst.highlight = src.highlight.clone();
    }
    // §17.3.2.12 w:em — value property, set-wins (same shape as highlight).
    if src.emphasis_mark.is_some() {
        dst.emphasis_mark = src.emphasis_mark.clone();
    }
    if src.emphasis_typography.is_some() {
        dst.emphasis_typography = src.emphasis_typography.clone();
    }
    if src.border.is_some() {
        dst.border = src.border.clone();
    }
    if src.border_typography.is_some() {
        dst.border_typography = src.border_typography.clone();
    }
    if src.rtl.is_some() {
        dst.rtl = src.rtl;
    }
    if src.snap_to_grid.is_some() {
        dst.snap_to_grid = src.snap_to_grid;
    }
    // §17.3.2.7 <w:cs/> — complex-script run toggle. Must mirror
    // apply_direct_run; without this arm a style-chain <w:cs/> (set in a
    // paragraph/character style rPr by parse_run_fmt) is silently dropped.
    if src.cs_toggle.is_some() {
        dst.cs_toggle = src.cs_toggle;
    }
    if src.font_family_cs.is_some() {
        dst.font_family_cs = src.font_family_cs.clone();
    }
    if src.font_family_cs_direct.is_some() || src.font_family_cs_theme.is_some() {
        dst.font_family_cs_direct = src.font_family_cs_direct.clone();
        dst.font_family_cs_theme = src.font_family_cs_theme.clone();
    }
    if src.bold_cs.is_some() {
        dst.bold_cs = src.bold_cs;
    }
    if src.italic_cs.is_some() {
        dst.italic_cs = src.italic_cs;
    }
    if src.lang_bidi.is_some() {
        dst.lang_bidi = src.lang_bidi.clone();
    }
    if src.lang_east_asia.is_some() {
        dst.lang_east_asia = src.lang_east_asia.clone();
    }
    // Run character-metric axes (§17.3.2.14 fitText / §17.3.2.35 spacing /
    // §17.3.2.43 w / §17.3.2.24 position / §17.3.2.19 kern). Each carries the
    // same "if omitted, inherit; if set, override" rule as the other run
    // properties, so a level that names the attribute wins and absence inherits.
    if src.char_spacing.is_some() {
        dst.char_spacing = src.char_spacing;
    }
    if src.fit_text.is_some() {
        dst.fit_text.clone_from(&src.fit_text);
    }
    if src.char_scale.is_some() {
        dst.char_scale = src.char_scale;
    }
    if src.position.is_some() {
        dst.position = src.position;
    }
    if src.position_typography.is_some() {
        dst.position_typography = src.position_typography.clone();
    }
    if src.kerning.is_some() {
        dst.kerning = src.kerning;
    }
    // East Asian typography (§17.3.2.10 eastAsianLayout) — each attribute is a
    // set-wins toggle/value like the run properties above: a level that names it
    // overrides, absence inherits.
    if src.east_asian_vert.is_some() {
        dst.east_asian_vert = src.east_asian_vert;
    }
    if src.east_asian_vert_compress.is_some() {
        dst.east_asian_vert_compress = src.east_asian_vert_compress;
    }
    if src.east_asian_combine.is_some() {
        dst.east_asian_combine = src.east_asian_combine;
    }
    if src.east_asian_combine_brackets.is_some() {
        dst.east_asian_combine_brackets = src.east_asian_combine_brackets.clone();
    }
    if src.combine_brackets_typography.is_some() {
        dst.combine_brackets_typography = src.combine_brackets_typography.clone();
    }

    // Complex-script font size resolution (ECMA-376 §17.3.2.18). Word treats a
    // directly-applied `w:sz` as also setting the complex-script size UNLESS the
    // same level supplies its own `w:szCs`: a run/style that sets only `w:sz`
    // renders its Arabic/Hebrew text at that size too (e.g. sample-7's underlined
    // title sets sz=36 with no szCs and Word draws the Arabic at 18pt, not at the
    // inherited docDefaults szCs=11pt). So at a level that sets sz-without-szCs,
    // the Latin size shadows the inherited cs size; an explicit szCs always wins.
    if src.font_size_cs_set_here {
        dst.font_size_cs = src.font_size_cs;
    } else if src.font_size_set_here {
        dst.font_size_cs = src.font_size;
    } else if src.font_size_cs.is_some() {
        // Inherited-only szCs (no sz/szCs literally on this level): carry it.
        dst.font_size_cs = src.font_size_cs;
    }
    // OR-propagate the "set here" markers. `apply_run` may be used to fold a
    // whole (rStyle) sub-chain into a single RunFmt that is later re-applied via
    // `apply_direct_run`; carrying these flags forward lets that later merge see
    // that the sub-chain set sz/szCs and apply the same mirroring rule.
    dst.font_size_set_here |= src.font_size_set_here;
    dst.font_size_cs_set_here |= src.font_size_cs_set_here;
}

pub fn parse_para_fmt(ppr: roxmltree::Node) -> ParaFmt {
    let mut fmt = ParaFmt::default();

    // Alignment
    if let Some(jc) = child_w(ppr, "jc") {
        fmt.alignment = attr_w(jc, "val");
    }

    // Spacing
    if let Some(sp) = child_w(ppr, "spacing") {
        if let Some(v) = attr_w(sp, "before") {
            fmt.space_before = Some(twips_to_pt(&v));
        }
        if let Some(v) = attr_w(sp, "after") {
            fmt.space_after = Some(twips_to_pt(&v));
        }
        if let Some(v) = attr_w(sp, "line") {
            let rule = attr_w(sp, "lineRule").unwrap_or_else(|| "auto".to_string());
            let raw: f64 = v.parse().unwrap_or(240.0);
            // OOXML encodes line spacing as:
            //   auto      → raw / 240   = multiplier (1.0 = single, 1.5 = 1½, 2.0 = double)
            //   atLeast   → raw / 20    = pt (minimum line height)
            //   exact     → raw / 20    = pt (exact line height)
            // Previously we reinterpreted auto > 720 (3× single) as atLeast-pt
            // to tame decorative-title overruns, but that was an empirical
            // work-around. ECMA-376 §17.6.5 w:docGrid (handled at render time)
            // already constrains large auto multipliers to a grid pitch when
            // the section enables a line grid, which is where those oversized
            // values are actually authored.
            let (val, effective_rule) = match rule.as_str() {
                "exact" => (raw / 20.0, "exact".to_string()),
                "atLeast" => (raw / 20.0, "atLeast".to_string()),
                _ => (raw / 240.0, "auto".to_string()),
            };
            fmt.line_spacing_val = Some(val);
            fmt.line_spacing_rule = Some(effective_rule);
            fmt.line_spacing_explicit = Some(true);
        }
    }

    // Indentation. ECMA-376 §17.3.1.12 allows both the older "left"/"right"
    // attributes and the logical "start"/"end" aliases. In LTR docs these are
    // identical; use either if present, with start/end taking precedence when
    // both appear (logical wins for bidi correctness).
    if let Some(ind) = child_w(ppr, "ind") {
        if let Some(v) = attr_w(ind, "left") {
            fmt.indent_left = Some(twips_to_pt(&v));
        }
        if let Some(v) = attr_w(ind, "start") {
            fmt.indent_left = Some(twips_to_pt(&v));
        }
        if let Some(v) = attr_w(ind, "right") {
            fmt.indent_right = Some(twips_to_pt(&v));
        }
        if let Some(v) = attr_w(ind, "end") {
            fmt.indent_right = Some(twips_to_pt(&v));
        }
        if let Some(v) = attr_w(ind, "firstLine") {
            fmt.indent_first = Some(twips_to_pt(&v));
        }
        // hanging overrides firstLine per §17.3.1.12 when both are present.
        if let Some(v) = attr_w(ind, "hanging") {
            fmt.indent_first = Some(-twips_to_pt(&v));
        }
    }

    // Numbering
    if let Some(pnpr) = child_w(ppr, "numPr") {
        // ilvl defaults to 0 when absent
        fmt.num_level = Some(match child_w(pnpr, "ilvl") {
            None => 0,
            // Preserve malformed/out-of-representation authored values as an
            // unsupported level for validation at consumption, not as an
            // absent level that silently selects zero. No raw string is kept.
            Some(level) => attr_w(level, "val")
                .and_then(|value| value.trim().parse().ok())
                .unwrap_or(u32::MAX),
        });
        if let Some(nid) = child_w(pnpr, "numId") {
            fmt.num_id = attr_w(nid, "val").and_then(|v| v.parse().ok());
        }
    }

    // Explicit tab stops (pPr/tabs/tab)
    if let Some(tabs_node) = child_w(ppr, "tabs") {
        let mut tabs: Vec<(f64, String, String)> = Vec::new();
        for t in children_w(tabs_node, "tab") {
            let val = attr_w(t, "val").unwrap_or_else(|| "left".to_string());
            let pos = match attr_w(t, "pos").map(|s| twips_to_pt(&s)) {
                Some(p) => p,
                None => continue,
            };
            let leader = attr_w(t, "leader").unwrap_or_else(|| "none".to_string());
            // `val="clear"` (§17.3.1.37 / §17.18.84) is preserved through the
            // style-chain merge (`apply_para` → `merge_tab_stops`), where it
            // REMOVES an inherited stop at this position; `merge_tab_stops` drops any
            // remaining `clear` (and the final `TabStop` conversion never emits
            // one). Keeping it here — rather than dropping at parse — is what
            // lets a direct `<w:tab val="clear" pos="X">` cancel a style stop
            // at X while other direct stops merge with the inherited set.
            tabs.push((pos, val, leader));
        }
        if !tabs.is_empty() {
            tabs.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal));
            fmt.tab_stops = Some(tabs);
        }
    }

    // pPr/rPr (run defaults within paragraph)
    if let Some(rpr) = child_w(ppr, "rPr") {
        fmt.run = parse_run_fmt(rpr);
    }

    // Paragraph shading
    if let Some(shd) = child_w(ppr, "shd") {
        if let Some(fill) = attr_w(shd, "fill") {
            if fill != "auto" && fill.len() == 6 {
                fmt.shading = Some(fill.to_lowercase());
            }
        }
    }

    // Page break before paragraph
    fmt.page_break_before = bool_prop(ppr, "pageBreakBefore");

    // Contextual spacing
    fmt.contextual_spacing = bool_prop(ppr, "contextualSpacing");

    // keepNext — keep this paragraph on the same page as the next one
    fmt.keep_next = bool_prop(ppr, "keepNext");

    // keepLines — do not split this paragraph's lines across pages
    fmt.keep_lines = bool_prop(ppr, "keepLines");

    // widowControl — avoid leaving a single line at page top/bottom
    // (ECMA-376 default: true; explicit value=0 disables).
    fmt.widow_control = bool_prop(ppr, "widowControl");

    // overflowPunct — allow one punctuation character past paragraph extents.
    // ECMA-376 §17.3.1.21 defines omission as true; retain Option here so an
    // explicit style/direct false participates correctly in the cascade.
    fmt.overflow_punct = bool_prop(ppr, "overflowPunct");

    // adjustRightInd — ECMA-376 §17.3.1.1. The setting participates in the
    // paragraph-style hierarchy and omission ultimately defaults to true.
    // Geometry is intentionally resolved later because the adjustment depends
    // on the active section document grid and paragraph placement width.
    fmt.adjust_right_ind = bool_prop(ppr, "adjustRightInd");

    // bidi — right-to-left paragraph (ECMA-376 §17.3.1.6). On-off toggle:
    // present (or w:val="1"/"true") = RTL, w:val="0"/"false" = LTR. Carried to
    // the model as the paragraph base direction; the renderer feeds it to the
    // UAX#9 pass and to start/end alignment-edge resolution.
    fmt.bidi = bool_prop(ppr, "bidi");

    // snapToGrid — ECMA-376 §17.3.1.32. On-off toggle (default on). When off,
    // the paragraph ignores the section's docGrid line pitch and uses natural
    // line metrics. Carried to the model so the renderer can skip grid snapping
    // for this paragraph.
    fmt.snap_to_grid = bool_prop(ppr, "snapToGrid");

    // outlineLvl — 0..8 marks this paragraph (or its style) as a heading.
    // ECMA-376 §17.3.1.20 lists only 0–8 and "no level" (absent). Word
    // attaches an implicit keepNext to heading paragraphs (Heading 1–9
    // styles) even when the style XML omits it, which we replicate at
    // the final paragraph build step.
    if let Some(lvl) = child_w(ppr, "outlineLvl") {
        if let Some(v) = attr_w(lvl, "val") {
            if let Ok(n) = v.parse::<u32>() {
                if n <= 8 {
                    fmt.outline_level = Some(n);
                }
            }
        }
    }

    // Paragraph borders (pBdr)
    if let Some(pbdr) = child_w(ppr, "pBdr") {
        use crate::types::{ParaBorderEdge, ParagraphBorders};
        let parse_edge = |name: &str| -> Option<ParaBorderEdge> {
            // An OMITTED edge element → None (it inherits from the style hierarchy,
            // §17.3.1.7). An edge element that IS present but specifies `val="nil"`
            // or `val="none"` is NOT omitted: it explicitly says "no border", so it
            // must CLEAR an inherited edge rather than inherit it. Keep it as a
            // present "cleared" edge (style normalized to "none" — nil/none are
            // synonyms) so the per-edge merge in `apply_para` overrides the inherited
            // edge; the renderer treats a "none" edge as no-paint and as equivalent
            // to an absent edge for border-box matching.
            let node = child_w(pbdr, name)?;
            let style = attr_w(node, "val").unwrap_or_else(|| "none".to_string());
            if style == "none" || style == "nil" {
                return Some(ParaBorderEdge {
                    style: "none".to_string(),
                    color: None,
                    width: 0.0,
                    space: 0.0,
                });
            }
            // §17.3.4 CT_Border: sz is in eighths of a point; space in points.
            // Neither has a normative default when the attribute is absent (it is
            // optional with no spec fallback), so the unwrap_or values below are
            // arbitrary fallbacks for the near-nonexistent "val set, sz/space
            // omitted" authoring; the renderer compares width/space when matching
            // adjacent borders, so they only need to be internally consistent.
            let width = attr_w(node, "sz")
                .and_then(|s| s.parse::<f64>().ok())
                .map(|v| v / 8.0)
                .unwrap_or(0.5);
            let space = attr_w(node, "space")
                .and_then(|s| s.parse::<f64>().ok())
                .unwrap_or(1.0);
            let color = attr_w(node, "color")
                .filter(|c| c != "auto")
                .map(|c| c.to_lowercase());
            Some(ParaBorderEdge {
                style,
                color,
                width,
                space,
            })
        };
        let borders = ParagraphBorders {
            top: parse_edge("top"),
            bottom: parse_edge("bottom"),
            left: parse_edge("left"),
            right: parse_edge("right"),
            between: parse_edge("between"),
        };
        let typography_borders = crate::types::ParagraphTypographyBordersWire {
            top: child_w(pbdr, "top").map(parse_ct_border_typography),
            right: child_w(pbdr, "right").map(parse_ct_border_typography),
            bottom: child_w(pbdr, "bottom").map(parse_ct_border_typography),
            left: child_w(pbdr, "left").map(parse_ct_border_typography),
            between: child_w(pbdr, "between").map(parse_ct_border_typography),
            // CT_PBdr includes `bar` even though the stable public ParagraphBorders
            // predates it. Keep it private rather than widening that API.
            bar: child_w(pbdr, "bar").map(parse_ct_border_typography),
        };
        // §17.3.1.7: a `between` border (the rule drawn between two adjacent
        // paragraphs sharing an identical border set) is a first-class edge — a
        // paragraph may declare ONLY `between` (internal rules, no outer box), so
        // it must keep the borders alive on its own. The renderer already consumes
        // `between` (the suppressed-top join), so omitting it here would starve a
        // valid border set.
        if borders.top.is_some()
            || borders.bottom.is_some()
            || borders.left.is_some()
            || borders.right.is_some()
            || borders.between.is_some()
        {
            fmt.para_borders = Some(borders);
        }
        if typography_borders.top.is_some()
            || typography_borders.right.is_some()
            || typography_borders.bottom.is_some()
            || typography_borders.left.is_some()
            || typography_borders.between.is_some()
            || typography_borders.bar.is_some()
        {
            fmt.paragraph_typography = Some(crate::types::ParagraphTypographyWire {
                borders: typography_borders,
            });
        }
    }

    // Text frame / drop cap (ECMA-376 §17.3.1.11 w:framePr). Presence makes the
    // paragraph part of a text frame; all attributes are optional with the
    // spec-defined defaults applied here so the renderer never re-derives them.
    if let Some(fp) = child_w(ppr, "framePr") {
        use crate::types::FramePr;
        let twips = |name: &str| attr_w(fp, name).map(|s| twips_to_pt(&s));
        fmt.frame_pr = Some(Box::new(FramePr {
            // CT_FramePr identity includes anchorLock. Retain it on the private
            // wire without widening the stable public FramePr projection.
            anchor_lock: on_off_attr(fp, "anchorLock").unwrap_or(false),
            // ST_DropCap default "none" (§17.18.20).
            drop_cap: attr_w(fp, "dropCap").unwrap_or_else(|| "none".to_string()),
            // §17.3.1.11 lines default 1.
            lines: attr_w(fp, "lines")
                .and_then(|v| v.parse().ok())
                .unwrap_or(1),
            // ST_Wrap default "around" (§17.3.1.11 / §17.18.104).
            wrap: attr_w(fp, "wrap").unwrap_or_else(|| "around".to_string()),
            // ST_HAnchor / ST_VAnchor default "page" (§17.3.1.11).
            h_anchor: attr_w(fp, "hAnchor").unwrap_or_else(|| "page".to_string()),
            v_anchor: attr_w(fp, "vAnchor").unwrap_or_else(|| "page".to_string()),
            // ST_HeightRule default "auto" (§17.18.37).
            h_rule: attr_w(fp, "hRule").unwrap_or_else(|| "auto".to_string()),
            // hSpace / vSpace default 0 (§17.3.1.11).
            h_space: twips("hSpace").unwrap_or(0.0),
            v_space: twips("vSpace").unwrap_or(0.0),
            // w/h/x/y are kept Option so the renderer can tell "auto" (absent)
            // from an explicit 0; §17.3.1.11 assumes 0 when consumed.
            w: twips("w"),
            h: twips("h"),
            x: twips("x"),
            y: twips("y"),
            // ST_XAlign / ST_YAlign — supersede x/y when present (§17.3.1.11).
            x_align: attr_w(fp, "xAlign"),
            y_align: attr_w(fp, "yAlign"),
        }));
    }

    fmt
}

fn typography_value<T>(
    raw: Option<String>,
    value: Option<T>,
) -> crate::types::TypographyValueWire<T> {
    use crate::types::{TypographyValueStatusWire, TypographyValueWire};
    let status = match (&raw, value.is_some()) {
        (None, _) => TypographyValueStatusWire::Missing,
        (Some(_), true) => TypographyValueStatusWire::Valid,
        (Some(_), false) => TypographyValueStatusWire::Invalid,
    };
    TypographyValueWire { status, raw, value }
}

fn normalized_string_attr(
    node: roxmltree::Node,
    name: &str,
    normalize: impl FnOnce(&str) -> Option<String>,
) -> crate::types::TypographyValueWire<String> {
    let raw = attr_w(node, name);
    let value = raw.as_deref().and_then(normalize);
    typography_value(raw, value)
}

fn on_off_value(node: roxmltree::Node, name: &str) -> crate::types::TypographyValueWire<bool> {
    let raw = attr_w(node, name);
    let value = raw.as_deref().and_then(|value| match value {
        "1" | "true" | "on" => Some(true),
        "0" | "false" | "off" => Some(false),
        _ => None,
    });
    typography_value(raw, value)
}

fn is_st_border_value(value: &str) -> bool {
    // ST_Border is intentionally validated at acquisition time. Keeping an
    // unknown lexical value as a valid border would erase the distinction
    // between conforming OOXML and a producer extension before layout can
    // diagnose it (ECMA-376-1 Strict, §17.18.2).
    const VALUES: &[&str] = &[
        "nil",
        "none",
        "single",
        "thick",
        "double",
        "dotted",
        "dashed",
        "dotDash",
        "dotDotDash",
        "triple",
        "thinThickSmallGap",
        "thickThinSmallGap",
        "thinThickThinSmallGap",
        "thinThickMediumGap",
        "thickThinMediumGap",
        "thinThickThinMediumGap",
        "thinThickLargeGap",
        "thickThinLargeGap",
        "thinThickThinLargeGap",
        "wave",
        "doubleWave",
        "dashSmallGap",
        "dashDotStroked",
        "threeDEmboss",
        "threeDEngrave",
        "outset",
        "inset",
        "apples",
        "archedScallops",
        "babyPacifier",
        "babyRattle",
        "balloons3Colors",
        "balloonsHotAir",
        "basicBlackDashes",
        "basicBlackDots",
        "basicBlackSquares",
        "basicThinLines",
        "basicWhiteDashes",
        "basicWhiteDots",
        "basicWhiteSquares",
        "basicWideInline",
        "basicWideMidline",
        "basicWideOutline",
        "bats",
        "birds",
        "birdsFlight",
        "cabins",
        "cakeSlice",
        "candyCorn",
        "celticKnotwork",
        "certificateBanner",
        "chainLink",
        "champagneBottle",
        "checkedBarBlack",
        "checkedBarColor",
        "checkered",
        "christmasTree",
        "circlesLines",
        "circlesRectangles",
        "classicalWave",
        "clocks",
        "compass",
        "confetti",
        "confettiGrays",
        "confettiOutline",
        "confettiStreamers",
        "confettiWhite",
        "cornerTriangles",
        "couponCutoutDashes",
        "couponCutoutDots",
        "crazyMaze",
        "creaturesButterfly",
        "creaturesFish",
        "creaturesInsects",
        "creaturesLadyBug",
        "crossStitch",
        "cup",
        "decoArch",
        "decoArchColor",
        "decoBlocks",
        "diamondsGray",
        "doubleD",
        "doubleDiamonds",
        "earth1",
        "earth2",
        "earth3",
        "eclipsingSquares1",
        "eclipsingSquares2",
        "eggsBlack",
        "fans",
        "film",
        "firecrackers",
        "flowersBlockPrint",
        "flowersDaisies",
        "flowersModern1",
        "flowersModern2",
        "flowersPansy",
        "flowersRedRose",
        "flowersRoses",
        "flowersTeacup",
        "flowersTiny",
        "gems",
        "gingerbreadMan",
        "gradient",
        "handmade1",
        "handmade2",
        "heartBalloon",
        "heartGray",
        "hearts",
        "heebieJeebies",
        "holly",
        "houseFunky",
        "hypnotic",
        "iceCreamCones",
        "lightBulb",
        "lightning1",
        "lightning2",
        "mapPins",
        "mapleLeaf",
        "mapleMuffins",
        "marquee",
        "marqueeToothed",
        "moons",
        "mosaic",
        "musicNotes",
        "northwest",
        "ovals",
        "packages",
        "palmsBlack",
        "palmsColor",
        "paperClips",
        "papyrus",
        "partyFavor",
        "partyGlass",
        "pencils",
        "people",
        "peopleWaving",
        "peopleHats",
        "poinsettias",
        "postageStamp",
        "pumpkin1",
        "pushPinNote2",
        "pushPinNote1",
        "pyramids",
        "pyramidsAbove",
        "quadrants",
        "rings",
        "safari",
        "sawtooth",
        "sawtoothGray",
        "scaredCat",
        "seattle",
        "shadowedSquares",
        "sharksTeeth",
        "shorebirdTracks",
        "skyrocket",
        "snowflakeFancy",
        "snowflakes",
        "sombrero",
        "southwest",
        "stars",
        "starsTop",
        "stars3d",
        "starsBlack",
        "starsShadowed",
        "sun",
        "swirligig",
        "tornPaper",
        "tornPaperBlack",
        "trees",
        "triangleParty",
        "triangles",
        "triangle1",
        "triangle2",
        "triangleCircle1",
        "triangleCircle2",
        "shapes1",
        "shapes2",
        "twistedLines1",
        "twistedLines2",
        "vine",
        "waveline",
        "weavingAngles",
        "weavingBraid",
        "weavingRibbon",
        "weavingStrips",
        "whiteFlowers",
        "woodwork",
        "xIllusions",
        "zanyTriangles",
        "zigZag",
        "zigZagStitch",
        "custom",
    ];
    VALUES.contains(&value)
}

fn is_st_theme_color_value(value: &str) -> bool {
    matches!(
        value,
        "dark1"
            | "light1"
            | "dark2"
            | "light2"
            | "accent1"
            | "accent2"
            | "accent3"
            | "accent4"
            | "accent5"
            | "accent6"
            | "hyperlink"
            | "followedHyperlink"
            | "none"
            | "background1"
            | "text1"
            | "background2"
            | "text2"
    )
}

fn parse_ct_border_typography(node: roxmltree::Node) -> crate::types::CtBorderTypographyWire {
    use crate::types::CtBorderTypographyWire;
    let lower = |value: &str| (!value.trim().is_empty()).then(|| value.to_lowercase());
    let hex_byte = |value: &str| {
        (value.len() == 2 && value.chars().all(|c| c.is_ascii_hexdigit()))
            .then(|| value.to_ascii_uppercase())
    };
    let number = |name: &str, divisor: f64| {
        let raw = attr_w(node, name);
        let value = raw
            .as_deref()
            .and_then(|value| value.parse::<f64>().ok())
            .filter(|value| value.is_finite() && *value >= 0.0)
            .map(|value| value / divisor);
        typography_value(raw, value)
    };
    CtBorderTypographyWire {
        // CT_Border/@val is required (§17.3.4). Keep a missing/empty lexical
        // value as such; the stable legacy border can still use its old fallback.
        val: normalized_string_attr(node, "val", |value| {
            is_st_border_value(value).then(|| value.to_string())
        }),
        color: normalized_string_attr(node, "color", lower),
        theme_color: normalized_string_attr(node, "themeColor", |value| {
            is_st_theme_color_value(value).then(|| value.to_string())
        }),
        theme_tint: normalized_string_attr(node, "themeTint", hex_byte),
        theme_shade: normalized_string_attr(node, "themeShade", hex_byte),
        size_pt: number("sz", 8.0),
        space_pt: number("space", 1.0),
        shadow: on_off_value(node, "shadow"),
        frame: on_off_value(node, "frame"),
    }
}

fn is_underline_value(value: &str) -> bool {
    matches!(
        value,
        "none"
            | "words"
            | "single"
            | "thick"
            | "double"
            | "dotted"
            | "dottedHeavy"
            | "dash"
            | "dashedHeavy"
            | "dashLong"
            | "dashLongHeavy"
            | "dotDash"
            | "dashDotHeavy"
            | "dotDotDash"
            | "dashDotDotHeavy"
            | "wave"
            | "wavyHeavy"
            | "wavyDouble"
    )
}

pub fn parse_run_fmt(rpr: roxmltree::Node) -> RunFmt {
    let mut fmt = RunFmt {
        bold: bool_prop(rpr, "b"),
        italic: bool_prop(rpr, "i"),
        strikethrough: bool_prop(rpr, "strike"),
        ..Default::default()
    };

    // Underline (§17.3.2.40). `w:u@val` is ST_Underline (§17.18.99); the bool
    // stays true for any non-"none" value so existing single-line paths keep
    // working, and the raw value is carried for the renderer's style dispatch
    // (skipping "single"/"none", which need no hint). `w:u@color` (§17.18.99 note)
    // is an underline-only colour override (hex 6 or the literal "auto").
    if let Some(u) = child_w(rpr, "u") {
        // §17.3.2.40: an omitted @val inherits the setting from the previous
        // style-hierarchy level; the element's color/theme attributes can still
        // override an inherited underline. Do not invent `single` merely from
        // element presence.
        if let Some(val) = attr_w(u, "val") {
            fmt.underline = Some(val != "none");
            fmt.underline_style = if val == "none" || val == "single" {
                None
            } else {
                Some(val)
            };
        }
        if let Some(color) = attr_w(u, "color") {
            // Lowercase like the sibling `color` field below: the renderer's
            // `underlineColor !== 'auto'` check (§17.3.2.40's `color="auto"`
            // sentinel) is a strict-case comparison, so a producer that emits
            // "Auto" must still be normalized to the lowercase sentinel here.
            fmt.underline_color = Some(color.to_lowercase());
        }
        let lower = |value: &str| (!value.trim().is_empty()).then(|| value.to_lowercase());
        let identity = |value: &str| (!value.trim().is_empty()).then(|| value.to_string());
        let hex_byte = |value: &str| {
            (value.len() == 2 && value.chars().all(|c| c.is_ascii_hexdigit()))
                .then(|| value.to_ascii_uppercase())
        };
        fmt.underline_typography = Some(crate::types::UnderlineTypographyWire {
            val: normalized_string_attr(u, "val", |value| {
                is_underline_value(value).then(|| value.to_string())
            }),
            color: normalized_string_attr(u, "color", lower),
            theme_color: normalized_string_attr(u, "themeColor", identity),
            theme_tint: normalized_string_attr(u, "themeTint", hex_byte),
            theme_shade: normalized_string_attr(u, "themeShade", hex_byte),
        });
    }

    // Font size — w:sz (§17.3.2.38) governs Latin and East Asian (CJK) text
    // ONLY. w:szCs (§17.3.2.39) is the complex-script (Arabic/Hebrew/RTL) size,
    // recorded separately below as font_size_cs; the renderer selects it for
    // complex-script runs. Do NOT fall back to szCs here: a non-complex run that
    // carries only szCs (common Word editing residue) must inherit its sz from
    // the style/docDefaults chain, not adopt the complex-script metric —
    // otherwise body text with a leftover szCs renders a size too large.
    fmt.font_size = child_w(rpr, "sz")
        .and_then(|sz| attr_w(sz, "val"))
        .and_then(|value| half_pt_to_pt(&value));

    // Complex-script font size (ECMA-376 §17.3.2.39 w:szCs, half-points).
    // Recorded independently of the sz/szCs fallback above so RTL runs can use
    // the complex-script metric without disturbing the existing Latin/CJK size.
    fmt.font_size_cs = child_w(rpr, "szCs")
        .and_then(|sz_cs| attr_w(sz_cs, "val"))
        .and_then(|value| half_pt_to_pt(&value));

    // Record which of w:sz / w:szCs were set AT THIS LEVEL so the style-chain
    // merge can mirror a directly-applied Latin size into the complex-script
    // size when no szCs accompanies it (Word's behaviour — §17.3.2.18). Only a
    // successfully normalized finite, non-negative value is an authored
    // formatting fact; malformed children remain absent so they cannot
    // manufacture a 0pt override of inheritance.
    fmt.font_size_set_here = fmt.font_size.is_some();
    fmt.font_size_cs_set_here = fmt.font_size_cs.is_some();

    // Complex-script bold / italic toggles (ECMA-376 §17.3.2.3 / §17.3.2.17).
    fmt.bold_cs = bool_prop(rpr, "bCs");
    fmt.italic_cs = bool_prop(rpr, "iCs");

    // Complex-script language tag (ECMA-376 §17.3.2.20 w:lang/@w:bidi). Lower-
    // cased; its primary subtag later decides European-digit AN classification.
    if let Some(lang) = child_w(rpr, "lang") {
        if let Some(bidi) = attr_w(lang, "bidi") {
            if !bidi.is_empty() {
                fmt.lang_bidi = Some(bidi.to_lowercase());
            }
        }
        if let Some(east_asia) = attr_w(lang, "eastAsia") {
            if !east_asia.is_empty() {
                fmt.lang_east_asia = Some(east_asia.to_lowercase());
            }
        }
    }

    // Color. An explicit `<w:color w:val="auto"/>` (ECMA-376 §17.3.2.6) does NOT
    // name a concrete color and is NOT "inherit": auto leaves the final color to
    // be decided from the effective background at render time (an
    // implementation-defined black/white pick; ECMA-376 defines no algorithm).
    // We
    // record it as `color=None` + `color_auto=true`. `color_auto` also breaks
    // inheritance in `apply_run` so an inherited style color (e.g. a run with
    // `rStyle="PlaceholderText"`, gray #808080) does not survive past an
    // explicit auto. An absent `<w:color>` element stays None with color_auto
    // false (pure inherit).
    if let Some(col) = child_w(rpr, "color") {
        let val = attr_w(col, "val").unwrap_or_default();
        if val == "auto" {
            fmt.color = None;
            fmt.color_auto = true;
        } else if !val.is_empty() {
            fmt.color = Some(val.to_lowercase());
        }
    }

    // Font family. ECMA-376 §17.3.2.26 rFonts supports both direct typeface
    // attributes (ascii/hAnsi/eastAsia/cs) and theme references (asciiTheme,
    // hAnsiTheme, eastAsiaTheme, cstheme). Theme refs are resolved post-parse
    // in parse_document once a Theme is available; here we just record the
    // reference string under the corresponding axis. ECMA-376 §17.3.2.26 says
    // the theme attribute supersedes the matching direct typeface attribute.
    if let Some(rf) = child_w(rpr, "rFonts") {
        fmt.font_hint = attr_w(rf, "hint");
        fmt.font_family_ascii_direct = attr_w(rf, "ascii");
        fmt.font_family_ascii_theme = attr_w(rf, "asciiTheme").map(|t| format!("@theme:{t}"));
        fmt.font_family_ascii = fmt
            .font_family_ascii_theme
            .clone()
            .or_else(|| fmt.font_family_ascii_direct.clone());
        fmt.font_family_high_ansi_direct = attr_w(rf, "hAnsi");
        fmt.font_family_high_ansi_theme = attr_w(rf, "hAnsiTheme").map(|t| format!("@theme:{t}"));
        fmt.font_family_high_ansi = fmt
            .font_family_high_ansi_theme
            .clone()
            .or_else(|| fmt.font_family_high_ansi_direct.clone());

        fmt.font_family_east_asia_direct = attr_w(rf, "eastAsia");
        fmt.font_family_east_asia_theme =
            attr_w(rf, "eastAsiaTheme").map(|t| format!("@theme:{t}"));
        fmt.font_family_east_asia = fmt
            .font_family_east_asia_theme
            .clone()
            .or_else(|| fmt.font_family_east_asia_direct.clone());

        // Complex-script typeface (§17.3.2.26 @cs / @cstheme).
        fmt.font_family_cs_direct = attr_w(rf, "cs");
        fmt.font_family_cs_theme = attr_w(rf, "cstheme").map(|t| format!("@theme:{t}"));
        fmt.font_family_cs = fmt
            .font_family_cs_theme
            .clone()
            .or_else(|| fmt.font_family_cs_direct.clone());
    }

    // Run shading (ECMA-376 §17.3.2.32 w:shd). We adopt `@w:fill` only; `@w:val`
    // (the pattern) and `@w:color` are not modeled. `val="clear"` (inverse
    // video) is exact since only the fill is visible, but `val="solid"` etc.
    // drop information by ignoring the pattern foreground.
    if let Some(shd) = child_w(rpr, "shd") {
        if let Some(fill) = attr_w(shd, "fill") {
            if fill != "auto" && fill.len() == 6 {
                fmt.background = Some(fill.to_lowercase());
            }
        }
    }

    // Vertical alignment (superscript / subscript)
    if let Some(va) = child_w(rpr, "vertAlign") {
        let raw = attr_w(va, "val");
        let value = raw.as_deref().and_then(|value| match value {
            "superscript" => Some("super".to_string()),
            "subscript" => Some("sub".to_string()),
            _ => None,
        });
        fmt.vert_align = value.clone();
        fmt.vertical_align_typography = Some(typography_value(raw, value));
    }

    // All caps / small caps
    fmt.all_caps = bool_prop(rpr, "caps");
    fmt.small_caps = bool_prop(rpr, "smallCaps");

    // Double strikethrough
    fmt.dstrike = bool_prop(rpr, "dstrike");

    // Hidden text. §17.3.2.41 w:vanish hides the run in the normal/print view we
    // render. §17.3.2.44 w:webHidden hides ONLY in web page view (§17.18.102),
    // so it is recorded separately and must NOT feed `vanish` — otherwise TOC
    // dot-leader tabs and PAGEREF page numbers (marked `<w:webHidden/>` by Word)
    // would be dropped from the rendered page.
    fmt.vanish = bool_prop(rpr, "vanish");
    fmt.web_hidden = bool_prop(rpr, "webHidden");

    // Highlight
    if let Some(hl) = child_w(rpr, "highlight") {
        fmt.highlight = attr_w(hl, "val").filter(|v| v != "none");
    }

    // Emphasis mark (ECMA-376 §17.3.2.12 w:em / §17.18.24 ST_Em). A single
    // ST_Em value ("dot" | "comma" | "circle" | "underDot") drawn over (or,
    // for underDot, under) each non-space character. `val="none"` = no mark, so
    // it filters to `None` exactly like `highlight`.
    if let Some(em) = child_w(rpr, "em") {
        fmt.emphasis_mark = attr_w(em, "val").filter(|v| v != "none");
        fmt.emphasis_typography = Some(normalized_string_attr(em, "val", |value| {
            matches!(value, "none" | "dot" | "comma" | "circle" | "underDot")
                .then(|| value.to_string())
        }));
    }

    // Run border (ECMA-376 §17.3.2.4 w:bdr) — drawn as a box around the run.
    // val="none"/"nil" means no border, so we drop it rather than carrying a
    // zero-style EdgeBorder.
    if let Some(bdr) = child_w(rpr, "bdr") {
        fmt.border_typography = Some(parse_ct_border_typography(bdr));
        let edge = parse_edge_border(bdr);
        if edge.style != "none" && edge.style != "nil" {
            fmt.border = Some(edge);
        }
    }

    // Complex-script / RTL run (ECMA-376 §17.3.2.30 w:rtl). On-off toggle.
    fmt.rtl = bool_prop(rpr, "rtl");
    // Character-grid participation (ECMA-376 §17.3.2.34 w:snapToGrid).
    fmt.snap_to_grid = bool_prop(rpr, "snapToGrid");
    // §17.3.2.7 w:cs — complex-script run toggle (distinct from rFonts@cs,
    // which is only a font SLOT and must not force cs formatting).
    fmt.cs_toggle = bool_prop(rpr, "cs");

    // Character-spacing adjustment (ECMA-376 §17.3.2.35 `<w:spacing w:val>`).
    // ST_SignedTwipsMeasure (twips, 1/20 pt), signed — pitch added AFTER each
    // character. NOTE: the run-level `<w:spacing>` element carries only `w:val`;
    // the identically-named paragraph `<w:spacing>` (§17.3.1.33) uses
    // before/after/line and is parsed by `parse_para_fmt`, never here (this is
    // the run `rPr` context). `w:val="0"` is a real "no extra pitch" override,
    // so `Some(0.0)` must survive to shadow an inherited positive value.
    if let Some(sp) = child_w(rpr, "spacing") {
        if let Some(v) = attr_w(sp, "val") {
            fmt.char_spacing = Some(twips_to_pt(&v));
        }
    }

    // Manual run width (ECMA-376 §17.3.2.14 `<w:fitText>`). The whole element is
    // one cascaded property. Keep w:val in ST_TwipsMeasure units; suffixed
    // ST_PositiveUniversalMeasure values are normalized back to twips. Preserve
    // w:id as its trimmed lexical XSD integer because it has arbitrary precision
    // and is used only for equality linking.
    if let Some(fit_text) = child_w(rpr, "fitText") {
        fmt.fit_text = attr_w(fit_text, "val")
            .and_then(|value| parse_measure_to_pt(&value, 20.0))
            .map(|points| points * 20.0)
            .filter(|value| *value >= 0.0)
            .map(|val| FitTextSpec {
                val,
                id: attr_w(fit_text, "id")
                    .map(|value| value.trim().to_string())
                    .filter(|value| !value.is_empty()),
            });
    }

    // Expanded/Compressed text scale (ECMA-376 §17.3.2.43 `<w:w w:val>`).
    // ST_TextScale (§17.18.95): a percentage of normal character width, 1%–600%.
    // Word writes either a bare integer (`67`) or a percent literal (`67%`);
    // accept both by stripping a trailing '%'. Stored as a fraction (67 → 0.67)
    // and clamped to the spec's [0.01, 6.0] range. A malformed value is ignored
    // (leaves inheritance intact) rather than defaulting to 1.0, which would
    // incorrectly shadow an inherited scale.
    if let Some(w) = child_w(rpr, "w") {
        if let Some(v) = attr_w(w, "val") {
            let trimmed = v.trim().trim_end_matches('%');
            if let Ok(pct) = trimmed.parse::<f64>() {
                if pct > 0.0 {
                    fmt.char_scale = Some((pct / 100.0).clamp(0.01, 6.0));
                }
            }
        }
    }

    // Vertically raised/lowered text (ECMA-376 §17.3.2.24 `<w:position w:val>`).
    // ST_SignedHpsMeasure (half-points), signed — positive = raised above the
    // baseline, negative = lowered. Converted to points here; the renderer adds
    // it as a baseline y-offset without changing the font size or line box.
    if let Some(pos) = child_w(rpr, "position") {
        let raw = attr_w(pos, "val");
        let value = raw.as_deref().and_then(signed_half_pt_to_pt);
        // Keep the stable public projection byte-for-byte compatible while the
        // private value carries invalid status for the replacement layout path.
        fmt.position = raw.as_ref().map(|_| value.unwrap_or(0.0));
        fmt.position_typography = Some(typography_value(raw, value));
    }

    // Font kerning threshold (ECMA-376 §17.3.2.19 `<w:kern w:val>`). ST_HpsMeasure
    // (half-points) — the SMALLEST font size that has kerning applied. The mere
    // presence of the element turns kerning on (subject to the threshold); Word's
    // hierarchy default is OFF. `w:val="0"` = kern at all sizes. Stored in points.
    if let Some(kern) = child_w(rpr, "kern") {
        if let Some(v) = attr_w(kern, "val") {
            fmt.kerning = half_pt_to_pt(&v);
        }
    }

    // East Asian typography settings (ECMA-376 §17.3.2.10 `<w:eastAsianLayout>`).
    // `w:vert` (horizontal-in-vertical / 縦中横) and `w:vertCompress` are ST_OnOff
    // toggles; `w:combine` (two-lines-in-one) + `w:combineBrackets` (§17.18.8) are
    // parsed for completeness — the two-lines-in-one draw is a follow-up (no
    // fixture). All four are Option so an unset attribute inherits through the
    // style chain like the other run properties. `w:id` (§17.18.10) only links
    // multiple runs into one eastAsianLayout region and does not affect a single
    // run's layout, so it is not modeled.
    if let Some(eal) = child_w(rpr, "eastAsianLayout") {
        fmt.east_asian_vert = on_off_attr(eal, "vert");
        fmt.east_asian_vert_compress = on_off_attr(eal, "vertCompress");
        fmt.east_asian_combine = on_off_attr(eal, "combine");
        fmt.east_asian_combine_brackets = attr_w(eal, "combineBrackets").filter(|v| v != "none");
        fmt.combine_brackets_typography =
            Some(normalized_string_attr(eal, "combineBrackets", |value| {
                matches!(value, "none" | "round" | "square" | "angle" | "curly")
                    .then(|| value.to_string())
            }));
    }

    fmt
}

// ===== Table style parsing =====

fn shd_fill(node: roxmltree::Node) -> Option<String> {
    child_w(node, "shd")
        .and_then(|s| attr_w(s, "fill"))
        .filter(|f| f != "auto" && f.len() == 6)
        .map(|f| f.to_lowercase())
}

fn parse_edge_border(node: roxmltree::Node) -> EdgeBorder {
    let style = attr_w(node, "val").unwrap_or_else(|| "none".to_string());
    let width = attr_w(node, "sz")
        .and_then(|v| v.parse::<f64>().ok())
        .map(|v| v / 8.0)
        .unwrap_or(0.5);
    let color = attr_w(node, "color")
        .filter(|c| c != "auto")
        .map(|c| c.to_lowercase());
    // CT_Border @w:space is in points (no eighths conversion), unlike @w:sz.
    let space = attr_w(node, "space")
        .and_then(|v| v.parse::<f64>().ok())
        .unwrap_or(0.0);
    EdgeBorder {
        width,
        color,
        style,
        space,
    }
}

fn parse_raw_tbl_borders(node: roxmltree::Node) -> RawTblBorders {
    let mut b = RawTblBorders::default();
    for edge in node.children().filter(|n| n.is_element()) {
        let e = parse_edge_border(edge);
        match edge.tag_name().name() {
            "top" => b.top = Some(e),
            "bottom" => b.bottom = Some(e),
            "left" | "start" => b.left = Some(e),
            "right" | "end" => b.right = Some(e),
            "insideH" => b.inside_h = Some(e),
            "insideV" => b.inside_v = Some(e),
            _ => {}
        }
    }
    b
}

fn merge_raw_borders(dst: &mut RawTblBorders, src: &RawTblBorders) {
    if src.top.is_some() {
        dst.top = src.top.clone();
    }
    if src.bottom.is_some() {
        dst.bottom = src.bottom.clone();
    }
    if src.left.is_some() {
        dst.left = src.left.clone();
    }
    if src.right.is_some() {
        dst.right = src.right.clone();
    }
    if src.inside_h.is_some() {
        dst.inside_h = src.inside_h.clone();
    }
    if src.inside_v.is_some() {
        dst.inside_v = src.inside_v.clone();
    }
}

fn table_width_wire(
    node: Option<roxmltree::Node>,
) -> Option<crate::types::TableWidthAcquisitionWire> {
    node.map(|n| crate::types::TableWidthAcquisitionWire {
        kind: attr_w(n, "type"),
        value: attr_w(n, "w"),
    })
}

fn table_margin_wire(node: roxmltree::Node) -> crate::types::TableMarginAcquisitionWire {
    crate::types::TableMarginAcquisitionWire {
        top: table_width_wire(child_w(node, "top")),
        bottom: table_width_wire(child_w(node, "bottom")),
        start: table_width_wire(child_w(node, "start")),
        end: table_width_wire(child_w(node, "end")),
        left: table_width_wire(child_w(node, "left")),
        right: table_width_wire(child_w(node, "right")),
    }
}

pub(crate) fn merge_table_margin_layer(
    target: &mut Option<crate::types::TableMarginAcquisitionWire>,
    source: &Option<crate::types::TableMarginAcquisitionWire>,
) {
    let Some(source) = source else { return };
    let target = target.get_or_insert_with(Default::default);
    if source.top.is_some() {
        target.top = source.top.clone();
    }
    if source.bottom.is_some() {
        target.bottom = source.bottom.clone();
    }
    if source.start.is_some() {
        target.start = source.start.clone();
    }
    if source.end.is_some() {
        target.end = source.end.clone();
    }
    if source.left.is_some() {
        target.left = source.left.clone();
    }
    if source.right.is_some() {
        target.right = source.right.clone();
    }
}

fn parse_tbl_style_def(style_node: roxmltree::Node, based_on: Option<String>) -> TableStyleDef {
    let mut def = TableStyleDef {
        based_on,
        ..Default::default()
    };
    if let Some(tbl_pr) = child_w(style_node, "tblPr") {
        if let Some(borders) = child_w(tbl_pr, "tblBorders") {
            def.borders = parse_raw_tbl_borders(borders);
        }
        // ECMA-376 §17.7.6.7 / §17.7.6.5: row/column band widths. A value of 0
        // would make banding ill-defined (division by zero in the parity walk),
        // so clamp to at least 1. Stored as Some(..) only when present so an
        // omitting derived style inherits the base via resolve_table_style.
        def.row_band_size = child_w(tbl_pr, "tblStyleRowBandSize")
            .and_then(|n| attr_w(n, "val"))
            .and_then(|v| v.parse::<usize>().ok())
            .map(|n| n.max(1));
        def.col_band_size = child_w(tbl_pr, "tblStyleColBandSize")
            .and_then(|n| attr_w(n, "val"))
            .and_then(|v| v.parse::<usize>().ok())
            .map(|n| n.max(1));
        // ECMA-376 §17.4.42 `<w:tblCellMar>`: per-edge default cell margins
        // for tables that inherit this style. Each edge is stored as
        // Option<f64> (points) so a derived style omitting an edge inherits
        // the base via `resolve_table_style`. §17.18.90 ST_TblWidth says
        // dxa = 1/20 pt; w:type other than dxa is ignored for margins.
        if let Some(m) = child_w(tbl_pr, "tblCellMar") {
            def.cell_margins = Some(table_margin_wire(m));
            let edge = |name: &str| -> Option<f64> {
                child_w(m, name)
                    .filter(|n| attr_w(*n, "type").map(|t| t == "dxa").unwrap_or(true))
                    .and_then(|n| attr_w(n, "w"))
                    .map(|v| twips_to_pt(&v))
            };
            def.cell_margin_top = edge("top");
            def.cell_margin_bottom = edge("bottom");
            // §17.4.34 / §17.4.35 use `start`/`end` in the schema; Word also
            // writes the legacy `left`/`right` aliases. Accept either.
            def.cell_margin_left = edge("left").or_else(|| edge("start"));
            def.cell_margin_right = edge("right").or_else(|| edge("end"));
        }
        def.cell_spacing = table_width_wire(child_w(tbl_pr, "tblCellSpacing"));
    }
    if let Some(tc_pr) = child_w(style_node, "tcPr") {
        def.cell_shd = shd_fill(tc_pr);
        def.cell_valign = child_w(tc_pr, "vAlign").and_then(|v| attr_w(v, "val"));
    }
    if let Some(tr_pr) = child_w(style_node, "trPr") {
        def.tbl_header = bool_prop(tr_pr, "tblHeader");
        def.cant_split = bool_prop(tr_pr, "cantSplit");
    }
    // ECMA-376 §17.7.6: a table style's top-level `<w:rPr>`/`<w:pPr>` are
    // whole-table run/paragraph defaults applied to every cell (e.g. Calendar 3
    // sets `<w:color w:val="7F7F7F"/>` so day numbers are gray, and `<w:jc
    // w:val="right"/>` so they right-align). These are also indexed in
    // `StyleMap::styles` for the resolve_para pPr path, but we keep a copy here
    // so the conditional-formatting resolution (resolve_table_cond) can layer
    // them below the conditional rPr/pPr without re-walking the named-style map.
    if let Some(rpr) = child_w(style_node, "rPr") {
        def.run = Some(parse_run_fmt(rpr));
    }
    if let Some(ppr) = child_w(style_node, "pPr") {
        def.para = Some(parse_para_fmt(ppr));
    }
    for sp in children_w(style_node, "tblStylePr") {
        let Some(typ) = attr_w(sp, "type") else {
            continue;
        };
        let mut cf = CondFmt::default();
        if let Some(tc_pr) = child_w(sp, "tcPr") {
            cf.shd = shd_fill(tc_pr);
            if let Some(borders) = child_w(tc_pr, "tcBorders") {
                cf.borders = parse_raw_tbl_borders(borders);
            }
        }
        if let Some(tbl_pr) = child_w(sp, "tblPr") {
            if let Some(borders) = child_w(tbl_pr, "tblBorders") {
                merge_raw_borders(&mut cf.borders, &parse_raw_tbl_borders(borders));
            }
            cf.cell_spacing = table_width_wire(child_w(tbl_pr, "tblCellSpacing"));
            cf.cell_margins = child_w(tbl_pr, "tblCellMar").map(table_margin_wire);
        }
        if let Some(tr_pr) = child_w(sp, "trPr") {
            cf.tbl_header = bool_prop(tr_pr, "tblHeader");
            cf.cant_split = bool_prop(tr_pr, "cantSplit");
        }
        // §17.7.6: the conditional block carries its own `<w:rPr>`/`<w:pPr>`
        // (e.g. firstRow `<w:color w:val="365F91"/>` + `<w:jc w:val="right"/>`).
        if let Some(rpr) = child_w(sp, "rPr") {
            cf.run = Some(parse_run_fmt(rpr));
        }
        if let Some(ppr) = child_w(sp, "pPr") {
            cf.para = Some(parse_para_fmt(ppr));
        }
        def.cond.insert(typ, cf);
    }
    def
}

#[cfg(test)]
mod tests {
    use super::*;
    use roxmltree::Document as XmlDoc;

    fn stop(pos: f64, al: &str, ldr: &str) -> (f64, String, String) {
        (pos, al.to_string(), ldr.to_string())
    }

    #[test]
    fn merge_tab_stops_unions_by_position_keeping_inherited_leader() {
        // §17.3.1.37 — a direct left tab must NOT drop the style's right leader
        // tab (issue #820 TOC): the two coexist because they sit at different
        // positions. Result is sorted by pos.
        let style = vec![
            stop(50.85, "left", "none"),
            stop(467.5, "right", "underscore"),
        ];
        let direct = vec![stop(195.05, "left", "none")];
        let merged = merge_tab_stops(&style, &direct);
        assert_eq!(
            merged,
            vec![
                stop(50.85, "left", "none"),
                stop(195.05, "left", "none"),
                stop(467.5, "right", "underscore"),
            ]
        );
    }

    #[test]
    fn merge_tab_stops_direct_overrides_at_same_position() {
        // A direct stop at (within epsilon of) an inherited position wins.
        let style = vec![stop(100.0, "left", "dot")];
        let direct = vec![stop(100.02, "right", "none")];
        let merged = merge_tab_stops(&style, &direct);
        assert_eq!(merged, vec![stop(100.02, "right", "none")]);
    }

    #[test]
    fn merge_tab_stops_clear_removes_inherited_and_never_emits() {
        // `val="clear"` at a position removes the inherited stop there and is
        // itself dropped from the result (§17.18.84).
        let style = vec![stop(100.0, "left", "none"), stop(200.0, "right", "dot")];
        let direct = vec![stop(100.0, "clear", "none")];
        let merged = merge_tab_stops(&style, &direct);
        assert_eq!(merged, vec![stop(200.0, "right", "dot")]);
    }

    #[test]
    fn overflow_punct_participates_in_paragraph_style_cascade() {
        let parse = |inner: &str| {
            let xml = format!(
                r#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{inner}</w:pPr>"#
            );
            let document = XmlDoc::parse(&xml).unwrap();
            parse_para_fmt(document.root_element())
        };

        let mut inherited = parse(r#"<w:overflowPunct w:val="0"/>"#);
        apply_para(&mut inherited, &parse("<w:keepNext/>"));
        assert_eq!(
            inherited.overflow_punct,
            Some(false),
            "an omitted child property inherits the parent style value"
        );

        apply_para(&mut inherited, &parse("<w:overflowPunct/>"));
        assert_eq!(
            inherited.overflow_punct,
            Some(true),
            "a later style/direct on-value overrides inherited false"
        );
    }

    #[test]
    fn adjust_right_ind_participates_in_paragraph_style_cascade() {
        let parse = |inner: &str| {
            let xml = format!(
                r#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{inner}</w:pPr>"#
            );
            let document = XmlDoc::parse(&xml).unwrap();
            parse_para_fmt(document.root_element())
        };

        let mut inherited = parse(r#"<w:adjustRightInd w:val="0"/>"#);
        apply_para(&mut inherited, &parse("<w:keepNext/>"));
        assert_eq!(
            inherited.adjust_right_ind,
            Some(false),
            "an omitted child property inherits the parent style value"
        );

        apply_para(&mut inherited, &parse("<w:adjustRightInd/>"));
        assert_eq!(
            inherited.adjust_right_ind,
            Some(true),
            "a later style/direct on-value overrides inherited false"
        );
    }

    fn run_fmt_from(rpr_xml: &str) -> RunFmt {
        let xml = format!(
            r#"<w:rPr xmlns:w="{ns}">{body}</w:rPr>"#,
            ns = W_NS,
            body = rpr_xml
        );
        let doc = XmlDoc::parse(&xml).unwrap();
        parse_run_fmt(doc.root_element())
    }

    #[test]
    fn rfonts_preserves_four_axes_and_theme_supersedes_direct_typeface() {
        let fmt = run_fmt_from(
            r#"<w:rFonts
                w:ascii="Ascii Direct" w:asciiTheme="majorAscii"
                w:hAnsi="HAnsi Direct" w:hAnsiTheme="minorHAnsi"
                w:eastAsia="EA Direct" w:eastAsiaTheme="majorEastAsia"
                w:cs="CS Direct" w:cstheme="minorBidi"/>"#,
        );

        assert_eq!(fmt.font_family_ascii.as_deref(), Some("@theme:majorAscii"));
        assert_eq!(
            fmt.font_family_high_ansi.as_deref(),
            Some("@theme:minorHAnsi")
        );
        assert_eq!(
            fmt.font_family_east_asia.as_deref(),
            Some("@theme:majorEastAsia")
        );
        assert_eq!(fmt.font_family_cs.as_deref(), Some("@theme:minorBidi"));
        assert_eq!(
            fmt.font_family_ascii_direct.as_deref(),
            Some("Ascii Direct")
        );
        assert_eq!(
            fmt.font_family_ascii_theme.as_deref(),
            Some("@theme:majorAscii")
        );
        assert_eq!(
            fmt.font_family_high_ansi_direct.as_deref(),
            Some("HAnsi Direct")
        );
        assert_eq!(
            fmt.font_family_high_ansi_theme.as_deref(),
            Some("@theme:minorHAnsi")
        );
    }

    #[test]
    fn newer_direct_axes_replace_inherited_theme_axes_as_whole_slots() {
        let mut effective = run_fmt_from(
            r#"<w:rFonts
                w:asciiTheme="majorAscii" w:hAnsiTheme="majorHAnsi"
                w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/>"#,
        );
        let direct = run_fmt_from(
            r#"<w:rFonts
                w:ascii="Run ASCII" w:hAnsi="Run HANSI"
                w:eastAsia="Run EA" w:cs="Run CS"/>"#,
        );
        apply_run(&mut effective, &direct);

        assert_eq!(effective.font_family_ascii.as_deref(), Some("Run ASCII"));
        assert_eq!(
            effective.font_family_high_ansi.as_deref(),
            Some("Run HANSI")
        );
        assert_eq!(effective.font_family_east_asia.as_deref(), Some("Run EA"));
        assert_eq!(effective.font_family_cs.as_deref(), Some("Run CS"));

        let themes = run_fmt_from(
            r#"<w:rFonts
                w:asciiTheme="minorAscii" w:hAnsiTheme="minorHAnsi"
                w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>"#,
        );
        apply_run(&mut effective, &themes);
        assert_eq!(
            effective.font_family_ascii.as_deref(),
            Some("@theme:minorAscii")
        );
        assert_eq!(
            effective.font_family_high_ansi.as_deref(),
            Some("@theme:minorHAnsi")
        );
        assert_eq!(
            effective.font_family_east_asia.as_deref(),
            Some("@theme:minorEastAsia")
        );
        assert_eq!(
            effective.font_family_cs.as_deref(),
            Some("@theme:minorBidi")
        );

        let ascii_only = run_fmt_from(r#"<w:rFonts w:asciiTheme="minorAscii"/>"#);
        apply_run(&mut effective, &ascii_only);
        assert_eq!(
            effective.font_family_ascii.as_deref(),
            Some("@theme:minorAscii")
        );
        assert_eq!(
            effective.font_family_high_ansi.as_deref(),
            Some("@theme:minorHAnsi")
        );
        assert_eq!(
            effective.font_family_east_asia.as_deref(),
            Some("@theme:minorEastAsia")
        );
        assert_eq!(
            effective.font_family_cs.as_deref(),
            Some("@theme:minorBidi")
        );
    }

    fn para_fmt_from(ppr_xml: &str) -> ParaFmt {
        let xml = format!(
            r#"<w:pPr xmlns:w="{ns}">{body}</w:pPr>"#,
            ns = W_NS,
            body = ppr_xml
        );
        let doc = XmlDoc::parse(&xml).unwrap();
        parse_para_fmt(doc.root_element())
    }

    #[test]
    fn frame_pr_preserves_anchor_lock_on_the_private_wire() {
        // ECMA-376 §17.3.1.11: anchorLock participates in the effective
        // CT_FramePr identity used to group adjacent framed paragraphs. Keep it
        // off the stable public FramePr projection while preserving the fact for
        // layout acquisition under an explicitly private wire name.
        let locked = para_fmt_from(r#"<w:framePr w:anchorLock="on"/>"#)
            .frame_pr
            .expect("framePr parses");
        assert!(locked.anchor_lock);
        let locked_wire = serde_json::to_value(&locked).expect("framePr serializes");
        assert_eq!(locked_wire["__anchorLock"], serde_json::json!(true));
        assert!(locked_wire.get("anchorLock").is_none());

        let unlocked = para_fmt_from(r#"<w:framePr/>"#)
            .frame_pr
            .expect("framePr parses");
        assert!(!unlocked.anchor_lock);
        let unlocked_wire = serde_json::to_value(&unlocked).expect("framePr serializes");
        assert!(unlocked_wire.get("__anchorLock").is_none());
    }

    // ── RB2 neutralization: a pathologically deep styles.xml is rejected by the
    //    depth pre-check in `parse_guarded` BEFORE roxmltree's recursive tree
    //    builder runs, so `StyleMap::parse` returns gracefully instead of trapping
    //    the whole parse on a stack overflow. This test runs on the DEFAULT
    //    (small) test-thread stack ON PURPOSE — if `StyleMap::parse` handed the
    //    5 000-deep XML straight to roxmltree it would overflow and abort the
    //    process here. That it returns at all is the guarantee. §17.7 (styles).
    #[test]
    fn deeply_nested_styles_xml_is_rejected_not_trapped() {
        let mut xml = format!(r#"<w:styles xmlns:w="{ns}">"#, ns = W_NS);
        // 5 000 levels — ~20× MAX_XML_DEPTH and well past roxmltree's ~1 000-deep
        // small-stack overflow threshold. Nest an arbitrary element (styles.xml is
        // shallow in the wild, so this can only be an attack).
        for _ in 0..5_000 {
            xml.push_str("<w:x>");
        }
        xml.push('y');
        for _ in 0..5_000 {
            xml.push_str("</w:x>");
        }
        xml.push_str("</w:styles>");

        // Must return (not trap). The rejected part yields an empty style map —
        // the "skip the part, keep the document" degradation contract.
        let map = StyleMap::parse(&xml);
        assert!(
            map.styles.is_empty(),
            "an over-deep styles.xml must be rejected, yielding no styles"
        );
    }

    #[test]
    fn circular_based_on_chain_terminates_and_applies_each_style_once() {
        // ECMA-376 Part 1 §17.7.4.3 defines style inheritance, while
        // MS-OI29500 §2.1.233 says Word requires the based-on chain not to be
        // circular. Invalid input must still degrade without overflowing the
        // WASM stack. The edge that returns Base to Derived is ignored.
        let xml = format!(
            r#"<w:styles xmlns:w="{ns}">
                <w:style w:type="paragraph" w:styleId="Base">
                    <w:basedOn w:val="Derived"/>
                    <w:pPr><w:spacing w:after="120"/></w:pPr>
                </w:style>
                <w:style w:type="paragraph" w:styleId="Derived">
                    <w:basedOn w:val="Base"/>
                    <w:rPr><w:b/></w:rPr>
                </w:style>
            </w:styles>"#,
            ns = W_NS,
        );

        let (paragraph, run) = StyleMap::parse(&xml).resolve_para(Some("Derived"), None);
        assert_eq!(paragraph.space_after, Some(6.0));
        assert_eq!(run.bold, Some(true));

        let self_referencing = format!(
            r#"<w:styles xmlns:w="{ns}">
                <w:style w:type="paragraph" w:styleId="Normal">
                    <w:basedOn w:val="Normal"/>
                    <w:rPr><w:sz w:val="20"/></w:rPr>
                </w:style>
            </w:styles>"#,
            ns = W_NS,
        );
        let (_, run) = StyleMap::parse(&self_referencing).resolve_para(Some("Normal"), None);
        assert_eq!(run.font_size, Some(10.0));
    }

    #[test]
    fn pbdr_merges_per_edge_over_inherited_box() {
        // Each pBdr EDGE inherits independently across the style hierarchy (bottom
        // §17.3.1.7, left §17.3.1.17, right §17.3.1.28, top §17.3.1.42, between
        // §17.3.1.5: "if this element is omitted ... its value is
        // determined by the setting previously set at any level"). So a level that
        // sets only `<w:bottom>` must KEEP the inherited top/left/right; the old
        // wholesale-replace dropped them. Mirrors the direct-over-style merge too
        // (apply_para is shared after PR #613).
        let mut base = para_fmt_from(
            r#"<w:pBdr>
                 <w:top w:val="single" w:sz="8" w:color="FF0000"/>
                 <w:left w:val="single" w:sz="8" w:color="FF0000"/>
                 <w:bottom w:val="single" w:sz="8" w:color="FF0000"/>
                 <w:right w:val="single" w:sz="8" w:color="FF0000"/>
               </w:pBdr>"#,
        );
        let direct = para_fmt_from(
            r#"<w:pBdr><w:bottom w:val="single" w:sz="24" w:color="0000FF"/></w:pBdr>"#,
        );
        apply_para(&mut base, &direct);
        let b = base.para_borders.expect("borders present after merge");
        // The three edges the direct level did not touch survive from the inherited box.
        assert!(b.top.is_some(), "inherited top edge kept");
        assert!(b.left.is_some(), "inherited left edge kept");
        assert!(b.right.is_some(), "inherited right edge kept");
        // The one edge the direct level set takes its value.
        let bottom = b.bottom.expect("bottom edge present");
        assert_eq!(
            bottom.color.as_deref(),
            Some("0000ff"),
            "direct bottom overrides the inherited bottom"
        );
        assert_eq!(bottom.width, 3.0, "direct bottom sz=24 → 3.0pt");
    }

    #[test]
    fn pbdr_with_no_inherited_box_is_unchanged() {
        // No inherited pBdr: a direct single-edge pBdr is taken as-is (per-edge
        // merge over an empty box is identical to the old wholesale behavior).
        let mut base = ParaFmt::default();
        let direct = para_fmt_from(
            r#"<w:pBdr><w:bottom w:val="single" w:sz="12" w:color="000000"/></w:pBdr>"#,
        );
        apply_para(&mut base, &direct);
        let b = base.para_borders.expect("borders present");
        assert!(b.bottom.is_some(), "direct bottom present");
        assert!(b.top.is_none() && b.left.is_none() && b.right.is_none());
    }

    #[test]
    fn explicit_nil_edge_is_kept_distinct_from_omitted() {
        // §17.3.1.7: only an OMITTED edge inherits. An EXPLICIT `<w:bottom
        // w:val="nil"/>` (or "none") is PRESENT and specifies "no border", so it
        // must survive parsing as a distinct "cleared" edge (style "none") rather
        // than collapse to None (which is reserved for an omitted edge that
        // inherits). nil/none are synonyms → normalized to "none".
        let fmt = para_fmt_from(r#"<w:pBdr><w:bottom w:val="nil"/></w:pBdr>"#);
        let b = fmt
            .para_borders
            .expect("an explicit-nil edge keeps the pBdr alive");
        let bottom = b
            .bottom
            .expect("explicit nil bottom is a present 'cleared' edge");
        assert_eq!(bottom.style, "none", "nil normalized to none");
        // The OMITTED edges inherit → None.
        assert!(b.top.is_none(), "omitted top stays None (inherits)");
        assert!(b.left.is_none() && b.right.is_none());
    }

    #[test]
    fn explicit_nil_edge_clears_an_inherited_edge() {
        // A level with a full box; a child sets only `<w:bottom w:val="nil"/>` to
        // REMOVE the inherited bottom. The other inherited edges survive (per-edge),
        // and the bottom is cleared (becomes a "none" edge), NOT inherited.
        let mut base = para_fmt_from(
            r#"<w:pBdr>
                 <w:top w:val="single" w:sz="8" w:color="FF0000"/>
                 <w:left w:val="single" w:sz="8" w:color="FF0000"/>
                 <w:bottom w:val="single" w:sz="8" w:color="FF0000"/>
                 <w:right w:val="single" w:sz="8" w:color="FF0000"/>
               </w:pBdr>"#,
        );
        let direct = para_fmt_from(r#"<w:pBdr><w:bottom w:val="nil"/></w:pBdr>"#);
        apply_para(&mut base, &direct);
        let b = base.para_borders.expect("borders present");
        assert!(
            b.top.is_some() && b.left.is_some() && b.right.is_some(),
            "other edges kept"
        );
        let bottom = b
            .bottom
            .expect("bottom is the explicit cleared edge, not None");
        assert_eq!(
            bottom.style, "none",
            "inherited bottom was CLEARED, not kept"
        );
    }

    #[test]
    fn explicit_color_auto_breaks_inheritance_and_defers_to_background() {
        // ECMA-376 §17.3.2.6: an explicit w:color="auto" is NOT "inherit" and is
        // NOT a concrete color either — it defers the final color to the
        // background-contrast resolution (implementation-defined; no normative
        // algorithm). We record it as
        // color=None + color_auto=true so the renderer can pick black/white from
        // the effective background. The intent of overriding an inherited style
        // color (e.g. PlaceholderText gray) is carried by `color_auto` in
        // `apply_run`, which clears `dst.color` when a child sets auto.
        let fmt = run_fmt_from(r#"<w:color w:val="auto"/>"#);
        assert_eq!(fmt.color, None);
        assert!(fmt.color_auto);
    }

    #[test]
    fn color_auto_breaks_inherited_concrete_color_on_merge() {
        // §17.3.2.6: a child run that sets w:color="auto" must drop an inherited
        // concrete color (e.g. PlaceholderText gray #808080), deferring to the
        // background-contrast pass rather than keeping the gray.
        let mut dst = RunFmt {
            color: Some("808080".to_string()),
            ..RunFmt::default()
        };
        let src = run_fmt_from(r#"<w:color w:val="auto"/>"#);
        apply_run(&mut dst, &src);
        assert_eq!(dst.color, None);
        assert!(dst.color_auto);
    }

    #[test]
    fn east_asia_font_hint_and_language_override_through_run_merge() {
        let mut dst = run_fmt_from(r#"<w:rFonts w:hint="default"/><w:lang w:eastAsia="ja-JP"/>"#);
        let src = run_fmt_from(r#"<w:rFonts w:hint="eastAsia"/><w:lang w:eastAsia="ZH-cn"/>"#);
        apply_run(&mut dst, &src);
        assert_eq!(dst.font_hint.as_deref(), Some("eastAsia"));
        assert_eq!(dst.lang_east_asia.as_deref(), Some("zh-cn"));
    }

    #[test]
    fn run_border_bdr_is_parsed() {
        // ECMA-376 §17.3.2.4 w:bdr — a run-level border ("box"). w:sz is in
        // eighths of a point (4 → 0.5pt); w:space is in points (1 → 1.0pt).
        let fmt = run_fmt_from(r#"<w:bdr w:val="single" w:sz="4" w:space="1" w:color="auto"/>"#);
        let b = fmt.border.expect("border should be Some");
        assert_eq!(b.style, "single");
        assert_eq!(b.width, 0.5);
        assert_eq!(b.space, 1.0);
        // color=auto on a border means "automatic" → recorded as None so the
        // renderer falls back to the default text color.
        assert_eq!(b.color, None);
    }

    #[test]
    fn run_shading_fill_sets_background() {
        // §17.3.2.32 w:shd/@w:fill — run shading fill becomes the run background.
        // Regression guard for the inverse-video case (black fill).
        let fmt = run_fmt_from(r#"<w:shd w:val="clear" w:color="auto" w:fill="000000"/>"#);
        assert_eq!(fmt.background.as_deref(), Some("000000"));
    }

    #[test]
    fn explicit_hex_color_is_lowercased() {
        let fmt = run_fmt_from(r#"<w:color w:val="FF0000"/>"#);
        assert_eq!(fmt.color.as_deref(), Some("ff0000"));
    }

    #[test]
    fn underline_color_is_lowercased_like_sibling_color() {
        // §17.3.2.40 w:u@color, hex case — must lowercase the same as the
        // sibling w:color@val field above.
        let fmt = run_fmt_from(r#"<w:u w:val="single" w:color="FF0000"/>"#);
        assert_eq!(fmt.underline_color.as_deref(), Some("ff0000"));
    }

    #[test]
    fn underline_without_val_inherits_instead_of_enabling_single() {
        let direct = run_fmt_from(r#"<w:u w:color="FF0000"/>"#);
        assert_eq!(direct.underline, None);
        assert_eq!(direct.underline_style, None);
        assert_eq!(direct.underline_color.as_deref(), Some("ff0000"));

        let mut inherited = run_fmt_from(r#"<w:u w:val="double"/>"#);
        apply_run(&mut inherited, &direct);
        assert_eq!(inherited.underline, Some(true));
        assert_eq!(inherited.underline_style.as_deref(), Some("double"));
        assert_eq!(inherited.underline_color.as_deref(), Some("ff0000"));
    }

    #[test]
    fn underline_color_auto_is_lowercased_so_renderer_sentinel_check_matches() {
        // A producer that writes the capitalized "Auto" sentinel must still
        // normalize to lowercase "auto": the TS renderer's underline-color
        // override guard (`underlineColor !== 'auto'`) is a strict-case
        // comparison, so an un-lowercased "Auto" would slip past it and be
        // treated as a literal (invalid) hex color instead of "follow the
        // glyph color".
        let fmt = run_fmt_from(r#"<w:u w:val="single" w:color="Auto"/>"#);
        assert_eq!(fmt.underline_color.as_deref(), Some("auto"));
    }

    #[test]
    fn sz_sets_latin_cjk_font_size() {
        // §17.3.2.38: w:sz (half-points) → the Latin/CJK font size.
        let fmt = run_fmt_from(r#"<w:sz w:val="20"/>"#);
        assert_eq!(fmt.font_size, Some(10.0));
    }

    #[test]
    fn sz_and_szcs_accept_positive_universal_measures() {
        // ECMA-376 §17.18.42: ST_HpsMeasure is the union of an unsigned
        // half-point count and ST_PositiveUniversalMeasure. Unit-bearing values
        // are absolute, so they must not be divided by two.
        let fmt = run_fmt_from(r#"<w:sz w:val="3.6mm"/><w:szCs w:val="12pt"/>"#);
        assert!((fmt.font_size.expect("w:sz parses") - 72.0 * 3.6 / 25.4).abs() < 1e-9);
        assert_eq!(fmt.font_size_cs, Some(12.0));
    }

    #[test]
    fn universal_measure_font_sizes_survive_docdefaults_and_basedon_cascades() {
        // §17.7.2 applies docDefaults below the paragraph style chain. Exercise
        // the lexical union through both inheritance sources rather than only
        // testing the conversion helper.
        let defaults = format!(
            r#"<w:styles xmlns:w="{ns}">
              <w:docDefaults><w:rPrDefault><w:rPr>
                <w:sz w:val="3.6mm"/><w:szCs w:val="12pt"/>
              </w:rPr></w:rPrDefault></w:docDefaults>
              <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
            </w:styles>"#,
            ns = W_NS,
        );
        let (_, inherited_defaults) = StyleMap::parse(&defaults).resolve_para(None, None);
        assert!(
            (inherited_defaults.font_size.expect("docDefaults w:sz") - 72.0 * 3.6 / 25.4).abs()
                < 1e-9
        );
        assert_eq!(inherited_defaults.font_size_cs, Some(12.0));

        let styles = format!(
            r#"<w:styles xmlns:w="{ns}">
              <w:docDefaults><w:rPrDefault><w:rPr>
                <w:sz w:val="18"/><w:szCs w:val="20"/>
              </w:rPr></w:rPrDefault></w:docDefaults>
              <w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>
              <w:style w:type="paragraph" w:styleId="Base">
                <w:basedOn w:val="Normal"/>
                <w:rPr><w:sz w:val="3.6mm"/><w:szCs w:val="12pt"/></w:rPr>
              </w:style>
              <w:style w:type="paragraph" w:styleId="Derived">
                <w:basedOn w:val="Base"/>
              </w:style>
            </w:styles>"#,
            ns = W_NS,
        );
        let (_, inherited_style) = StyleMap::parse(&styles).resolve_para(Some("Derived"), None);
        assert!(
            (inherited_style.font_size.expect("basedOn w:sz") - 72.0 * 3.6 / 25.4).abs() < 1e-9
        );
        assert_eq!(inherited_style.font_size_cs, Some(12.0));

        // A malformed value is not a valid formatting value and must not become
        // an authored 0pt override that blocks the inherited style.
        for invalid in [
            "invalid", "-0", "-0pt", "-2", "-2pt", "1.5", "1e2", "+12", "+12pt",
        ] {
            let mut inherited = inherited_style.clone();
            let direct = run_fmt_from(&format!(
                r#"<w:sz w:val="{invalid}"/><w:szCs w:val="{invalid}"/>"#
            ));
            apply_run(&mut inherited, &direct);
            assert!(
                (inherited
                    .font_size
                    .expect("style survives invalid direct w:sz")
                    - 72.0 * 3.6 / 25.4)
                    .abs()
                    < 1e-9,
                "{invalid:?} must not block w:sz inheritance"
            );
            assert_eq!(
                inherited.font_size_cs,
                Some(12.0),
                "{invalid:?} must not block w:szCs inheritance"
            );
        }

        for valid_zero in ["0", "0pt"] {
            let mut inherited = inherited_style.clone();
            let direct = run_fmt_from(&format!(
                r#"<w:sz w:val="{valid_zero}"/><w:szCs w:val="{valid_zero}"/>"#
            ));
            apply_run(&mut inherited, &direct);
            assert_eq!(inherited.font_size, Some(0.0));
            assert_eq!(inherited.font_size_cs, Some(0.0));
        }
    }

    #[test]
    fn szcs_alone_does_not_set_latin_cjk_font_size() {
        // §17.3.2.39: w:szCs is the complex-script size only. A non-complex run
        // carrying ONLY szCs (Word editing residue) must leave font_size None so
        // it inherits sz from the style/docDefaults chain — it must NOT adopt the
        // complex-script metric (the bug that rendered body text at 12pt instead
        // of the inherited 10pt). szCs is still recorded as font_size_cs for
        // complex-script runs.
        let fmt = run_fmt_from(r#"<w:szCs w:val="24"/>"#);
        assert_eq!(fmt.font_size, None);
        assert_eq!(fmt.font_size_cs, Some(12.0));
    }

    #[test]
    fn absent_color_element_stays_none_to_inherit() {
        let fmt = run_fmt_from(r#"<w:b/>"#);
        assert_eq!(fmt.color, None);
    }

    #[test]
    fn vanish_hides_run_in_print_layout() {
        // ECMA-376 §17.3.2.41 w:vanish — "Hidden Text". Hidden in the normal
        // (print/page) view we render, so it sets the `vanish` flag the parser
        // uses to skip the run.
        let fmt = run_fmt_from(r#"<w:vanish/>"#);
        assert_eq!(fmt.vanish, Some(true));
        assert_eq!(fmt.web_hidden, None);
    }

    #[test]
    fn web_hidden_does_not_vanish_in_print_layout() {
        // ECMA-376 §17.3.2.44 w:webHidden — text hidden ONLY in *web page view*
        // (§17.18.102), NOT in the normal/print layout this renderer produces.
        // Conflating it with §17.3.2.41 w:vanish wrongly dropped TOC page numbers
        // and dot-leader tab runs (which Word marks `<w:webHidden/>`) from the
        // rendered page. webHidden must leave `vanish` unset so the run renders;
        // the flag is preserved separately for a future web-view mode.
        let fmt = run_fmt_from(r#"<w:webHidden/>"#);
        assert_eq!(fmt.vanish, None);
        assert_eq!(fmt.web_hidden, Some(true));
    }

    /// Minimal styles.xml mirroring sample-11's TOC1 (bold) / TOC2 (italic)
    /// paragraph styles: both basedOn Normal with the toggle in the style's
    /// direct `<w:rPr>`.
    fn toc_style_map() -> StyleMap {
        let xml = format!(
            r#"<w:styles xmlns:w="{ns}">
              <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
                <w:name w:val="Normal"/>
              </w:style>
              <w:style w:type="paragraph" w:styleId="TOC1">
                <w:name w:val="toc 1"/><w:basedOn w:val="Normal"/>
                <w:rPr><w:b/><w:bCs/><w:sz w:val="20"/></w:rPr>
              </w:style>
              <w:style w:type="paragraph" w:styleId="TOC2">
                <w:name w:val="toc 2"/><w:basedOn w:val="Normal"/>
                <w:rPr><w:i/><w:iCs/><w:sz w:val="20"/></w:rPr>
              </w:style>
            </w:styles>"#,
            ns = W_NS
        );
        StyleMap::parse(&xml)
    }

    #[test]
    fn toc1_style_run_base_is_bold() {
        // ECMA-376 §17.7.4 / §17.7.2: a paragraph style's direct `<w:rPr><w:b/>`
        // resolves onto the run BASE for every run in a TOC1 paragraph, so the
        // entry text, dot leader and page number all inherit bold unless a run
        // overrides it directly. (sample-11 p6 ToC, report #5.)
        let sm = toc_style_map();
        let (_para, run) = sm.resolve_para(Some("TOC1"), None);
        assert_eq!(run.bold, Some(true));
    }

    #[test]
    fn toc2_style_run_base_is_italic() {
        // §17.7.4 / §17.7.2: TOC2's `<w:rPr><w:i/>` resolves onto the run base —
        // "Inline formatting" etc. inherit italic. (sample-11 p6 ToC, report #6.)
        let sm = toc_style_map();
        let (_para, run) = sm.resolve_para(Some("TOC2"), None);
        assert_eq!(run.italic, Some(true));
    }

    // ===== Table style conditional rPr/pPr (§17.7.6 / §17.7.2) =====

    /// A minimal styles.xml with a Calendar-3-like table style: a whole-table
    /// rPr gray color + a firstRow conditional rPr blue color and a firstRow
    /// pPr right-jc. Mirrors the real Calendar 3 (styleId "Calendar3") subset
    /// the parser must resolve.
    fn calendar_style_map() -> StyleMap {
        let xml = format!(
            r#"<w:styles xmlns:w="{ns}">
              <w:style w:type="table" w:styleId="Calendar3">
                <w:name w:val="Calendar 3"/>
                <w:pPr><w:jc w:val="right"/></w:pPr>
                <w:rPr><w:color w:val="7F7F7F"/></w:rPr>
                <w:tblStylePr w:type="firstRow">
                  <w:pPr><w:jc w:val="right"/></w:pPr>
                  <w:rPr><w:color w:val="365F91"/><w:sz w:val="44"/></w:rPr>
                </w:tblStylePr>
              </w:style>
            </w:styles>"#,
            ns = W_NS
        );
        StyleMap::parse(&xml)
    }

    #[test]
    fn table_style_parses_whole_table_and_conditional_run_para() {
        let sm = calendar_style_map();
        let def = sm.resolve_table_style("Calendar3");
        // Whole-table rPr/pPr captured.
        assert_eq!(def.run.as_ref().unwrap().color.as_deref(), Some("7f7f7f"));
        assert_eq!(
            def.para.as_ref().unwrap().alignment.as_deref(),
            Some("right")
        );
        // firstRow conditional rPr/pPr captured.
        let fr = def.cond.get("firstRow").unwrap();
        assert_eq!(fr.run.as_ref().unwrap().color.as_deref(), Some("365f91"));
        assert_eq!(fr.run.as_ref().unwrap().font_size, Some(22.0));
        assert_eq!(
            fr.para.as_ref().unwrap().alignment.as_deref(),
            Some("right")
        );
    }

    #[test]
    fn firstrow_conditional_run_color_is_inherited_by_cell_base() {
        // §17.7.6 + §17.7.2: a cell paragraph in the firstRow inherits the
        // conditional rPr color (365F91) as its run BASE — this is the calendar
        // header "Sun/Mon/…" blue.
        let sm = calendar_style_map();
        let cond = {
            let def = sm.resolve_table_style("Calendar3");
            def.cond.get("firstRow").cloned()
        };
        let (para, run) = sm.resolve_para_cond(None, Some("Calendar3"), cond.as_ref());
        assert_eq!(run.color.as_deref(), Some("365f91"));
        // firstRow pPr right-jc also resolves onto the paragraph.
        assert_eq!(para.alignment.as_deref(), Some("right"));
    }

    #[test]
    fn body_row_inherits_whole_table_color_not_conditional() {
        // A row with no conditional (cond = None) inherits only the whole-table
        // rPr gray (7F7F7F) — the calendar day numbers — never the firstRow blue.
        let sm = calendar_style_map();
        let (_para, run) = sm.resolve_para_cond(None, Some("Calendar3"), None);
        assert_eq!(run.color.as_deref(), Some("7f7f7f"));
    }

    #[test]
    fn direct_run_color_overrides_conditional_base() {
        // §17.7.2 ordering: the table conditional rPr is a BASE below direct
        // run formatting. A run that carries its OWN w:color must win over the
        // firstRow conditional color. We resolve the conditional base, then
        // layer a direct rPr the way the run walk does (set-value wins).
        let sm = calendar_style_map();
        let cond = {
            let def = sm.resolve_table_style("Calendar3");
            def.cond.get("firstRow").cloned()
        };
        let (_para, base_run) = sm.resolve_para_cond(None, Some("Calendar3"), cond.as_ref());
        assert_eq!(base_run.color.as_deref(), Some("365f91"));

        // Direct rPr on the run: explicit red. apply_run mirrors the
        // set-value-wins merge that apply_direct_run performs for the run walk.
        let mut fmt = base_run.clone();
        let direct = run_fmt_from(r#"<w:color w:val="FF0000"/>"#);
        apply_run(&mut fmt, &direct);
        assert_eq!(
            fmt.color.as_deref(),
            Some("ff0000"),
            "direct run color must override the conditional base"
        );
    }

    // ===== Table-style tblCellMar inheritance (§17.4.42 / §17.7.6) =====

    /// styles.xml mirroring sample-3: the default table style "TableNormal"
    /// (`w:default="1"`) carries `<w:tblPr><w:tblCellMar>` with left/right=108
    /// twips. ECMA-376 §17.4.42 says a table whose `<w:tblPr>` omits
    /// `<w:tblCellMar>` inherits the margins from its associated table style;
    /// when no style is set, the spec maps the table to the default table
    /// style. The parser must therefore (a) capture per-edge margins on
    /// `TableStyleDef`, (b) flatten them through the basedOn chain, and (c)
    /// expose the default table style so `parser.rs` can fall back to it when
    /// no `<w:tblStyle>` is present on the table.
    fn default_table_style_map() -> StyleMap {
        let xml = format!(
            r#"<w:styles xmlns:w="{ns}">
              <w:style w:type="table" w:default="1" w:styleId="TableNormal">
                <w:name w:val="Normal Table"/>
                <w:tblPr>
                  <w:tblCellMar>
                    <w:top w:w="0" w:type="dxa"/>
                    <w:left w:w="108" w:type="dxa"/>
                    <w:bottom w:w="0" w:type="dxa"/>
                    <w:right w:w="108" w:type="dxa"/>
                  </w:tblCellMar>
                </w:tblPr>
              </w:style>
            </w:styles>"#,
            ns = W_NS
        );
        StyleMap::parse(&xml)
    }

    #[test]
    fn table_style_captures_tbl_cell_mar_per_edge() {
        // §17.4.42: a table style's `<w:tblPr><w:tblCellMar>` defines defaults
        // for tables that inherit from it. The parser must surface each edge
        // on TableStyleDef so a table that omits `<w:tblCellMar>` can fall
        // back to these values (sample-3 root cause). 108 twips = 5.4 pt.
        let sm = default_table_style_map();
        let def = sm.resolve_table_style("TableNormal");
        assert_eq!(def.cell_margin_top, Some(0.0));
        assert_eq!(def.cell_margin_bottom, Some(0.0));
        assert_eq!(def.cell_margin_left, Some(5.4));
        assert_eq!(def.cell_margin_right, Some(5.4));
    }

    #[test]
    fn default_table_style_id_is_exposed() {
        // §17.7.4: a `<w:style w:type="table" w:default="1">` is the implicit
        // style for tables that omit `<w:tblStyle>`. The parser must expose
        // the styleId so callers can resolve the default style chain
        // (sample-3's TableNormal carries the 108-twips cell margins that
        // make tcMar-absent cells line up).
        let sm = default_table_style_map();
        assert_eq!(sm.default_table_style_id(), Some("TableNormal"));
    }

    #[test]
    fn tbl_cell_mar_inherits_through_based_on_chain() {
        // §17.7.6 + §17.4.42: a derived table style that omits `<w:tblCellMar>`
        // inherits the per-edge values from its base. A derived style that
        // sets ONLY one edge keeps the base values for the rest.
        let xml = format!(
            r#"<w:styles xmlns:w="{ns}">
              <w:style w:type="table" w:default="1" w:styleId="TableNormal">
                <w:tblPr>
                  <w:tblCellMar>
                    <w:left w:w="108" w:type="dxa"/>
                    <w:right w:w="108" w:type="dxa"/>
                  </w:tblCellMar>
                </w:tblPr>
              </w:style>
              <w:style w:type="table" w:styleId="Derived">
                <w:basedOn w:val="TableNormal"/>
                <w:tblPr>
                  <w:tblCellMar>
                    <w:left w:w="200" w:type="dxa"/>
                  </w:tblCellMar>
                </w:tblPr>
              </w:style>
            </w:styles>"#,
            ns = W_NS
        );
        let sm = StyleMap::parse(&xml);
        let def = sm.resolve_table_style("Derived");
        // Derived overrides left (200 twips = 10pt) …
        assert_eq!(def.cell_margin_left, Some(10.0));
        // … but right inherits TableNormal (108 twips = 5.4pt).
        assert_eq!(def.cell_margin_right, Some(5.4));
    }

    #[test]
    fn conditional_run_para_default_to_none_when_absent() {
        // A table style with no rPr/pPr anywhere yields None — no panics, no
        // spurious base layer (so cells inherit docDefaults/paragraph style).
        let xml = format!(
            r#"<w:styles xmlns:w="{ns}">
              <w:style w:type="table" w:styleId="Plain">
                <w:name w:val="Plain"/>
                <w:tblStylePr w:type="firstRow">
                  <w:tcPr><w:shd w:val="clear" w:fill="CCCCCC"/></w:tcPr>
                </w:tblStylePr>
              </w:style>
            </w:styles>"#,
            ns = W_NS
        );
        let sm = StyleMap::parse(&xml);
        let def = sm.resolve_table_style("Plain");
        assert!(def.run.is_none());
        assert!(def.para.is_none());
        let fr = def.cond.get("firstRow").unwrap();
        assert!(fr.run.is_none());
        assert!(fr.para.is_none());
        // shd still parses (existing behavior unchanged).
        assert_eq!(fr.shd.as_deref(), Some("cccccc"));
    }

    // ── WD4: run-level character metrics (§17.3.2.35 / .43 / .24 / .19) ──────

    #[test]
    fn char_spacing_parses_signed_twips_to_pt() {
        // §17.3.2.35: val is ST_SignedTwipsMeasure (twips = 1/20 pt). The spec
        // example `<w:spacing w:val="200"/>` == 10 pt of extra pitch.
        let f = run_fmt_from(r#"<w:spacing w:val="200"/>"#);
        assert_eq!(f.char_spacing, Some(10.0));
        // Negative (tighter) — sample-1/3/4/5/14 style, e.g. -10 twips = -0.5 pt.
        let f = run_fmt_from(r#"<w:spacing w:val="-10"/>"#);
        assert_eq!(f.char_spacing, Some(-0.5));
        // val="0" is a real "no extra pitch" override, not absence.
        let f = run_fmt_from(r#"<w:spacing w:val="0"/>"#);
        assert_eq!(f.char_spacing, Some(0.0));
        // Absent ⇒ inherit (None).
        let f = run_fmt_from(r#"<w:b/>"#);
        assert_eq!(f.char_spacing, None);
    }

    #[test]
    fn fit_text_parses_twips_and_signed_id() {
        // ECMA-376 §17.3.2.14: w:val is ST_TwipsMeasure and remains in twips;
        // w:id is ST_DecimalNumber and may therefore be a large signed value.
        let f = run_fmt_from(r#"<w:fitText w:val="2400" w:id="-1431456512"/>"#);
        let fit_text = f.fit_text.expect("fitText");
        assert_eq!(fit_text.val, 2400.0);
        assert_eq!(fit_text.id.as_deref(), Some("-1431456512"));
    }

    #[test]
    fn fit_text_cascades_as_one_composite_property() {
        // ECMA-376 §17.3.2.14: a direct `<w:fitText>` ELEMENT replaces the
        // inherited one as a UNIT. Omitting `w:id` on the direct element means
        // "this run does not link with any other run" — the style level's id
        // must NOT survive underneath the direct val (val and id are one
        // composite property, not two independently-cascading axes).
        let mut base = run_fmt_from(r#"<w:fitText w:val="2400" w:id="7"/>"#);
        let over = run_fmt_from(r#"<w:fitText w:val="1200"/>"#);
        apply_run(&mut base, &over);
        assert_eq!(
            base.fit_text.as_ref().map(|fit_text| fit_text.val),
            Some(1200.0)
        );
        assert!(
            base.fit_text.and_then(|fit_text| fit_text.id).is_none(),
            "a direct fitText WITHOUT w:id must clear the inherited id"
        );
    }

    #[test]
    fn fit_text_id_preserves_arbitrary_precision_decimals() {
        // §17.18.10 ST_DecimalNumber is an XSD integer with NO range facet —
        // arbitrary precision. An id wider than i64 / an f64 mantissa must
        // still survive parsing so equality linking of consecutive runs works.
        let f = run_fmt_from(r#"<w:fitText w:val="2400" w:id="123456789012345678901234567890"/>"#);
        assert!(
            f.fit_text.and_then(|fit_text| fit_text.id).is_some(),
            "an arbitrary-precision w:id must be preserved for equality linking"
        );
    }

    #[test]
    fn fit_text_val_accepts_positive_universal_measures() {
        // ST_TwipsMeasure (§22.9.2.14) = ST_UnsignedDecimalNumber |
        // ST_PositiveUniversalMeasure — a plain number is twips, but the
        // suffixed forms ('1in', '12pt', '2cm', …) are equally valid.
        let inch = run_fmt_from(r#"<w:fitText w:val="1in" w:id="1"/>"#);
        assert_eq!(
            inch.fit_text.map(|fit_text| fit_text.val),
            Some(1440.0),
            "1in = 1440 twips"
        );
        let pt = run_fmt_from(r#"<w:fitText w:val="12pt" w:id="1"/>"#);
        assert_eq!(
            pt.fit_text.map(|fit_text| fit_text.val),
            Some(240.0),
            "12pt = 240 twips"
        );
    }

    #[test]
    fn run_spacing_does_not_collide_with_para_spacing() {
        // The paragraph `<w:spacing before/after/line>` (§17.3.1.33) must never
        // populate the run char_spacing, and the run `<w:spacing w:val>` must
        // never touch paragraph spacing. Different elements, same tag name.
        let p = para_fmt_from(
            r#"<w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>"#,
        );
        assert_eq!(p.space_before, Some(12.0));
        assert_eq!(p.space_after, Some(6.0));
        assert_eq!(
            p.run.char_spacing, None,
            "para spacing must not set run char_spacing"
        );
    }

    #[test]
    fn char_scale_parses_percent_bare_and_literal() {
        // §17.3.2.43 / ST_TextScale (§17.18.95): percentage of normal width.
        // Word writes a bare integer (sample-13 `w:val="80"`, sample-26 `"67"`).
        let f = run_fmt_from(r#"<w:w w:val="80"/>"#);
        assert_eq!(f.char_scale, Some(0.80));
        let f = run_fmt_from(r#"<w:w w:val="67"/>"#);
        assert!((f.char_scale.unwrap() - 0.67).abs() < 1e-9);
        // The `%` literal form is also valid ST_TextScale.
        let f = run_fmt_from(r#"<w:w w:val="200%"/>"#);
        assert_eq!(f.char_scale, Some(2.0));
        // Clamp to the [1%, 600%] range.
        let f = run_fmt_from(r#"<w:w w:val="1000"/>"#);
        assert_eq!(f.char_scale, Some(6.0));
        // Malformed / zero ⇒ ignored (leaves inheritance), not defaulted to 1.0.
        let f = run_fmt_from(r#"<w:w w:val="0"/>"#);
        assert_eq!(f.char_scale, None);
        let f = run_fmt_from(r#"<w:b/>"#);
        assert_eq!(f.char_scale, None);
    }

    #[test]
    fn position_parses_signed_half_points_to_pt() {
        // §17.3.2.24: val is ST_SignedHpsMeasure (half-points). Spec example
        // `<w:position w:val="24"/>` == 12 pt raised above the baseline.
        let f = run_fmt_from(r#"<w:position w:val="24"/>"#);
        assert_eq!(f.position, Some(12.0));
        // Negative = lowered (sample-11 uses -10 half-pt = -5 pt).
        let f = run_fmt_from(r#"<w:position w:val="-10"/>"#);
        assert_eq!(f.position, Some(-5.0));
        let f = run_fmt_from(r#"<w:position w:val="12pt"/>"#);
        assert_eq!(f.position, Some(12.0));
        assert_eq!(
            f.position_typography.and_then(|value| value.value),
            Some(12.0),
            "the private layout input uses the same universal-measure conversion"
        );
        let f = run_fmt_from(r#"<w:position w:val="-3.6mm"/>"#);
        let expected = -72.0 * 3.6 / 25.4;
        assert!((f.position.expect("signed universal projection") - expected).abs() < 1e-9);
        assert!(
            (f.position_typography
                .and_then(|value| value.value)
                .expect("signed universal private value")
                - expected)
                .abs()
                < 1e-9
        );
        for invalid in ["invalid", "1.5", "1e2", "+12pt"] {
            let f = run_fmt_from(&format!(r#"<w:position w:val="{invalid}"/>"#));
            assert_eq!(
                f.position,
                Some(0.0),
                "the stable public projection keeps its legacy invalid-value fallback"
            );
            let private = f.position_typography.expect("authored position fact");
            assert_eq!(
                private.status,
                crate::types::TypographyValueStatusWire::Invalid
            );
            assert_eq!(private.value, None);
        }
        let f = run_fmt_from(r#"<w:b/>"#);
        assert_eq!(f.position, None);
    }

    #[test]
    fn kern_parses_half_point_threshold() {
        // §17.3.2.19: val is ST_HpsMeasure (half-points) — the smallest font
        // size that gets kerning. Spec example `<w:kern w:val="28"/>` == 14 pt.
        let f = run_fmt_from(r#"<w:kern w:val="28"/>"#);
        assert_eq!(f.kerning, Some(14.0));
        // val="0" (common in Word documents) = kern at every size — presence,
        // not absence, so it must be Some(0.0) to keep kerning enabled.
        let f = run_fmt_from(r#"<w:kern w:val="0"/>"#);
        assert_eq!(f.kerning, Some(0.0));
        let f = run_fmt_from(r#"<w:kern w:val="12pt"/>"#);
        assert_eq!(f.kerning, Some(12.0));
        let f = run_fmt_from(r#"<w:kern w:val="0pt"/>"#);
        assert_eq!(f.kerning, Some(0.0));
        for invalid in [
            "invalid", "-0", "-0pt", "-2", "-2pt", "1.5", "1e2", "+12", "+12pt",
        ] {
            let f = run_fmt_from(&format!(r#"<w:kern w:val="{invalid}"/>"#));
            assert_eq!(
                f.kerning, None,
                "{invalid:?} must not become an authored kerning threshold"
            );

            let mut inherited = run_fmt_from(r#"<w:kern w:val="24"/>"#);
            apply_run(&mut inherited, &f);
            assert_eq!(
                inherited.kerning,
                Some(12.0),
                "{invalid:?} must not block an inherited kerning threshold"
            );
        }
        // Absent ⇒ None (inherit; the hierarchy default is kerning OFF).
        let f = run_fmt_from(r#"<w:b/>"#);
        assert_eq!(f.kerning, None);
    }

    #[test]
    fn char_metrics_merge_set_over_unset_and_inherit() {
        // The four axes follow the shared "set overrides, absent inherits" rule
        // in `apply_run` (used by both the style cascade and apply_direct_run).
        let mut base = run_fmt_from(
            r#"<w:fitText w:val="2400" w:id="-10"/><w:spacing w:val="200"/><w:w w:val="90"/><w:position w:val="4"/><w:kern w:val="24"/>"#,
        );
        // A later level that only re-sets spacing must keep the inherited w /
        // position / kern.
        let over = run_fmt_from(r#"<w:spacing w:val="-20"/>"#);
        apply_run(&mut base, &over);
        assert_eq!(base.char_spacing, Some(-1.0)); // overridden
        assert_eq!(
            base.fit_text.as_ref().map(|fit_text| fit_text.val),
            Some(2400.0)
        ); // inherited
        assert_eq!(
            base.fit_text
                .as_ref()
                .and_then(|fit_text| fit_text.id.as_deref()),
            Some("-10")
        ); // inherited
        assert_eq!(base.char_scale, Some(0.90)); // inherited
        assert_eq!(base.position, Some(2.0)); // inherited (8 half-pt = 4 pt? no: val=4 → 2pt)
        assert_eq!(base.kerning, Some(12.0)); // inherited (24 half-pt = 12 pt)

        let fit_override = run_fmt_from(r#"<w:fitText w:val="1200" w:id="-20"/>"#);
        apply_run(&mut base, &fit_override);
        assert_eq!(
            base.fit_text.as_ref().map(|fit_text| fit_text.val),
            Some(1200.0)
        );
        assert_eq!(
            base.fit_text.and_then(|fit_text| fit_text.id),
            Some("-20".to_string())
        );
    }
}
