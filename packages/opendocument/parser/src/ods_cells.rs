//! Table-cell styles (ODF Part 3 §16.2, §17.18 table-cell properties, §17.6
//! paragraph and §16.27 text properties) as SpreadsheetML cell formats
//! (ECMA-376 Part 1 §18.8.45 `xf`), with interned fonts, fills, borders and
//! number formats.

use crate::odf_props::{self, get, hex_color, length_pt, Style, StyleStore};
use crate::ods_number::{DataKind, DataStyles, FormatCode};
use std::collections::HashMap;
use xlsx_model::{Border, BorderEdge, CellXf, Dxf, Fill, Font, NumFmt, RunFont, Styles};

/// The first number-format id free for custom codes (§18.8.30).
const FIRST_CUSTOM_NUMFMT: u32 = 164;

pub struct XfBuilder {
    /// Every interned entry (and its key) is charged here before it is
    /// stored: the interner is the one place styles enter the model.
    budget: crate::ods::BudgetRef,
    styles: Styles,
    fonts: HashMap<String, u32>,
    fills: HashMap<String, u32>,
    borders: HashMap<String, u32>,
    num_fmts: HashMap<String, u32>,
    xfs: HashMap<String, u32>,
    /// xf of each (cell, row default, column default) style triple, keyed
    /// by the styles' identities (0: none), so lookups never allocate.
    by_style: HashMap<(usize, usize, usize), u32>,
    dxfs: HashMap<String, u32>,
    indent_pt: Vec<f64>,
    /// Data kind of each xf's number format (`None`: General).
    kinds: Vec<Option<DataKind>>,
    default_font: Font,
}

/// The styles that may give a table cell its format: its own
/// `table:style-name`, its row's and its column's
/// `table:default-cell-style-name` (the first present applies).
#[derive(Clone, Copy, Default)]
pub struct CellStyles<'a> {
    pub cell: Option<&'a str>,
    pub row: Option<&'a str>,
    pub column: Option<&'a str>,
}

/// The resolved properties of one cell style (or the default).
pub struct CellFormat {
    pub font: Font,
    pub fill: Fill,
    pub border: Border,
    pub num_fmt: Option<(FormatCode, DataKind)>,
    pub xf: CellXf,
    pub indent_pt: f64,
}

impl XfBuilder {
    /// A builder whose xf 0 is the `Default` cell style (§16.2: the style
    /// of cells that name none), as SpreadsheetML's xf 0 is the Normal
    /// format of unstyled cells.
    pub(crate) fn new(
        store: &StyleStore,
        data: &mut DataStyles,
        faces: &HashMap<String, String>,
        budget: crate::ods::BudgetRef,
    ) -> Result<Self, String> {
        let chain = store.chain("table-cell", Some("Default"));
        let default_font = font_of(&chain, faces, None);
        // The builder's own copy of the default font.
        budget.bytes(crate::retained_size(&default_font))?;
        let mut builder = Self {
            budget,
            styles: Styles::default(),
            fonts: HashMap::new(),
            fills: HashMap::new(),
            borders: HashMap::new(),
            num_fmts: HashMap::new(),
            xfs: HashMap::new(),
            by_style: HashMap::new(),
            dxfs: HashMap::new(),
            indent_pt: Vec::new(),
            kinds: Vec::new(),
            default_font: default_font.clone(),
        };
        // §18.8.21 / §18.8.20: fill 0 is none, fill 1 gray125 by convention.
        builder.intern_font(default_font)?;
        builder.intern_fill(Fill {
            pattern_type: "none".into(),
            ..Fill::default()
        })?;
        builder.intern_fill(Fill {
            pattern_type: "gray125".into(),
            ..Fill::default()
        })?;
        builder.intern_border(Border::default())?;
        let default_xf = builder.xf_for(store, data, faces, CellStyles::default())?;
        debug_assert_eq!(default_xf, 0);
        Ok(builder)
    }

    /// Admit one new interned entry's owned heap and its map key (callers
    /// admit its vector slot with [`crate::ods::Budget::grow`]).
    fn admit<T: ooxml_common::chart::RetainedBytes>(
        &self,
        value: &T,
        key_text: usize,
    ) -> Result<(), String> {
        self.budget.bytes(
            usize::try_from(value.heap_bytes())
                .unwrap_or(usize::MAX)
                .saturating_add(crate::ods::hash_entry_bytes::<String, u32>(key_text)),
        )
    }

    /// (family, size pt, bold, italic) of the default cell font: the host
    /// measures its digit width.
    pub fn default_font(&self) -> (String, f64, bool, bool) {
        (
            self.default_font
                .name
                .clone()
                .unwrap_or_else(|| "Liberation Sans".into()),
            if self.default_font.size > 0.0 {
                self.default_font.size
            } else {
                10.0
            },
            self.default_font.bold,
            self.default_font.italic,
        )
    }

    pub fn default_font_model(&self) -> &Font {
        &self.default_font
    }

    /// The xf of a table cell, memoized by its style's identity. The cell's
    /// style is its own, else its row's default cell style, else its
    /// column's (§19.615 / §19.619), else `Default`; its properties are that
    /// style's chain. LibreOffice reads a cell this way: it does not search
    /// a styled cell's missing properties (its number format above all) in
    /// the row and column defaults, and the files it writes rely on that
    /// (measured over the private corpus, where searching them changes 81
    /// displayed cells away from the original workbooks' text).
    pub fn xf_for(
        &mut self,
        store: &StyleStore,
        data: &mut DataStyles,
        faces: &HashMap<String, String>,
        styles: CellStyles<'_>,
    ) -> Result<u32, String> {
        let name = styles.cell.or(styles.row).or(styles.column);
        // An unknown name resolves to the family default alone.
        let style = store.find("table-cell", name.unwrap_or("Default"));
        let key = (
            style.map_or(0, |s| s as *const Style as usize),
            usize::from(name.is_some()),
            0,
        );
        if let Some(index) = self.by_style.get(&key) {
            return Ok(*index);
        }
        let found: Vec<&Style> = style.into_iter().collect();
        let chain = store.chain_through("table-cell", &found);
        let format = resolve(&chain, data, faces);
        let index = self.intern_format(format)?;
        self.budget
            .bytes(crate::ods::hash_entry_bytes::<(usize, usize, usize), u32>(
                0,
            ))?;
        self.by_style.insert(key, index);
        Ok(index)
    }

    /// What an empty cell with this xf paints: its fill and borders. Two
    /// xfs with the same signature draw the same blank cell.
    pub fn blank_signature(&self, xf: u32) -> (u32, u32) {
        self.styles
            .cell_xfs
            .get(xf as usize)
            .map_or((0, 0), |xf| (xf.fill_id, xf.border_id))
    }

    /// The number-format kind of an xf (for value-type defaults).
    pub fn kind(&self, xf: u32) -> Option<DataKind> {
        self.kinds.get(xf as usize).copied().flatten()
    }

    /// Intern `xf` with its number format replaced by `code` (a cell whose
    /// value type has no matching data style).
    pub fn with_num_fmt(&mut self, xf: u32, code: &str, kind: DataKind) -> Result<u32, String> {
        let Some(base) = self.styles.cell_xfs.get(xf as usize) else {
            return Ok(xf);
        };
        let mut copy = clone_xf(base);
        copy.num_fmt_id = self.intern_num_fmt(&FormatCode::plain(code))?;
        let indent = self.indent_pt.get(xf as usize).copied().unwrap_or(0.0);
        self.intern_xf(copy, indent, Some(kind))
    }

    fn intern_format(&mut self, format: CellFormat) -> Result<u32, String> {
        let mut xf = format.xf;
        xf.font_id = self.intern_font(format.font)?;
        xf.fill_id = self.intern_fill(format.fill)?;
        xf.border_id = self.intern_border(format.border)?;
        let kind = format.num_fmt.as_ref().map(|(_, kind)| *kind);
        xf.num_fmt_id = match &format.num_fmt {
            Some((format, _)) => self.intern_num_fmt(format)?,
            None => 0,
        };
        self.intern_xf(xf, format.indent_pt, kind)
    }

    fn intern_xf(
        &mut self,
        xf: CellXf,
        indent_pt: f64,
        kind: Option<DataKind>,
    ) -> Result<u32, String> {
        let key = format!(
            "{}|{}",
            serde_json::to_string(&xf).unwrap_or_default(),
            indent_pt
        );
        if let Some(found) = self.xfs.get(&key) {
            return Ok(*found);
        }
        // The xf, its key, and its indent and kind slots.
        self.admit(&xf, key.len())?;
        self.budget.grow(&self.styles.cell_xfs)?;
        self.budget.grow(&self.indent_pt)?;
        self.budget.grow(&self.kinds)?;
        let index = self.styles.cell_xfs.len() as u32;
        self.styles.cell_xfs.push(xf);
        self.indent_pt.push(indent_pt);
        self.kinds.push(kind);
        self.xfs.insert(key, index);
        Ok(index)
    }

    pub fn intern_font(&mut self, font: Font) -> Result<u32, String> {
        let key = serde_json::to_string(&font).unwrap_or_default();
        if let Some(found) = self.fonts.get(&key) {
            return Ok(*found);
        }
        self.admit(&font, key.len())?;
        self.budget.grow(&self.styles.fonts)?;
        self.styles.fonts.push(font);
        let index = self.styles.fonts.len() as u32 - 1;
        self.fonts.insert(key, index);
        Ok(index)
    }

    fn intern_fill(&mut self, fill: Fill) -> Result<u32, String> {
        let key = serde_json::to_string(&fill).unwrap_or_default();
        if let Some(found) = self.fills.get(&key) {
            return Ok(*found);
        }
        self.admit(&fill, key.len())?;
        self.budget.grow(&self.styles.fills)?;
        self.styles.fills.push(fill);
        let index = self.styles.fills.len() as u32 - 1;
        self.fills.insert(key, index);
        Ok(index)
    }

    fn intern_border(&mut self, border: Border) -> Result<u32, String> {
        let key = serde_json::to_string(&border).unwrap_or_default();
        if let Some(found) = self.borders.get(&key) {
            return Ok(*found);
        }
        self.admit(&border, key.len())?;
        self.budget.grow(&self.styles.borders)?;
        self.styles.borders.push(border);
        let index = self.styles.borders.len() as u32 - 1;
        self.borders.insert(key, index);
        Ok(index)
    }

    /// The number-format id of `format`: built-in ids for General and the
    /// built-in short date / date-time, an interned custom code otherwise
    /// (a General code that divides is a custom code: built-ins carry no
    /// divisors).
    pub fn intern_num_fmt(&mut self, format: &FormatCode) -> Result<u32, String> {
        if format.is_general() {
            return Ok(0);
        }
        if format.divisors.is_empty() {
            if let Some(id) = format
                .code
                .strip_prefix("#builtin:")
                .and_then(|id| id.parse::<u32>().ok())
            {
                return Ok(id);
            }
        }
        let key = format!("{}\u{0}{:?}", format.code, format.divisors);
        if let Some(id) = self.num_fmts.get(&key) {
            return Ok(*id);
        }
        let id = FIRST_CUSTOM_NUMFMT + self.styles.num_fmts.len() as u32;
        let num_fmt = NumFmt {
            num_fmt_id: id,
            format_code: crate::ods_number::literal_code(&format.code),
            section_divisors: format.divisors.clone(),
        };
        self.admit(&num_fmt, key.len())?;
        self.budget.grow(&self.styles.num_fmts)?;
        self.styles.num_fmts.push(num_fmt);
        self.num_fmts.insert(key, id);
        Ok(id)
    }

    /// A differential format for a conditional style (§16.3 map target).
    pub fn dxf_for(
        &mut self,
        store: &StyleStore,
        data: &mut DataStyles,
        faces: &HashMap<String, String>,
        style: &str,
    ) -> Result<u32, String> {
        if let Some(found) = self.dxfs.get(style) {
            return Ok(*found);
        }
        // Only the properties the conditional style sets itself (and its
        // parents below `Default`) differ from the cell's own format.
        let chain: Vec<&Style> = store
            .chain("table-cell", Some(style))
            .into_iter()
            .filter(|s| s.name != "Default" && !s.name.is_empty())
            .collect();
        let has = |key: &str| get(&chain, key).is_some();
        let font_set = [
            "text:fo:color",
            "text:fo:font-weight",
            "text:fo:font-style",
            "text:style:text-underline-style",
            "text:style:text-line-through-style",
        ]
        .iter()
        .any(|k| has(k));
        let full = store.chain("table-cell", Some(style));
        let format = resolve(&full, data, faces);
        let dxf = Dxf {
            font: font_set.then(|| format.font.clone()),
            fill: has("cell:fo:background-color").then(|| format.fill.clone()),
            border: [
                "cell:fo:border",
                "cell:fo:border-top",
                "cell:fo:border-bottom",
                "cell:fo:border-left",
                "cell:fo:border-right",
            ]
            .iter()
            .any(|k| has(k))
            .then(|| format.border.clone()),
            num_fmt: chain
                .iter()
                .any(|s| s.data_style.is_some())
                .then_some(format.num_fmt.as_ref())
                .flatten()
                .map(|(format, _)| {
                    Ok::<_, String>(NumFmt {
                        num_fmt_id: self.intern_num_fmt(format)?,
                        format_code: crate::ods_number::literal_code(&format.code),
                        section_divisors: format.divisors.clone(),
                    })
                })
                .transpose()?,
            font_toggles: font_set.then(|| xlsx_model::DxfFontToggles {
                bold: has("text:fo:font-weight").then_some(format.font.bold),
                italic: has("text:fo:font-style").then_some(format.font.italic),
                underline: has("text:style:text-underline-style").then_some(format.font.underline),
                strike: has("text:style:text-line-through-style").then_some(format.font.strike),
                automatic_color: None,
            }),
        };
        self.admit(&dxf, style.len())?;
        self.budget.grow(&self.styles.dxfs)?;
        let index = self.styles.dxfs.len() as u32;
        self.styles.dxfs.push(dxf);
        self.dxfs.insert(style.to_string(), index);
        Ok(index)
    }

    /// The run font of a `text:span` style over a cell's font.
    pub fn font_of_xf(&self, xf: u32) -> Font {
        self.styles
            .cell_xfs
            .get(xf as usize)
            .and_then(|xf| self.styles.fonts.get(xf.font_id as usize))
            .cloned()
            .unwrap_or_default()
    }

    pub fn finish(self) -> (Styles, Vec<f64>) {
        (self.styles, self.indent_pt)
    }
}

fn clone_xf(xf: &CellXf) -> CellXf {
    CellXf {
        font_id: xf.font_id,
        fill_id: xf.fill_id,
        border_id: xf.border_id,
        num_fmt_id: xf.num_fmt_id,
        align_h: xf.align_h.clone(),
        align_v: xf.align_v.clone(),
        wrap_text: xf.wrap_text,
        indent: xf.indent,
        text_rotation: xf.text_rotation,
        shrink_to_fit: xf.shrink_to_fit,
        reading_order: xf.reading_order,
        own_font_color: xf.own_font_color,
    }
}

/// A rich-text run's font from a resolved context font.
pub fn run_font_of(font: &Font) -> RunFont {
    RunFont {
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        strike: font.strike,
        size: Some(font.size),
        color: font.color.clone(),
        name: font.name.clone(),
        underline_style: font.underline_style.clone(),
        vert_align: font.vert_align.clone(),
        ..RunFont::default()
    }
}

/// Resolve one cell style chain.
pub fn resolve(
    chain: &[&Style],
    data: &mut DataStyles,
    faces: &HashMap<String, String>,
) -> CellFormat {
    let font = font_of(chain, faces, None);
    let fill = match get(chain, "cell:fo:background-color").and_then(hex_color) {
        Some(color) => Fill {
            pattern_type: "solid".into(),
            fg_color: Some(color.clone()),
            bg_color: Some(color),
            gradient: None,
        },
        None => Fill {
            pattern_type: "none".into(),
            ..Fill::default()
        },
    };
    let edge = |side: &str| -> Option<BorderEdge> {
        let value = get(chain, &format!("cell:fo:border-{side}"))
            .or_else(|| get(chain, "cell:fo:border"))?;
        border_edge(value)
    };
    let border = Border {
        left: edge("left"),
        right: edge("right"),
        top: edge("top"),
        bottom: edge("bottom"),
        diagonal_up: get(chain, "cell:style:diagonal-bl-tr").and_then(border_edge),
        diagonal_down: get(chain, "cell:style:diagonal-tl-br").and_then(border_edge),
        horizontal: None,
        vertical: None,
    };
    let num_fmt = chain
        .iter()
        .find_map(|s| s.data_style.as_deref())
        .and_then(|name| {
            if data.is_general(name) {
                None
            } else {
                data.code(name)
            }
        })
        .filter(|(format, _)| !format.is_general());
    // §20.250 style:text-align-source: `value-type` aligns by the value's
    // type (SpreadsheetML General); `fix` uses fo:text-align.
    let fixed = get(chain, "cell:style:text-align-source") != Some("value-type");
    let align_h = if get(chain, "cell:style:repeat-content") == Some("true") {
        Some("fill".to_string())
    } else if fixed {
        get(chain, "par:fo:text-align").and_then(|value| {
            Some(
                match value {
                    "start" | "left" => "left",
                    "center" => "center",
                    "end" | "right" => "right",
                    "justify" => {
                        if get(chain, "par:fo:text-align-last") == Some("justify") {
                            "distributed"
                        } else {
                            "justify"
                        }
                    }
                    _ => return None,
                }
                .to_string(),
            )
        })
    } else {
        None
    };
    let align_v = match get(chain, "cell:loext:vertical-justify") {
        Some("distribute") | Some("distributed") => Some("distributed".to_string()),
        Some("justify") => Some("justify".to_string()),
        _ => match get(chain, "cell:style:vertical-align") {
            Some("top") => Some("top".to_string()),
            Some("middle") => Some("center".to_string()),
            Some("bottom") => Some("bottom".to_string()),
            _ => None,
        },
    };
    let rotation = get(chain, "cell:style:rotation-angle")
        .and_then(|v| v.trim_end_matches("deg").trim().parse::<f64>().ok())
        .map(|a| a.rem_euclid(360.0).round() as u32)
        .and_then(|a| match a {
            0 => None,
            1..=90 => Some(a),
            270..=359 => Some(90 + (360 - a)),
            // Upside-down text has no SpreadsheetML angle; the nearest
            // representable reading direction is kept.
            91..=180 => Some(90),
            _ => Some(180),
        });
    let stacked = get(chain, "cell:style:direction") == Some("ttb");
    let indent_pt = if align_h.as_deref().is_none_or(|a| a == "left") {
        get(chain, "par:fo:margin-left")
            .and_then(length_pt)
            .unwrap_or(0.0)
            .max(0.0)
    } else {
        0.0
    };
    let writing = get(chain, "par:style:writing-mode");
    let reading_order = match writing {
        Some("rl-tb") | Some("rl") => Some(2),
        Some("lr-tb") | Some("lr") => None,
        _ => None,
    };
    let own_font_color = font.color.is_some();
    CellFormat {
        font,
        fill,
        border,
        num_fmt,
        indent_pt,
        xf: CellXf {
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            num_fmt_id: 0,
            align_h,
            align_v,
            wrap_text: get(chain, "cell:fo:wrap-option") == Some("wrap"),
            indent: None,
            text_rotation: if stacked { Some(255) } else { rotation },
            shrink_to_fit: get(chain, "cell:style:shrink-to-fit") == Some("true"),
            reading_order,
            own_font_color,
        },
    }
}

/// The font of a style chain over `base` (or the defaults).
pub fn font_of(chain: &[&Style], faces: &HashMap<String, String>, base: Option<&Font>) -> Font {
    let mut font = base.cloned().unwrap_or(Font {
        size: 10.0,
        ..Font::default()
    });
    if let Some(name) = get(chain, "text:style:font-name") {
        font.name = Some(faces.get(name).cloned().unwrap_or_else(|| name.to_string()));
    } else if let Some(family) = get(chain, "text:fo:font-family") {
        font.name = Some(odf_props::unquote_family(family));
    }
    if let Some(size) = get(chain, "text:fo:font-size") {
        if let Some(pt) = length_pt(size) {
            font.size = pt;
        } else if let Some(fraction) = odf_props::percent(size) {
            font.size *= fraction;
        }
    }
    if let Some(weight) = get(chain, "text:fo:font-weight") {
        font.bold = weight == "bold" || weight.parse::<u32>().is_ok_and(|w| w >= 600);
    }
    if let Some(style) = get(chain, "text:fo:font-style") {
        font.italic = matches!(style, "italic" | "oblique");
    }
    if let Some(underline) = get(chain, "text:style:text-underline-style") {
        font.underline = underline != "none";
        font.underline_style = (font.underline
            && get(chain, "text:style:text-underline-type") == Some("double"))
        .then(|| "double".to_string());
    }
    if let Some(strike) = get(chain, "text:style:text-line-through-style") {
        font.strike = strike != "none";
    }
    // The nearest style that sets either the colour or the automatic
    // (window) colour decides (§20.209 / §20.391, per-attribute cascade
    // read as LibreOffice reads it: one font colour property).
    for style in chain {
        let window = style.properties.get("text:style:use-window-font-color");
        let color = style.properties.get("text:fo:color");
        if window.is_none() && color.is_none() {
            continue;
        }
        if window.is_some_and(|w| w == "true") && color.is_none() {
            font.color = None;
        } else if let Some(color) = color.and_then(|c| hex_color(c)) {
            font.color = Some(color);
        } else if window.is_some_and(|w| w == "true") {
            font.color = None;
        }
        break;
    }
    if let Some(position) = get(chain, "text:style:text-position") {
        let first = position.split_whitespace().next().unwrap_or("");
        font.vert_align = match first {
            "super" => Some("superscript".into()),
            "sub" => Some("subscript".into()),
            value => match odf_props::percent(value) {
                Some(v) if v > 0.0 => Some("superscript".into()),
                Some(v) if v < 0.0 => Some("subscript".into()),
                _ => None,
            },
        };
    }
    font
}

/// `0.74pt solid #000000` (§20.183 fo:border) as a SpreadsheetML edge.
/// Widths follow LibreOffice's SpreadsheetML line weights: hair 0.05pt,
/// thin 0.75pt, medium 1.75pt, thick 2.5pt.
pub fn border_edge(value: &str) -> Option<BorderEdge> {
    let mut width = None;
    let mut style = None;
    let mut color = None;
    for token in value.split_whitespace() {
        if let Some(c) = hex_color(token) {
            color = Some(c);
        } else if let Some(pt) = length_pt(token) {
            width = Some(pt);
        } else {
            match token {
                "thin" => width = Some(0.75),
                "medium" => width = Some(1.75),
                "thick" => width = Some(2.5),
                other => style = Some(other.to_string()),
            }
        }
    }
    let style = style.unwrap_or_else(|| "solid".into());
    if style == "none" || style == "hidden" || width == Some(0.0) {
        return None;
    }
    let width = width.unwrap_or(0.75);
    let weight = if width <= 0.25 {
        0
    } else if width < 1.25 {
        1
    } else if width < 2.2 {
        2
    } else {
        3
    };
    let name = match style.as_str() {
        "double" | "double-thin" | "thin-thick" | "thick-thin" | "inset" | "outset" | "groove"
        | "ridge" => "double",
        "dotted" => {
            if weight >= 2 {
                "mediumDashDotDot"
            } else {
                "dotted"
            }
        }
        "dashed" | "fine-dashed" | "long-dash" => {
            if weight >= 2 {
                "mediumDashed"
            } else {
                "dashed"
            }
        }
        "dash-dot" => {
            if weight >= 2 {
                "mediumDashDot"
            } else {
                "dashDot"
            }
        }
        "dash-dot-dot" => {
            if weight >= 2 {
                "mediumDashDotDot"
            } else {
                "dashDotDot"
            }
        }
        _ => match weight {
            0 => "hair",
            1 => "thin",
            2 => "medium",
            _ => "thick",
        },
    };
    Some(BorderEdge {
        style: name.into(),
        color,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn border_weights_and_styles() {
        let edge = |v: &str| border_edge(v).map(|e| (e.style, e.color));
        assert_eq!(
            edge("0.74pt solid #000000"),
            Some(("thin".into(), Some("000000".into())))
        );
        assert_eq!(edge("1.76pt solid #ff0000").unwrap().0, "medium");
        assert_eq!(edge("2.49pt solid #000000").unwrap().0, "thick");
        assert_eq!(edge("0.05pt solid #000000").unwrap().0, "hair");
        assert_eq!(edge("1.76pt double-thin #3f3f3f").unwrap().0, "double");
        assert_eq!(
            edge("1.76pt dash-dot-dot #000000").unwrap().0,
            "mediumDashDotDot"
        );
        assert_eq!(edge("none"), None);
    }
}
