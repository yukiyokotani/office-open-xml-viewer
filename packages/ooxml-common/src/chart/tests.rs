use super::*;
use roxmltree::Document;

fn root_of(xml: &str) -> Document<'_> {
    Document::parse(xml).expect("parse fixture")
}

// Test resolver: returns the schemeClr@val verbatim, or the srgbClr@val
// uppercased. Just enough to drive `extract_data_label_font_color`.
struct StubResolver;
impl ColorResolver for StubResolver {
    fn resolve_solid_fill(&self, node: Node) -> Option<String> {
        for c in node.children().filter(|n| n.is_element()) {
            match c.tag_name().name() {
                "srgbClr" => return c.attribute("val").map(|v| v.to_uppercase()),
                "schemeClr" => return c.attribute("val").map(|v| v.to_string()),
                _ => {}
            }
        }
        None
    }
}

const C_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

// Minimal theme-aware resolver for chart entry-point tests: resolves
// `<a:srgbClr>` verbatim (uppercased, matching real resolvers' hex
// normalization) and `<a:schemeClr>` against a small fixed table covering
// the slots real decks use for chart text/borders. Also overrides the
// theme major/minor Latin font hooks so CH10 theme-fallback fields can be
// exercised without pulling in a crate's full theme parser.
struct FixtureResolver;

impl ColorResolver for FixtureResolver {
    fn resolve_solid_fill(&self, node: Node) -> Option<String> {
        let c = node
            .children()
            .find(|n| n.is_element() && matches!(n.tag_name().name(), "srgbClr" | "schemeClr"))?;
        match c.tag_name().name() {
            "srgbClr" => c.attribute("val").map(|v| v.to_uppercase()),
            "schemeClr" => match c.attribute("val")? {
                "accent1" => Some("4472C4".to_string()),
                "accent2" => Some("ED7D31".to_string()),
                "accent3" => Some("A5A5A5".to_string()),
                "tx1" | "dk1" => Some("000000".to_string()),
                "bg1" | "lt1" => Some("FFFFFF".to_string()),
                _ => None,
            },
            _ => None,
        }
    }

    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        match name {
            "accent1" => Some("4472C4".to_string()),
            "accent2" => Some("ED7D31".to_string()),
            "accent3" => Some("A5A5A5".to_string()),
            "tx1" | "dk1" => Some("000000".to_string()),
            "bg1" | "lt1" => Some("FFFFFF".to_string()),
            _ => None,
        }
    }

    fn theme_major_font_latin(&self) -> Option<String> {
        Some("Calibri Light".to_string())
    }

    fn theme_minor_font_latin(&self) -> Option<String> {
        Some("Calibri".to_string())
    }

    fn resolve_series_accent(&self, idx: usize) -> Option<String> {
        // Cycle a 6-accent palette exactly like the docx resolver so chartEx
        // box/sunburst tests can assert the branch/series colors.
        const ACCENTS: [&str; 6] = ["5B9BD5", "ED7D31", "A5A5A5", "FFC000", "4472C4", "70AD47"];
        Some(ACCENTS[idx % 6].to_string())
    }
}

struct XlsxCompatibilityFixtureResolver;

impl ColorResolver for XlsxCompatibilityFixtureResolver {
    fn resolve_solid_fill(&self, node: Node) -> Option<String> {
        FixtureResolver.resolve_solid_fill(node)
    }

    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        FixtureResolver.resolve_scheme_color(name)
    }

    fn resolve_series_accent(&self, idx: usize) -> Option<String> {
        FixtureResolver.resolve_series_accent(idx)
    }

    fn implicit_outline_only_negative_column_style(&self) -> bool {
        true
    }
}

struct DarkRedFixtureResolver;

impl ColorResolver for DarkRedFixtureResolver {
    fn resolve_solid_fill(&self, node: Node) -> Option<String> {
        FixtureResolver.resolve_solid_fill(node)
    }

    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        match name {
            "tx1" | "dk1" => Some("240000".to_string()),
            _ => FixtureResolver.resolve_scheme_color(name),
        }
    }

    fn resolve_series_accent(&self, idx: usize) -> Option<String> {
        FixtureResolver.resolve_series_accent(idx)
    }
}

struct WhiteChartFixtureResolver;

impl ColorResolver for WhiteChartFixtureResolver {
    fn resolve_solid_fill(&self, node: Node) -> Option<String> {
        FixtureResolver.resolve_solid_fill(node)
    }

    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        FixtureResolver.resolve_scheme_color(name)
    }

    fn resolve_series_accent(&self, idx: usize) -> Option<String> {
        FixtureResolver.resolve_series_accent(idx)
    }

    fn default_chart_bg(&self) -> Option<String> {
        Some("FFFFFF".to_string())
    }

    fn default_plot_area_bg(&self) -> Option<String> {
        Some("FFFFFF".to_string())
    }
}

struct FormatSchemeFixtureResolver {
    format_scheme: crate::theme::ThemeFormatScheme,
}

impl ColorResolver for FormatSchemeFixtureResolver {
    fn resolve_solid_fill(&self, node: Node) -> Option<String> {
        FixtureResolver.resolve_solid_fill(node)
    }

    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        FixtureResolver.resolve_scheme_color(name)
    }

    fn resolve_series_accent(&self, idx: usize) -> Option<String> {
        FixtureResolver.resolve_series_accent(idx)
    }

    fn theme_format_scheme(&self) -> Option<&crate::theme::ThemeFormatScheme> {
        Some(&self.format_scheme)
    }

    fn default_chart_bg(&self) -> Option<String> {
        Some("FFFFFF".to_string())
    }
}

fn chart_space_of(xml: &str) -> Document<'_> {
    Document::parse(xml).expect("parse chartSpace fixture")
}

// These tests exercise the Word host's evidenced dark-text extension;
// the generic resolver deliberately leaves that extension disabled.
struct WordContrastFixtureResolver;
impl ColorResolver for WordContrastFixtureResolver {
    fn resolve_solid_fill(&self, node: Node) -> Option<String> {
        FixtureResolver.resolve_solid_fill(node)
    }
    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        FixtureResolver.resolve_scheme_color(name)
    }
    fn office_dark_text_contrast_applies(&self, style: u8) -> bool {
        (41..=48).contains(&style)
    }
    fn office_dark_title_contrast_applies(&self, style: u8) -> bool {
        (41..=48).contains(&style)
    }
}

// ─────────────────────────────────────────────────────────────────────
// Direct unit tests for the extractors moved from the xlsx parser into
// this shared module. These call the functions themselves (not through
// `parse_chart_part`) so a regression in one is pinpointed rather than
// surfacing only as a diff in a much larger golden `ChartModel`.
// ─────────────────────────────────────────────────────────────────────

// ── CT_Boolean bare-element defaults (issue #806) ───────────────────────
//
// dml-chart.xsd defines `CT_Boolean` with `val` `default="true"`. Every
// element typed `CT_Boolean` (delete / show* / showLeaderLines / marker /
// noEndCap / autoTitleDeleted / …) therefore means TRUE when the element is
// PRESENT but the `val` attribute is OMITTED. Office always writes an
// explicit `val="0|1"`, so these probes drive the bare form that only a
// hand-authored / third-party file emits — the latent divergence the issue
// flags. `val="0"` must still read false, and an ABSENT element keeps its
// own semantic default (off).

// ── `parse_chartex_part` direct contract tests ──────────────────────────
//
// The chartEx counterpart to the `parse_chart_part_*` tests above. These
// call the shared `parse_chartex_part` (not the pptx wrapper) so a
// regression in the chartEx structure walk — categories, values, subtotal
// indices, series/per-label colours (resolved through the `ColorResolver`),
// axis visibility, gap-width fraction→percent conversion, and the theme
// fallback faces — is pinpointed here. `FixtureResolver` resolves
// `<a:schemeClr val="accent1">`→`4472C4`, `tx1`→`000000`, and reports
// `Calibri Light` / `Calibri` as the theme major/minor faces.

const CX_NS: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const CS_NS: &str = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";

struct FormulaOnlyTreemapResolver;

impl ChartReferenceResolver for FormulaOnlyTreemapResolver {
    fn resolve_strings(&mut self, _formula: &str) -> Option<Vec<String>> {
        None
    }

    fn resolve_numbers(&mut self, formula: &str) -> Option<Vec<Option<f64>>> {
        (formula == "_xlchart.v1.2").then(|| vec![Some(50.0), Some(30.0), Some(20.0)])
    }

    fn resolve_number_format(&mut self, formula: &str) -> Option<String> {
        (formula == "_xlchart.v1.2").then(|| "#,##0".to_string())
    }

    fn resolve_string_levels(&mut self, formula: &str) -> Option<Vec<Vec<String>>> {
        (formula == "_xlchart.v1.0").then(|| {
            vec![
                vec!["North".into(), "South".into(), "East".into()],
                vec!["Americas".into(), "Americas".into(), "Asia".into()],
            ]
        })
    }
}

struct FormulaOnlyFlatResolver;

impl ChartReferenceResolver for FormulaOnlyFlatResolver {
    fn resolve_strings(&mut self, formula: &str) -> Option<Vec<String>> {
        (formula == "_xlchart.name").then(|| vec!["Authored series".to_string()])
    }

    fn resolve_numbers(&mut self, formula: &str) -> Option<Vec<Option<f64>>> {
        (formula == "_xlchart.values").then(|| vec![Some(3.0), Some(2.0), Some(1.0)])
    }
}

// ── CH15: chartEx structured layout parsing ──────────────────────────────

struct FormulaOnlyBoxResolver;

impl ChartReferenceResolver for FormulaOnlyBoxResolver {
    fn resolve_strings(&mut self, formula: &str) -> Option<Vec<String>> {
        (formula == "_xlchart.name").then(|| vec!["Adaptation".to_string()])
    }

    fn resolve_numbers(&mut self, formula: &str) -> Option<Vec<Option<f64>>> {
        match formula {
            "_xlchart.v1.1" => Some(vec![Some(1.0), Some(2.0), Some(3.0)]),
            "_xlchart.v1.3" => Some(vec![Some(4.0), Some(5.0), Some(6.0)]),
            _ => None,
        }
    }
}

// ── CH13: 3D flattening / stock / ofPie type detection ───────────────────

// Build a minimal `<c:chartSpace>` whose plot area holds a single
// chart-group element (`group_xml`) with one series. Used by the CH13
// type-detection probes below.
fn chart_space_with_group(group_xml: &str) -> String {
    format!(
        r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart><c:plotArea>
                {group_xml}
                <c:catAx><c:axId val="1"/><c:axPos val="b"/></c:catAx>
                <c:valAx><c:axId val="2"/><c:axPos val="l"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>"#
    )
}

const CH13_SER: &str = r#"<c:ser><c:idx val="0"/>
        <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:cat>
        <c:val><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt></c:numCache></c:val>
      </c:ser>"#;

// A resolver that DOES supply the default series accent palette (like the
// real docx/xlsx resolvers), used to pin the §21.2.2.227 `<c:varyColors>`
// per-slice accent fill. `FixtureResolver` returns `None` for accents so it
// cannot exercise this path.
struct AccentResolver;
impl ColorResolver for AccentResolver {
    fn resolve_solid_fill(&self, node: Node) -> Option<String> {
        node.children()
            .find(|n| n.is_element() && n.tag_name().name() == "srgbClr")
            .and_then(|n| attr(&n, "val"))
            .map(|v| v.to_uppercase())
    }
    fn resolve_series_accent(&self, idx: usize) -> Option<String> {
        // Six-accent cycle, matching `theme.accent[(idx % 6) + 1]`.
        const ACCENTS: [&str; 6] = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"];
        Some(ACCENTS[idx % 6].to_string())
    }
}

// A named single series in `<c:tx>` — reused by the auto-title tests. `idx`
// distinguishes the two series in the multi-series fixture.
fn named_ser(idx: u32, name: &str) -> String {
    format!(
        r#"<c:ser><c:idx val="{idx}"/>
              <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>{name}</c:v></c:pt></c:strCache></c:strRef></c:tx>
              <c:cat><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:cat>
              <c:val><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt></c:numCache></c:val>
            </c:ser>"#
    )
}

// A `<c:chart>` with an optional `<c:autoTitleDeleted val=…>`, an empty
// `<c:title>` frame, and the given series in a bar plot area. Models the
// observed auto-title shape: the title frame exists but has no `<c:tx>`, so
// the synthesized title comes from the sole series name.
fn chart_space_auto_title(auto_title_deleted: Option<&str>, sers: &str) -> String {
    let atd = auto_title_deleted
        .map(|v| format!(r#"<c:autoTitleDeleted val="{v}"/>"#))
        .unwrap_or_default();
    format!(
        r#"<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}">
              <c:chart>
                <c:title><c:txPr><a:bodyPr/><a:lstStyle/><a:p/></c:txPr></c:title>
                {atd}
                <c:plotArea>
                  <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>{sers}</c:barChart>
                  <c:catAx><c:axId val="1"/><c:axPos val="b"/></c:catAx>
                  <c:valAx><c:axId val="2"/><c:axPos val="l"/></c:valAx>
                </c:plotArea>
              </c:chart>
            </c:chartSpace>"#
    )
}

struct FormulaResolver;

impl ChartReferenceResolver for FormulaResolver {
    fn resolve_strings(&mut self, formula: &str) -> Option<Vec<String>> {
        Some(match formula {
            "Name" => vec!["Resolved series".into()],
            "X" => vec!["1".into(), "2".into()],
            "CachedCats" => vec!["live value must not win".into()],
            _ => return None,
        })
    }

    fn resolve_numbers(&mut self, formula: &str) -> Option<Vec<Option<f64>>> {
        Some(match formula {
            "Y" => vec![Some(10.0), Some(20.0)],
            "Size" => vec![Some(3.0), Some(5.0)],
            _ => return None,
        })
    }
}

struct CountingCategoryResolver {
    string_calls: usize,
}

impl ChartReferenceResolver for CountingCategoryResolver {
    fn resolve_strings(&mut self, formula: &str) -> Option<Vec<String>> {
        self.string_calls += 1;
        (formula == "Cats").then(|| vec!["A".into(), "B".into(), "C".into()])
    }

    fn resolve_numbers(&mut self, _formula: &str) -> Option<Vec<Option<f64>>> {
        None
    }
}

struct HiddenSourceResolver;

impl ChartReferenceResolver for HiddenSourceResolver {
    fn resolve_strings(&mut self, _formula: &str) -> Option<Vec<String>> {
        None
    }

    fn resolve_numbers(&mut self, _formula: &str) -> Option<Vec<Option<f64>>> {
        None
    }

    fn resolve_hidden(&mut self, formula: &str) -> Option<Vec<bool>> {
        match formula {
            "Cats" => Some(vec![false, true, false, false]),
            "Values" => Some(vec![false, false, true, false]),
            _ => None,
        }
    }
}

#[path = "axis_tests.rs"]
mod axis_tests;
#[path = "cache_tests.rs"]
mod cache_tests;
#[path = "chartex_tests.rs"]
mod chartex_tests;
#[path = "classic_tests.rs"]
mod classic_tests;
#[path = "labels_tests.rs"]
mod labels_tests;
#[path = "model_tests.rs"]
mod model_tests;
#[path = "style_tests.rs"]
mod style_tests;
