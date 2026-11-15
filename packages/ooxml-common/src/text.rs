//! Shared DrawingML text-body helpers used by the pptx and xlsx parsers.
//!
//! Both hosts embed the same DrawingML text grammar: a paragraph's line spacing
//! rides in `<a:lnSpc>` (ECMA-376 §21.1.2.2.5) and a text body's autofit mode in
//! a `<a:bodyPr>` child (`<a:spAutoFit>` / `<a:normAutofit>` / `<a:noAutofit>`,
//! §21.1.2.1.1-.4). These leaves were previously read inline in each parser with
//! byte-identical serde shapes; sharing the type + leaf parse keeps the two
//! formats' line-spacing / autofit handling identical.
//!
//! Following the `parse_src_rect` precedent in [`crate::blip`], each caller
//! *locates* the node (`<a:lnSpc>` / `<a:bodyPr>`) and keeps its own inheritance
//! and defaults; the shared function only parses the located leaf.

use roxmltree::Node;
use serde::{Deserialize, Serialize};

use crate::units::coordinate32_to_emu;

/// Paragraph line spacing (`<a:lnSpc>`, ECMA-376 §21.1.2.2.5): a percentage of
/// the natural single line, or an absolute per-line height in points. The serde
/// shape (`{"type":"pct","val":..}` / `{"type":"pts","val":..}`) mirrors core's
/// TS `SpaceLine`, so the pptx and xlsx JSON stays byte-identical.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[cfg_attr(feature = "schema", derive(schemars::JsonSchema))]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum SpaceLine {
    /// `<a:spcPct@val>` (e.g. 100000 = 100%, 150000 = 150%).
    Pct { val: f64 },
    /// `<a:spcPts@val>` in points (the raw ST_TextSpacingPoint hundredths-of-a-
    /// point value is divided by 100 by this parser, matching core / pptx).
    Pts { val: f64 },
}

/// Parse a located `<a:lnSpc>` node (ECMA-376 §21.1.2.2.5). A `<a:spcPct@val>`
/// child yields the raw percentage (e.g. `150000`); otherwise a `<a:spcPts@val>`
/// child yields points (its raw hundredths-of-a-point `@val` divided by 100).
/// Returns `None` when neither child carries a parseable `@val`. The caller
/// passes the located `<a:lnSpc>` node and keeps its own inheritance
/// (lstStyle / master / body defaults).
pub fn parse_lnspc(ln_spc: Node<'_, '_>) -> Option<SpaceLine> {
    if let Some(v) = ln_spc
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spcPct")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok())
    {
        return Some(SpaceLine::Pct { val: v });
    }
    ln_spc
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spcPts")
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok())
        .map(|v| SpaceLine::Pts { val: v / 100.0 })
}

/// Parse a located `<a:bodyPr>` node's autofit child (ECMA-376 §21.1.2.1.1-.4).
/// Returns `None` when the `<a:bodyPr>` has *no* autofit child, so the caller
/// applies its own default (pptx: the inherited placeholder value; both formats
/// then fall back to `none`).
/// Otherwise returns `Some((auto_fit, font_scale, ln_spc_reduction))`:
///
/// - `<a:spAutoFit>` → `("sp", None, None)`
/// - `<a:normAutofit fontScale? lnSpcReduction?>` → `("norm", fontScale?, lnSpcReduction?)`,
///   each scale being the raw ST_Percentage (1000ths of a percent) divided by
///   100000 to a fraction (e.g. `62500` → `0.625`)
/// - `<a:noAutofit>` → `("none", None, None)`
///
/// The OOXML spelling is `normAutofit` (lowercase `f`).
pub fn parse_autofit(body_pr: Node<'_, '_>) -> Option<(String, Option<f64>, Option<f64>)> {
    for c in body_pr.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "spAutoFit" => return Some(("sp".to_owned(), None, None)),
            "normAutofit" => {
                let font_scale = c
                    .attribute("fontScale")
                    .and_then(|v| v.parse::<f64>().ok())
                    .map(|v| v / 100_000.0);
                let ln_spc_reduction = c
                    .attribute("lnSpcReduction")
                    .and_then(|v| v.parse::<f64>().ok())
                    .map(|v| v / 100_000.0);
                return Some(("norm".to_owned(), font_scale, ln_spc_reduction));
            }
            "noAutofit" => return Some(("none".to_owned(), None, None)),
            _ => {}
        }
    }
    None
}

/// ECMA-376 §21.1.2.1.1 default text insets, in EMU. The left/right inset
/// defaults to 91440 EMU (0.1 in = 7.2 pt = 9.6 px @96dpi); the top/bottom inset
/// to 45720 EMU (0.05 in = 3.6 pt = 4.8 px). Single source of truth for the
/// `CT_TextBodyProperties` inset defaults the parsers apply when `<a:bodyPr>`
/// omits `lIns`/`rIns` / `tIns`/`bIns`.
pub const DEFAULT_INS_LR_EMU: i64 = 91_440;
/// See [`DEFAULT_INS_LR_EMU`].
pub const DEFAULT_INS_TB_EMU: i64 = 45_720;

/// The subset of `<a:bodyPr>` (ECMA-376 §21.1.2.1.1 `CT_TextBodyProperties`)
/// that the pptx and xlsx parsers share: the vertical anchor, wrap mode, text
/// direction, the four text insets (EMU), and the autofit mode (+ its stored
/// normAutofit scales). Multi-column attributes (`numCol` / `spcCol` /
/// `rtlCol`) are pptx-only and stay in that parser.
///
/// Every field is already defaulted — [`parse_body_pr`] layers the bodyPr's own
/// attributes over a caller-supplied [`BodyPrDefaults`] (which carries each
/// host's inheritance resolution). No serde: each parser
/// maps these onto its own model type (pptx `TextBody`, xlsx `ShapeText`).
#[derive(Debug, Clone, PartialEq)]
pub struct BodyPr {
    /// `@anchor` — vertical alignment (`t`/`ctr`/`b`/`just`/`dist`).
    pub anchor: String,
    /// `@wrap` — `square` (wrap to width) or `none`.
    pub wrap: String,
    /// `@vert` — text direction (`horz`/`vert`/`vert270`/`eaVert`/…).
    pub vert: String,
    /// `@lIns` — left inset (EMU).
    pub l_ins: i64,
    /// `@tIns` — top inset (EMU).
    pub t_ins: i64,
    /// `@rIns` — right inset (EMU).
    pub r_ins: i64,
    /// `@bIns` — bottom inset (EMU).
    pub b_ins: i64,
    /// Autofit mode (`sp`/`norm`/`none`) from the autofit child (or the
    /// default when there is no child). See [`parse_autofit`].
    pub auto_fit: String,
    /// `<a:normAutofit@fontScale>` as a fraction (62500 → 0.625). `None` unless
    /// a `<a:normAutofit>` child carries it.
    pub font_scale: Option<f64>,
    /// `<a:normAutofit@lnSpcReduction>` as a fraction. `None` unless present.
    pub ln_spc_reduction: Option<f64>,
}

/// Defaults for each `<a:bodyPr>` field, applied by [`parse_body_pr`] when the
/// attribute (or autofit child) is absent. The caller pre-resolves these from
/// its own inheritance chain (e.g. pptx: inherited placeholder anchor → spec
/// default; xlsx: just the spec default). Use
/// [`BodyPrDefaults::spec`] for the bare ECMA-376 defaults.
#[derive(Debug, Clone, PartialEq)]
pub struct BodyPrDefaults {
    pub anchor: String,
    pub wrap: String,
    pub vert: String,
    pub l_ins: i64,
    pub t_ins: i64,
    pub r_ins: i64,
    pub b_ins: i64,
    /// Autofit mode used when the bodyPr has no autofit child.
    pub auto_fit: String,
}

impl BodyPrDefaults {
    /// The bare ECMA-376 §21.1.2.1.1 defaults: anchor `t`, wrap `square`, vert
    /// `horz`, insets [`DEFAULT_INS_LR_EMU`] / [`DEFAULT_INS_TB_EMU`], autofit
    /// `none`. A host with no inheritance / theme layer (xlsx) uses this
    /// directly; pptx overrides individual fields with its resolved values.
    pub fn spec() -> Self {
        Self {
            anchor: "t".to_owned(),
            wrap: "square".to_owned(),
            vert: "horz".to_owned(),
            l_ins: DEFAULT_INS_LR_EMU,
            t_ins: DEFAULT_INS_TB_EMU,
            r_ins: DEFAULT_INS_LR_EMU,
            b_ins: DEFAULT_INS_TB_EMU,
            auto_fit: "none".to_owned(),
        }
    }
}

/// Unqualified (no-namespace) `ST_Coordinate32` attribute on `<a:bodyPr>`.
fn attr_coordinate32(node: Node<'_, '_>, local: &str) -> Option<i64> {
    node.attributes()
        .find(|a| a.name() == local && a.namespace().is_none())
        .and_then(|a| coordinate32_to_emu(a.value()))
}

/// Unqualified (no-namespace) string attribute on `<a:bodyPr>`.
fn attr_str(node: Node<'_, '_>, local: &str) -> Option<String> {
    node.attributes()
        .find(|a| a.name() == local && a.namespace().is_none())
        .map(|a| a.value().to_owned())
}

/// Parse a located `<a:bodyPr>` node's shared attributes over `defaults`
/// (ECMA-376 §21.1.2.1.1). Each attribute uses the bodyPr's own value when
/// present, else the corresponding `defaults` field; the autofit mode comes
/// from the autofit child ([`parse_autofit`]) when present, else
/// `defaults.auto_fit` (with `font_scale` / `ln_spc_reduction` only from a
/// `<a:normAutofit>` child). The caller passes the located `<a:bodyPr>` and its
/// pre-resolved [`BodyPrDefaults`], keeping host-specific inheritance out of the
/// shared grammar.
pub fn parse_body_pr(body_pr: Node<'_, '_>, defaults: &BodyPrDefaults) -> BodyPr {
    let (auto_fit, font_scale, ln_spc_reduction) =
        parse_autofit(body_pr).unwrap_or_else(|| (defaults.auto_fit.clone(), None, None));
    BodyPr {
        anchor: attr_str(body_pr, "anchor").unwrap_or_else(|| defaults.anchor.clone()),
        wrap: attr_str(body_pr, "wrap").unwrap_or_else(|| defaults.wrap.clone()),
        vert: attr_str(body_pr, "vert").unwrap_or_else(|| defaults.vert.clone()),
        l_ins: attr_coordinate32(body_pr, "lIns").unwrap_or(defaults.l_ins),
        t_ins: attr_coordinate32(body_pr, "tIns").unwrap_or(defaults.t_ins),
        r_ins: attr_coordinate32(body_pr, "rIns").unwrap_or(defaults.r_ins),
        b_ins: attr_coordinate32(body_pr, "bIns").unwrap_or(defaults.b_ins),
        auto_fit,
        font_scale,
        ln_spc_reduction,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use roxmltree::Document;

    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// `<a:spcPct@val>` → a percent SpaceLine carrying the raw `@val`.
    #[test]
    fn parse_lnspc_reads_pct_raw_val() {
        let xml = format!(r#"<a:lnSpc xmlns:a="{A_NS}"><a:spcPct val="150000"/></a:lnSpc>"#);
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            parse_lnspc(doc.root_element()),
            Some(SpaceLine::Pct { val: 150000.0 })
        );
    }

    /// `<a:spcPts@val>` → a points SpaceLine (raw hundredths of a point / 100).
    #[test]
    fn parse_lnspc_reads_pts_hundredths_of_point() {
        let xml = format!(r#"<a:lnSpc xmlns:a="{A_NS}"><a:spcPts val="1800"/></a:lnSpc>"#);
        let doc = Document::parse(&xml).unwrap();
        // 1800 hundredths of a point → 18 pt.
        assert_eq!(
            parse_lnspc(doc.root_element()),
            Some(SpaceLine::Pts { val: 18.0 })
        );
    }

    /// spcPct takes precedence over spcPts, and an empty/absent-val lnSpc → None.
    #[test]
    fn parse_lnspc_none_when_no_parseable_child() {
        let empty = format!(r#"<a:lnSpc xmlns:a="{A_NS}"/>"#);
        assert!(parse_lnspc(Document::parse(&empty).unwrap().root_element()).is_none());
        // spcPct with no @val falls through and (here) so does absent spcPts.
        let no_val = format!(r#"<a:lnSpc xmlns:a="{A_NS}"><a:spcPct/></a:lnSpc>"#);
        assert!(parse_lnspc(Document::parse(&no_val).unwrap().root_element()).is_none());
    }

    /// The serde shape matches the two enums it replaces: tag "type", camelCase
    /// variant tags, plain `val`.
    #[test]
    fn space_line_serializes_with_type_tag_and_camelcase() {
        let pct = serde_json::to_value(SpaceLine::Pct { val: 150000.0 }).unwrap();
        assert_eq!(pct["type"], "pct");
        assert_eq!(pct["val"], 150000.0);
        let pts = serde_json::to_value(SpaceLine::Pts { val: 18.0 }).unwrap();
        assert_eq!(pts["type"], "pts");
        assert_eq!(pts["val"], 18.0);
    }

    /// `<a:normAutofit fontScale lnSpcReduction>` → ("norm", scale/100000,
    /// reduction/100000).
    #[test]
    fn parse_autofit_normautofit_reads_scales() {
        let xml = format!(
            r#"<a:bodyPr xmlns:a="{A_NS}"><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>"#
        );
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            parse_autofit(doc.root_element()),
            Some(("norm".to_owned(), Some(0.625), Some(0.20)))
        );
    }

    /// `<a:normAutofit>` with no scale attributes → ("norm", None, None).
    #[test]
    fn parse_autofit_normautofit_without_scales() {
        let xml = format!(r#"<a:bodyPr xmlns:a="{A_NS}"><a:normAutofit/></a:bodyPr>"#);
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            parse_autofit(doc.root_element()),
            Some(("norm".to_owned(), None, None))
        );
    }

    /// `<a:spAutoFit>` → ("sp", None, None).
    #[test]
    fn parse_autofit_spautofit() {
        let xml = format!(r#"<a:bodyPr xmlns:a="{A_NS}"><a:spAutoFit/></a:bodyPr>"#);
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            parse_autofit(doc.root_element()),
            Some(("sp".to_owned(), None, None))
        );
    }

    /// `<a:noAutofit>` → ("none", None, None).
    #[test]
    fn parse_autofit_noautofit() {
        let xml = format!(r#"<a:bodyPr xmlns:a="{A_NS}"><a:noAutofit/></a:bodyPr>"#);
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            parse_autofit(doc.root_element()),
            Some(("none".to_owned(), None, None))
        );
    }

    /// A `<a:bodyPr>` with NO autofit child → None, so the caller applies its own
    /// default (pptx: inherited placeholder value; xlsx: "none").
    #[test]
    fn parse_autofit_none_when_no_child() {
        let xml = format!(r#"<a:bodyPr xmlns:a="{A_NS}" anchor="ctr" wrap="square"/>"#);
        let doc = Document::parse(&xml).unwrap();
        assert!(parse_autofit(doc.root_element()).is_none());
    }

    // ── parse_body_pr ───────────────────────────────────────────────────────

    /// The spec defaults expose ECMA-376 §21.1.2.1.1 values.
    #[test]
    fn body_pr_spec_defaults() {
        let d = BodyPrDefaults::spec();
        assert_eq!(d.anchor, "t");
        assert_eq!(d.wrap, "square");
        assert_eq!(d.vert, "horz");
        assert_eq!(d.l_ins, 91_440);
        assert_eq!(d.r_ins, 91_440);
        assert_eq!(d.t_ins, 45_720);
        assert_eq!(d.b_ins, 45_720);
        assert_eq!(d.auto_fit, "none");
        assert_eq!(DEFAULT_INS_LR_EMU, 91_440);
        assert_eq!(DEFAULT_INS_TB_EMU, 45_720);
    }

    /// An empty `<a:bodyPr>` yields exactly the passed defaults.
    #[test]
    fn parse_body_pr_uses_defaults_when_absent() {
        let xml = format!(r#"<a:bodyPr xmlns:a="{A_NS}"/>"#);
        let doc = Document::parse(&xml).unwrap();
        let b = parse_body_pr(doc.root_element(), &BodyPrDefaults::spec());
        assert_eq!(b.anchor, "t");
        assert_eq!(b.wrap, "square");
        assert_eq!(b.vert, "horz");
        assert_eq!(b.l_ins, 91_440);
        assert_eq!(b.t_ins, 45_720);
        assert_eq!(b.r_ins, 91_440);
        assert_eq!(b.b_ins, 45_720);
        assert_eq!(b.auto_fit, "none");
        assert_eq!(b.font_scale, None);
        assert_eq!(b.ln_spc_reduction, None);
    }

    /// Explicit attributes override defaults; the autofit child sets auto_fit +
    /// its scales.
    #[test]
    fn parse_body_pr_reads_attrs_and_autofit_child() {
        let xml = format!(
            r#"<a:bodyPr xmlns:a="{A_NS}" anchor="ctr" wrap="none" vert="vert270" lIns="0" tIns="1pt" rIns="0" bIns="-0.5pt"><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let b = parse_body_pr(doc.root_element(), &BodyPrDefaults::spec());
        assert_eq!(b.anchor, "ctr");
        assert_eq!(b.wrap, "none");
        assert_eq!(b.vert, "vert270");
        assert_eq!(b.l_ins, 0);
        assert_eq!(b.t_ins, 12_700);
        assert_eq!(b.r_ins, 0);
        assert_eq!(b.b_ins, -6_350);
        assert_eq!(b.auto_fit, "norm");
        assert_eq!(b.font_scale, Some(0.625));
        assert_eq!(b.ln_spc_reduction, Some(0.20));
    }

    /// Non-spec defaults (e.g. pptx's theme-resolved anchor) are honored when the
    /// attribute is absent, but overridden when present. Also: the autofit
    /// default applies only when there is no autofit child.
    #[test]
    fn parse_body_pr_honors_custom_defaults() {
        let defaults = BodyPrDefaults {
            anchor: "b".to_owned(),
            wrap: "square".to_owned(),
            vert: "horz".to_owned(),
            l_ins: 100_000,
            t_ins: 50_000,
            r_ins: 100_000,
            b_ins: 50_000,
            auto_fit: "sp".to_owned(),
        };
        // Absent attrs → custom defaults; no autofit child → default auto_fit.
        let xml = format!(r#"<a:bodyPr xmlns:a="{A_NS}"/>"#);
        let doc = Document::parse(&xml).unwrap();
        let b = parse_body_pr(doc.root_element(), &defaults);
        assert_eq!(b.anchor, "b");
        assert_eq!(b.l_ins, 100_000);
        assert_eq!(b.auto_fit, "sp");
        // Present anchor attr wins over the custom default.
        let xml2 = format!(r#"<a:bodyPr xmlns:a="{A_NS}" anchor="t"/>"#);
        let doc2 = Document::parse(&xml2).unwrap();
        assert_eq!(parse_body_pr(doc2.root_element(), &defaults).anchor, "t");
    }
}
