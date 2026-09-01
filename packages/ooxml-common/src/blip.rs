//! Shared DrawingML blip helpers used by the docx, pptx and xlsx parsers.
//!
//! Every OOXML host references embedded media through the same DrawingML
//! `<a:blip>` element (ECMA-376 §20.1.8.13). A blip's raster image rides in its
//! `@r:embed` relationship; when the picture is a vector graphic, Microsoft's
//! 2016 SVG extension (MS-ODRAWXML) nests an `<asvg:svgBlip r:embed="…">` inside
//! the blip's `<a:extLst>`, and the raster `@r:embed` becomes a *fallback* for
//! SVG-incapable clients. These helpers were previously re-implemented (or, for
//! the SVG extension, only implemented in pptx) per parser; sharing them keeps
//! the three formats' blip handling identical.

use crate::ns::relationships;
use roxmltree::Node;
use serde::{Deserialize, Serialize};

/// Transitional relationships namespace (`r:`) — where a blip's `embed` / `link`
/// rIds live. Runtime lookups accept the Strict URI too (see [`blip_embed_rid`]);
/// this constant is kept for the crate's test fixtures.
pub const R_NS: &str = relationships::TRANSITIONAL;

/// Microsoft 2016 SVG extension URI (MS-ODRAWXML). An `<a:blip>` wrapping a
/// vector image carries `<a:extLst><a:ext uri="{96DAC541-…}"><asvg:svgBlip
/// r:embed="…"/></a:ext></a:extLst>`.
pub const SVG_BLIP_EXT_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
pub const SVG_BLIP_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

/// Resolve a blip-like node's `r:embed` relationship id (the raster image, or
/// an `svgBlip`'s vector target). Reads the `embed` attribute in the
/// relationships namespace — Transitional or Strict (ISO/IEC 29500) — tolerating
/// the literal `r:embed` form some producers emit without binding the namespace.
pub fn blip_embed_rid(node: &Node<'_, '_>) -> Option<String> {
    node.attribute((relationships::TRANSITIONAL, "embed"))
        .or_else(|| node.attribute((relationships::STRICT, "embed")))
        .or_else(|| node.attribute("r:embed"))
        .map(str::to_string)
}

/// Resolve the relationship id of the vector original from an `<a:blip>`'s
/// `asvg:svgBlip` extension. Both the registered extension URI and Microsoft
/// SVG namespace are required; an unrelated extension may legally reuse the
/// same local name. The producer's prefix (`asvg:` etc.) does not matter. Returns
/// `None` when the blip carries no svgBlip extension (the common, raster-only
/// case).
pub fn svg_blip_rid(blip: Node<'_, '_>) -> Option<String> {
    blip.children()
        .find(|n| n.is_element() && n.tag_name().name() == "extLst")?
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "ext")
        .filter(|ext| ext.attribute("uri") == Some(SVG_BLIP_EXT_URI))
        .find_map(|ext| {
            ext.children().find(|n| {
                n.is_element()
                    && n.tag_name().name() == "svgBlip"
                    && n.tag_name().namespace() == Some(SVG_BLIP_NS)
            })
        })
        .and_then(|svg_blip| blip_embed_rid(&svg_blip))
}

/// Map a part name / path extension to its MIME type, covering the image, audio
/// and video parts OOXML blips reference. Unknown extensions fall back to
/// `application/octet-stream`. This is the single source of truth for all three
/// parsers — notably it is the only place `svg` ⇒ `image/svg+xml` is decided.
pub fn mime_from_ext(path: &str) -> &'static str {
    match path
        .rsplit('.')
        .next()
        .unwrap_or("")
        .to_ascii_lowercase()
        .as_str()
    {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "svg" => "image/svg+xml",
        "webp" => "image/webp",
        "tif" | "tiff" => "image/tiff",
        // Windows Metafile / Enhanced Metafile. Office embeds these for charts
        // and diagrams; the renderer rasterizes WMF via a minimal player and
        // skips EMF (a follow-up). The conventional MIME types per IANA / Windows.
        "wmf" => "image/wmf",
        "emf" => "image/emf",
        "mp3" => "audio/mpeg",
        "m4a" => "audio/mp4",
        "wav" => "audio/wav",
        "aac" => "audio/aac",
        "wma" => "audio/x-ms-wma",
        "flac" => "audio/flac",
        "ogg" | "oga" => "audio/ogg",
        "mp4" | "m4v" => "video/mp4",
        "mov" | "qt" => "video/quicktime",
        "avi" => "video/x-msvideo",
        "wmv" => "video/x-ms-wmv",
        "mpg" | "mpeg" => "video/mpeg",
        "3gp" => "video/3gpp",
        "mkv" => "video/x-matroska",
        "webm" => "video/webm",
        "ogv" => "video/ogg",
        _ => "application/octet-stream",
    }
}

/// ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop (DrawingML
/// `CT_RelativeRect`). Each edge inset is a signed fraction of the *source*
/// bitmap, measured from that edge, so the visible source region is
/// `[l, t, 1−r, 1−b]`. The raw attributes are `ST_Percentage` in 1000ths of a
/// percent; the parser divides by 100000 to a fraction so renderers need no unit
/// knowledge. `ST_Percentage` has no fixed range; absent edges default to `0`.
/// Shared by the docx, pptx and xlsx parsers so the three formats crop identically.
#[derive(Serialize, Deserialize, Debug, Clone, Copy, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SrcRect {
    pub l: f64,
    pub t: f64,
    pub r: f64,
    pub b: f64,
}

/// Parse `<a:srcRect l t r b>` from a `<*:blipFill>` node (the parent of the
/// `<a:blip>`), matched by namespace-local name so a `p:`, `a:` or `xdr:`
/// blipFill all work. Each edge is divided by 100000 to a fraction; an absent
/// edge defaults to `0`. Returns `None` when there is no `srcRect` or all four
/// edges are zero (no crop), so an uncropped picture never forces a renderer
/// onto the sub-rectangle draw path.
pub fn parse_src_rect(blip_fill: Node<'_, '_>) -> Option<SrcRect> {
    let sr = blip_fill
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "srcRect")?;
    let read = |name: &str| -> f64 {
        sr.attribute(name)
            .and_then(crate::units::drawingml_percentage_to_fraction)
            .unwrap_or(0.0)
    };
    let rect = SrcRect {
        l: read("l"),
        t: read("t"),
        r: read("r"),
        b: read("b"),
    };
    if rect.l.abs() < 1e-9 && rect.t.abs() < 1e-9 && rect.r.abs() < 1e-9 && rect.b.abs() < 1e-9 {
        None
    } else {
        Some(rect)
    }
}

/// Parse `<a:blip><a:alphaModFix amt="…"/></a:blip>` from a `<*:blipFill>` node
/// (ECMA-376 §20.1.8.6 CT_AlphaModulateFixedEffect). `amt` is an
/// ST_PositivePercentage in 1000ths of a percent, so `amt/100000` is the alpha
/// scale fraction (0.0–1.0). The effect multiplies the picture's opacity by that
/// fraction — the renderer applies it via `globalAlpha`.
///
/// Returns `None` when there is no `alphaModFix` OR when the amount is ≥ 100%
/// (fully opaque = nothing to apply), so an unaffected picture never forces the
/// renderer onto the alpha path. Negative amounts clamp to 0. Shared by the
/// pptx and xlsx parsers (docx can carry it too) so the three formats read the
/// blip alpha identically.
pub fn parse_blip_alpha(blip_fill: Node<'_, '_>) -> Option<f64> {
    let blip = blip_fill
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "blip")?;
    let mut found = false;
    let mut frac = 1.0;
    for effect in blip
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "alphaModFix")
    {
        let amt = effect
            .attribute("amt")
            .map(crate::units::drawingml_percentage_to_fraction)
            .unwrap_or(Some(1.0))?;
        frac *= amt.max(0.0);
        found = true;
    }
    if !found {
        return None;
    }
    if frac >= 1.0 {
        None
    } else {
        Some(frac)
    }
}

/// A `<a:duotone>` image effect (ECMA-376 §20.1.8.23) resolved to its two
/// endpoint colours. `clr1` is the first `EG_ColorChoice` child (the dark
/// endpoint, luminance 0), `clr2` the second (the light endpoint, luminance 1) —
/// the spec says the effect "combines clr1 and clr2 through a linear
/// interpolation" per pixel, so pixel luminance is the interpolation factor.
/// Both are 6-char uppercase hex WITHOUT a `#`; any per-colour transforms
/// (lumMod/lumOff/tint/satMod/shade/…) are already baked in by
/// [`crate::color::parse_color_node`]. Shared across the docx/pptx/xlsx parsers.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct Duotone {
    /// First colour child — dark endpoint. 6-char hex, no `#`.
    pub clr1: String,
    /// Second colour child — light endpoint. 6-char hex, no `#`.
    pub clr2: String,
}

/// Parse `<a:blipFill>`'s `<a:duotone>` (§20.1.8.23), resolving its two
/// `EG_ColorChoice` children through the shared DrawingML colour grammar
/// ([`crate::color::parse_color_node`]) so all per-colour transforms are applied.
/// `resolver` supplies the host's `<a:schemeClr>` palette lookup and `tint_mode`
/// selects the Word/PowerPoint `<a:tint>` interpretation, exactly as every other
/// colour path in the parser does.
///
/// `<a:duotone>` is a sibling of `<a:blip>` inside `<a:blipFill>` (it is one of
/// CT_BlipFillProperties' blip-effect children). Returns `None` when there is no
/// duotone or fewer than two resolvable colours (a malformed effect is dropped
/// rather than half-applied). The returned hexes carry NO `#` and are uppercase.
pub fn parse_blip_duotone<R: crate::color::ThemeResolver + ?Sized>(
    blip_fill: Node<'_, '_>,
    resolver: &R,
    tint_mode: crate::color::TintMode,
) -> Option<Duotone> {
    // `<a:duotone>` lives on the blip itself (a blip-effect child), with the
    // srcRect/stretch/tile as its siblings. Some producers place it directly
    // under blipFill; accept either by searching descendants of blipFill's
    // immediate children shallowly.
    let duotone = blip_fill
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "duotone")
        .or_else(|| {
            blip_fill
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "blip")
                .and_then(|blip| {
                    blip.children()
                        .find(|n| n.is_element() && n.tag_name().name() == "duotone")
                })
        })?;
    // The two EG_ColorChoice children, in document order. Each child IS a colour
    // element (srgbClr/schemeClr/prstClr/…) — not a container wrapping one — so it
    // is classified with `color_source_from_element` and resolved (transforms and
    // all) via `resolve_color_source`, reusing the exact colour grammar every
    // other path uses.
    let mut colors = duotone
        .children()
        .filter(|n| n.is_element())
        .filter_map(|node| crate::color::color_source_from_element(node))
        .filter_map(|source| crate::color::resolve_color_source(source, resolver, tint_mode));
    let clr1 = colors.next()?;
    let clr2 = colors.next()?;
    Some(Duotone {
        clr1: clr1.to_uppercase(),
        clr2: clr2.to_uppercase(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use roxmltree::Document;

    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const ASVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

    // ISO/IEC 29500 Strict: a blip declaring the relationships namespace under
    // the `http://purl.oclc.org/ooxml/officeDocument/relationships` URI still
    // resolves its `r:embed` rId (the `relationships::STRICT` branch of
    // `blip_embed_rid`).
    #[test]
    fn blip_embed_rid_reads_strict_relationships_ns() {
        let r_strict = crate::ns::relationships::STRICT;
        let xml = format!(r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{r_strict}" r:embed="rIdStrict"/>"#);
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            blip_embed_rid(&doc.root_element()).as_deref(),
            Some("rIdStrict")
        );
    }

    #[test]
    fn svg_blip_rid_reads_extension_regardless_of_prefix() {
        // The svgBlip uses a different prefix (asvg:) on purpose: matching is by
        // namespace-local name, so the prefix must not matter.
        let xml = format!(
            r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:asvg="{ASVG_NS}" r:embed="rIdPng">
                 <a:extLst>
                   <a:ext uri="{SVG_BLIP_EXT_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>
                 </a:extLst>
               </a:blip>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let blip = doc.root_element();
        assert_eq!(blip_embed_rid(&blip).as_deref(), Some("rIdPng"));
        assert_eq!(svg_blip_rid(blip).as_deref(), Some("rIdSvg"));
    }

    #[test]
    fn svg_blip_rid_requires_registered_uri_and_namespace() {
        for (uri, namespace) in [
            ("{00000000-0000-0000-0000-000000000000}", ASVG_NS),
            (SVG_BLIP_EXT_URI, "urn:example:not-microsoft-svg"),
        ] {
            let xml = format!(
                r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:x="{namespace}">
                     <a:extLst><a:ext uri="{uri}"><x:svgBlip r:embed="rIdWrong"/></a:ext></a:extLst>
                   </a:blip>"#
            );
            let doc = Document::parse(&xml).unwrap();
            assert_eq!(svg_blip_rid(doc.root_element()), None);
        }
    }

    #[test]
    fn parse_src_rect_reads_fractions_and_defaults_absent_edges() {
        // l/r given, t/b absent → fractions = raw/100000, absent ⇒ 0.
        let xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip/><a:srcRect l="32560" r="3829"/></a:blipFill>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let sr = parse_src_rect(doc.root_element()).expect("srcRect present");
        assert!((sr.l - 0.3256).abs() < 1e-9);
        assert!((sr.r - 0.03829).abs() < 1e-9);
        assert_eq!(sr.t, 0.0);
        assert_eq!(sr.b, 0.0);

        let strict_xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip/><a:srcRect l="25%" r="-10%"/></a:blipFill>"#
        );
        let strict = Document::parse(&strict_xml).unwrap();
        let sr = parse_src_rect(strict.root_element()).expect("Strict srcRect present");
        assert_eq!(sr.l, 0.25);
        assert_eq!(sr.r, -0.1);
    }

    #[test]
    fn parse_src_rect_none_when_absent_or_all_zero() {
        let none = format!(r#"<a:blipFill xmlns:a="{A_NS}"><a:blip/></a:blipFill>"#);
        assert!(parse_src_rect(Document::parse(&none).unwrap().root_element()).is_none());
        let zero = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:srcRect l="0" t="0" r="0" b="0"/></a:blipFill>"#
        );
        assert!(parse_src_rect(Document::parse(&zero).unwrap().root_element()).is_none());
    }

    #[test]
    fn svg_blip_only_has_no_raster_embed() {
        // sample-12 shape: the blip has no r:embed at all, only the svgBlip ext.
        let xml = format!(
            r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:asvg="{ASVG_NS}">
                 <a:extLst>
                   <a:ext uri="{SVG_BLIP_EXT_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>
                 </a:extLst>
               </a:blip>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let blip = doc.root_element();
        assert_eq!(blip_embed_rid(&blip), None);
        assert_eq!(svg_blip_rid(blip).as_deref(), Some("rIdSvg"));
    }

    #[test]
    fn plain_blip_has_no_svg() {
        let xml = format!(r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{R_NS}" r:embed="rId1"/>"#);
        let doc = Document::parse(&xml).unwrap();
        let blip = doc.root_element();
        assert_eq!(blip_embed_rid(&blip).as_deref(), Some("rId1"));
        assert_eq!(svg_blip_rid(blip), None);
    }

    #[test]
    fn mime_table_covers_svg_and_common_rasters() {
        assert_eq!(mime_from_ext("../media/image2.svg"), "image/svg+xml");
        assert_eq!(mime_from_ext("a/b/image1.PNG"), "image/png");
        assert_eq!(mime_from_ext("x.jpeg"), "image/jpeg");
        assert_eq!(mime_from_ext("x.webp"), "image/webp");
        assert_eq!(mime_from_ext("x.tif"), "image/tiff");
        assert_eq!(mime_from_ext("x.TIFF"), "image/tiff");
        assert_eq!(mime_from_ext("noext"), "application/octet-stream");
    }

    #[test]
    fn mime_table_covers_metafiles() {
        // WMF/EMF blips (charts, diagrams) get the conventional metafile MIMEs,
        // not the application/octet-stream fallback, so the renderer's content
        // sniff has a sensible MIME to pair with each part.
        assert_eq!(mime_from_ext("word/media/image1.wmf"), "image/wmf");
        assert_eq!(mime_from_ext("ppt/media/image3.emf"), "image/emf");
        // Case-insensitive, like the raster entries.
        assert_eq!(mime_from_ext("a/b/CHART.WMF"), "image/wmf");
        assert_eq!(mime_from_ext("x.EMF"), "image/emf");
    }

    // A minimal ThemeResolver mapping a couple of scheme names, so the duotone
    // scheme-color path is exercised without a full parser.
    struct DuoResolver;
    impl crate::color::ThemeResolver for DuoResolver {
        fn resolve_scheme_color(&self, name: &str) -> Option<String> {
            match name {
                "accent1" => Some("4472C4".to_owned()),
                _ => None,
            }
        }
    }

    #[test]
    fn parse_blip_alpha_reads_amt_fraction() {
        // amt=70000 → 0.70 (the sample-9.xlsx alphaModFix).
        let xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip r:embed="rId1" xmlns:r="{R_NS}"><a:alphaModFix amt="70000"/></a:blip></a:blipFill>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let a = parse_blip_alpha(doc.root_element()).expect("alpha present");
        assert!((a - 0.70).abs() < 1e-9);

        let strict = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip><a:alphaModFix amt="70%"/></a:blip></a:blipFill>"#
        );
        let a = parse_blip_alpha(Document::parse(&strict).unwrap().root_element())
            .expect("Strict alpha present");
        assert!((a - 0.70).abs() < 1e-9);
    }

    #[test]
    fn parse_blip_alpha_none_when_absent_or_fully_opaque() {
        let none = format!(r#"<a:blipFill xmlns:a="{A_NS}"><a:blip/></a:blipFill>"#);
        assert!(parse_blip_alpha(Document::parse(&none).unwrap().root_element()).is_none());
        // amt=100000 (100%) ⇒ nothing to apply.
        let opaque = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip><a:alphaModFix amt="100000"/></a:blip></a:blipFill>"#
        );
        assert!(parse_blip_alpha(Document::parse(&opaque).unwrap().root_element()).is_none());
        let defaulted = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip><a:alphaModFix/></a:blip></a:blipFill>"#
        );
        assert!(parse_blip_alpha(Document::parse(&defaulted).unwrap().root_element()).is_none());

        for amt in ["99990", "99999"] {
            let xml = format!(
                r#"<a:blipFill xmlns:a="{A_NS}"><a:blip><a:alphaModFix amt="{amt}"/></a:blip></a:blipFill>"#
            );
            assert!(
                parse_blip_alpha(Document::parse(&xml).unwrap().root_element()).is_some(),
                "a non-identity authored alpha must be retained: {amt}"
            );
        }
    }

    #[test]
    fn parse_blip_duotone_resolves_two_colours_with_transforms() {
        // The sample-9.xlsx effect: clr1=black (prstClr), clr2=srgbClr DAB6BA with
        // lumMod/lumOff/tint/satMod transforms. clr1 is the dark endpoint.
        let xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}">
                 <a:blip r:embed="rId1" xmlns:r="{R_NS}"/>
                 <a:duotone>
                   <a:prstClr val="black"/>
                   <a:srgbClr val="DAB6BA">
                     <a:lumMod val="20000"/><a:lumOff val="80000"/>
                     <a:tint val="45000"/><a:satMod val="400000"/>
                   </a:srgbClr>
                 </a:duotone>
               </a:blipFill>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let duo = parse_blip_duotone(
            doc.root_element(),
            &DuoResolver,
            crate::color::TintMode::PowerPointLinear,
        )
        .expect("duotone present");
        // clr1 = black.
        assert_eq!(duo.clr1, "000000");
        // clr2 is DAB6BA lightened by the transforms → a light pink (all channels
        // high, R the largest). We assert the shape, not an exact hex (the exact
        // value depends on the shared transform math, reported separately).
        let r = u8::from_str_radix(&duo.clr2[0..2], 16).unwrap();
        let g = u8::from_str_radix(&duo.clr2[2..4], 16).unwrap();
        let b = u8::from_str_radix(&duo.clr2[4..6], 16).unwrap();
        assert!(r > 200, "clr2 R should be high (light pink), got {r}");
        assert!(
            r >= g && r >= b,
            "clr2 should be pink (R largest): {}",
            duo.clr2
        );
    }

    #[test]
    fn parse_blip_duotone_none_when_absent_or_single_colour() {
        let none = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip r:embed="rId1" xmlns:r="{R_NS}"/></a:blipFill>"#
        );
        assert!(parse_blip_duotone(
            Document::parse(&none).unwrap().root_element(),
            &DuoResolver,
            crate::color::TintMode::PowerPointLinear
        )
        .is_none());
        // Only one colour child ⇒ not a valid duotone; dropped.
        let one = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:duotone><a:srgbClr val="112233"/></a:duotone></a:blipFill>"#
        );
        assert!(parse_blip_duotone(
            Document::parse(&one).unwrap().root_element(),
            &DuoResolver,
            crate::color::TintMode::PowerPointLinear
        )
        .is_none());
    }
}
