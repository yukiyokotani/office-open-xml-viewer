//! Theme + colour-map parsing: `<a:theme>` colour/font scheme extraction, the
//! master `<p:clrMap>` / `<p:clrMapOvr>` logical-name remapping, and the pptx
//! `ThemeResolver` (`PptxSchemeResolver`) that resolves `<a:schemeClr>` names to
//! theme-slot hexes. Extracted verbatim from `lib.rs`. Shared XML helpers
//! (`child`, `attr`) stay in `lib.rs` and are imported here.

use crate::parse_preflighted_pptx_xml;
use crate::{attr, child, parse_rels, read_zip_str, resolve_path, PptxZip};
use ooxml_common::rels::relationship_part_path;
use ooxml_common::theme::ThemeFormatScheme;
use serde::ser::Serializer;
use std::collections::HashMap;
use std::ops::{Deref, DerefMut};

const THEME_REL_PREFIX: &str = "+themeRel-";
const RAW_SCHEME_PREFIX: &str = "+rawScheme-";

/// PPTX host adapter for the shared DrawingML theme model. The historic flat
/// map remains the color/font/object-default carrier; `format_scheme` is kept
/// beside it because style recipes are structured XML, not synthetic map keys.
#[derive(Debug, Clone, Default)]
pub(crate) struct PptxTheme {
    values: HashMap<String, String>,
    pub(crate) format_scheme: ThemeFormatScheme,
    format_scheme_present: bool,
    pub(crate) chart_images: ooxml_common::chart::ChartImageRelationships,
}

impl PptxTheme {
    pub(crate) fn from_xml(xml: &str) -> Self {
        let format_scheme = if parse_preflighted_pptx_xml(xml).is_ok() {
            ThemeFormatScheme::parse(xml)
        } else {
            ThemeFormatScheme::default()
        };
        Self {
            values: parse_theme_colors(xml),
            format_scheme,
            format_scheme_present: true,
            chart_images: ooxml_common::chart::ChartImageRelationships::default(),
        }
    }
}

impl Deref for PptxTheme {
    type Target = HashMap<String, String>;

    fn deref(&self) -> &Self::Target {
        &self.values
    }
}

impl DerefMut for PptxTheme {
    fn deref_mut(&mut self) -> &mut Self::Target {
        &mut self.values
    }
}

impl serde::Serialize for PptxTheme {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        self.values.serialize(serializer)
    }
}

/// Theme inputs accepted by shape parsing. Unit tests and callers that build a
/// palette directly continue to use a `HashMap`; package-backed parsing passes
/// `PptxTheme` and therefore exposes the format scheme as well.
pub(crate) trait PptxThemeSource {
    fn colors(&self) -> &HashMap<String, String>;
    fn format_scheme(&self) -> Option<&ThemeFormatScheme> {
        None
    }
    fn chart_images(&self) -> Option<&ooxml_common::chart::ChartImageRelationships> {
        None
    }
}

impl PptxThemeSource for HashMap<String, String> {
    fn colors(&self) -> &HashMap<String, String> {
        self
    }
}

impl PptxThemeSource for PptxTheme {
    fn colors(&self) -> &HashMap<String, String> {
        &self.values
    }

    fn format_scheme(&self) -> Option<&ThemeFormatScheme> {
        self.format_scheme_present.then_some(&self.format_scheme)
    }

    fn chart_images(&self) -> Option<&ooxml_common::chart::ChartImageRelationships> {
        Some(&self.chart_images)
    }
}

pub(crate) fn theme_relationship_path<'a>(
    theme: &'a HashMap<String, String>,
    relationship_id: &str,
) -> Option<&'a str> {
    theme
        .get(&format!("{THEME_REL_PREFIX}{relationship_id}"))
        .map(String::as_str)
}

/// Parse a theme part together with relationships owned by that theme.
///
/// DrawingML style-matrix fills can embed images. Their relationship IDs are
/// scoped to the theme part, not to the slide, layout, or master that later
/// references the style. Retain the resolved package paths beside the existing
/// flat theme map so deferred `fillRef` / `bgRef` resolution uses the correct
/// OPC source part.
pub(crate) fn parse_theme_part(theme_path: &str, zip: &mut PptxZip) -> PptxTheme {
    let theme_xml = read_zip_str(zip, theme_path).unwrap_or_default();
    let mut theme = PptxTheme::from_xml(&theme_xml);
    let rels_xml = read_zip_str(zip, &relationship_part_path(theme_path)).unwrap_or_default();
    let theme_dir = theme_path.rsplit_once('/').map_or("", |(dir, _)| dir);

    theme.chart_images.insert_part_relationships(
        ooxml_common::chart::ChartImageSource::Theme,
        theme_path,
        &rels_xml,
    );

    for (relationship_id, target) in parse_rels(&rels_xml) {
        let path = resolve_path(theme_dir, &target);
        if zip.index_for_name(&path).is_some() {
            theme.insert(format!("{THEME_REL_PREFIX}{relationship_id}"), path);
        }
    }
    theme
}

#[cfg(test)]
mod relationship_tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;

    #[test]
    fn theme_part_relationships_are_resolved_from_the_theme_directory() {
        let mut bytes = Vec::new();
        {
            let mut archive = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = SimpleFileOptions::default();
            archive.start_file("ppt/theme/theme1.xml", options).unwrap();
            archive
                .write_all(
                    b"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>",
                )
                .unwrap();
            archive
                .start_file("ppt/theme/_rels/theme1.xml.rels", options)
                .unwrap();
            archive.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/background.png"/></Relationships>"#).unwrap();
            archive
                .start_file("ppt/media/background.png", options)
                .unwrap();
            archive.write_all(b"png").unwrap();
            archive.finish().unwrap();
        }

        let mut zip = PptxZip::new(Cursor::new(bytes)).expect("open package");
        let theme = parse_theme_part("ppt/theme/theme1.xml", &mut zip);
        assert_eq!(
            theme_relationship_path(&theme, "rId1"),
            Some("ppt/media/background.png")
        );
    }
}

/// Parse the color scheme from a theme XML file.
/// Returns a map: scheme slot name (e.g. "dk1", "lt1", "acc1") → hex string.
///
/// The clrScheme and fontScheme are parsed by the shared
/// [`ooxml_common::theme`] grammar; this function keeps pptx's flat merged-map
/// storage (colors, `+mj-lt`/`+mn-*` font keys and
/// `+txDef`/`+spDef` object defaults all in one `HashMap<String, String>`)
/// because ~30 call sites look these up by string key. Existing fill/effect
/// fragment adapters and objectDefaults stay local. Structured line recipes
/// are retained by [`PptxTheme`] instead of being flattened into width
/// sentinels.
pub(crate) fn parse_theme_colors(xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    // Enforce the PPTX-local defense-in-depth node ceiling before any shared
    // theme adapter reparses this already lexically-preflighted input.
    let doc = match parse_preflighted_pptx_xml(xml) {
        Ok(d) => d,
        Err(_) => return map,
    };

    // Color slots: shared parse, kept RAW (pptx applies no case-folding, unlike
    // docx/xlsx). prstClr is now resolved via the shared preset table.
    for (slot, hex) in ooxml_common::theme::ThemeColorScheme::parse(xml).iter() {
        map.insert(slot.to_owned(), hex.to_owned());
        // Logical PresentationML color names overlap the raw theme slot names
        // for accent1..accent6/hlink/folHlink. Keep an immutable copy so a
        // non-identity clrMap can be reapplied without destroying the source
        // palette, and style-matrix recipes can resolve fixed scheme colors
        // without accidentally applying the presentation mapping twice.
        map.insert(format!("{RAW_SCHEME_PREFIX}{slot}"), hex.to_owned());
    }

    // Font scheme: shared parse, mapped onto pptx's `+mj-*` / `+mn-*` keys.
    let fonts = ooxml_common::theme::ThemeFonts::parse(xml);
    for (group, prefix) in [(&fonts.major, "+mj"), (&fonts.minor, "+mn")] {
        for (face, axis) in [(&group.latin, "lt"), (&group.ea, "ea"), (&group.cs, "cs")] {
            if let Some(typeface) = face {
                map.insert(format!("{prefix}-{axis}"), typeface.clone());
            }
        }
    }

    let root = doc.root_element();

    // Parse fmtScheme so style references can resolve against the theme's format
    // style matrix (ECMA-376 §20.1.4.1.18). Fill entries stay as small owned XML
    // fragments because phClr is supplied by each individual fillRef/bgRef and
    // therefore cannot be resolved while the theme is parsed. Only a referenced
    // fragment is reparsed later; ordinary explicit fills keep their current fast
    // path. Line recipes live in the shared `ThemeFormatScheme` sidecar.
    if let Some(fmt_scheme) = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "fmtScheme")
    {
        for (list_name, key_prefix) in [
            ("fillStyleLst", "+fillStyle"),
            ("bgFillStyleLst", "+bgFillStyle"),
        ] {
            if let Some(style_list) = child(fmt_scheme, list_name) {
                for (i, fill) in style_list.children().filter(|n| n.is_element()).enumerate() {
                    if let Some(fragment) = xml.get(fill.range()) {
                        map.insert(format!("{key_prefix}-{}", i + 1), fragment.to_owned());
                    }
                }
            }
        }
        if let Some(effect_style_lst) = child(fmt_scheme, "effectStyleLst") {
            for (i, effect_style) in effect_style_lst
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "effectStyle")
                .enumerate()
            {
                if let Some(fragment) = xml.get(effect_style.range()) {
                    map.insert(format!("+effectStyle-{}", i + 1), fragment.to_owned());
                }
            }
        }
    }

    // Parse <a:objectDefaults> per ECMA-376 §20.1.6.7. PowerPoint stores
    // `<a:txDef>` (text-box default), `<a:spDef>` (shape default) and
    // `<a:lnDef>` (line default) here; their `<a:bodyPr>` settings are the
    // last-resort fallback below master/layout/slide for any text body that
    // doesn't override the attribute. Sample-2 hides a `<a:spAutoFit/>`
    // inside its txDef — without inheriting it, every text box in the
    // template defaults to "noAutofit + wrap=square" (the spec literal),
    // which makes mixed-size runs like "20代" wrap unnecessarily and
    // reproduces the slide-13 regression on every similar deck. We parse
    // the bodyPr attributes (and the autoFit child element) into namespaced
    // theme keys so `parse_text_body` can fall back through them without
    // changing function signatures.
    if let Some(obj_defaults) = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "objectDefaults")
    {
        let read_def = |def_name: &str, key_prefix: &str, map: &mut HashMap<String, String>| {
            let Some(def) = child(obj_defaults, def_name) else {
                return;
            };
            let Some(body_pr) = child(def, "bodyPr") else {
                return;
            };
            // Plain attributes — copy verbatim; consumers parse the strings.
            for attr_name in [
                "wrap",
                "anchor",
                "anchorCtr",
                "vert",
                "rtlCol",
                "lIns",
                "rIns",
                "tIns",
                "bIns",
                "numCol",
                "spcCol",
                "vertOverflow",
                "horzOverflow",
                "spcFirstLastPara",
                "rot",
                "upright",
                "fromWordArt",
                "forceAA",
                "compatLnSpc",
            ] {
                if let Some(v) = attr(&body_pr, attr_name) {
                    map.insert(format!("{key_prefix}-bodyPr-{attr_name}"), v);
                }
            }
            // Auto-fit is a child element, not an attribute. Encode as
            // `{prefix}-autoFit` → "sp" | "norm" | "none".
            let auto_fit = if child(body_pr, "spAutoFit").is_some() {
                "sp"
            } else if child(body_pr, "normAutofit").is_some() {
                "norm"
            } else {
                "none"
            };
            // Only record when explicit; "none" is the spec default and
            // recording it would make every consumer see `Some("none")` even
            // for themes that didn't say anything.
            if auto_fit != "none" {
                map.insert(format!("{key_prefix}-autoFit"), auto_fit.to_owned());
            }
        };
        read_def("txDef", "+txDef", &mut map);
        read_def("spDef", "+spDef", &mut map);
    }

    map
}

/// Bake a slide master's `<p:clrMap>` (ECMA-376 §19.3.1.6) into a theme map so
/// that logical scheme names (bg1/tx1/bg2/tx2/accent1..6/hlink/folHlink) can be
/// resolved by a direct `theme.get(name)` lookup later.
///
/// `<p:clrMap>` maps each logical name to a theme color-scheme slot
/// (dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink). We resolve that indirection
/// here and insert `theme[logical] = theme[slot]` for every logical name. This
/// keeps `parse_color_node_tint`'s `schemeClr` handling a single map lookup.
///
/// When `<p:clrMap>` is absent (or an attribute is missing) the PowerPoint
/// default mapping is applied: bg1=lt1, tx1=dk1, bg2=lt2, tx2=dk2, accentN
/// identity, hlink/folHlink identity. The raw slot keys (dk1, lt1, …) added by
/// `parse_theme_colors` are left untouched, so canonical lookups still work.
/// The 12 logical scheme names of a `CT_ColorMapping` (ECMA-376 §19.3.1.6),
/// paired with the scheme slot each maps to by default when a `<p:clrMap>`
/// attribute is absent: bg1=lt1, tx1=dk1, bg2=lt2, tx2=dk2, accentN identity,
/// hlink/folHlink identity. Shared by `<p:clrMap>` (§19.3.1.6) and
/// `<a:overrideClrMapping>` (§20.1.6.8), which carry the identical attribute set.
/// Aliases `ooxml_common::color::SCHEME_DEFAULT_SLOTS` so the default §19.3.1.6
/// table lives in exactly one place across the workspace.
pub(crate) const CLR_MAP_LOGICALS: &[(&str, &str)] = ooxml_common::color::SCHEME_DEFAULT_SLOTS;

/// Read the 12 `CT_ColorMapping` attributes (§19.3.1.6) from `node` into an owned
/// `{logical → slot}` map. Works for both `<p:clrMap>` and
/// `<a:overrideClrMapping>` (same attribute set). roxmltree borrows the doc, so
/// we resolve into an owned map here rather than return the node.
pub(crate) fn parse_clr_map_node(node: roxmltree::Node<'_, '_>) -> HashMap<String, String> {
    let mut m: HashMap<String, String> = HashMap::new();
    for (logical, _) in CLR_MAP_LOGICALS {
        if let Some(slot) = attr(&node, logical) {
            m.insert((*logical).to_owned(), slot);
        }
    }
    m
}

/// Apply a `{logical → slot}` color mapping to `theme`, inserting
/// `theme[logical] = theme[slot]` for every logical name (falling back to the
/// default slot from `CLR_MAP_LOGICALS` when the mapping omits an attr). Reads
/// the raw scheme slot keys (dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink) and
/// writes the logical keys, leaving the raw slots untouched — so the same
/// `theme` can be re-baked later with an override mapping that again resolves
/// against the intact raw slots. `clr_map` = `None` applies the all-default
/// PowerPoint mapping.
pub(crate) fn apply_clr_map(
    theme: &mut HashMap<String, String>,
    clr_map: Option<&HashMap<String, String>>,
) {
    for (logical, default_slot) in CLR_MAP_LOGICALS {
        // Resolve the slot this logical name points at (mapping value, else default).
        let slot = clr_map
            .and_then(|m| m.get(*logical).cloned())
            .unwrap_or_else(|| (*default_slot).to_owned());
        // theme[logical] = theme[slot] when the slot has a hex; otherwise skip
        // (leaves any prior value, and the canonical fallback still applies).
        if let Some(hex) = theme
            .get(&format!("{RAW_SCHEME_PREFIX}{slot}"))
            .or_else(|| theme.get(&slot))
            .cloned()
        {
            theme.insert((*logical).to_owned(), hex);
        }
    }
}

pub(crate) fn bake_clr_map(theme: &mut HashMap<String, String>, master_xml: Option<&str>) {
    // Find the master's <p:clrMap> element (direct child of <p:sldMaster>) and
    // resolve its 12 logical→slot attrs, then apply.
    let clr_map = master_xml.and_then(|xml| {
        let doc = parse_preflighted_pptx_xml(xml).ok()?;
        let node = child(doc.root_element(), "clrMap")?;
        Some(parse_clr_map_node(node))
    });
    apply_clr_map(theme, clr_map.as_ref());
}

/// Resolve a slide-or-layout's `<p:clrMapOvr>` color-mapping override
/// (ECMA-376 §19.3.1.7 CT_ColorMappingOverride). Returns:
/// - `Some(map)` when `<a:overrideClrMapping>` is present — its 12 logical→slot
///   attrs replace the master's mapping for this slide/layout (§20.1.6.8).
/// - `None` when `<a:masterClrMapping/>` is present (§20.1.6.6) or there is no
///   `<p:clrMapOvr>` at all — both mean "inherit the master's clrMap".
pub(crate) fn parse_clr_map_ovr(xml: &str) -> Option<HashMap<String, String>> {
    // Fast reject: `<p:clrMapOvr>` is absent on the vast majority of slides, so
    // skip a full second parse of the (often largest) slide XML part when the
    // element name does not even appear. A substring false-positive is harmless
    // — we then parse and find no `overrideClrMapping`, returning `None` as usual.
    if !xml.contains("clrMapOvr") {
        return None;
    }
    let doc = parse_preflighted_pptx_xml(xml).ok()?;
    // <p:clrMapOvr> is a direct child of <p:sld> / <p:sldLayout> (right after
    // <p:cSld>); the choice inside is masterClrMapping XOR overrideClrMapping.
    let ovr = child(doc.root_element(), "clrMapOvr")?;
    let override_node = child(ovr, "overrideClrMapping")?;
    Some(parse_clr_map_node(override_node))
}

/// Resolve a theme typeface reference (e.g. "+mj-lt") to the actual font family name.
/// If the typeface starts with '+' and has a matching entry in the theme map (added by
/// parse_theme_colors from the fontScheme), returns the resolved name; otherwise returns
/// the original string unchanged.
pub(crate) fn resolve_theme_typeface(typeface: &str, theme: &HashMap<String, String>) -> String {
    if typeface.starts_with('+') {
        if let Some(resolved) = theme.get(typeface) {
            return resolved.clone();
        }
    }
    typeface.to_string()
}

/// Resolves a `<a:schemeClr val>` name to its base theme hex the PowerPoint
/// way, for the shared [`ooxml_common::color::parse_color_node`]. The color
/// grammar (srgbClr/sysClr/prstClr/schemeClr + transforms) is shared; only this
/// theme-slot lookup is pptx-specific.
pub(crate) struct PptxSchemeResolver<'a> {
    pub(crate) theme: &'a HashMap<String, String>,
}

/// Resolve the theme's authored color-scheme slots without applying the
/// PresentationML logical color map. This is required after a `clrMap` maps an
/// overlapping name such as `accent1` to another slot: format-scheme recipes
/// and chart-local color maps refer to the underlying scheme slot, not the
/// already-mapped logical color.
pub(crate) struct PptxRawSchemeResolver<'a> {
    pub(crate) theme: &'a HashMap<String, String>,
}

impl ooxml_common::color::ThemeResolver for PptxRawSchemeResolver<'_> {
    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        let slot = ooxml_common::color::default_scheme_slot(name);
        self.theme
            .get(&format!("{RAW_SCHEME_PREFIX}{slot}"))
            .or_else(|| self.theme.get(slot))
            .cloned()
    }
}

impl ooxml_common::color::ThemeResolver for PptxSchemeResolver<'_> {
    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        // Per ECMA-376 §19.3.1.6 the master's <p:clrMap> remaps logical
        // names (bg1/tx1/bg2/tx2/accentN/hlink/folHlink) to theme slots.
        // `bake_clr_map` pre-bakes those logical names into the theme
        // map, so try a direct lookup FIRST — this honors clrMap (e.g.
        // tx1="lt1"). Fall back to the canonical alias only when the
        // logical name was not baked (no master / unmapped name), so a
        // missing clrMap still resolves tx1→dk1, bg1→lt1, etc.
        if let Some(hex) = self.theme.get(name) {
            return Some(hex.clone());
        }
        // Canonical logical→slot fallback, per the default §19.3.1.6
        // clrMap (shared table: ooxml_common::color::SCHEME_DEFAULT_SLOTS).
        // The helper also passes raw slot names (dk1/lt1/…) and accents
        // through unchanged.
        let canonical: &str = match name {
            // phClr = "placeholder color" (inherits from layout).
            // Approximate as the primary dark text color. Not part of
            // §19.3.1.6, so it stays a local special case.
            "phClr" => "dk1",
            other => ooxml_common::color::default_scheme_slot(other),
        };
        self.theme.get(canonical).cloned()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};

    #[test]
    fn color_map_swaps_are_resolved_from_the_immutable_scheme_palette() {
        let mut theme = parse_theme_colors(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:themeElements><a:clrScheme name="swap">
                <a:dk1><a:srgbClr val="000000"/></a:dk1>
                <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                <a:dk2><a:srgbClr val="111111"/></a:dk2>
                <a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
                <a:accent1><a:srgbClr val="AA0000"/></a:accent1>
                <a:accent2><a:srgbClr val="00BB00"/></a:accent2>
                <a:accent3><a:srgbClr val="0000CC"/></a:accent3>
                <a:accent4><a:srgbClr val="444444"/></a:accent4>
                <a:accent5><a:srgbClr val="555555"/></a:accent5>
                <a:accent6><a:srgbClr val="666666"/></a:accent6>
                <a:hlink><a:srgbClr val="777777"/></a:hlink>
                <a:folHlink><a:srgbClr val="888888"/></a:folHlink>
              </a:clrScheme></a:themeElements>
            </a:theme>"#,
        );
        let mapping = HashMap::from([
            ("accent1".to_owned(), "accent2".to_owned()),
            ("accent2".to_owned(), "accent1".to_owned()),
        ]);

        apply_clr_map(&mut theme, Some(&mapping));
        assert_eq!(theme.get("accent1").map(String::as_str), Some("00BB00"));
        assert_eq!(theme.get("accent2").map(String::as_str), Some("AA0000"));
        assert_eq!(
            ooxml_common::color::ThemeResolver::resolve_scheme_color(
                &PptxRawSchemeResolver { theme: &theme },
                "accent1",
            ),
            Some("AA0000".to_owned())
        );

        // Reapplying a different map must still use the authored palette, not
        // the values written by the first mapping.
        let second = HashMap::from([("accent1".to_owned(), "accent3".to_owned())]);
        apply_clr_map(&mut theme, Some(&second));
        assert_eq!(theme.get("accent1").map(String::as_str), Some("0000CC"));
        assert_eq!(theme.get("accent2").map(String::as_str), Some("00BB00"));
    }

    #[test]
    fn absent_theme_and_broken_theme_part_keep_distinct_chart_semantics() {
        assert!(PptxTheme::default().format_scheme().is_none());

        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default();
            writer.start_file("placeholder", options).unwrap();
            writer.write_all(b"x").unwrap();
            writer.finish().unwrap();
        }
        let mut archive = PptxZip::new(Cursor::new(bytes)).unwrap();
        let broken = parse_theme_part("ppt/theme/missing.xml", &mut archive);
        assert!(broken.format_scheme().is_some());
        assert!(matches!(
            broken.format_scheme.lookup_fill_ref(1),
            ooxml_common::theme::StyleMatrixLookup::Missing
        ));
    }

    #[test]
    fn shallow_theme_xml_parses_colors() {
        // A normal theme parses through the PPTX preflighted parser unchanged (sanity: the
        // guard does not reject legitimate parts).
        let xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:themeElements><a:clrScheme name="Office">
              <a:dk1><a:srgbClr val="000000"/></a:dk1>
              <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
            </a:clrScheme></a:themeElements></a:theme>"#;
        let map = parse_theme_colors(xml);
        assert_eq!(map.get("dk1").map(String::as_str), Some("000000"));
        assert_eq!(map.get("lt1").map(String::as_str), Some("FFFFFF"));
    }

    // ── RB2 neutralization (roxmltree layer): a pathologically deep theme/part
    //    XML is rejected by the depth pre-check in the PPTX parser BEFORE
    //    roxmltree's recursive tree builder runs, so `parse_theme_colors` returns
    //    gracefully instead of trapping. This runs on the DEFAULT (small)
    //    test-thread stack ON PURPOSE: handing 5 000-deep XML straight to
    //    roxmltree here would overflow and abort the process. That it returns is
    //    the guarantee. (The shape-walk DepthGuard is covered separately in
    //    shape.rs on a generous stack.)
    #[test]
    fn deeply_nested_theme_xml_is_rejected_not_trapped() {
        let mut xml = String::from(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
        );
        // 5 000 levels — ~20× MAX_XML_DEPTH and far past roxmltree's ~1 000-deep
        // small-stack overflow threshold.
        for _ in 0..5_000 {
            xml.push_str("<a:x>");
        }
        for _ in 0..5_000 {
            xml.push_str("</a:x>");
        }
        xml.push_str("</a:theme>");

        // Must return (not trap); the rejected part yields no colors.
        let map = parse_theme_colors(&xml);
        assert!(
            map.is_empty(),
            "an over-deep theme XML must be rejected, yielding no colors"
        );
    }
}
