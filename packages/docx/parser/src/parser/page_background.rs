//! ECMA-376 Part 1 §17.2.1 `w:background`: the fill of every page surface,
//! behind all other document content.
//!
//! Shown only while settings carry §17.15.1.26 `w:displayBackgroundShape`
//! (CT_OnOff; omitted means not displayed). Windows Word 365 PDF output agrees:
//! a non-white `w:color` paints the whole page with displayBackgroundShape on.
//! The solid colour comes from `w:color` or, taking precedence, `w:themeColor`
//! with `w:themeTint`/`w:themeShade`. Transitional VML children (Part 4
//! §19.1.2.2 `v:background`) and a §17.3.3.9 `w:drawing` child are background
//! shapes layered above that colour in document order.
use super::*;

/// ECMA-376 §17.15.1.20 `w:clrSchemeMapping`: each ST_ThemeColor reference
/// (§17.18.97) reads one mapping attribute whose ST_WmlColorSchemeIndex
/// (§17.18.103) names the theme part slot. Omitted attributes use the
/// defaults listed for each attribute (bg1→light1, t1→dark1, bg2→light2,
/// t2→dark2, every other attribute maps to the same-named colour).
#[derive(Debug, Clone, Default)]
pub(super) struct ColorSchemeMapping {
    authored: HashMap<&'static str, String>,
}

const MAPPING_DEFAULTS: [(&str, &str); 12] = [
    ("bg1", "light1"),
    ("t1", "dark1"),
    ("bg2", "light2"),
    ("t2", "dark2"),
    ("accent1", "accent1"),
    ("accent2", "accent2"),
    ("accent3", "accent3"),
    ("accent4", "accent4"),
    ("accent5", "accent5"),
    ("accent6", "accent6"),
    ("hyperlink", "hyperlink"),
    ("followedHyperlink", "followedHyperlink"),
];

impl ColorSchemeMapping {
    pub(super) fn parse(settings_root: roxmltree::Node) -> Self {
        let Some(node) = child_w(settings_root, "clrSchemeMapping") else {
            return Self::default();
        };
        let authored = MAPPING_DEFAULTS
            .iter()
            .filter_map(|(attribute, _)| {
                attr_w(node, attribute)
                    .filter(|value| wml_scheme_slot(value).is_some())
                    .map(|value| (*attribute, value))
            })
            .collect();
        Self { authored }
    }

    fn mapped(&self, attribute: &str) -> Option<&str> {
        self.authored
            .get(attribute)
            .map(String::as_str)
            .or_else(|| {
                MAPPING_DEFAULTS
                    .iter()
                    .find(|(name, _)| *name == attribute)
                    .map(|(_, value)| *value)
            })
    }
}

/// §17.18.97: the clrSchemeMapping attribute an ST_ThemeColor reads. `none`
/// (and an unknown value) names no theme colour.
fn mapping_attribute(theme_color: &str) -> Option<&'static str> {
    Some(match theme_color {
        "dark1" | "text1" => "t1",
        "light1" | "background1" => "bg1",
        "dark2" | "text2" => "t2",
        "light2" | "background2" => "bg2",
        "accent1" => "accent1",
        "accent2" => "accent2",
        "accent3" => "accent3",
        "accent4" => "accent4",
        "accent5" => "accent5",
        "accent6" => "accent6",
        "hyperlink" => "hyperlink",
        "followedHyperlink" => "followedHyperlink",
        _ => return None,
    })
}

/// §17.18.103 ST_WmlColorSchemeIndex → the DrawingML `a:clrScheme` slot.
fn wml_scheme_slot(index: &str) -> Option<&'static str> {
    Some(match index {
        "dark1" => "dk1",
        "light1" => "lt1",
        "dark2" => "dk2",
        "light2" => "lt2",
        "accent1" => "accent1",
        "accent2" => "accent2",
        "accent3" => "accent3",
        "accent4" => "accent4",
        "accent5" => "accent5",
        "accent6" => "accent6",
        "hyperlink" => "hlink",
        "followedHyperlink" => "folHlink",
        _ => return None,
    })
}

/// ST_UcharHexNumber (§17.18.98): exactly two hexadecimal digits.
fn uchar_hex(value: Option<String>) -> Option<u8> {
    let value = value?;
    (value.len() == 2)
        .then(|| u8::from_str_radix(&value, 16).ok())
        .flatten()
}

/// §17.2.1 themeTint / themeShade: transform the theme colour's HSL
/// luminance, `L' = L·tint + (1 − tint)` or `L' = L·shade`, with the stored
/// byte read as a fraction of 255. A supplied tint makes the shade ignored.
fn apply_wml_tint_or_shade(hex: &str, tint: Option<u8>, shade: Option<u8>) -> Option<String> {
    let channel = |index: usize| {
        hex.get(index..index + 2)
            .and_then(|pair| u8::from_str_radix(pair, 16).ok())
            .map(|value| value as f64 / 255.0)
    };
    let (r, g, b) = (channel(0)?, channel(2)?, channel(4)?);
    let (h, l, s) = ooxml_common::color::rgb_to_hls(r, g, b);
    let l = match (tint, shade) {
        (Some(tint), _) => {
            let tint = tint as f64 / 255.0;
            l * tint + (1.0 - tint)
        }
        (None, Some(shade)) => l * (shade as f64 / 255.0),
        (None, None) => l,
    };
    let (r, g, b) = ooxml_common::color::hls_to_rgb(h, l.clamp(0.0, 1.0), s);
    let byte = |channel: f64| (channel.clamp(0.0, 1.0) * 255.0).round() as u8;
    Some(format!("{:02X}{:02X}{:02X}", byte(r), byte(g), byte(b)))
}

/// Resolve a WordprocessingML theme colour reference (`w:themeColor` with its
/// tint/shade siblings) through §17.15.1.20 and the theme part. `None` when
/// the reference names no colour or the theme does not define the slot.
pub(super) fn resolve_wml_theme_color(
    node: roxmltree::Node,
    mapping: &ColorSchemeMapping,
    theme: &ThemeColors,
) -> Option<String> {
    let attribute = mapping_attribute(&attr_w(node, "themeColor")?)?;
    let slot = wml_scheme_slot(mapping.mapped(attribute)?)?;
    let base = theme.resolve(slot)?;
    apply_wml_tint_or_shade(
        &base,
        uchar_hex(attr_w(node, "themeTint")),
        uchar_hex(attr_w(node, "themeShade")),
    )
}

/// §17.2.1 solid background colour: `themeColor` (when it resolves) wins over
/// `color`; `auto` (the attribute default) leaves the consumer's page colour,
/// which is this library's white page; neither attribute means no background.
fn background_color(
    node: roxmltree::Node,
    mapping: &ColorSchemeMapping,
    theme: &ThemeColors,
) -> Option<String> {
    resolve_wml_theme_color(node, mapping, theme).or_else(|| {
        attr_w(node, "color")
            .filter(|value| {
                value.len() == 6 && value.chars().all(|character| character.is_ascii_hexdigit())
            })
            .map(|value| value.to_ascii_uppercase())
    })
}

pub(super) struct PageBackgroundContext<'a> {
    pub(super) mapping: &'a ColorSchemeMapping,
    pub(super) theme: &'a ThemeColors,
    pub(super) style_map: &'a StyleMap,
    pub(super) num_map: &'a mut NumberingMap,
    pub(super) media_map: &'a HashMap<String, String>,
    pub(super) chart_map: &'a ChartMap,
    pub(super) rel_map: &'a HashMap<String, String>,
}

fn push_diagnostic(
    diagnostics: &mut Vec<ParseDiagnostic>,
    code: &str,
    severity: DiagnosticSeverity,
) {
    if diagnostics.iter().any(|diagnostic| diagnostic.code == code) {
        return;
    }
    diagnostics.push(ParseDiagnostic {
        code: code.to_string(),
        severity,
        part: DOCUMENT_PART.to_string(),
        path: Vec::new(),
    });
}

/// Part 4 §19.1.2.2: the VML background fill follows the ordinary §19.1.2.5
/// brush rules of `fillcolor` / `filled` / `v:fill` (fillcolor defaults to
/// white). A solid brush's opacity is carried in the colour's alpha byte.
/// `o:targetscreensize` only names the screen resolution an editor prepared
/// the fill for, and `o:bwmode`/`o:bwpure`/`o:bwnormal` select black-and-white
/// output renditions; neither changes the colour page this library paints.
fn vml_background_fill(
    background: roxmltree::Node,
    media_map: &HashMap<String, String>,
    diagnostics: &mut Vec<ParseDiagnostic>,
) -> Option<ShapeFill> {
    let resolved = vml_fill::resolve(background, None, media_map);
    if resolved.incomplete {
        push_diagnostic(
            diagnostics,
            PARSE_DIAGNOSTIC_CODE_UNSUPPORTED_VML_CONTENT,
            PARSE_DIAGNOSTIC_SEVERITY_UNSUPPORTED_VML_CONTENT,
        );
    }
    match resolved.fill? {
        ShapeFill::Solid { color } => Some(ShapeFill::Solid {
            color: vml_color_alpha(color, resolved.opacity),
        }),
        fill => Some(fill),
    }
}

/// A page fill from one §17.3.3.9 drawing payload: a picture stretches over
/// the page (§20.1.8.56 stretch is the only picture fill mode), and a shape
/// contributes its fill. Other payloads are not fills.
fn drawing_background_fill(run: DocRun) -> Result<Option<ShapeFill>, ()> {
    match run {
        DocRun::Image(image) => {
            // Picture crop, transparency and recolour effects keep their
            // blip meaning; placement transforms do not apply to a page fill.
            if image.color_replace_from.is_some() {
                return Err(());
            }
            Ok(Some(ShapeFill::Image {
                image_path: image.image_path,
                mime_type: image.mime_type,
                svg_image_path: image.svg_image_path,
                src_rect: image.src_rect,
                fill_rect: None,
                tile: None,
                alpha: image.alpha,
                duotone: image.duotone,
                blip_effects: image.blip_effects,
                vml: None,
            }))
        }
        DocRun::Shape(shape) => match shape.fill {
            // DrawingML tiled blip fills (§20.1.8.58) are not rendered by the
            // shared shape image path either; report rather than stretch.
            Some(ShapeFill::Image { tile: Some(_), .. }) => Err(()),
            Some(ShapeFill::Solid { color }) => Ok(Some(ShapeFill::Solid {
                color: vml_color_alpha(color, shape.fill_opacity),
            })),
            fill => Ok(fill),
        },
        _ => Err(()),
    }
}

/// Acquire the §17.2.1 background. `display` is the effective §17.15.1.26
/// `w:displayBackgroundShape`; when it is off nothing is painted.
pub(super) fn parse_page_background(
    node: roxmltree::Node,
    display: bool,
    context: PageBackgroundContext,
    diagnostics: &mut Vec<ParseDiagnostic>,
) -> Option<PageBackgroundWire> {
    if !display {
        return None;
    }
    let mut background = PageBackgroundWire {
        color: background_color(node, context.mapping, context.theme),
        fills: Vec::new(),
    };
    for child in node.children().filter(roxmltree::Node::is_element) {
        if child.has_tag_name(("urn:schemas-microsoft-com:vml", "background")) {
            background
                .fills
                .extend(vml_background_fill(child, context.media_map, diagnostics));
        } else if is_w_ns(child.tag_name().namespace()) && child.tag_name().name() == "drawing" {
            for container in drawing_own_containers(child) {
                let mut runs = parse_drawing_container(
                    context.style_map,
                    context.num_map,
                    container,
                    context.media_map,
                    context.chart_map,
                    context.rel_map,
                    context.theme,
                    DepthGuard::root(),
                );
                let fill = match (runs.pop(), runs.is_empty()) {
                    (None, _) => Ok(None),
                    (Some(run), true) => drawing_background_fill(run),
                    // A group or canvas is several drawables, not one fill.
                    (Some(_), false) => Err(()),
                };
                match fill {
                    Ok(fill) => background.fills.extend(fill),
                    Err(()) => push_diagnostic(
                        diagnostics,
                        PARSE_DIAGNOSTIC_CODE_UNSUPPORTED_BACKGROUND_DRAWING,
                        PARSE_DIAGNOSTIC_SEVERITY_UNSUPPORTED_BACKGROUND_DRAWING,
                    ),
                }
            }
        }
    }
    (background.color.is_some() || !background.fills.is_empty()).then_some(background)
}

#[cfg(test)]
mod tests {
    use super::*;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn theme() -> ThemeColors {
        ThemeColors::parse(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="t">
                <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                <a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
                <a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>
                <a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>
                <a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>
                <a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>
            </a:clrScheme></a:themeElements></a:theme>"#,
        )
    }

    fn parse(background: &str, settings: &str, display: bool) -> Option<PageBackgroundWire> {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{background}</w:document>"#
        );
        let doc = XmlDoc::parse(&xml).unwrap();
        let settings_xml = format!(r#"<w:settings xmlns:w="{W}">{settings}</w:settings>"#);
        let settings = XmlDoc::parse(&settings_xml).unwrap();
        let mapping = ColorSchemeMapping::parse(settings.root_element());
        let theme = theme();
        let mut num_map = NumberingMap::default();
        let media = HashMap::from([("tex".to_string(), "word/media/tex.png".to_string())]);
        let mut diagnostics = Vec::new();
        parse_page_background(
            doc.root_element().first_element_child().unwrap(),
            display,
            PageBackgroundContext {
                mapping: &mapping,
                theme: &theme,
                style_map: &StyleMap::default(),
                num_map: &mut num_map,
                media_map: &media,
                chart_map: &ChartMap::default(),
                rel_map: &HashMap::new(),
            },
            &mut diagnostics,
        )
    }

    #[test]
    fn hex_colour_paints_only_while_background_shapes_are_displayed() {
        let painted = parse(r#"<w:background w:color="22551b"/>"#, "", true).unwrap();
        assert_eq!(painted.color.as_deref(), Some("22551B"));
        assert!(painted.fills.is_empty());
        assert!(parse(r#"<w:background w:color="22551B"/>"#, "", false).is_none());
        // `auto` (the attribute default) and an absent colour paint nothing.
        assert!(parse(r#"<w:background w:color="auto"/>"#, "", true).is_none());
        assert!(parse(r#"<w:background/>"#, "", true).is_none());
    }

    #[test]
    fn theme_colour_wins_and_applies_the_spec_tint_and_shade_examples() {
        // §17.2.1 examples: accent 4F81BD tinted 0x99 is 95B3D7. Its shade
        // example (C0504D, 0xBF) rounds the intermediate HSL to (0.48, 0.53)
        // and prints 943634; the exact formula gives 953735. The stored
        // `w:color` is ignored.
        let tint = parse(
            r#"<w:background w:color="000000" w:themeColor="accent1" w:themeTint="99"/>"#,
            "",
            true,
        );
        assert_eq!(tint.unwrap().color.as_deref(), Some("95B3D7"));
        let shade = parse(
            r#"<w:background w:themeColor="accent2" w:themeShade="BF" w:themeTint="zz"/>"#,
            "",
            true,
        );
        assert_eq!(shade.unwrap().color.as_deref(), Some("953735"));
        // Tint supersedes shade.
        let both = parse(
            r#"<w:background w:themeColor="accent1" w:themeTint="99" w:themeShade="BF"/>"#,
            "",
            true,
        );
        assert_eq!(both.unwrap().color.as_deref(), Some("95B3D7"));
    }

    #[test]
    fn theme_colour_reads_the_settings_colour_scheme_mapping() {
        let default = parse(
            r#"<w:background w:color="FFFFFF" w:themeColor="background1"/>"#,
            "",
            true,
        );
        assert_eq!(default.unwrap().color.as_deref(), Some("FFFFFF"));
        let mapped = parse(
            r#"<w:background w:color="FFFFFF" w:themeColor="background1"/>"#,
            r#"<w:clrSchemeMapping w:bg1="dark2"/>"#,
            true,
        );
        assert_eq!(mapped.unwrap().color.as_deref(), Some("1F497D"));
        // `none` names no theme colour, so the RGB colour applies.
        let none = parse(
            r#"<w:background w:color="ABCDEF" w:themeColor="none"/>"#,
            "",
            true,
        );
        assert_eq!(none.unwrap().color.as_deref(), Some("ABCDEF"));
    }

    #[test]
    fn vml_background_fills_layer_above_the_colour() {
        let gradient = parse(
            r##"<w:background w:color="FFF0E1"><v:background fillcolor="#fff0e1 [3201]"><v:fill type="gradient" color2="#4f81bd [3204]" angle="90"/></v:background></w:background>"##,
            "",
            true,
        )
        .unwrap();
        assert_eq!(gradient.color.as_deref(), Some("FFF0E1"));
        let fill = serde_json::to_value(&gradient.fills).unwrap();
        assert_eq!(fill[0]["fillType"], "gradient");
        assert_eq!(fill[0]["stops"][0]["color"], "fff0e1");
        assert_eq!(fill[0]["stops"][1]["color"], "4f81bd");

        let tile = parse(
            r#"<w:background><v:background><v:fill r:id="tex" type="tile"/></v:background></w:background>"#,
            "",
            true,
        )
        .unwrap();
        assert!(tile.color.is_none());
        let fill = serde_json::to_value(&tile.fills).unwrap();
        assert_eq!(fill[0]["fillType"], "image");
        assert_eq!(fill[0]["imagePath"], "word/media/tex.png");
        assert_eq!(fill[0]["vml"]["repeat"], true);

        let translucent = parse(
            r##"<w:background><v:background fillcolor="red"><v:fill opacity=".5"/></v:background></w:background>"##,
            "",
            true,
        )
        .unwrap();
        assert_eq!(translucent.fills.len(), 1);
        assert!(matches!(&translucent.fills[0], ShapeFill::Solid { color } if color == "ff000080"));
        let unfilled = parse(
            r#"<w:background w:color="00FF00"><v:background filled="f"/></w:background>"#,
            "",
            true,
        )
        .unwrap();
        assert!(unfilled.fills.is_empty());
    }

    fn package(background: &str, settings: &str) -> Vec<u8> {
        use std::io::{Cursor, Write};
        let theme = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="t"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>"#;
        let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rPic" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/bg.png"/></Relationships>"#;
        let document = format!(
            r#"<w:document xmlns:w="{W}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">{background}<w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#
        );
        let settings = format!(r#"<w:settings xmlns:w="{W}">{settings}</w:settings>"#);
        let mut bytes = Vec::new();
        {
            let mut archive = zip::ZipWriter::new(Cursor::new(&mut bytes));
            crate::parser::write_test_content_types(&mut archive);
            let options = zip::write::SimpleFileOptions::default();
            for (path, content) in [
                ("word/document.xml", document.as_bytes()),
                ("word/_rels/document.xml.rels", rels.as_bytes()),
                ("word/settings.xml", settings.as_bytes()),
                ("word/theme/theme1.xml", theme.as_bytes()),
                ("word/media/bg.png", b"png".as_slice()),
            ] {
                archive.start_file(path, options).unwrap();
                archive.write_all(content).unwrap();
            }
            archive.finish().unwrap();
        }
        bytes
    }

    /// Parse through the production bounded cursor and the whole-part DOM
    /// path; both must acquire the same background wire.
    fn parse_package(background: &str, settings: &str) -> serde_json::Value {
        let bytes = package(background, settings);
        let mut zip = open_zip(bytes.clone()).unwrap();
        let streamed = zip.run_operation("background", parse_streamed).unwrap();
        let mut zip = open_zip(bytes).unwrap();
        let native = zip
            .run_operation("background", super::super::parse)
            .unwrap();
        let streamed = serde_json::to_value(&streamed).unwrap();
        let native = serde_json::to_value(&native).unwrap();
        assert_eq!(streamed["__pageBackground"], native["__pageBackground"]);
        assert_eq!(streamed["diagnostics"], native["diagnostics"]);
        streamed
    }

    #[test]
    fn package_background_follows_display_background_shape_on_off_semantics() {
        let background =
            r#"<w:background w:color="FFF0E1" w:themeColor="accent1" w:themeTint="99"/>"#;
        for (settings, expected) in [
            ("<w:displayBackgroundShape/>", Some("95B3D7")),
            (r#"<w:displayBackgroundShape w:val="1"/>"#, Some("95B3D7")),
            (r#"<w:displayBackgroundShape w:val="false"/>"#, None),
            (r#"<w:displayBackgroundShape w:val="0"/>"#, None),
            ("", None),
        ] {
            let document = parse_package(background, settings);
            assert_eq!(
                document["__pageBackground"]["color"].as_str(),
                expected,
                "{settings}"
            );
        }
        // No background element: nothing, even when displayed.
        assert!(parse_package("", "<w:displayBackgroundShape/>")
            .get("__pageBackground")
            .is_none());
    }

    #[test]
    fn package_background_resolves_vml_and_drawingml_children_in_order() {
        let picture = r#"<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="bg"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="bg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rPic"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#;
        let document = parse_package(
            &format!(
                r##"<w:background w:color="22551B"><v:background id="_x0000_s1025" o:bwmode="white" fillcolor="#22551b [3204]" o:targetscreensize="1024,768"><v:fill r:id="rPic" o:title="tex" type="frame"/></v:background>{picture}</w:background>"##
            ),
            "<w:displayBackgroundShape/>",
        );
        let background = &document["__pageBackground"];
        assert_eq!(background["color"], "22551B");
        assert_eq!(background["fills"][0]["fillType"], "image");
        assert_eq!(background["fills"][0]["imagePath"], "word/media/bg.png");
        assert_eq!(background["fills"][0]["vml"]["repeat"], false);
        assert_eq!(background["fills"][1]["fillType"], "image");
        assert!(background["fills"][1].get("vml").is_none());
        assert!(document.get("diagnostics").is_none());
    }

    #[test]
    fn package_background_reports_a_drawing_that_is_not_a_page_fill() {
        let chart_like = r#"<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="g"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="missing"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#;
        let document = parse_package(
            &format!(r#"<w:background w:color="22551B">{chart_like}</w:background>"#),
            "<w:displayBackgroundShape/>",
        );
        assert_eq!(document["__pageBackground"]["color"], "22551B");
        assert_eq!(
            document["diagnostics"][0]["code"],
            "UNSUPPORTED_BACKGROUND_DRAWING"
        );
    }
}
