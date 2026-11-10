//! Effective VML image properties (ECMA-376 Part 4 §§19.1.2.10/.11/.19/.20).
//! An instance overrides only properties it supplies, including when its
//! imagedata child contains editor metadata but no resource or crop facts.
use super::*;

fn node<'a, 'input>(host: roxmltree::Node<'a, 'input>) -> Option<roxmltree::Node<'a, 'input>> {
    if host.has_tag_name(("urn:schemas-microsoft-com:vml", "image")) {
        Some(host)
    } else {
        vml_direct_child(host, "imagedata")
    }
}

pub(super) fn present(shape: roxmltree::Node, definition: Option<roxmltree::Node>) -> bool {
    node(shape).is_some() || definition.and_then(node).is_some()
}

pub(super) fn property<'a>(
    shape: roxmltree::Node<'a, '_>,
    definition: Option<roxmltree::Node<'a, '_>>,
    name: &str,
) -> Option<&'a str> {
    let read = |host| {
        node(host)
            .and_then(|n| n.attribute(name))
            // AG_ShapeAttributes also exposes chromakey on the drawable.
            .or_else(|| {
                (name == "chromakey")
                    .then(|| host.attribute(name))
                    .flatten()
            })
    };
    read(shape).or_else(|| definition.and_then(read))
}

pub(super) fn incomplete(shape: roxmltree::Node, definition: Option<roxmltree::Node>) -> bool {
    let attr = |name| property(shape, definition, name);
    // Part 4 specifies endpoints and names, but does not define the transfer
    // functions/order for gamma or the bilevel threshold. Do not guess Office
    // pixel math. Retain the available picture and disclose these effects
    // until varied Office-produced transfer/order controls exist. MS-OE376
    // §2.1.1701(a) further distinguishes Office bilevel from the ECMA
    // per-channel conversion. Neither interpretation is silently substituted.
    // gain/blacklevel are disclosed only where `luminance` withholds them.
    attr("gamma").is_some_and(|raw| parse_vml_opacity(raw) != Some(1.0))
        || (authored_luminance(shape, definition) && luminance(shape, definition).is_none())
        || attr("bilevel").is_some_and(|raw| parse_vml_true_false(raw) != Some(false))
        || attr("grayscale").is_some_and(|raw| parse_vml_true_false(raw).is_none())
        || attr("chromakey").is_some_and(|raw| vml_color_hex6(raw).is_none())
}

fn authored_luminance(shape: roxmltree::Node, definition: Option<roxmltree::Node>) -> bool {
    [("gain", 1.0), ("blacklevel", 0.0)]
        .iter()
        .any(|(name, default)| {
            property(shape, definition, name)
                .is_some_and(|raw| parse_vml_opacity(raw) != Some(*default))
        })
}

/// `gain` (image intensity) and `blacklevel` (image brightness), Part 4
/// §19.1.2.11, as the shared `<a:lum>` picture correction (§20.1.8.42).
///
/// Part 4 names both attributes without a transfer function. They are the
/// VML serialization of the Office Art picture properties pictureContrast and
/// pictureBrightness (MS-ODRAW §§2.3.23.11-12) and carry the same 16.16 raw
/// values: Word's Watermark dialog writes Washout as gain="19661f"
/// blacklevel="22938f" (0x4ccd, 0x599a); PowerPoint converts the binary pair
/// 0x4ccd/0x599a to `<a:lum bright="70000" contrast="-70000"/>`, which is also
/// what Word writes for Washout on a DrawingML picture. The shared luminance
/// transfer was measured from PowerPoint PDF exports of a gray ramp under a
/// bright x contrast grid (0, +-35%, +-70%, +-100%), for `<a:lum>` and for the
/// binary properties, the latter with the stored contrast used directly as
/// the slope k (core `blip-effects.ts`; legacy PPT picture projection). So
/// gain is k (contrast = k - 1 for k <= 1, 1 - 1/k above) and blacklevel is
/// half the brightness fraction (raw 0x8000 = bright 1).
///
/// Office's order relative to gamma, grayscale, bilevel and chromakey is
/// unmeasured, so such a combination withholds the correction and stays
/// disclosed through `incomplete` rather than guessing an order.
fn luminance(shape: roxmltree::Node, definition: Option<roxmltree::Node>) -> Option<BlipEffect> {
    let attr = |name| property(shape, definition, name);
    let value = |name, default: f64| match attr(name) {
        None => Some(default),
        Some(raw) => parse_vml_opacity(raw).filter(|n| n.is_finite()),
    };
    let gain = value("gain", 1.0)?;
    let blacklevel = value("blacklevel", 0.0)?;
    if gain == 1.0 && blacklevel == 0.0 {
        return None;
    }
    let combined = attr("gamma").is_some_and(|raw| parse_vml_opacity(raw) != Some(1.0))
        || attr("bilevel").is_some_and(|raw| parse_vml_true_false(raw) != Some(false))
        || attr("grayscale").is_some_and(|raw| parse_vml_true_false(raw) != Some(false))
        || attr("chromakey").is_some();
    if combined {
        return None;
    }
    let k = gain.max(0.0);
    let contrast = if k <= 1.0 { k - 1.0 } else { 1.0 - 1.0 / k };
    Some(BlipEffect::Luminance {
        bright: (blacklevel * 2.0).clamp(-1.0, 1.0),
        contrast: contrast.clamp(-1.0, 1.0),
    })
}

/// Resolve the effective imagedata pixel effects onto a picture built by any
/// VML image path (grouped leaf, standalone pict, or OLE preview).
pub(super) fn apply_effects(
    picture: &mut ImageRun,
    shape: roxmltree::Node,
    definition: Option<roxmltree::Node>,
) {
    picture.color_replace_from = property(shape, definition, "chromakey").and_then(vml_color_hex6);
    // VML requires grayscale without prescribing a luminance weighting. Use
    // the shared library grayscale ramp, preserving source alpha; this makes
    // no claim about Office's exact intermediate grayscale byte values.
    picture.duotone = property(shape, definition, "grayscale")
        .and_then(parse_vml_true_false)
        .filter(|on| *on)
        .map(|_| Duotone {
            clr1: "000000".into(),
            clr2: "FFFFFF".into(),
        });
    picture.blip_effects = luminance(shape, definition).into_iter().collect();
}

pub(super) fn acquire(
    shape: roxmltree::Node,
    definition: Option<roxmltree::Node>,
    media: &HashMap<String, String>,
) -> Option<ImageRun> {
    // Source aliases select one reference at each cascade level. An authored
    // instance src must supersede an inherited relationship, and an unresolved
    // instance reference must not silently resurrect the template picture.
    let reference = |host| {
        node(host).filter(|n| vml_image_relationship(*n).is_some() || n.attribute("src").is_some())
    };
    let image = reference(shape).or_else(|| definition.and_then(reference))?;
    let path = vml_image_source(image, media)?;
    let style = shape.attribute("style").unwrap_or("");
    let coordinates = VmlStyleCoordinates::for_node(shape);
    let mut picture = vml_image_run(
        shape,
        definition,
        path,
        coordinates.css_length(style, "width")?,
        coordinates.css_length(style, "height")?,
        image,
    );
    let fraction = |name| {
        property(shape, definition, name)
            .and_then(|raw| {
                raw.trim()
                    .strip_suffix('%')
                    .and_then(|s| s.parse::<f64>().ok())
                    .map(|n| n / 100.0)
                    .or_else(|| parse_vml_opacity(raw))
            })
            .filter(|n| n.is_finite())
            .unwrap_or(0.0)
    };
    let crop = SrcRect {
        l: fraction("cropleft"),
        t: fraction("croptop"),
        r: fraction("cropright"),
        b: fraction("cropbottom"),
    };
    picture.src_rect =
        (crop.l != 0.0 || crop.t != 0.0 || crop.r != 0.0 || crop.b != 0.0).then_some(crop);
    Some(picture)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn source_crop_and_effects_cascade_individually_including_explicit_resets() {
        let doc = XmlDoc::parse(r##"<v:group xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <v:shapetype><v:imagedata r:id="pic" cropleft="50%" croptop="25%" grayscale="t" chromakey="red" gamma="2"/></v:shapetype>
          <v:shape style="width:20;height:10"><v:imagedata title="metadata" cropleft="0" grayscale="f" gamma="1"/></v:shape>
          <v:image style="width:20;height:10" src="media/instance.png"/>
          <v:image style="width:20;height:10" src="media/reset.png" grayscale="f" gamma="100%"/>
        </v:group>"##).unwrap();
        let nodes: Vec<_> = doc
            .root_element()
            .children()
            .filter(|n| n.is_element())
            .collect();
        let media = HashMap::from([("pic".into(), "word/media/template.png".into())]);
        let picture = acquire(nodes[1], Some(nodes[0]), &media).unwrap();
        assert_eq!(picture.image_path, "word/media/template.png");
        let crop = picture.src_rect.unwrap();
        assert_eq!((crop.l, crop.t), (0.0, 0.25));
        assert!(picture.duotone.is_none());
        assert_eq!(picture.color_replace_from.as_deref(), Some("ff0000"));
        assert!(!incomplete(nodes[1], Some(nodes[0])));
        let picture = acquire(nodes[2], Some(nodes[0]), &media).unwrap();
        assert_eq!(picture.image_path, "word/media/instance.png");
        assert!(picture.duotone.is_some());
        assert!(incomplete(nodes[2], Some(nodes[0])));
        let reset = acquire(nodes[3], Some(nodes[0]), &media).unwrap();
        assert!(reset.duotone.is_none());
        assert!(!vml_geometry::incomplete_leaf(nodes[3], Some(nodes[0])));
    }

    #[test]
    fn gain_and_blacklevel_map_to_luminance_unless_combined_with_unordered_effects() {
        let doc = XmlDoc::parse(
            r##"<v:group xmlns:v="urn:schemas-microsoft-com:vml">
          <v:shapetype><v:imagedata src="media/a.png" gain="2"/></v:shapetype>
          <v:image style="width:20;height:10" src="media/a.png" blacklevel="-0.25"/>
          <v:image style="width:20;height:10" src="media/a.png" gain="19661f" grayscale="t"/>
          <v:image style="width:20;height:10" src="media/a.png" gain="1" blacklevel="0"/>
        </v:group>"##,
        )
        .unwrap();
        let nodes: Vec<_> = doc
            .root_element()
            .children()
            .filter(|n| n.is_element())
            .collect();
        let media = HashMap::new();
        // Inherited gain 2 (slope k = 2 -> contrast 1 - 1/k) with instance
        // blacklevel -0.25 (bright -0.5).
        let picture = acquire(nodes[1], Some(nodes[0]), &media).unwrap();
        assert_eq!(
            picture.blip_effects,
            vec![BlipEffect::Luminance {
                bright: -0.5,
                contrast: 0.5
            }]
        );
        assert!(!incomplete(nodes[1], Some(nodes[0])));
        // Unmeasured order against grayscale: withheld and disclosed.
        let combined = acquire(nodes[2], None, &media).unwrap();
        assert!(combined.blip_effects.is_empty());
        assert!(combined.duotone.is_some());
        assert!(incomplete(nodes[2], None));
        // Explicit defaults are no correction.
        let neutral = acquire(nodes[3], None, &media).unwrap();
        assert!(neutral.blip_effects.is_empty());
        assert!(!incomplete(nodes[3], None));
    }
}
