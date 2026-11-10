//! Legacy VML (`urn:schemas-microsoft-com:vml`, ECMA-376 Part 4 §19) helpers
//! shared by the parsers that read an OLE object's preview picture from VML.
//!
//! Word keeps that preview inline (`<w:object><v:shape><v:imagedata>`), Excel
//! in the sheet's `vmlDrawingN.vml` keyed by `oleObject@shapeId`, and
//! PowerPoint in the slide's `vmlDrawingN.vml` keyed by `oleObj@spid`. The
//! shape lookup differs per host; the `<v:imagedata>` grammar does not.

use crate::blip::SrcRect;
use crate::ns::relationships;
use crate::rels::{RelTarget, TargetMode};
use std::collections::BTreeMap;

/// `urn:schemas-microsoft-com:office:office` (the `o:` prefix).
pub const OFFICE_NS: &str = "urn:schemas-microsoft-com:office:office";

/// The relationship id naming the picture of a `<v:imagedata>` preview
/// (Part 4 §19.1.2.11): Office writes `o:relid`; `r:id` (Transitional or
/// Strict) is the other relationship-valued form. A namespace-less `relid` is
/// accepted for producers that drop the prefix. The id resolves through the
/// rels of the part that holds the VML, not the host part.
///
/// The VML `id` attribute is an object identifier, never a relationship id, so
/// it is not consulted here.
pub fn imagedata_rel_id<'a>(imagedata: roxmltree::Node<'a, '_>) -> Option<&'a str> {
    imagedata
        .attribute((OFFICE_NS, "relid"))
        .or_else(|| imagedata.attribute("relid"))
        .or_else(|| imagedata.attribute((relationships::TRANSITIONAL, "id")))
        .or_else(|| imagedata.attribute((relationships::STRICT, "id")))
}

/// OPC image relationship type (ECMA-376 Part 1 §15.2.14), Transitional and
/// Strict.
pub const IMAGE_REL_TYPES: &[&str] = &[
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image",
];

/// Resolve a `<v:imagedata>` preview to the image part it names, through the
/// rels of `vml_part` (the part holding the VML). Only an *internal*
/// relationship whose type is the image relationship qualifies: an External
/// relationship names no package part (ECMA-376 Part 2 §9.3), even when its
/// `Target` happens to look like a relative part path, and a relationship of
/// any other type names a part that is not a picture.
pub fn imagedata_image_part(
    imagedata: roxmltree::Node,
    rels: &BTreeMap<String, RelTarget>,
    vml_part: &str,
) -> Option<String> {
    let rel = rels.get(imagedata_rel_id(imagedata)?)?;
    if rel.mode != TargetMode::Internal {
        return None;
    }
    if !IMAGE_REL_TYPES.contains(&rel.relationship_type.as_deref()?) {
        return None;
    }
    rel.resolve_part(vml_part)
}

/// `<v:imagedata>` crop insets (Part 4 §19.1.2.11 `cropleft`/`croptop`/
/// `cropright`/`cropbottom`) as a DrawingML-style source rectangle. The values
/// are fractions, percentages, or signed 16.16 fixed-point values (`f`
/// suffix). Negative insets are preserved for the srcRect pipeline; clamping
/// here would discard authored outsets. `None` when no edge is cropped.
pub fn imagedata_crop(imagedata: roxmltree::Node) -> Option<SrcRect> {
    let fraction = |name| {
        imagedata
            .attribute(name)
            .and_then(parse_vml_fraction)
            .filter(|n| n.is_finite())
            .unwrap_or(0.0)
    };
    let crop = SrcRect {
        l: fraction("cropleft"),
        t: fraction("croptop"),
        r: fraction("cropright"),
        b: fraction("cropbottom"),
    };
    (crop.l != 0.0 || crop.t != 0.0 || crop.r != 0.0 || crop.b != 0.0).then_some(crop)
}

/// VML fraction grammar: `N%`, `Nf` (16.16 fixed point), or a plain decimal.
fn parse_vml_fraction(raw: &str) -> Option<f64> {
    let raw = raw.trim();
    if let Some(num) = raw.strip_suffix('%') {
        return num.trim().parse::<f64>().ok().map(|n| n / 100.0);
    }
    if let Some(num) = raw.strip_suffix('f') {
        return num.trim().parse::<f64>().ok().map(|n| n / 65536.0);
    }
    raw.parse::<f64>().ok()
}

/// The `<v:imagedata>` of the VML shape whose `o:spid` equals `spid`.
///
/// PresentationML connects `<p:oleObj spid>` (ECMA-376 Part 1 §19.3.2.4) to the
/// legacy drawing shape by its Office shape id `o:spid`; the VML `id` is the
/// shape's name-derived object id and is not the key.
pub fn imagedata_for_spid<'a, 'i>(
    vml: &'a roxmltree::Document<'i>,
    spid: &str,
) -> Option<roxmltree::Node<'a, 'i>> {
    let shape = vml.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "shape"
            && n.attribute((OFFICE_NS, "spid")) == Some(spid)
    })?;
    shape
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "imagedata")
}

#[cfg(test)]
mod tests {
    use super::*;

    const VML: &str = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <v:shape id="_x0000_s1026" o:spid="_x0000_s2000"><v:imagedata o:relid="rIdWrong"/></v:shape>
      <v:shape id="Object_x0020_3" o:spid="_x0000_s1026">
        <v:imagedata o:relid="rId1" cropleft="10%" croptop="6554f" cropright=".25" o:title=""/>
      </v:shape>
      <v:shape id="Object_x0020_4" o:spid="_x0000_s1027"><v:imagedata r:id="rId2"/></v:shape>
      <v:shape id="Object_x0020_5" o:spid="_x0000_s1028"><v:imagedata id="rId9"/></v:shape>
    </xml>"##;

    #[test]
    fn spid_lookup_uses_office_spid_not_vml_id() {
        let doc = roxmltree::Document::parse(VML).unwrap();
        let data = imagedata_for_spid(&doc, "_x0000_s1026").unwrap();
        assert_eq!(imagedata_rel_id(data), Some("rId1"));
        let crop = imagedata_crop(data).unwrap();
        assert_eq!((crop.l, crop.r, crop.b), (0.1, 0.25, 0.0));
        assert!((crop.t - 6554.0 / 65536.0).abs() < 1e-12);

        let data = imagedata_for_spid(&doc, "_x0000_s1027").unwrap();
        assert_eq!(imagedata_rel_id(data), Some("rId2"));
        assert_eq!(imagedata_crop(data), None);

        // The VML object `id` is never a relationship id.
        let data = imagedata_for_spid(&doc, "_x0000_s1028").unwrap();
        assert_eq!(imagedata_rel_id(data), None);

        assert!(imagedata_for_spid(&doc, "_x0000_s9999").is_none());
    }

    #[test]
    fn imagedata_image_part_requires_internal_image_relationship() {
        let doc = roxmltree::Document::parse(VML).unwrap();
        let data = imagedata_for_spid(&doc, "_x0000_s1026").unwrap(); // o:relid="rId1"
        let rels = |xml: &str| {
            crate::rels::parse_rels(&format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{xml}</Relationships>"#
            ))
        };
        let part = "ppt/drawings/vmlDrawing1.vml";
        for ty in IMAGE_REL_TYPES {
            let map = rels(&format!(r#"<Relationship Id="rId1" Type="{ty}" Target="../media/image1.png"/>"#));
            assert_eq!(imagedata_image_part(data, &map, part).as_deref(), Some("ppt/media/image1.png"));
        }
        // External: the relative-looking Target must not resolve to a package part.
        let external = rels(&format!(
            r#"<Relationship Id="rId1" Type="{}" Target="../media/image1.png" TargetMode="External"/>"#,
            IMAGE_REL_TYPES[0]
        ));
        assert_eq!(imagedata_image_part(data, &external, part), None);
        // Wrong type, missing type, missing relationship.
        let ole = rels(r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../media/image1.png"/>"#);
        assert_eq!(imagedata_image_part(data, &ole, part), None);
        let untyped = rels(r#"<Relationship Id="rId1" Target="../media/image1.png"/>"#);
        assert_eq!(imagedata_image_part(data, &untyped, part), None);
        assert_eq!(imagedata_image_part(data, &rels(""), part), None);
    }
}
