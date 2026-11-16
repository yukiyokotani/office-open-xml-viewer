//! OfficeArt shape fills (MS-ODRAW 2.3.7) as the XLSX model's shape fills,
//! as the XLSX parser reads the DrawingML fills of ECMA-376 20.1.8: solid,
//! gradient (`a:gradFill`), pattern (`a:pattFill`) and picture (`a:blipFill`).
//!
//! The reference for each mapping is Excel's own output of a workbook that
//! Excel authored with that fill (the fill colors, focus, angle and fill-to
//! rectangle of each gradient type against the gradient Excel draws):
//!
//! - msofillShade and msofillShadeScale are linear gradients through the
//!   shared OfficeArt rules (`officeart::paint::Paint::linear_shade`);
//! - msofillShadeCenter is a rectangular path gradient (`a:path path="rect"`)
//!   that starts at the fillTo rectangle (MS-ODRAW 2.3.7.16-19, fractions of
//!   the shape) with the first color of the shade and ends at the shape's
//!   bounds. Excel writes all four fillTo values 0.5 for "from center" and
//!   leaves them 0 for "from corner";
//! - msofillPattern paints its 8 x 8 bitmap in the fill and back colors: the
//!   preset pattern with those cells (see `pattern`);
//! - msofillTexture repeats its picture at the picture's own size from the
//!   shape's top-left corner and msofillPicture stretches it over the shape.

use super::{color, pattern, Table};
use crate::officeart::{gradient, paint::Paint};
use crate::xls::{drawing_anchors::ShapeSource, styles, unsupported};
use ooxml_common::fill::{FillRect, GradStop};

/// A fill's picture (fillBlip, a one-based BLIP store index) whose media is
/// read with the sheet pictures: the bitmap that names a pattern fill's
/// preset, or the picture of an image fill, whose MIME type is still unset.
pub(super) struct Blip {
    pub(super) store: u32,
}

/// Stops of one shade as model gradient stops: `RRGGBB`, or `RRGGBBAA` when
/// an opacity applies (the convention of the shared gradient parser).
fn stops(
    shade: &crate::officeart::paint::LinearShade,
    styles: &styles::Styles<'_>,
) -> Vec<GradStop> {
    shade
        .projection
        .stops
        .iter()
        .zip(&shade.alphas)
        .map(|(stop, alpha)| {
            let [r, g, b] = color(stop.color, [0xff, 0xff, 0xff], styles);
            let mut text = format!("{r:02X}{g:02X}{b:02X}");
            if *alpha != 65_536 {
                let byte = ((u64::from(*alpha) * 255 + 32_768) / 65_536) as u8;
                text.push_str(&format!("{byte:02X}"));
            }
            GradStop {
                position: stop.position(),
                color: text,
            }
        })
        .collect()
}

/// The fill of a filled shape, and the picture it still waits for.
pub(super) fn project(
    paint: &Paint,
    table: &Table,
    source: &ShapeSource,
    styles: &styles::Styles<'_>,
) -> Result<(xlsx_model::ShapeFill, Option<Blip>), String> {
    let non_solid = || unsupported("non-solid XLS shape fills are not projected");
    // fUseRect draws the fill in its own rectangle (fillRect*), which no
    // model fill states.
    if paint.fill_rect == Some(true) {
        return Err(non_solid());
    }
    let kind = paint.fill_type.unwrap_or(0);
    let front = color(
        paint.fill.unwrap_or(0x00ff_ffff),
        [0xff, 0xff, 0xff],
        styles,
    );
    let hex = |[r, g, b]: [u8; 3]| format!("{r:02X}{g:02X}{b:02X}");
    match kind {
        0 => Ok((
            xlsx_model::ShapeFill::Solid {
                color: super::rgb(front, paint.fill_alpha.unwrap_or(65_536)),
            },
            None,
        )),
        1..=3 => {
            let store = paint
                .fill_blip
                .filter(|store| *store != 0)
                .ok_or_else(|| crate::rejection::corrupt("XLS picture fill without its picture"))?;
            let alpha = paint.fill_alpha.filter(|alpha| *alpha != 65_536);
            if kind == 1 {
                let back = color(
                    paint.fill_back.unwrap_or(0x00ff_ffff),
                    [0xff, 0xff, 0xff],
                    styles,
                );
                // Named once the bitmap is read (`Shapes::attach_images`).
                return Ok((
                    xlsx_model::ShapeFill::Pattern {
                        fg: hex(front),
                        bg: hex(back),
                        preset: String::new(),
                    },
                    Some(Blip { store }),
                ));
            }
            // A texture whose tile has an authored size or origin
            // (fillWidth/fillHeight, fillOrigin*) is drawn at the picture's
            // own size from the top-left corner.
            if kind == 2
                && (paint.fill_width.is_some_and(|size| size != 0)
                    || paint.fill_height.is_some_and(|size| size != 0)
                    || paint.fill_origins.iter().any(|v| v.is_some_and(|v| v != 0)))
            {
                crate::rejection::approximated("non-solid XLS shape fills are not projected", ());
            }
            Ok((
                xlsx_model::ShapeFill::Image {
                    image_path: format!("legacy-xls/image/{store}"),
                    mime_type: String::new(),
                    layout: if kind == 2 {
                        xlsx_model::ShapeImageLayout::Tile
                    } else {
                        xlsx_model::ShapeImageLayout::Stretch
                    },
                    width: None,
                    height: None,
                    align: [0.0, 0.0],
                    offset: [0.0, 0.0],
                    shift: None,
                    alpha: alpha.map(|alpha| f64::from(alpha) / 65_536.0),
                },
                Some(Blip { store }),
            ))
        }
        4 | 5 | 7 => {
            let mut colors = gradient::Borrowed::default();
            if let Some((_, data)) = source.complex.iter().find(|(id, _)| *id == 0x197) {
                colors.set(data.as_slice());
            } else if let Some((_, value)) = source
                .properties
                .iter()
                .find(|(opid, _)| *opid & 0xbfff == 0x197)
            {
                // A scalar fillShadeColors holds no array.
                colors.scalar(*value);
            }
            // The center shade takes its stops from the same rules as the
            // linear shades; only its path differs.
            let mut linear = *paint;
            if kind == 5 {
                linear.fill_type = Some(7);
                linear.fill_angle = None;
            }
            let mut work = 100_000usize;
            let mut bytes = 1 << 20;
            let shade = linear
                .linear_shade(&colors, true, &mut work, &mut bytes)?
                .ok_or_else(non_solid)?;
            let stops = stops(&shade, styles);
            let fill = if kind == 5 {
                // fillToLeft/Top/Right/Bottom: 16.16 fractions of the shape.
                let to = |id: u16| {
                    f64::from(table.values.get(&id).copied().unwrap_or(0) as i32) / 65_536.0
                };
                let (left, top, right, bottom) = (to(0x18d), to(0x18e), to(0x18f), to(0x190));
                xlsx_model::ShapeFill::Gradient {
                    stops,
                    angle: 0.0,
                    grad_type: "radial".into(),
                    scaled: None,
                    path: Some("rect".into()),
                    // ECMA-376 20.1.8.31 fillToRect: insets from each edge.
                    fill_to_rect: Some(FillRect {
                        l: left,
                        t: top,
                        r: 1.0 - right,
                        b: 1.0 - bottom,
                    }),
                    tile_rect: None,
                    flip: None,
                    rot_with_shape: Some(shade.rotate_with_shape),
                }
            } else {
                xlsx_model::ShapeFill::Gradient {
                    stops,
                    angle: shade.projection.angle.degrees(),
                    grad_type: "linear".into(),
                    scaled: Some(shade.scaled),
                    path: None,
                    fill_to_rect: None,
                    tile_rect: None,
                    flip: None,
                    rot_with_shape: Some(shade.rotate_with_shape),
                }
            };
            Ok((fill, None))
        }
        // msofillShadeShape, msofillShadeTitle and msofillBackground.
        _ => Err(non_solid()),
    }
}

/// The single color the model also keeps for a fill, as the XLSX parser
/// derives it (`xlsx_fill_fallback_color`).
pub(super) fn fallback_color(fill: &xlsx_model::ShapeFill) -> Option<String> {
    match fill {
        xlsx_model::ShapeFill::Solid { color } => Some(color.clone()),
        xlsx_model::ShapeFill::Gradient { stops, .. } => stops
            .iter()
            .rev()
            .find(|stop| !stop.color.ends_with("00"))
            .or_else(|| stops.last())
            .map(|stop| format!("#{}", stop.color.to_uppercase())),
        xlsx_model::ShapeFill::Pattern { fg, .. } => Some(format!("#{}", fg.to_uppercase())),
        xlsx_model::ShapeFill::Hatch { color, .. } => Some(color.clone()),
        xlsx_model::ShapeFill::Image { .. } => None,
    }
}

/// Name a pattern fill by its bitmap, or keep the tone of a bitmap that is
/// no preset's artwork.
pub(super) fn pattern_preset(bytes: &[u8], extension: &str) -> String {
    const SITE: &str = "non-solid XLS shape fills are not projected";
    let cells = (extension == "png")
        .then(|| pattern::cells(bytes))
        .flatten();
    match cells {
        Some(cells) => match pattern::preset(cells) {
            Some(name) => name.to_owned(),
            None => crate::rejection::approximated(SITE, pattern::nearest_tone(cells)).to_owned(),
        },
        // A bitmap this reader does not decode: an even mix of both colors.
        None => crate::rejection::approximated(SITE, "pct50").to_owned(),
    }
}
