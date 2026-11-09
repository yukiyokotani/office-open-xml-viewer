//! Drawings of a table (ODF Part 3 §10): frames with images, embedded chart
//! objects and text boxes, and shapes, anchored to a cell (the shape is a
//! child of `table:table-cell`; `svg:x` / `svg:y` are offsets in that cell
//! and `table:end-cell-address` / `table:end-x` / `table:end-y` its end) or
//! to the table (`table:shapes`; `svg:x` / `svg:y` from the table origin).
//! Both become SpreadsheetML two-cell anchors (ECMA-376 Part 1 §20.5.2.33)
//! on the reader's pixel grid.

use super::table::Anchored;
use super::{parse_cell, Context, OdsSheet, MAX_COLS, MAX_DRAWINGS, MAX_ROWS};
use crate::odf_props::{get, hex_color, length_pt, percent, Style};
use crate::package::OdfPackage;
use crate::xml::{self, NS_DRAW, NS_SVG, NS_TEXT, NS_XLINK};
use std::collections::BTreeMap;
use xlsx_model::{ChartAnchor, DrawingAnchorTag, ImageAnchor, ShapeAnchor, ShapeInfo, Worksheet};

const NS_TABLE: &str = crate::xml::NS_TABLE;
pub const EMU_PER_PX: i64 = 9525;
const EMU_PER_PT: f64 = 12_700.0;

/// The rendered pixel grid of one worksheet.
pub struct Grid {
    columns: Vec<(u32, u32, u32)>,
    rows: BTreeMap<u32, u32>,
    default_row_px: u32,
}

impl Grid {
    pub fn new(columns: &[(u32, u32, u32)], worksheet: &Worksheet, default_row_px: u32) -> Self {
        let rows = worksheet
            .row_heights
            .iter()
            .map(|(row, pt)| (*row, super::pt_to_px(*pt)))
            .collect();
        Self {
            columns: columns.to_vec(),
            rows,
            default_row_px,
        }
    }

    fn col_px(&self, col: u32) -> u32 {
        self.columns
            .iter()
            .find(|(first, last, _)| *first <= col && col <= *last)
            .map_or(64, |(_, _, px)| *px)
    }

    /// Left edge of 0-based column `col0`.
    pub fn col_left(&self, col0: u32) -> i64 {
        let mut x = 0i64;
        let mut covered = 0u32;
        for (first, last, px) in &self.columns {
            if *first > col0 {
                break;
            }
            let end = (*last).min(col0);
            x += i64::from(end - first + 1) * i64::from(*px);
            covered = end;
        }
        x + i64::from(col0.saturating_sub(covered)) * 64
    }

    pub fn row_top(&self, row0: u32) -> i64 {
        let mut y = i64::from(row0) * i64::from(self.default_row_px);
        if row0 == 0 {
            return 0;
        }
        for (_, px) in self.rows.range(1..=row0) {
            y += i64::from(*px) - i64::from(self.default_row_px);
        }
        y
    }

    /// (0-based column, offset px) containing `x`.
    pub fn locate_x(&self, x: i64) -> (u32, i64) {
        let mut left = 0i64;
        for col in 1..=MAX_COLS {
            let width = i64::from(self.col_px(col));
            if x < left + width || col == MAX_COLS {
                return (col - 1, x - left);
            }
            left += width;
        }
        (MAX_COLS - 1, 0)
    }

    pub fn locate_y(&self, y: i64) -> (u32, i64) {
        // Binary search on the monotonic row offsets.
        let (mut lo, mut hi) = (0u32, MAX_ROWS - 1);
        while lo < hi {
            let mid = lo + (hi - lo).div_ceil(2);
            if self.row_top(mid) <= y {
                lo = mid;
            } else {
                hi = mid - 1;
            }
        }
        (lo, y - self.row_top(lo))
    }
}

/// An anchor rectangle in markers (0-based) and EMU offsets.
#[derive(Clone, Copy, Debug)]
pub struct Rect {
    pub from_col: u32,
    pub from_col_off: i64,
    pub from_row: u32,
    pub from_row_off: i64,
    pub to_col: u32,
    pub to_col_off: i64,
    pub to_row: u32,
    pub to_row_off: i64,
    pub width_emu: i64,
    pub height_emu: i64,
}

/// A chart frame awaiting its object (resolved once every table is read).
pub struct PendingChart {
    pub sheet: usize,
    pub rect: Rect,
    pub z_order: u64,
    pub object: String,
    pub replacement: Option<String>,
}

fn pt_attr(node: roxmltree::Node<'_, '_>, ns: &str, name: &str) -> Option<f64> {
    node.attribute((ns, name)).and_then(length_pt)
}

fn px(pt: f64) -> i64 {
    (pt * 96.0 / 72.0).round() as i64
}

/// `rotate (r) translate (x y)` (§19.229 draw:transform): the angle in
/// degrees clockwise and the translated origin in points.
fn transform(value: &str) -> Option<(f64, f64, f64)> {
    let mut angle = 0.0;
    let mut origin = None;
    let mut rest = value;
    while let Some(open) = rest.find('(') {
        let name = rest[..open].trim();
        let close = rest[open..].find(')')? + open;
        let args: Vec<&str> = rest[open + 1..close]
            .split(|c: char| c == ',' || c.is_whitespace())
            .filter(|s| !s.is_empty())
            .collect();
        match name {
            "rotate" => angle = -args.first()?.parse::<f64>().ok()?.to_degrees(),
            "translate" => {
                origin = Some((
                    length_pt(args.first()?)?,
                    args.get(1).and_then(|v| length_pt(v)).unwrap_or(0.0),
                ))
            }
            _ => {}
        }
        rest = &rest[close + 1..];
    }
    let (x, y) = origin?;
    Some((angle, x, y))
}

/// Place a drawing: its rectangle on the grid and its rotation.
fn place(
    node: roxmltree::Node<'_, '_>,
    cell: Option<(u32, u32)>,
    grid: &Grid,
    sheet_names: &[String],
) -> Option<(Rect, f64)> {
    let width = pt_attr(node, NS_SVG, "width");
    let height = pt_attr(node, NS_SVG, "height");
    let (mut x, mut y, mut w, mut h, mut rotation) = (
        pt_attr(node, NS_SVG, "x").unwrap_or(0.0),
        pt_attr(node, NS_SVG, "y").unwrap_or(0.0),
        width.unwrap_or(0.0),
        height.unwrap_or(0.0),
        0.0,
    );
    if width.is_none() || height.is_none() {
        // Lines and connectors: svg:x1/y1/x2/y2 (§19.570-§19.575).
        let (x1, y1, x2, y2) = (
            pt_attr(node, NS_SVG, "x1")?,
            pt_attr(node, NS_SVG, "y1")?,
            pt_attr(node, NS_SVG, "x2")?,
            pt_attr(node, NS_SVG, "y2")?,
        );
        x = x1.min(x2);
        y = y1.min(y2);
        w = (x2 - x1).abs();
        h = (y2 - y1).abs();
    }
    if let Some((angle, tx, ty)) = node.attribute((NS_DRAW, "transform")).and_then(transform) {
        // The translated origin is the rotated top-left corner; SpreadsheetML
        // anchors the unrotated box around the same centre.
        let radians = angle.to_radians();
        let (cx, cy) = (
            tx + (w / 2.0) * radians.cos() - (h / 2.0) * radians.sin(),
            ty + (w / 2.0) * radians.sin() + (h / 2.0) * radians.cos(),
        );
        x = cx - w / 2.0;
        y = cy - h / 2.0;
        rotation = angle;
    }
    let (origin_x, origin_y) = match cell {
        Some((row0, col0)) => (grid.col_left(col0), grid.row_top(row0)),
        None => (0, 0),
    };
    let left = origin_x + px(x);
    let top = origin_y + px(y);
    let (from_col, from_col_off) = grid.locate_x(left.max(0));
    let (from_row, from_row_off) = grid.locate_y(top.max(0));
    let end = node
        .attribute((NS_TABLE, "end-cell-address"))
        .and_then(|address| {
            let (sheet, reference) = super::split_sheet(address);
            let _ = (sheet, sheet_names);
            parse_cell(reference)
        })
        .filter(|_| rotation == 0.0);
    let (to_col, to_col_off, to_row, to_row_off) = match end {
        Some((row, col)) => (
            col - 1,
            px(pt_attr(node, NS_TABLE, "end-x").unwrap_or(0.0)),
            row - 1,
            px(pt_attr(node, NS_TABLE, "end-y").unwrap_or(0.0)),
        ),
        None => {
            let (c, co) = grid.locate_x((left + px(w)).max(0));
            let (r, ro) = grid.locate_y((top + px(h)).max(0));
            (c, co, r, ro)
        }
    };
    Some((
        Rect {
            from_col,
            from_col_off: from_col_off * EMU_PER_PX,
            from_row,
            from_row_off: from_row_off * EMU_PER_PX,
            to_col,
            to_col_off: to_col_off * EMU_PER_PX,
            to_row,
            to_row_off: to_row_off * EMU_PER_PX,
            width_emu: (w * EMU_PER_PT).round() as i64,
            height_emu: (h * EMU_PER_PT).round() as i64,
        },
        rotation,
    ))
}

fn z_index(node: roxmltree::Node<'_, '_>, fallback: usize) -> u64 {
    node.attribute((NS_DRAW, "z-index"))
        .and_then(|v| v.parse::<u64>().ok())
        .unwrap_or(fallback as u64)
}

/// A package path from an `xlink:href` (`./Pictures/a.png` → `Pictures/a.png`).
pub fn package_path(href: &str) -> Option<String> {
    if href.contains("://") || href.starts_with('/') || href.contains("..") {
        return None;
    }
    let path = href.trim_start_matches("./");
    (!path.is_empty()).then(|| path.to_string())
}

/// MIME type of an image the renderer can decode, from its first bytes.
pub fn image_mime(bytes: &[u8]) -> Option<&'static str> {
    if bytes.starts_with(b"\x89PNG") {
        Some("image/png")
    } else if bytes.starts_with(&[0xFF, 0xD8]) {
        Some("image/jpeg")
    } else if bytes.starts_with(b"GIF8") {
        Some("image/gif")
    } else if bytes.starts_with(b"BM") {
        Some("image/bmp")
    } else if bytes.starts_with(b"RIFF") && bytes.get(8..12) == Some(b"WEBP") {
        Some("image/webp")
    } else if bytes.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A]) || bytes.starts_with(&[1, 0, 9, 0]) {
        Some("image/x-wmf")
    } else if bytes.get(40..44) == Some(b" EMF") {
        Some("image/x-emf")
    } else if bytes.starts_with(b"II*\0") || bytes.starts_with(b"MM\0*") {
        Some("image/tiff")
    } else {
        let head = String::from_utf8_lossy(&bytes[..bytes.len().min(512)]);
        (head.contains("<svg") || head.trim_start().starts_with("<?xml") && head.contains("svg"))
            .then_some("image/svg+xml")
    }
}

/// Pixel size of a PNG, JPEG or GIF (for crop fractions).
fn image_size(bytes: &[u8]) -> Option<(f64, f64)> {
    let be32 = |at: usize| -> Option<u32> {
        Some(u32::from_be_bytes(bytes.get(at..at + 4)?.try_into().ok()?))
    };
    let be16 = |at: usize| -> Option<u16> {
        Some(u16::from_be_bytes(bytes.get(at..at + 2)?.try_into().ok()?))
    };
    if bytes.starts_with(b"\x89PNG") {
        return Some((f64::from(be32(16)?), f64::from(be32(20)?)));
    }
    if bytes.starts_with(b"GIF8") {
        let w = u16::from_le_bytes(bytes.get(6..8)?.try_into().ok()?);
        let h = u16::from_le_bytes(bytes.get(8..10)?.try_into().ok()?);
        return Some((f64::from(w), f64::from(h)));
    }
    if bytes.starts_with(&[0xFF, 0xD8]) {
        let mut at = 2usize;
        while at + 9 < bytes.len() {
            if bytes[at] != 0xFF {
                return None;
            }
            let marker = bytes[at + 1];
            let length = usize::from(be16(at + 2)?);
            if matches!(marker, 0xC0..=0xC3 | 0xC5..=0xC7 | 0xC9..=0xCB | 0xCD..=0xCF) {
                return Some((f64::from(be16(at + 7)?), f64::from(be16(at + 5)?)));
            }
            at += 2 + length;
        }
    }
    None
}

pub(crate) fn read_image_bytes(package: &OdfPackage, path: &str) -> Result<Vec<u8>, String> {
    package
        .session()
        .run_operation("odf-image", |operation| operation.read_bytes(path))
}

/// Read the drawings of one sheet. Chart frames are returned pending.
pub(crate) fn read_drawings(
    ctx: &mut Context,
    package: &OdfPackage,
    anchored: &[Anchored<'_, '_>],
    grid: &Grid,
    worksheet: &mut Worksheet,
    sheet: usize,
) -> Result<Vec<PendingChart>, String> {
    let mut charts = Vec::new();
    for (order, item) in anchored.iter().enumerate() {
        read_drawing(
            ctx,
            package,
            item.node,
            item.cell,
            grid,
            worksheet,
            sheet,
            order,
            &mut charts,
        )?;
    }
    Ok(charts)
}

#[allow(clippy::too_many_arguments)]
fn read_drawing(
    ctx: &mut Context,
    package: &OdfPackage,
    node: roxmltree::Node<'_, '_>,
    cell: Option<(u32, u32)>,
    grid: &Grid,
    worksheet: &mut Worksheet,
    sheet: usize,
    order: usize,
    charts: &mut Vec<PendingChart>,
) -> Result<(), String> {
    if node.tag_name().namespace() != Some(NS_DRAW) {
        return Ok(());
    }
    ctx.drawing_count += 1;
    if ctx.drawing_count > MAX_DRAWINGS {
        return Err(super::resource("drawing count"));
    }
    let name = node.tag_name().name();
    if name == "a" {
        for child in node.children().filter(|n| n.is_element()) {
            read_drawing(
                ctx, package, child, cell, grid, worksheet, sheet, order, charts,
            )?;
        }
        return Ok(());
    }
    if name == "g" {
        // A group's members carry their own absolute geometry.
        for child in node.children().filter(|n| n.is_element()) {
            read_drawing(
                ctx, package, child, cell, grid, worksheet, sheet, order, charts,
            )?;
        }
        return Ok(());
    }
    let Some((rect, rotation)) = place(node, cell, grid, &ctx.sheet_names) else {
        return Ok(());
    };
    let z = z_index(node, order);
    let style_chain = ctx
        .store
        .chain("graphic", node.attribute((NS_DRAW, "style-name")));
    if name == "frame" {
        let object = node
            .children()
            .find(|n| xml::is(*n, NS_DRAW, "object"))
            .and_then(|o| o.attribute((NS_XLINK, "href")))
            .and_then(package_path);
        let images: Vec<String> = node
            .children()
            .filter(|n| xml::is(*n, NS_DRAW, "image"))
            .filter_map(|n| n.attribute((NS_XLINK, "href")))
            .filter_map(package_path)
            .collect();
        if let Some(object) = object {
            charts.push(PendingChart {
                sheet,
                rect,
                z_order: z,
                object,
                replacement: images.first().cloned(),
            });
            return Ok(());
        }
        if node
            .children()
            .any(|n| xml::is(n, NS_DRAW, "object-ole") || xml::is(n, NS_DRAW, "plugin"))
            || !images.is_empty()
        {
            push_image(
                &mut ctx.budget,
                &mut ctx.images,
                package,
                &images,
                rect,
                rotation,
                z,
                &style_chain,
                worksheet,
            )?;
            return Ok(());
        }
        if let Some(text_box) = node.children().find(|n| xml::is(*n, NS_DRAW, "text-box")) {
            let shape = shape_info(ctx, text_box, &style_chain, "rect", z, rotation, rect);
            push_shape_checked(ctx, package, worksheet, rect, shape)?;
        }
        return Ok(());
    }
    let preset = match name {
        "rect" => {
            if pt_attr(node, NS_DRAW, "corner-radius").is_some_and(|r| r > 0.0) {
                "roundRect"
            } else {
                "rect"
            }
        }
        "ellipse" | "circle" => "ellipse",
        "line" | "connector" => "line",
        "custom-shape" => {
            let geometry = node
                .children()
                .find(|n| xml::is(*n, NS_DRAW, "enhanced-geometry"));
            match geometry.and_then(|g| g.attribute((NS_DRAW, "type"))) {
                Some(kind) => preset_name(kind).unwrap_or("custom"),
                None => "custom",
            }
        }
        "polygon" | "polyline" | "path" => "rect",
        _ => return Ok(()),
    };
    let mut shape = shape_info(ctx, node, &style_chain, preset, z, rotation, rect);
    if name == "line" || name == "connector" {
        let (x1, y1, x2, y2) = (
            pt_attr(node, NS_SVG, "x1").unwrap_or(0.0),
            pt_attr(node, NS_SVG, "y1").unwrap_or(0.0),
            pt_attr(node, NS_SVG, "x2").unwrap_or(0.0),
            pt_attr(node, NS_SVG, "y2").unwrap_or(0.0),
        );
        shape.flip_h = x2 < x1;
        shape.flip_v = y2 < y1;
        shape.fill = None;
        shape.fill_color = None;
    }
    if preset == "custom" {
        // §19.145: the enhanced path itself is the geometry.
        let size = (
            pt_attr(node, NS_SVG, "width").unwrap_or(0.0),
            pt_attr(node, NS_SVG, "height").unwrap_or(0.0),
        );
        shape.geom = match super::shape_path::custom_geometry(node, size) {
            Some(paths) => xlsx_model::ShapeGeom::Custom { paths },
            None => xlsx_model::ShapeGeom::Preset {
                name: "rect".into(),
                adj: Vec::new(),
            },
        };
    }
    if matches!(name, "polygon" | "polyline" | "path") {
        if let Some(geom) = polygon_geometry(node, name) {
            shape.geom = geom;
        }
        if name == "polyline" {
            shape.fill = None;
            shape.fill_color = None;
        }
    }
    push_shape_checked(ctx, package, worksheet, rect, shape)
}

/// LibreOffice writes imported OOXML preset shapes as `ooxml-<preset>`
/// (§19.145 draw:type is application-defined); its own presets use the
/// names below.
fn preset_name(kind: &str) -> Option<&'static str> {
    if let Some(preset) = kind.strip_prefix("ooxml-") {
        return OOXML_PRESETS.iter().copied().find(|p| *p == preset);
    }
    Some(match kind {
        "rectangle" => "rect",
        "round-rectangle" => "roundRect",
        "ellipse" | "circle" => "ellipse",
        "isosceles-triangle" => "triangle",
        "right-triangle" => "rtTriangle",
        "diamond" => "diamond",
        "pentagon" => "pentagon",
        "hexagon" => "hexagon",
        "octagon" => "octagon",
        "parallelogram" => "parallelogram",
        "trapezoid" => "trapezoid",
        "cross" => "plus",
        "star5" => "star5",
        "star4" => "star4",
        "right-arrow" => "rightArrow",
        "left-arrow" => "leftArrow",
        "up-arrow" => "upArrow",
        "down-arrow" => "downArrow",
        "can" => "can",
        "cube" => "cube",
        "heart" => "heart",
        "smiley" => "smileyFace",
        "sun" => "sun",
        "moon" => "moon",
        "flowchart-process" => "flowChartProcess",
        "flowchart-decision" => "flowChartDecision",
        "flowchart-terminator" => "flowChartTerminator",
        "round-rectangular-callout" => "wedgeRoundRectCallout",
        "rectangular-callout" => "wedgeRectCallout",
        "round-callout" => "wedgeEllipseCallout",
        _ => return None,
    })
}

const OOXML_PRESETS: &[&str] = &[
    "rect",
    "roundRect",
    "ellipse",
    "triangle",
    "rtTriangle",
    "diamond",
    "pentagon",
    "hexagon",
    "heptagon",
    "octagon",
    "decagon",
    "dodecagon",
    "parallelogram",
    "trapezoid",
    "plus",
    "star4",
    "star5",
    "star6",
    "star7",
    "star8",
    "star10",
    "star12",
    "star16",
    "star24",
    "star32",
    "rightArrow",
    "leftArrow",
    "upArrow",
    "downArrow",
    "leftRightArrow",
    "upDownArrow",
    "can",
    "cube",
    "heart",
    "smileyFace",
    "sun",
    "moon",
    "line",
    "straightConnector1",
    "bentConnector3",
    "curvedConnector3",
    "snip1Rect",
    "snip2SameRect",
    "round1Rect",
    "round2SameRect",
    "chevron",
    "homePlate",
    "flowChartProcess",
    "flowChartDecision",
    "flowChartTerminator",
    "flowChartAlternateProcess",
    "flowChartDocument",
    "flowChartConnector",
    "flowChartPredefinedProcess",
    "flowChartInputOutput",
    "wedgeRectCallout",
    "wedgeRoundRectCallout",
    "wedgeEllipseCallout",
    "cloudCallout",
    "cloud",
    "frame",
    "donut",
    "noSmoking",
    "blockArc",
    "arc",
    "chord",
    "pie",
    "bevel",
    "foldedCorner",
    "lightningBolt",
    "plaque",
    "teardrop",
    "halfFrame",
    "corner",
    "diagStripe",
    "leftBrace",
    "rightBrace",
    "leftBracket",
    "rightBracket",
    "bracePair",
    "bracketPair",
    "notchedRightArrow",
    "stripedRightArrow",
    "bentArrow",
    "uturnArrow",
    "curvedRightArrow",
    "curvedLeftArrow",
    "curvedUpArrow",
    "curvedDownArrow",
    "circularArrow",
    "rightArrowCallout",
    "leftArrowCallout",
    "upArrowCallout",
    "downArrowCallout",
    "ribbon",
    "ribbon2",
    "ellipseRibbon",
    "ellipseRibbon2",
    "verticalScroll",
    "horizontalScroll",
    "wave",
    "doubleWave",
    "irregularSeal1",
    "irregularSeal2",
    "actionButtonBlank",
    "mathPlus",
    "mathMinus",
    "mathMultiply",
    "mathDivide",
    "mathEqual",
    "mathNotEqual",
    "pentagon",
    "borderCallout1",
    "borderCallout2",
    "callout1",
    "callout2",
    "accentCallout1",
    "flowChartManualInput",
    "flowChartManualOperation",
    "flowChartPreparation",
    "flowChartMultidocument",
    "flowChartOffpageConnector",
    "flowChartOnlineStorage",
    "flowChartMagneticDisk",
    "flowChartSort",
    "flowChartExtract",
    "flowChartMerge",
    "flowChartDelay",
    "flowChartCollate",
    "flowChartOr",
    "flowChartSummingJunction",
    "flowChartPunchedCard",
    "flowChartPunchedTape",
    "flowChartDisplay",
    "flowChartMagneticTape",
    "flowChartMagneticDrum",
    "flowChartInternalStorage",
    "leftRightUpArrow",
    "quadArrow",
    "leftUpArrow",
    "bentUpArrow",
    "swooshArrow",
    "gear6",
    "gear9",
    "funnel",
    "pieWedge",
    "snipRoundRect",
    "snip2DiagRect",
    "round2DiagRect",
    "triangle",
    "can",
];

fn polygon_geometry(node: roxmltree::Node<'_, '_>, name: &str) -> Option<xlsx_model::ShapeGeom> {
    let points = node.attribute((NS_DRAW, "points"))?;
    let view: Vec<f64> = node
        .attribute((NS_SVG, "viewBox"))?
        .split_whitespace()
        .filter_map(|v| v.parse().ok())
        .collect();
    let (vw, vh) = (*view.get(2)?, *view.get(3)?);
    let (vx, vy) = (view[0], view[1]);
    let mut commands = Vec::new();
    for (i, pair) in points.split_whitespace().enumerate() {
        let (x, y) = pair.split_once(',')?;
        let (x, y) = (x.parse::<f64>().ok()? - vx, y.parse::<f64>().ok()? - vy);
        commands.push(if i == 0 {
            xlsx_model::PathCmd::MoveTo { x, y }
        } else {
            xlsx_model::PathCmd::LineTo { x, y }
        });
    }
    if name == "polygon" {
        commands.push(xlsx_model::PathCmd::Close);
    }
    Some(xlsx_model::ShapeGeom::Custom {
        paths: vec![xlsx_model::PathInfo {
            w: vw.max(1.0),
            h: vh.max(1.0),
            fill: None,
            stroke: true,
            commands,
        }],
    })
}

/// Fill and stroke of a graphic style chain (§20.130 draw:fill, §20.164
/// draw:stroke), and the shape's text.
fn shape_info(
    ctx: &Context,
    node: roxmltree::Node<'_, '_>,
    chain: &[&Style],
    preset: &str,
    z: u64,
    rotation: f64,
    _rect: Rect,
) -> ShapeInfo {
    let fill = shape_fill(ctx.fills.resolve(chain));
    let fill_color = match &fill {
        Some(xlsx_model::ShapeFill::Solid { color }) => Some(color.clone()),
        _ => None,
    };
    let stroke = get(chain, "graphic:draw:stroke");
    let stroke_color = if stroke == Some("none") {
        None
    } else {
        get(chain, "graphic:svg:stroke-color")
            .and_then(shape_color)
            .or_else(|| Some("#000000".into()))
    };
    let stroke_width = get(chain, "graphic:svg:stroke-width")
        .and_then(length_pt)
        .map(|pt| (pt * EMU_PER_PT).round() as i64)
        .unwrap_or(0)
        .max(if stroke_color.is_some() { 9525 } else { 0 });
    let text = shape_text(ctx, node, chain);
    ShapeInfo {
        z_order: z,
        x: 0.0,
        y: 0.0,
        w: 1.0,
        h: 1.0,
        rot: rotation,
        flip_h: false,
        flip_v: false,
        fill_color,
        fill,
        stroke_color,
        stroke_width,
        stroke_fill: None,
        stroke_dash_style: (stroke == Some("dash")).then(|| "dash".to_string()),
        stroke_custom_dash: Vec::new(),
        stroke_line_cap: None,
        stroke_line_join: None,
        stroke_miter_limit: None,
        stroke_alignment: None,
        stroke_cmpd: None,
        stroke_head_end: marker(chain, "start"),
        stroke_tail_end: marker(chain, "end"),
        geom: xlsx_model::ShapeGeom::Preset {
            name: preset.to_string(),
            adj: Vec::new(),
        },
        text,
    }
}

fn marker(chain: &[&Style], end: &str) -> Option<xlsx_model::ShapeLineEnd> {
    let name = get(chain, &format!("graphic:draw:marker-{end}"))?;
    if name.is_empty() {
        return None;
    }
    let lower = name.to_ascii_lowercase();
    let kind = if lower.contains("oval") || lower.contains("circle") {
        "oval"
    } else if lower.contains("diamond") || lower.contains("square") {
        "diamond"
    } else if lower.contains("stealth") {
        "stealth"
    } else if lower.contains("line") {
        "arrow"
    } else {
        "triangle"
    };
    Some(xlsx_model::ShapeLineEnd {
        r#type: kind.into(),
        w: "med".into(),
        len: "med".into(),
    })
}

fn shape_text(
    ctx: &Context,
    node: roxmltree::Node<'_, '_>,
    chain: &[&Style],
) -> Option<xlsx_model::ShapeText> {
    let paragraphs: Vec<_> = node
        .children()
        .filter(|n| xml::is(*n, NS_TEXT, "p") || xml::is(*n, NS_TEXT, "h"))
        .collect();
    if paragraphs.is_empty() {
        return None;
    }
    let base_font = super::cells::font_of(chain, &ctx.faces, None);
    let mut out = Vec::new();
    for p in paragraphs {
        let p_chain = ctx
            .store
            .chain("paragraph", p.attribute((NS_TEXT, "style-name")));
        let align =
            match get(&p_chain, "par:fo:text-align").or_else(|| get(chain, "par:fo:text-align")) {
                Some("center") => "ctr",
                Some("end") | Some("right") => "r",
                Some("justify") => "just",
                _ => "l",
            };
        let p_font = super::cells::font_of(&p_chain, &ctx.faces, Some(&base_font));
        let mut runs = Vec::new();
        collect_shape_runs(ctx, p, &p_font, &mut runs);
        out.push(xlsx_model::ShapeParagraph {
            align: align.into(),
            rtl: false,
            mar_l: None,
            mar_r: None,
            indent: None,
            def_tab_sz: None,
            tab_stops: Vec::new(),
            space_line: None,
            space_before: None,
            space_after: None,
            runs,
        });
    }
    let inset = |side: &str| {
        get(chain, &format!("graphic:fo:padding-{side}"))
            .or_else(|| get(chain, "graphic:fo:padding"))
            .and_then(length_pt)
            .map(|pt| (pt * EMU_PER_PT).round() as i64)
            .unwrap_or(91_440)
    };
    Some(xlsx_model::ShapeText {
        vert: "horz".into(),
        anchor_ctr: get(chain, "graphic:draw:textarea-horizontal-align") == Some("center"),
        spc_first_last_para: false,
        anchor: match get(chain, "graphic:draw:textarea-vertical-align") {
            Some("middle") => "ctr",
            Some("bottom") => "b",
            _ => "t",
        }
        .into(),
        wrap: if get(chain, "graphic:fo:wrap-option") == Some("no-wrap") {
            "none".into()
        } else {
            "square".into()
        },
        auto_fit: "none".into(),
        font_scale: None,
        ln_spc_reduction: None,
        l_ins: inset("left"),
        t_ins: inset("top"),
        r_ins: inset("right"),
        b_ins: inset("bottom"),
        paragraphs: out,
    })
}

fn collect_shape_runs(
    ctx: &Context,
    node: roxmltree::Node<'_, '_>,
    font: &xlsx_model::Font,
    runs: &mut Vec<xlsx_model::ShapeTextRun>,
) {
    for child in node.children() {
        if child.is_text() {
            push_text_run(runs, child.text().unwrap_or(""), font);
            continue;
        }
        if child.tag_name().namespace() != Some(NS_TEXT) {
            continue;
        }
        match child.tag_name().name() {
            "span" => {
                let chain = ctx
                    .store
                    .chain("text", child.attribute((NS_TEXT, "style-name")));
                let span = super::cells::font_of(&chain, &ctx.faces, Some(font));
                collect_shape_runs(ctx, child, &span, runs);
            }
            "s" => {
                let count = child
                    .attribute((NS_TEXT, "c"))
                    .and_then(|c| c.parse::<usize>().ok())
                    .unwrap_or(1)
                    .min(4096);
                push_text_run(runs, &" ".repeat(count), font);
            }
            "tab" => push_text_run(runs, "\t", font),
            "line-break" => runs.push(xlsx_model::ShapeTextRun::Break),
            _ => collect_shape_runs(ctx, child, font, runs),
        }
    }
}

fn push_text_run(runs: &mut Vec<xlsx_model::ShapeTextRun>, text: &str, font: &xlsx_model::Font) {
    if text.is_empty() {
        return;
    }
    runs.push(xlsx_model::ShapeTextRun::Text {
        text: text.to_string(),
        bold: font.bold,
        italic: font.italic,
        size: font.size,
        spacing: None,
        color: font.color.as_ref().map(|c| format!("#{c}")),
        font_face: font.name.clone(),
        font_face_ea: None,
        font_face_cs: None,
    });
}

fn push_shape_checked(
    ctx: &mut Context,
    package: &OdfPackage,
    worksheet: &mut Worksheet,
    rect: Rect,
    mut shape: ShapeInfo,
) -> Result<(), String> {
    admit_fill_image(package, &mut ctx.images, &mut shape)?;
    push_shape(&mut ctx.budget, worksheet, rect, shape)
}

fn push_shape(
    budget: &mut super::Budget,
    worksheet: &mut Worksheet,
    rect: Rect,
    shape: ShapeInfo,
) -> Result<(), String> {
    budget.bytes(1024)?;
    worksheet.shape_groups.push(ShapeAnchor {
        from_col: rect.from_col,
        from_col_off: rect.from_col_off,
        from_row: rect.from_row,
        from_row_off: rect.from_row_off,
        to_col: rect.to_col,
        to_col_off: rect.to_col_off,
        to_row: rect.to_row,
        to_row_off: rect.to_row_off,
        edit_as: None,
        native_ext_cx: rect.width_emu,
        native_ext_cy: rect.height_emu,
        anchor_tag: Some(DrawingAnchorTag::TwoCellAnchor),
        anchor_ext_cx: None,
        anchor_ext_cy: None,
        shapes: vec![shape],
    });
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_image(
    budget: &mut super::Budget,
    admitted: &mut std::collections::BTreeSet<String>,
    package: &OdfPackage,
    images: &[String],
    rect: Rect,
    rotation: f64,
    z: u64,
    chain: &[&Style],
    worksheet: &mut Worksheet,
) -> Result<(), String> {
    // A raster rendition first, an SVG original beside it.
    let mut raster = None;
    let mut svg = None;
    for path in images {
        if !package.contains(path) {
            continue;
        }
        let bytes = read_image_bytes(package, path)?;
        match image_mime(&bytes) {
            Some("image/svg+xml") => {
                svg.get_or_insert((path.clone(), bytes));
            }
            Some(mime) => {
                raster.get_or_insert((path.clone(), mime, bytes));
            }
            None => {}
        }
    }
    let (image_path, mime_type, bytes) = match (&raster, &svg) {
        (Some((path, mime, bytes)), _) => (path.clone(), mime.to_string(), bytes.as_slice()),
        (None, Some((path, bytes))) => {
            (path.clone(), "image/svg+xml".to_string(), bytes.as_slice())
        }
        // StarView metafiles (LibreOffice object replacements) have no
        // browser decoder.
        (None, None) => return Ok(()),
    };
    let src_rect = get(chain, "graphic:fo:clip").and_then(|clip| crop(clip, bytes));
    let alpha = get(chain, "graphic:draw:image-opacity")
        .and_then(percent)
        .filter(|a| *a < 1.0);
    admitted.insert(image_path.clone());
    if let Some((path, _)) = &svg {
        admitted.insert(path.clone());
    }
    budget.bytes(512)?;
    let mirror = get(chain, "graphic:style:mirror").unwrap_or("none");
    worksheet.images.push(ImageAnchor {
        z_order: z,
        from_col: rect.from_col,
        from_col_off: rect.from_col_off,
        from_row: rect.from_row,
        from_row_off: rect.from_row_off,
        to_col: rect.to_col,
        to_col_off: rect.to_col_off,
        to_row: rect.to_row,
        to_row_off: rect.to_row_off,
        edit_as: None,
        anchor_tag: Some(DrawingAnchorTag::TwoCellAnchor),
        anchor_ext_cx: None,
        anchor_ext_cy: None,
        native_ext_cx: rect.width_emu,
        native_ext_cy: rect.height_emu,
        rotation: (rotation != 0.0).then_some(rotation),
        flip_h: mirror.contains("horizontal").then_some(true),
        flip_v: mirror.contains("vertical").then_some(true),
        image_path,
        mime_type,
        svg_image_path: raster
            .as_ref()
            .and(svg.as_ref())
            .map(|(path, _)| path.clone()),
        src_rect,
        alpha,
        duotone: None,
        blip_effects: Vec::new(),
    });
    Ok(())
}

/// `fo:clip` `rect(top, right, bottom, left)` (§20.179), lengths measured on
/// the image at its natural size (96 DPI), as SpreadsheetML crop fractions.
fn crop(value: &str, bytes: &[u8]) -> Option<xlsx_model::SrcRect> {
    let inner = value.trim().strip_prefix("rect(")?.strip_suffix(')')?;
    let parts: Vec<f64> = inner
        .split(|c: char| c == ',' || c.is_whitespace())
        .filter(|s| !s.is_empty())
        .map(|s| if s == "auto" { Some(0.0) } else { length_pt(s) })
        .collect::<Option<_>>()?;
    if parts.len() != 4 || parts.iter().all(|v| *v == 0.0) {
        return None;
    }
    let (w, h) = image_size(bytes)?;
    // Natural size in points at 96 DPI.
    let (w, h) = (w * 0.75, h * 0.75);
    Some(xlsx_model::SrcRect {
        t: parts[0] / h,
        r: parts[1] / w,
        b: parts[2] / h,
        l: parts[3] / w,
    })
}

/// Resolve pending chart frames once every table is read (charts address
/// cells of any table).
pub(crate) fn resolve_charts(
    sheets: &mut [OdsSheet],
    pending: Vec<PendingChart>,
    package: &OdfPackage,
    ctx: &mut Context,
    warnings: &mut Vec<String>,
) -> Result<(), String> {
    for chart in pending {
        let resolver = SheetResolver { sheets, ctx };
        let model = crate::odf_chart::read_chart_object(package, &chart.object, Some(&resolver));
        let model = match model {
            Ok(model) => model,
            Err(error) if error.starts_with("OpenDocument spreadsheet exceeds") => {
                return Err(error)
            }
            Err(_) => None,
        };
        let rect = chart.rect;
        match model {
            Some(model) => {
                ctx.budget.bytes(4096)?;
                sheets[chart.sheet].worksheet.charts.push(ChartAnchor {
                    z_order: chart.z_order,
                    from_col: rect.from_col,
                    from_col_off: rect.from_col_off,
                    from_row: rect.from_row,
                    from_row_off: rect.from_row_off,
                    to_col: rect.to_col,
                    to_col_off: rect.to_col_off,
                    to_row: rect.to_row,
                    to_row_off: rect.to_row_off,
                    chart: model,
                });
            }
            None => {
                // An embedded object that is not a chart (a formula or a
                // nested document) is drawn from its replacement image when
                // the package holds one in a browser-decodable format.
                warnings.push("ods:object-replacement".into());
                let images: Vec<String> = chart.replacement.into_iter().collect();
                let mut worksheet = std::mem::replace(
                    &mut sheets[chart.sheet].worksheet,
                    Worksheet::placeholder("", String::new()),
                );
                let result = push_image(
                    &mut ctx.budget,
                    &mut ctx.images,
                    package,
                    &images,
                    rect,
                    0.0,
                    chart.z_order,
                    &[],
                    &mut worksheet,
                );
                sheets[chart.sheet].worksheet = worksheet;
                result?;
            }
        }
    }
    Ok(())
}

/// Cell ranges of the workbook for chart data (§18.3.5 cell range address).
struct SheetResolver<'a> {
    sheets: &'a [OdsSheet],
    ctx: &'a Context,
}

impl crate::odf_chart::RangeResolver for SheetResolver<'_> {
    fn cells(&self, address: &str) -> Option<Vec<crate::odf_chart::CellData>> {
        let (sheet, (r1, c1), (r2, c2)) = super::parse_range(address)?;
        let sheet = self
            .sheets
            .iter()
            .find(|s| sheet.as_deref().is_none_or(|name| name == s.name))?;
        let count = (u64::from(r2 - r1) + 1) * (u64::from(c2 - c1) + 1);
        if count > 1_048_576 {
            return None;
        }
        let _ = &self.ctx;
        let mut out = Vec::with_capacity(count as usize);
        for row in r1..=r2 {
            for col in c1..=c2 {
                let cell = super::find_cell(&sheet.rows, row, col);
                out.push(match cell.map(|c| &c.value) {
                    Some(xlsx_model::CellValue::Number { number }) => {
                        crate::odf_chart::CellData::Number(*number)
                    }
                    Some(xlsx_model::CellValue::Text { text, .. }) => {
                        crate::odf_chart::CellData::Text(text.clone())
                    }
                    Some(xlsx_model::CellValue::Bool { bool }) => {
                        crate::odf_chart::CellData::Number(if *bool { 1.0 } else { 0.0 })
                    }
                    _ => crate::odf_chart::CellData::Empty,
                });
            }
        }
        Some(out)
    }

    fn hidden(&self, _address: &str) -> Option<Vec<bool>> {
        None
    }
}

/// An area fill (`odf_fill`) in the workbook drawing model. A picture fill
/// is admitted by the caller once its bytes are known to decode.
fn shape_fill(fill: crate::odf_fill::Fill) -> Option<xlsx_model::ShapeFill> {
    use crate::odf_fill::{Fill, GradientShape};
    Some(match fill {
        Fill::None => return None,
        Fill::Solid(hex) => xlsx_model::ShapeFill::Solid {
            color: format!("#{hex}"),
        },
        Fill::Gradient(gradient) => {
            let (angle, grad_type, path, fill_to_rect) = match gradient.shape {
                GradientShape::Linear { angle } => (angle, "linear", None, None),
                GradientShape::Path { path, focus } => {
                    (0.0, "radial", Some(path.to_string()), Some(focus))
                }
            };
            xlsx_model::ShapeFill::Gradient {
                stops: gradient.stops,
                angle,
                grad_type: grad_type.into(),
                scaled: None,
                path,
                fill_to_rect,
                tile_rect: None,
                flip: None,
                rot_with_shape: Some(true),
            }
        }
        Fill::Hatch(hatch) => xlsx_model::ShapeFill::Hatch {
            color: format!("#{}", hatch.color),
            background: hatch.background.map(|c| format!("#{c}")),
            angle: hatch.angle,
            distance_pt: hatch.distance_pt,
            lines: hatch.lines,
        },
        Fill::Picture(picture) => xlsx_model::ShapeFill::Image {
            image_path: picture.path,
            mime_type: String::new(),
            tile: picture.tile.is_some(),
            tile_width_pt: picture.tile.and_then(|t| t.0),
            tile_height_pt: picture.tile.and_then(|t| t.1),
        },
    })
}

/// Admit a picture fill's image (its MIME type from the bytes); a fill
/// whose image is missing or not browser-decodable paints nothing.
fn admit_fill_image(
    package: &OdfPackage,
    admitted: &mut std::collections::BTreeSet<String>,
    shape: &mut ShapeInfo,
) -> Result<(), String> {
    let Some(xlsx_model::ShapeFill::Image {
        image_path,
        mime_type,
        ..
    }) = shape.fill.as_mut()
    else {
        return Ok(());
    };
    let mime = if package.contains(image_path) {
        image_mime(&read_image_bytes(package, image_path)?)
    } else {
        None
    };
    match mime {
        Some(mime) => {
            *mime_type = mime.to_string();
            admitted.insert(image_path.clone());
        }
        None => shape.fill = None,
    }
    Ok(())
}

/// Drawing colours in the workbook model are CSS `#RRGGBB` (as the
/// SpreadsheetML drawing parser emits them), unlike cell-style colours.
fn shape_color(value: &str) -> Option<String> {
    hex_color(value).map(|hex| format!("#{hex}"))
}
