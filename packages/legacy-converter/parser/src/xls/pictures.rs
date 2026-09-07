//! Passive XLS pictures -> ordinary SpreadsheetDrawing parts. Coordinates use
//! MS-XLS 2.5.193 cell fractions and ECMA-376 18.3.1.13 measured digit widths.
//! No legacy layout policy is delegated to the OOXML renderer.
use super::{
    drawing_anchors::{self, CellCorner, DrawingAnchor},
    drawing_media, Record, SheetData,
};
use std::collections::{BTreeMap, BTreeSet};

#[derive(Default)]
pub(super) struct Pictures {
    anchors: BTreeMap<usize, Vec<DrawingAnchor>>,
    images: Vec<(u32, &'static str, Vec<u8>)>,
    unsupported_images: bool,
}

pub(super) struct Parts {
    pub xml: Vec<(String, String)>,
    pub media: Vec<(String, Vec<u8>)>,
    pub types: String,
    pub sheets: BTreeSet<usize>,
}

impl Pictures {
    pub fn prepare(records: &[Record<'_>], tabs: &[usize]) -> Result<Self, String> {
        let sheet_ids: BTreeMap<_, _> = tabs.iter().enumerate().map(|(i, &tab)| (tab, i)).collect();
        let mut anchors = drawing_anchors::projectable(records)?;
        anchors.retain(|a| a.picture.is_some() && sheet_ids.contains_key(&a.sheet));
        let indices = anchors
            .iter()
            .filter_map(|a| a.picture.map(|p| p.store_index))
            .collect();
        let images = drawing_media::selected(records, &indices)?;
        let supported: BTreeSet<_> = images.iter().map(|i| i.0).collect();
        let mut by_sheet = BTreeMap::<_, Vec<_>>::new();
        for a in anchors {
            if a.picture
                .is_some_and(|p| supported.contains(&p.store_index))
            {
                by_sheet.entry(sheet_ids[&a.sheet]).or_default().push(a);
            }
        }
        Ok(Self {
            anchors: by_sheet,
            images,
            unsupported_images: supported.len() != indices.len(),
        })
    }

    pub fn is_empty(&self) -> bool {
        self.anchors.is_empty()
    }

    pub fn has_unsupported_images(&self) -> bool {
        self.unsupported_images
    }

    pub fn emit(
        self,
        sheets: &[(String, SheetData)],
        mdw: f64,
        warnings: &mut Vec<String>,
    ) -> Parts {
        let mut parts = Parts {
            xml: vec![],
            media: vec![],
            types: String::new(),
            sheets: BTreeSet::new(),
        };
        let mut used = BTreeSet::new();
        let mut omitted = false;
        // Resource governance, not a layout threshold. Prefixes are built once
        // per drawing sheet, never once per picture, and dropped after that sheet.
        let mut work = 2_000_000usize;
        let extensions: BTreeMap<_, _> = self.images.iter().map(|i| (i.0, i.1)).collect();
        for (sheet_index, anchors) in self.anchors {
            let sheet = &sheets[sheet_index].1;
            // Do not invent a required sheetFormatPr row height or infer which
            // window's doubled formula-display grid owns the saved rectangle.
            if !sheet.geometry.has_sheet_defaults() || sheet.views.displays_formulas() {
                omitted = true;
                continue;
            }
            let max_row = anchors
                .iter()
                .map(|a| a.from.row.max(a.to.row))
                .max()
                .unwrap_or(0);
            let max_col = anchors
                .iter()
                .map(|a| a.from.column.max(a.to.column))
                .max()
                .unwrap_or(0);
            let cost = usize::from(max_row) + usize::from(max_col) + anchors.len() + 4;
            let Some(left) = work.checked_sub(cost) else {
                omitted = true;
                continue;
            };
            work = left;
            let columns = prefix(max_col, |c| sheet.geometry.column_emu(c, mdw));
            let rows = prefix(max_row, |r| sheet.geometry.row_emu(r));
            let mut xml = String::from("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            let mut rels = String::from("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            let mut count = 0;
            for anchor in anchors {
                let Some(picture) = anchor.picture else {
                    continue;
                };
                let locate = |c: CellCorner| -> Option<(i64, i64, i64, i64)> {
                    let col = usize::from(c.column);
                    let row = usize::from(c.row);
                    let x = columns[col]?;
                    let y = rows[row]?;
                    let dx = (columns[col + 1]? - x) * f64::from(c.dx) / 1024.0;
                    let dy = (rows[row + 1]? - y) * f64::from(c.dy) / 256.0;
                    Some((
                        x.round() as i64,
                        y.round() as i64,
                        dx.round() as i64,
                        dy.round() as i64,
                    ))
                };
                let (Some((x, y, dx, dy)), Some((tx, ty, tdx, tdy))) =
                    (locate(anchor.from), locate(anchor.to))
                else {
                    omitted = true;
                    continue;
                };
                let (cx, cy) = (tx + tdx - x - dx, ty + tdy - y - dy);
                if cx <= 0 || cy <= 0 {
                    omitted = true;
                    continue;
                }
                let crop = picture
                    .crop
                    .map(|v| (f64::from(v) * 100_000.0 / 65536.0).round());
                if crop
                    .iter()
                    .any(|v| *v < f64::from(i32::MIN) || *v > f64::from(i32::MAX))
                {
                    omitted = true;
                    continue;
                }
                let id = picture.store_index;
                let Some(ext) = extensions.get(&id) else {
                    omitted = true;
                    continue;
                };
                let edit_as = match anchor.behavior {
                    0 => "twoCell",
                    2 => "oneCell",
                    3 => "absolute",
                    _ => {
                        omitted = true;
                        continue;
                    }
                };
                count += 1;
                used.insert(id);
                xml.push_str(&format!("<xdr:twoCellAnchor editAs=\"{edit_as}\">"));
                for (tag, corner, ox, oy) in
                    [("from", anchor.from, dx, dy), ("to", anchor.to, tdx, tdy)]
                {
                    xml.push_str(&format!("<xdr:{tag}><xdr:col>{}</xdr:col><xdr:colOff>{ox}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{oy}</xdr:rowOff></xdr:{tag}>", corner.column, corner.row));
                }
                let rotation = (f64::from(picture.rotation) * 60000.0 / 65536.0).round() as i64;
                let (flip_h, flip_v) = (
                    u8::from(anchor.shape_flags & 64 != 0),
                    u8::from(anchor.shape_flags & 128 != 0),
                );
                xml.push_str(&format!("<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"{count}\" name=\"Picture {count}\"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"rId{count}\"/><a:srcRect t=\"{}\" b=\"{}\" l=\"{}\" r=\"{}\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm rot=\"{rotation}\" flipH=\"{flip_h}\" flipV=\"{flip_v}\"><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>", crop[0], crop[1], crop[2], crop[3], x + dx, y + dy));
                rels.push_str(&format!("<Relationship Id=\"rId{count}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image{id}.{ext}\"/>"));
            }
            if count == 0 {
                continue;
            }
            let id = sheet_index + 1;
            xml.push_str("</xdr:wsDr>");
            rels.push_str("</Relationships>");
            parts.types.push_str(&format!("<Override PartName=\"/xl/drawings/drawing{id}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>"));
            parts
                .xml
                .push((format!("xl/drawings/drawing{id}.xml"), xml));
            parts
                .xml
                .push((format!("xl/drawings/_rels/drawing{id}.xml.rels"), rels));
            parts.xml.push((format!("xl/worksheets/_rels/sheet{id}.xml.rels"), format!("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"legacyDrawing\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing{id}.xml\"/></Relationships>")));
            parts.sheets.insert(sheet_index);
        }
        for (id, ext, bytes) in self.images {
            if !used.contains(&id) {
                continue;
            }
            let mime = match ext {
                "png" => "image/png",
                "jpg" | "jpeg" => "image/jpeg",
                "emf" => "image/x-emf",
                "wmf" => "image/wmf",
                _ => continue,
            };
            let name = format!("xl/media/image{id}.{ext}");
            parts.types.push_str(&format!(
                "<Override PartName=\"/{name}\" ContentType=\"{mime}\"/>"
            ));
            parts.media.push((name, bytes));
        }
        if omitted {
            warnings.push("legacy-xls:unresolved-picture-geometry-omitted".into());
        }
        parts
    }
}

fn prefix(last: u16, mut dimension: impl FnMut(u16) -> Option<f64>) -> Vec<Option<f64>> {
    let mut result = Vec::with_capacity(usize::from(last) + 2);
    result.push(Some(0.0));
    for i in 0..=last {
        result.push(result[usize::from(i)].and_then(|v| dimension(i).map(|d| v + d)));
    }
    result
}

#[cfg(test)]
mod tests {
    use super::super::drawing_anchors::PictureReference;
    use super::*;

    fn sheet() -> SheetData {
        let mut sheet = SheetData::default();
        sheet
            .geometry
            .read(&Record {
                kind: 0x225,
                offset: 0,
                data: &[0, 0, 44, 1],
            })
            .unwrap();
        sheet
            .geometry
            .read(&Record {
                kind: 0x99,
                offset: 0,
                data: &[0, 10],
            })
            .unwrap();
        sheet
    }
    fn anchor() -> DrawingAnchor {
        DrawingAnchor {
            sheet: 0,
            shape_id: 1,
            shape_flags: 64 | 128,
            object_id: 1,
            object_type: 8,
            object_flags: 0,
            group_depth: 1,
            behavior: 2,
            from: CellCorner {
                column: 0,
                row: 0,
                dx: -512,
                dy: -128,
            },
            to: CellCorner {
                column: 2,
                row: 3,
                dx: 256,
                dy: 64,
            },
            picture: Some(PictureReference {
                store_index: 1,
                crop: [32768, -16384, 16384, 0],
                rotation: -90 * 65536,
                clipboard_format: 9,
                auto_picture: true,
            }),
        }
    }
    fn pictures(anchors: Vec<DrawingAnchor>) -> Pictures {
        Pictures {
            anchors: BTreeMap::from([(0, anchors)]),
            images: vec![(1, "png", vec![1, 2, 3])],
            unsupported_images: false,
        }
    }

    #[test]
    fn deduplicates_media_and_emits_signed_crops_rotation_flips_and_relationships() {
        let mut second = anchor();
        second.shape_id = 2;
        second.object_id = 2;
        let mut warnings = Vec::new();
        let parts =
            pictures(vec![anchor(), second]).emit(&[("S".into(), sheet())], 7.0, &mut warnings);
        assert!(warnings.is_empty());
        assert_eq!(parts.media.len(), 1);
        assert_eq!(parts.media[0].0, "xl/media/image1.png");
        let xml = &parts.xml[0].1;
        assert_eq!(xml.matches("<xdr:pic>").count(), 2);
        assert!(xml.contains("rot=\"-5400000\" flipH=\"1\" flipV=\"1\""));
        assert!(xml.contains("<a:srcRect t=\"50000\" b=\"-25000\" l=\"25000\" r=\"0\"/>"));
        assert!(xml.contains("<xdr:colOff>-333375</xdr:colOff>"));
        assert!(xml.contains("r:embed=\"rId1\"") && xml.contains("r:embed=\"rId2\""));
        assert_eq!(
            parts.xml[1]
                .1
                .matches("Target=\"../media/image1.png\"")
                .count(),
            2
        );
    }

    #[test]
    fn mixed_supported_and_unsupported_catalog_entries_keep_supported_output() {
        let pictures = Pictures {
            anchors: BTreeMap::from([(0, vec![anchor()])]),
            images: vec![(1, "png", vec![1, 2, 3])],
            unsupported_images: true,
        };
        assert!(pictures.has_unsupported_images());
        let mut warnings = vec!["legacy-xls:invalid-or-unsupported-pictures-omitted".into()];
        let parts = pictures.emit(&[("S".into(), sheet())], 7.0, &mut warnings);
        assert_eq!(parts.media, [("xl/media/image1.png".into(), vec![1, 2, 3])]);
        assert_eq!(
            warnings,
            ["legacy-xls:invalid-or-unsupported-pictures-omitted"]
        );
    }

    #[test]
    fn refuses_missing_dimensions_and_never_emits_unreferenced_media() {
        let mut warnings = Vec::new();
        let parts = pictures(vec![anchor()]).emit(
            &[("S".into(), SheetData::default())],
            7.0,
            &mut warnings,
        );
        assert!(parts.xml.is_empty() && parts.media.is_empty());
        assert_eq!(warnings, ["legacy-xls:unresolved-picture-geometry-omitted"]);
        let mut bad = anchor();
        bad.to = bad.from;
        let parts = pictures(vec![bad]).emit(&[("S".into(), sheet())], 7.0, &mut vec![]);
        assert!(parts.media.is_empty());
        assert_eq!(
            prefix(2, |i| (i != 1).then_some(10.0)),
            [Some(0.0), Some(10.0), None, None]
        );
    }

    #[test]
    fn bounds_total_geometry_work_across_many_sheets() {
        let sheets: Vec<_> = (0..32).map(|_| ("S".into(), sheet())).collect();
        let mut picture = anchor();
        picture.to.row = 65535;
        let pictures = Pictures {
            anchors: (0..32).map(|i| (i, vec![picture])).collect(),
            images: vec![(1, "png", vec![1])],
            unsupported_images: false,
        };
        let mut warnings = vec![];
        let parts = pictures.emit(&sheets, 7.0, &mut warnings);
        assert_eq!(parts.sheets.len(), 30);
        assert_eq!(parts.media.len(), 1);
        assert_eq!(warnings, ["legacy-xls:unresolved-picture-geometry-omitted"]);
    }

    #[test]
    fn measured_defaults_only_change_sheets_that_receive_drawings() {
        use super::super::{minimal_styles, styles::NormalFont, PreparedXls};
        use std::io::{Cursor, Read};
        let prepared = PreparedXls {
            sheets: vec![("Picture".into(), sheet()), ("Cells".into(), sheet())],
            styles: minimal_styles(),
            date1904: false,
            window_count: 0,
            warnings: vec![
                "legacy-xls:drawings-conditional-formatting-and-external-links-omitted".into(),
            ],
            font: Some(NormalFont {
                name: "F".into(),
                size_points: 11.0,
                bold: false,
                italic: false,
            }),
            pictures: pictures(vec![anchor()]),
        };
        let result = prepared.finish(10000, Some(7.0)).unwrap();
        let mut zip = zip::ZipArchive::new(Cursor::new(result.bytes)).unwrap();
        for (id, measured) in [(1, true), (2, false)] {
            let mut xml = String::new();
            zip.by_name(&format!("xl/worksheets/sheet{id}.xml"))
                .unwrap()
                .read_to_string(&mut xml)
                .unwrap();
            assert_eq!(xml.contains("defaultColWidth=\"10\""), measured);
            assert_eq!(xml.contains("r:id=\"legacyDrawing\""), measured);
        }
    }
}
