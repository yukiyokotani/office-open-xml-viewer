//! Passive XLS pictures -> XLSX-model image anchors and image resources.
//! Coordinates use MS-XLS 2.5.193 cell fractions and ECMA-376 18.3.1.13
//! measured digit widths. No legacy layout policy is delegated to the renderer.
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
    /// Store entries drawn by grouped pictures, which the shape projection
    /// places; their media is retained with the sheet pictures'.
    grouped: BTreeSet<u32>,
}

pub(super) struct ResolvedPictures {
    sheets: BTreeMap<usize, Vec<ResolvedPicture>>,
    images: Vec<(u32, &'static str, Vec<u8>)>,
}

struct ResolvedPicture {
    from: CellCorner,
    to: CellCorner,
    dx: i64,
    dy: i64,
    tdx: i64,
    tdy: i64,
    cx: i64,
    cy: i64,
    crop: [f64; 4],
    rotation: i64,
    flip_h: bool,
    flip_v: bool,
    store_index: u32,
    extension: &'static str,
    edit_as: &'static str,
    /// Document (paint) order among all of the sheet's drawing objects.
    order: u64,
}

pub(super) struct NativePictures {
    pub sheets: BTreeMap<usize, Vec<xlsx_model::ImageAnchor>>,
    pub resources: BTreeMap<String, Vec<u8>>,
}

impl Pictures {
    pub fn prepare(
        records: &[Record<'_>],
        tabs: &[usize],
        raster: crate::officeart::raster::Raster,
        grouped: &BTreeSet<u32>,
    ) -> Result<Self, String> {
        let sheet_ids: BTreeMap<_, _> = tabs.iter().enumerate().map(|(i, &tab)| (tab, i)).collect();
        let mut anchors = drawing_anchors::projectable(records)?;
        anchors.retain(|a| a.picture.is_some() && sheet_ids.contains_key(&a.sheet));
        let mut indices: BTreeSet<u32> = anchors
            .iter()
            .filter_map(|a| a.picture.map(|p| p.store_index))
            .collect();
        indices.extend(grouped);
        let images = drawing_media::selected(records, &indices, raster)?;
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
            grouped: grouped.clone(),
        })
    }

    pub fn is_empty(&self) -> bool {
        self.anchors.is_empty()
    }

    /// The file extension of a store entry's decoded media.
    pub fn extension(&self, id: u32) -> Option<&'static str> {
        self.images
            .iter()
            .find(|image| image.0 == id)
            .map(|image| image.1)
    }

    pub fn has_unsupported_images(&self) -> bool {
        self.unsupported_images
    }

    /// Resolve source geometry once; model DTOs are built by `into_models`.
    pub(super) fn resolve(
        self,
        sheets: &[(String, SheetData)],
        mdw: f64,
        warnings: &mut Vec<String>,
    ) -> ResolvedPictures {
        let mut resolved_sheets = BTreeMap::new();
        let mut used = self.grouped.clone();
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
            let mut resolved = Vec::new();
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
                used.insert(id);
                resolved.push(ResolvedPicture {
                    from: anchor.from,
                    to: anchor.to,
                    dx,
                    dy,
                    tdx,
                    tdy,
                    cx,
                    cy,
                    crop,
                    rotation: (f64::from(picture.rotation) * 60000.0 / 65536.0).round() as i64,
                    flip_h: anchor.shape_flags & 64 != 0,
                    flip_v: anchor.shape_flags & 128 != 0,
                    store_index: id,
                    extension: ext,
                    edit_as,
                    order: anchor.order,
                });
            }
            if !resolved.is_empty() {
                resolved_sheets.insert(sheet_index, resolved);
            }
        }
        let images = self
            .images
            .into_iter()
            .filter(|image| used.contains(&image.0))
            .collect();
        if omitted {
            warnings.push("legacy-xls:unresolved-picture-geometry-omitted".into());
        }
        ResolvedPictures {
            sheets: resolved_sheets,
            images,
        }
    }
}

impl ResolvedPictures {
    /// Charge retained payload and model slots, excluding allocator bookkeeping.
    /// Source anchor/media limits bound map entry counts separately.
    pub(super) fn into_models(self, budget: &mut usize) -> Result<NativePictures, String> {
        fn charge(budget: &mut usize, bytes: usize) -> Result<(), String> {
            *budget = budget
                .checked_sub(bytes)
                .ok_or_else(|| super::unsupported("XLS picture model byte budget exceeded"))?;
            Ok(())
        }
        fn key(id: u32, budget: &mut usize) -> Result<String, String> {
            charge(
                budget,
                "legacy-xls/image/".len() + id.max(1).ilog10() as usize + 1,
            )?;
            Ok(format!("legacy-xls/image/{id}"))
        }
        let mut sheets = BTreeMap::new();
        for (index, values) in self.sheets {
            charge(
                budget,
                std::mem::size_of::<(usize, Vec<xlsx_model::ImageAnchor>)>(),
            )?;
            let bytes = values
                .len()
                .checked_mul(std::mem::size_of::<xlsx_model::ImageAnchor>())
                .ok_or_else(|| super::unsupported("XLS picture model byte budget exceeded"))?;
            charge(budget, bytes)?;
            let mut anchors = Vec::new();
            anchors
                .try_reserve_exact(values.len())
                .map_err(|_| super::unsupported("XLS picture model allocation failed"))?;
            for value in values {
                let image_path = key(value.store_index, budget)?;
                let mime = ooxml_common::blip::mime_from_ext(value.extension);
                charge(budget, mime.len() + value.edit_as.len())?;
                anchors.push(xlsx_model::ImageAnchor {
                    // OfficeArt document order, shared with charts and
                    // shapes, not an artificial XML byte offset.
                    z_order: value.order,
                    from_col: u32::from(value.from.column),
                    from_row: u32::from(value.from.row),
                    from_col_off: value.dx,
                    from_row_off: value.dy,
                    to_col: u32::from(value.to.column),
                    to_row: u32::from(value.to.row),
                    to_col_off: value.tdx,
                    to_row_off: value.tdy,
                    edit_as: Some(value.edit_as.into()),
                    native_ext_cx: value.cx,
                    native_ext_cy: value.cy,
                    rotation: (value.rotation != 0).then_some(value.rotation as f64 / 60000.0),
                    flip_h: value.flip_h.then_some(true),
                    flip_v: value.flip_v.then_some(true),
                    image_path,
                    mime_type: mime.into(),
                    svg_image_path: None,
                    src_rect: value.crop.iter().any(|v| *v != 0.0).then_some(
                        ooxml_common::blip::SrcRect {
                            t: value.crop[0] / 100000.0,
                            b: value.crop[1] / 100000.0,
                            l: value.crop[2] / 100000.0,
                            r: value.crop[3] / 100000.0,
                        },
                    ),
                    alpha: None,
                    duotone: None,
                });
            }
            sheets.insert(index, anchors);
        }
        let mut resources = BTreeMap::new();
        for (id, _, bytes) in self.images {
            charge(budget, std::mem::size_of::<(String, Vec<u8>)>())?;
            charge(budget, bytes.capacity())?;
            resources.insert(key(id, budget)?, bytes);
        }
        Ok(NativePictures { sheets, resources })
    }
}

pub(super) fn prefix(last: u16, mut dimension: impl FnMut(u16) -> Option<f64>) -> Vec<Option<f64>> {
    let mut result = Vec::with_capacity(usize::from(last) + 2);
    result.push(Some(0.0));
    for i in 0..=last {
        result.push(result[usize::from(i)].and_then(|v| dimension(i).map(|d| v + d)));
    }
    result
}

#[cfg(test)]
pub(super) fn session_fixture() -> (Pictures, SheetData, &'static str, Vec<u8>) {
    use super::drawing_anchors::PictureReference;
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
            kind: 0x55,
            offset: 0,
            data: &[10, 0],
        })
        .unwrap();
    let anchor = DrawingAnchor {
        sheet: 0,
        shape_id: 1,
        shape_flags: 0,
        object_id: 1,
        object_type: 8,
        object_flags: 0,
        group_depth: 1,
        order: 1,
        shape: None,
        members: Vec::new(),
        behavior: 2,
        chart: None,
        from: CellCorner {
            column: 0,
            row: 0,
            dx: 0,
            dy: 0,
        },
        to: CellCorner {
            column: 1,
            row: 1,
            dx: 0,
            dy: 0,
        },
        picture: Some(PictureReference {
            store_index: 7,
            crop: [0; 4],
            rotation: 0,
            clipboard_format: 9,
            auto_picture: true,
        }),
    };
    let bytes = vec![1, 2, 3, 4];
    (
        Pictures {
            anchors: BTreeMap::from([(0, vec![anchor])]),
            images: vec![(7, "png", bytes.clone())],
            unsupported_images: false,
            grouped: BTreeSet::new(),
        },
        sheet,
        "legacy-xls/image/7",
        bytes,
    )
}

#[cfg(test)]
mod tests {
    use super::super::drawing_anchors::PictureReference;
    use super::*;

    /// A model budget that never binds.
    fn unbounded() -> usize {
        usize::MAX
    }

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
            order: 1,
            shape: None,
            members: Vec::new(),
            behavior: 2,
            chart: None,
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
    /// Two pictures in OfficeArt document order.
    fn two() -> Vec<DrawingAnchor> {
        let mut second = anchor();
        second.order = 2;
        vec![anchor(), second]
    }
    fn pictures(anchors: Vec<DrawingAnchor>) -> Pictures {
        Pictures {
            anchors: BTreeMap::from([(0, anchors)]),
            images: vec![(1, "png", vec![1, 2, 3])],
            unsupported_images: false,
            grouped: BTreeSet::new(),
        }
    }

    #[test]
    fn native_picture_anchors_resolve_measured_cell_geometry() {
        // DxGCol 10 digits at mdw 7: trunc((2560 + trunc(128 / 7)) / 256 * 7)
        // = 70 px per column (ECMA-376 18.3.1.13); 300-twip rows.
        let (column, row) = (70 * 9525, 300 * 635);
        let mut budget = usize::MAX;
        let native = pictures(two())
            .resolve(&[("S".into(), sheet())], 7.0, &mut Vec::new())
            .into_models(&mut budget)
            .unwrap();
        let anchors = &native.sheets[&0];
        assert_eq!(anchors.len(), 2);
        let first = &anchors[0];
        assert_eq!((first.from_col, first.from_row), (0, 0));
        assert_eq!((first.to_col, first.to_row), (2, 3));
        // MS-XLS 2.5.193: dx in 1/1024 of the cell width, dy in 1/256 of
        // its height, rounded to whole EMUs.
        assert_eq!(
            (first.from_col_off, first.from_row_off),
            (-column / 2, -row / 2)
        );
        assert_eq!(
            (first.to_col_off, first.to_row_off),
            (((column as f64) / 4.0).round() as i64, row / 4)
        );
        assert_eq!(
            (first.native_ext_cx, first.native_ext_cy),
            (
                2 * column + first.to_col_off - first.from_col_off,
                3 * row + first.to_row_off - first.from_row_off
            )
        );
        assert_eq!(first.edit_as.as_deref(), Some("oneCell"));
        assert_eq!(first.image_path, "legacy-xls/image/1");
        // OfficeArt document order is the paint order.
        assert_eq!((first.z_order, anchors[1].z_order), (1, 2));
    }

    #[test]
    fn native_picture_models_preserve_transforms_and_charge_exact_payload_budget() {
        let resolve = || pictures(two()).resolve(&[("S".into(), sheet())], 7.0, &mut Vec::new());
        let mut remaining = usize::MAX;
        let native = resolve().into_models(&mut remaining).unwrap();
        let required = usize::MAX - remaining;
        assert_eq!(native.resources.len(), 1);
        assert_eq!(native.resources["legacy-xls/image/1"], [1, 2, 3]);
        let anchors = &native.sheets[&0];
        assert_eq!(anchors.len(), 2);
        assert!(anchors[0].z_order < anchors[1].z_order);
        assert_eq!(anchors[0].from_col_off, -333375);
        assert_eq!(anchors[0].rotation, Some(-90.0));
        assert_eq!(anchors[0].flip_h, Some(true));
        assert_eq!(anchors[0].flip_v, Some(true));
        assert_eq!(anchors[0].mime_type, "image/png");
        assert_eq!(
            anchors[0].src_rect.unwrap(),
            ooxml_common::blip::SrcRect {
                t: 0.5,
                b: -0.25,
                l: 0.25,
                r: 0.0,
            }
        );
        let mut exact = required;
        assert!(resolve().into_models(&mut exact).is_ok());
        assert_eq!(exact, 0);
        assert!(resolve().into_models(&mut (required - 1)).is_err());
    }

    #[test]
    fn pictures_sharing_a_store_entry_share_one_resource() {
        let mut second = anchor();
        second.shape_id = 2;
        second.object_id = 2;
        let mut warnings = Vec::new();
        let native = pictures(vec![anchor(), second])
            .resolve(&[("S".into(), sheet())], 7.0, &mut warnings)
            .into_models(&mut unbounded())
            .unwrap();
        assert!(warnings.is_empty());
        assert_eq!(native.resources.len(), 1);
        let anchors = &native.sheets[&0];
        assert_eq!(anchors.len(), 2);
        assert!(anchors
            .iter()
            .all(|anchor| anchor.image_path == "legacy-xls/image/1"));
    }

    #[test]
    fn refuses_missing_dimensions_and_never_retains_unreferenced_media() {
        let mut warnings = Vec::new();
        let native = pictures(vec![anchor()])
            .resolve(&[("S".into(), SheetData::default())], 7.0, &mut warnings)
            .into_models(&mut unbounded())
            .unwrap();
        assert!(native.sheets.is_empty() && native.resources.is_empty());
        assert_eq!(warnings, ["legacy-xls:unresolved-picture-geometry-omitted"]);
        let mut bad = anchor();
        bad.to = bad.from;
        let native = pictures(vec![bad])
            .resolve(&[("S".into(), sheet())], 7.0, &mut vec![])
            .into_models(&mut unbounded())
            .unwrap();
        assert!(native.resources.is_empty());
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
            anchors: (0..32).map(|i| (i, vec![picture.clone()])).collect(),
            images: vec![(1, "png", vec![1])],
            unsupported_images: false,
            grouped: BTreeSet::new(),
        };
        let mut warnings = vec![];
        let native = pictures
            .resolve(&sheets, 7.0, &mut warnings)
            .into_models(&mut unbounded())
            .unwrap();
        assert_eq!(native.sheets.len(), 30);
        assert_eq!(native.resources.len(), 1);
        assert_eq!(warnings, ["legacy-xls:unresolved-picture-geometry-omitted"]);
    }
}
