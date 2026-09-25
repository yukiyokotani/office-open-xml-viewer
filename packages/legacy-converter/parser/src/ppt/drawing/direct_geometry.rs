//! Direct projection of decoded OfficeArt paths into the renderer model.
//! ECMA-376 20.1.9.14 path coordinates are normalized by each path's w/h.
use crate::officeart::geometry::{Decoded, DecodedCommand};
use pptx_model::{PathCmd, PathPaint};

/// `fill`/`stroke` say whether any path is filled or stroked; `paint` carries
/// the per-path flags (ECMA-376 20.1.9.15) when the paths differ. MS-ODRAW
/// path escapes only switch a path's fill or line off (msopathEscape noFill
/// and noLine), so a projected fill mode is `none` or `norm`.
#[derive(Debug)]
pub(super) struct ModelGeometry {
    pub paths: Vec<Vec<PathCmd>>,
    pub fill: bool,
    pub stroke: bool,
    pub paint: Option<Vec<PathPaint>>,
}

pub(super) fn project(
    decoded: &Decoded,
    model_budget: &mut usize,
) -> Result<ModelGeometry, String> {
    let mut paths = decoded.paths();
    let Some(first) = paths.next() else {
        return Ok(ModelGeometry {
            paths: Vec::new(),
            fill: true,
            stroke: true,
            paint: None,
        });
    };
    let flags = (first.fill(), first.stroke());
    let mut mixed = false;
    let mut command_count = first.commands().len();
    for path in paths {
        mixed |= (path.fill(), path.stroke()) != flags;
        command_count = command_count
            .checked_add(path.commands().len())
            .ok_or_else(|| unsupported("OfficeArt custom-geometry model budget exceeded"))?;
    }
    let allocation = command_count
        .checked_mul(std::mem::size_of::<PathCmd>())
        .and_then(|bytes| {
            decoded
                .paths()
                .len()
                .checked_mul(std::mem::size_of::<Vec<PathCmd>>())
                .and_then(|outer| bytes.checked_add(outer))
        })
        .ok_or_else(|| unsupported("OfficeArt custom-geometry model budget exceeded"))?;
    *model_budget = model_budget
        .checked_sub(allocation)
        .ok_or_else(|| unsupported("OfficeArt custom-geometry model budget exceeded"))?;

    let width = decoded.width() as f64;
    let height = decoded.height() as f64;
    let mut result = Vec::new();
    result
        .try_reserve_exact(decoded.paths().len())
        .map_err(|_| unsupported("OfficeArt custom-geometry allocation failed"))?;
    for path in decoded.paths() {
        let mut commands = Vec::new();
        commands
            .try_reserve_exact(path.commands().len())
            .map_err(|_| unsupported("OfficeArt custom-geometry allocation failed"))?;
        for command in path.commands() {
            commands.push(match command {
                DecodedCommand::Move([x, y]) => PathCmd::MoveTo {
                    x: x as f64 / width,
                    y: y as f64 / height,
                },
                DecodedCommand::Line([x, y]) => PathCmd::LineTo {
                    x: x as f64 / width,
                    y: y as f64 / height,
                },
                DecodedCommand::Cubic([[x1, y1], [x2, y2], [x, y]]) => PathCmd::CubicBezTo {
                    x1: x1 as f64 / width,
                    y1: y1 as f64 / height,
                    x2: x2 as f64 / width,
                    y2: y2 as f64 / height,
                    x: x as f64 / width,
                    y: y as f64 / height,
                },
                DecodedCommand::Close => PathCmd::Close,
            });
        }
        result.push(commands);
    }
    if !mixed {
        return Ok(ModelGeometry {
            paths: result,
            fill: flags.0,
            stroke: flags.1,
            paint: None,
        });
    }
    let paint: Vec<PathPaint> = decoded
        .paths()
        .map(|path| PathPaint {
            fill: (!path.fill()).then(|| "none".to_owned()),
            stroke: path.stroke(),
        })
        .collect();
    Ok(ModelGeometry {
        paths: result,
        fill: paint.iter().any(|p| p.fill.is_none()),
        stroke: paint.iter().any(|p| p.stroke),
        paint: Some(paint),
    })
}

fn unsupported(message: &str) -> String {
    format!("UNSUPPORTED:{message}")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::officeart::geometry::Geometry;

    fn array(size: u16, body: Vec<u8>) -> Vec<u8> {
        let count = (body.len() / usize::from(size)) as u16;
        [
            count.to_le_bytes().as_slice(),
            count.to_le_bytes().as_slice(),
            size.to_le_bytes().as_slice(),
            &body,
        ]
        .concat()
    }

    fn vertices(points: &[[i32; 2]]) -> Vec<u8> {
        array(
            8,
            points
                .iter()
                .flat_map(|[x, y]| [x.to_le_bytes(), y.to_le_bytes()].concat())
                .collect(),
        )
    }

    fn segments(words: &[u16]) -> Vec<u8> {
        array(
            2,
            words.iter().flat_map(|word| word.to_le_bytes()).collect(),
        )
    }

    #[test]
    fn projects_signed_non_square_cubics_closes_and_all_subpaths() {
        let vertices = vertices(&[[-20, 20], [0, 40], [50, 60], [110, 220], [10, 20], [20, 40]]);
        let segments = segments(&[0x4000, 0x2001, 0x6001, 0x8000, 0x4000, 1, 0x6001, 0x8000]);
        let mut geometry = Geometry::default();
        geometry.complex(0x145, vertices.as_slice());
        geometry.complex(0x146, segments.as_slice());
        for (id, value) in [(0x140, -10i32), (0x141, 20), (0x142, 90), (0x143, 220)] {
            geometry.scalar(id, value as u32).unwrap();
        }
        let decoded = geometry.decode(&mut 100).unwrap().unwrap();
        let required = 6 * std::mem::size_of::<PathCmd>() + 2 * std::mem::size_of::<Vec<PathCmd>>();
        let mut model_budget = required;
        let model = project(&decoded, &mut model_budget).unwrap();
        assert_eq!(model_budget, 0);
        assert!(model.fill && model.stroke);
        assert_eq!(model.paths.len(), 2);
        match model.paths[0][0] {
            PathCmd::MoveTo { x, y } => assert_eq!((x, y), (-0.1, 0.0)),
            _ => panic!("expected move"),
        }
        match model.paths[0][1] {
            PathCmd::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x,
                y,
            } => {
                assert_eq!((x1, y1, x2, y2, x, y), (0.1, 0.1, 0.6, 0.2, 1.2, 1.0));
            }
            _ => panic!("expected cubic"),
        }
        assert!(matches!(model.paths[0][2], PathCmd::Close));
        assert!(matches!(model.paths[1][0], PathCmd::MoveTo { .. }));
        assert!(matches!(model.paths[1][1], PathCmd::LineTo { .. }));
        assert!(matches!(model.paths[1][2], PathCmd::Close));

        let mut exhausted = required - 1;
        assert!(project(&decoded, &mut exhausted).is_err());
        assert_eq!(exhausted, required - 1);
    }

    #[test]
    fn mixed_path_paint_projects_per_path_flags_and_empty_source_decodes_to_none() {
        let vertices = vertices(&[[0, 0], [1, 1], [2, 2], [3, 3]]);
        let segments = segments(&[0x4000, 1, 0xaa00, 0x8000, 0x4000, 1, 0x6001, 0x8000]);
        let mut geometry = Geometry::default();
        geometry.complex(0x145, vertices.as_slice());
        geometry.complex(0x146, segments.as_slice());
        let decoded = geometry.decode(&mut 100).unwrap().unwrap();
        // Path 1 ends with msopathEscape noFill (0xaa00); path 2 is closed
        // and painted normally.
        let model = project(&decoded, &mut 1000).unwrap();
        assert!(model.fill && model.stroke);
        assert_eq!(
            model.paint,
            Some(vec![
                PathPaint {
                    fill: Some("none".into()),
                    stroke: true
                },
                PathPaint {
                    fill: None,
                    stroke: true
                },
            ])
        );
        assert!(Geometry::default().decode(&mut 0).unwrap().is_none());
    }
}
