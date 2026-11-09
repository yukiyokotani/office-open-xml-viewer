//! Extruded custom shapes (ODF 1.3 Part 3 §19.157–§19.172
//! `draw:extrusion-*`, `dr3d:projection`) mapped onto the DrawingML 3-D
//! scene the PPTX renderer draws (`scene3d` camera, `sp3d` extrusion).
//!
//! The mapping inverts LibreOffice's DrawingML import
//! (oox `scene3dhelper.cxx`), which is how these properties were produced:
//! - the camera rotation (lat, lon, rev) becomes the extrusion
//!   `RotateAngle` pair and the shape rotation through
//!   `getAPIAnglesFromOOXAngle` (Z(rev)·X(lat)·Y(lon), decomposed); it is
//!   inverted numerically, with the shape's own 2-D rotation folded into
//!   the revolution (the shape is then not rotated in 2-D);
//! - the preset's projection, skew (oblique presets) and view point
//!   (perspective presets; a field of view `fov` gives
//!   `z = 15976 / tan(fov / 2)` in 1/100 mm) select the camera preset;
//! - the depth pair is (extrusionH, z / extrusionH) in 1/100 mm.

use crate::odp_style::Nd;
use crate::odp_units::{color, length_emu};
use crate::xml::NS_DRAW;
use pptx_model::{Camera3d, LightRig, Rot3d, Scene3d, Sp3d};

pub const NS_DR3D: &str = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";

/// LibreOffice's `getAPIAnglesFromOOXAngle`, all angles in radians:
/// (lat, lon, rev) → (fX, fY, fZ).
fn forward(lat: f64, lon: f64, rev: f64) -> (f64, f64, f64) {
    type M = [[f64; 3]; 3];
    let mul = |a: &M, b: &M| -> M {
        let mut o = [[0.0; 3]; 3];
        for (i, row) in o.iter_mut().enumerate() {
            for (j, cell) in row.iter_mut().enumerate() {
                *cell = (0..3).map(|k| a[i][k] * b[k][j]).sum();
            }
        }
        o
    };
    let (sl, cl) = lat.sin_cos();
    let (so, co) = lon.sin_cos();
    let (sr, cr) = rev.sin_cos();
    let x: M = [[1.0, 0.0, 0.0], [0.0, cl, sl], [0.0, -sl, cl]];
    let y: M = [[co, 0.0, -so], [0.0, 1.0, 0.0], [so, 0.0, co]];
    let z: M = [[cr, sr, 0.0], [-sr, cr, 0.0], [0.0, 0.0, 1.0]];
    let t = mul(&mul(&z, &x), &y);
    let fy = (-t[0][2]).clamp(-1.0, 1.0).asin();
    let c = fy.cos();
    if c.abs() < 1e-9 {
        (0.0, fy, t[2][1].atan2(t[1][1]))
    } else {
        (
            (-t[1][2] / c).atan2(t[2][2] / c),
            fy,
            (t[0][1] / c).atan2(t[0][0] / c),
        )
    }
}

fn wrap(a: f64) -> f64 {
    let tau = std::f64::consts::TAU;
    let mut v = a % tau;
    if v > std::f64::consts::PI {
        v -= tau;
    } else if v < -std::f64::consts::PI {
        v += tau;
    }
    v
}

/// Invert [`forward`]: the (lat, lon, rev) whose decomposition is
/// (fx, fy, fz). Damped Newton iterations from several starts.
pub fn invert(fx: f64, fy: f64, fz: f64) -> (f64, f64, f64) {
    let err = |p: (f64, f64, f64)| {
        let (a, b, c) = forward(p.0, p.1, p.2);
        (wrap(a - fx), wrap(b - fy), wrap(c - fz))
    };
    let norm = |e: (f64, f64, f64)| e.0.abs() + e.1.abs() + e.2.abs();
    let mut best = (0.0, 0.0, 0.0);
    let mut best_err = f64::INFINITY;
    for start in [(-fx, fy, fz), (fx, fy, fz), (0.0, 0.0, 0.0), (-fx, -fy, fz)] {
        let mut p = start;
        for _ in 0..60 {
            let e = err(p);
            let n = norm(e);
            if n < 1e-12 {
                break;
            }
            let h = 1e-6;
            let col = |dp: (f64, f64, f64)| {
                let e2 = err((p.0 + dp.0, p.1 + dp.1, p.2 + dp.2));
                (
                    wrap(e2.0 - e.0) / h,
                    wrap(e2.1 - e.1) / h,
                    wrap(e2.2 - e.2) / h,
                )
            };
            let (c0, c1, c2) = (col((h, 0.0, 0.0)), col((0.0, h, 0.0)), col((0.0, 0.0, h)));
            // Solve J·d = −e (Cramer's rule).
            let j = [[c0.0, c1.0, c2.0], [c0.1, c1.1, c2.1], [c0.2, c1.2, c2.2]];
            let det = |m: [[f64; 3]; 3]| {
                m[0][0] * (m[1][1] * m[2][2] - m[1][2] * m[2][1])
                    - m[0][1] * (m[1][0] * m[2][2] - m[1][2] * m[2][0])
                    + m[0][2] * (m[1][0] * m[2][1] - m[1][1] * m[2][0])
            };
            let d = det(j);
            if d.abs() < 1e-12 {
                break;
            }
            let rhs = [-e.0, -e.1, -e.2];
            let solve = |k: usize| {
                let mut m = j;
                for r in 0..3 {
                    m[r][k] = rhs[r];
                }
                det(m) / d
            };
            let step = (solve(0), solve(1), solve(2));
            p = (p.0 + step.0, p.1 + step.1, p.2 + step.2);
        }
        let n = norm(err(p));
        if n < best_err {
            best_err = n;
            best = p;
        }
    }
    (wrap(best.0), wrap(best.1), wrap(best.2))
}

/// The DrawingML scene of an extruded custom shape. `odf_rotation` is the
/// shape's `draw:transform` rotation in degrees counter-clockwise (LibreOffice
/// writes the decomposed fZ there); the caller then draws the shape without
/// its own 2-D rotation. `None` when the geometry is not extruded.
pub fn scene(
    eg: Nd<'_>,
    graphic: &crate::odp_style::Props<'_>,
    odf_rotation: f64,
) -> Option<(Scene3d, Sp3d)> {
    if eg.attribute((NS_DRAW, "extrusion")) != Some("true") {
        return None;
    }
    let pair = |name: &str| -> Option<(String, String)> {
        let v = eg.attribute((NS_DRAW, name))?;
        let mut it = v.split_whitespace();
        Some((it.next()?.to_string(), it.next().unwrap_or("0").to_string()))
    };
    let (ax, ay) = pair("extrusion-rotation-angle")
        .map(|(a, b)| {
            (
                a.parse::<f64>().unwrap_or(0.0),
                b.parse::<f64>().unwrap_or(0.0),
            )
        })
        .unwrap_or((0.0, 0.0));
    let (lat, lon, rev) = invert(ax.to_radians(), ay.to_radians(), odf_rotation.to_radians());
    let deg = |r: f64| {
        let d = r.to_degrees().rem_euclid(360.0);
        if (d - 360.0).abs() < 1e-6 {
            0.0
        } else {
            d
        }
    };
    let rot = Rot3d {
        lat: deg(lat),
        lon: deg(lon),
        rev: deg(rev),
    };
    let parallel = eg.attribute((NS_DR3D, "projection")) == Some("parallel");
    let (skew_amount, skew_angle) = pair("extrusion-skew")
        .map(|(a, b)| {
            (
                a.parse::<f64>().unwrap_or(0.0),
                b.parse::<f64>().unwrap_or(0.0),
            )
        })
        .unwrap_or((50.0, 45.0));
    let side = |angle: f64| -> &'static str {
        match angle.round() as i64 {
            90 => "Bottom",
            45 => "BottomLeft",
            135 => "BottomRight",
            -360 | 0 | 360 => "Left",
            180 | -180 => "Right",
            -90 | 270 => "Top",
            -45 | 315 => "TopLeft",
            -135 | 225 => "TopRight",
            _ => "TopRight",
        }
    };
    let mut fov = None;
    let prst = if parallel {
        if skew_amount.abs() < 1e-9 {
            "orthographicFront".to_string()
        } else if (skew_amount - 30.0).abs() < 1e-6 {
            format!("oblique{}", side(skew_angle))
        } else {
            format!("legacyOblique{}", side(skew_angle))
        }
    } else {
        let vp: Vec<f64> = eg
            .attribute((NS_DRAW, "extrusion-viewpoint"))
            .unwrap_or("3.5cm -3.5cm 25cm")
            .split_whitespace()
            .filter_map(length_emu)
            .map(|e| e / 360.0)
            .collect();
        let (vx, vy, vz) = (
            vp.first().copied().unwrap_or(0.0),
            vp.get(1).copied().unwrap_or(0.0),
            vp.get(2).copied().unwrap_or(25_000.0),
        );
        if vx.abs() > 1.0 || vy.abs() > 1.0 {
            let v = match (vx.signum() as i64, vy.signum() as i64) {
                (0, -1) => "Top",
                (0, 1) => "Bottom",
                (-1, 0) => "Left",
                (1, 0) => "Right",
                (-1, -1) => "TopLeft",
                (1, -1) => "TopRight",
                (-1, 1) => "BottomLeft",
                (1, 1) => "BottomRight",
                _ => "Front",
            };
            format!("legacyPerspective{v}")
        } else {
            if vz > 0.0 {
                fov = Some((2.0 * (15976.0 / vz).atan()).to_degrees());
            }
            "perspectiveFront".to_string()
        }
    };
    let (depth, fraction) = pair("extrusion-depth")
        .map(|(a, b)| {
            (
                length_emu(&a).unwrap_or(0.0),
                b.parse::<f64>().unwrap_or(0.0),
            )
        })
        .unwrap_or((0.0, 0.0));
    let extrusion_clr = (eg.attribute((NS_DRAW, "extrusion-color")) == Some("true"))
        .then(|| graphic.get(NS_DRAW, "secondary-fill-color").and_then(color))
        .flatten();
    let material = if eg.attribute((NS_DR3D, "shade-mode")) == Some("flat") {
        "flat"
    } else if eg.attribute((NS_DRAW, "extrusion-metal")) == Some("true") {
        "metal"
    } else {
        "warmMatte"
    };
    let identity = rot.lat == 0.0 && rot.lon == 0.0 && rot.rev == 0.0;
    Some((
        Scene3d {
            camera: Camera3d {
                prst,
                fov,
                zoom: None,
                rot: (!identity).then_some(rot),
            },
            light_rig: Some(LightRig {
                rig: "threePt".into(),
                dir: "t".into(),
                rot: None,
            }),
        },
        Sp3d {
            z: (fraction * depth).round() as i64,
            extrusion_h: depth.round() as i64,
            contour_w: 0,
            contour_clr: None,
            extrusion_clr,
            prst_material: material.to_string(),
            bevel_t: None,
            bevel_b: None,
        },
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reproduces_libreoffice_export_of_a_known_camera() {
        // A legacy camera with lat 25°, lon 35°, rev 17° is written by
        // LibreOffice as rotation angles (-33.894°, 26.570°) and a shape
        // rotation of 0.30072 rad.
        let (fx, fy, fz) = forward(25f64.to_radians(), 35f64.to_radians(), 17f64.to_radians());
        assert!(
            (fx.to_degrees() + 33.8939).abs() < 1e-3,
            "{}",
            fx.to_degrees()
        );
        assert!(
            (fy.to_degrees() - 26.5705).abs() < 1e-3,
            "{}",
            fy.to_degrees()
        );
        assert!((fz - 0.300720).abs() < 1e-3, "{fz}");
    }

    #[test]
    fn inverts_libreoffice_angle_decomposition() {
        for (lat, lon, rev) in [
            (25.0f64, 35.0f64, 17.0f64),
            (0.0, 30.0, 0.0),
            (-20.0, 45.0, 300.0),
            (10.0, 0.0, 0.0),
        ] {
            let (fx, fy, fz) = forward(lat.to_radians(), lon.to_radians(), rev.to_radians());
            let (a, b, c) = invert(fx, fy, fz);
            let (x2, y2, z2) = forward(a, b, c);
            assert!(
                (wrap(x2 - fx)).abs() < 1e-6
                    && (wrap(y2 - fy)).abs() < 1e-6
                    && (wrap(z2 - fz)).abs() < 1e-6
            );
        }
    }
}
