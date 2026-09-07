//! Bounded projection of classic OfficeArt linear-gradient parameters.
//!
//! MS-ODRAW 2.2.50, 2.2.51, 2.2.61, 2.3.7.14, 2.3.7.15, 2.3.7.26,
//! and 2.3.7.32 define the signed focus, 16.16 angle and shade-color fields.
//! The two-leg focus projection is behavior observed across controlled Office
//! renders; it is not claimed as a normative MS-ODRAW rule.

use std::mem::size_of;

use super::{super::unsupported, ShadeColor};

pub(crate) const SOURCE_ONE: u32 = 65_536;
pub(crate) const POSITION_DENOMINATOR: u64 = 6_553_600;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) struct ProjectedShadeStop {
    pub color: u32,
    position_numerator: u64,
}

impl ProjectedShadeStop {
    pub(crate) fn position(self) -> f64 {
        self.position_numerator as f64 / POSITION_DENOMINATOR as f64
    }

    /// Quantize to DrawingML's 100,000 position units. Round-half-up is an
    /// explicit compatibility output policy, not an Office-derived rule.
    pub(crate) fn position_units(self) -> u32 {
        round_half_up(self.position_numerator, POSITION_DENOMINATOR, 100_000) as u32
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) struct RationalAngle {
    numerator: i64,
    denominator: u32,
}

impl RationalAngle {
    pub(crate) fn degrees(self) -> f64 {
        self.numerator as f64 / f64::from(self.denominator)
    }

    /// Quantize to DrawingML's 60,000 angle units per degree. Round-half-up is
    /// an explicit compatibility output policy, not an Office-derived rule.
    pub(crate) fn angle_units(self) -> i32 {
        let units = round_half_up(self.numerator as u64, u64::from(self.denominator), 60_000);
        (units % (360 * 60_000)) as i32
    }
}

#[derive(Debug, Eq, PartialEq)]
pub(crate) struct Projection {
    pub stops: Vec<ProjectedShadeStop>,
    pub angle: RationalAngle,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) struct Requirements {
    pub work_iterations: usize,
    pub scratch_bytes: usize,
    pub output_bytes: usize,
}

fn round_half_up(numerator: u64, denominator: u64, units: u64) -> u64 {
    let scaled = u128::from(numerator) * u128::from(units);
    ((scaled + u128::from(denominator) / 2) / u128::from(denominator)) as u64
}

fn charge(budget: &mut usize, amount: usize, message: &str) -> Result<(), String> {
    *budget = budget
        .checked_sub(amount)
        .ok_or_else(|| unsupported(message))?;
    Ok(())
}

fn normalized_len(authored: &[ShadeColor]) -> Result<usize, String> {
    let add_front = authored.first().is_none_or(|stop| stop.position != 0);
    let add_back = authored
        .last()
        .is_none_or(|stop| stop.position != SOURCE_ONE);
    authored
        .len()
        .checked_add(usize::from(add_front) + usize::from(add_back))
        .ok_or_else(|| unsupported("OfficeArt shade stop count overflow"))
}

pub(crate) fn scratch_bytes(authored: &[ShadeColor]) -> Result<usize, String> {
    normalized_len(authored)?
        .checked_mul(size_of::<ShadeColor>())
        .ok_or_else(|| unsupported("OfficeArt shade scratch size overflow"))
}

pub(crate) fn requirements(authored: &[ShadeColor], focus: i32) -> Result<Requirements, String> {
    if !(-100..=100).contains(&focus) {
        return Err(unsupported("OfficeArt shade focus is outside -100..100"));
    }
    let normalized_len = normalized_len(authored)?;
    let output_len = if focus == 0 || focus.abs() == 100 {
        normalized_len
    } else {
        normalized_len
            .checked_mul(2)
            .and_then(|value| value.checked_sub(1))
            .ok_or_else(|| unsupported("OfficeArt projected shade stop count overflow"))?
    };
    Ok(Requirements {
        work_iterations: authored
            .len()
            .checked_add(normalized_len)
            .and_then(|value| value.checked_add(output_len))
            .ok_or_else(|| unsupported("OfficeArt shade projection work overflow"))?,
        scratch_bytes: scratch_bytes(authored)?,
        output_bytes: output_len
            .checked_mul(size_of::<ProjectedShadeStop>())
            .ok_or_else(|| unsupported("OfficeArt shade output size overflow"))?,
    })
}

pub(crate) fn project(
    authored: &[ShadeColor],
    scalar_front: u32,
    scalar_back: u32,
    focus: i32,
    angle_16_16: i32,
    work_budget: &mut usize,
    scratch_byte_budget: &mut usize,
    output_byte_budget: &mut usize,
) -> Result<Projection, String> {
    let add_front = authored.first().is_none_or(|stop| stop.position != 0);
    let add_back = authored
        .last()
        .is_none_or(|stop| stop.position != SOURCE_ONE);
    let normalized_len = normalized_len(authored)?;
    let requirements = requirements(authored, focus)?;
    let output_len = requirements.output_bytes / size_of::<ProjectedShadeStop>();

    // Charge all caller-owned resource budgets before either allocation.
    charge(
        work_budget,
        requirements.work_iterations,
        "OfficeArt shade projection work budget exceeded",
    )?;
    let mut previous = 0;
    for (index, stop) in authored.iter().enumerate() {
        if stop.position > SOURCE_ONE || (index != 0 && stop.position < previous) {
            return Err(unsupported(
                "OfficeArt shade stops are not ordered 16.16 fractions",
            ));
        }
        previous = stop.position;
    }
    charge(
        scratch_byte_budget,
        requirements.scratch_bytes,
        "OfficeArt shade projection scratch budget exceeded",
    )?;
    charge(
        output_byte_budget,
        requirements.output_bytes,
        "OfficeArt shade projection output budget exceeded",
    )?;

    let mut normalized = Vec::new();
    normalized
        .try_reserve_exact(normalized_len)
        .map_err(|_| unsupported("OfficeArt shade scratch allocation failed"))?;
    if add_front {
        normalized.push(ShadeColor {
            color: scalar_front,
            position: 0,
        });
    }
    normalized.extend_from_slice(authored);
    if add_back {
        normalized.push(ShadeColor {
            color: scalar_back,
            position: SOURCE_ONE,
        });
    }

    let mut stops = Vec::new();
    stops
        .try_reserve_exact(output_len)
        .map_err(|_| unsupported("OfficeArt shade output allocation failed"))?;
    let forward = |stop: ShadeColor, scale: u64, offset: u64| ProjectedShadeStop {
        color: stop.color,
        position_numerator: offset + scale * u64::from(stop.position),
    };
    let reverse = |stop: ShadeColor, scale: u64, offset: u64| ProjectedShadeStop {
        color: stop.color,
        position_numerator: offset + scale * u64::from(SOURCE_ONE - stop.position),
    };

    match focus {
        -100 | 100 => stops.extend(normalized.iter().copied().map(|stop| forward(stop, 100, 0))),
        0 => stops.extend(
            normalized
                .iter()
                .rev()
                .copied()
                .map(|stop| reverse(stop, 100, 0)),
        ),
        value if value > 0 => {
            let center = value as u64;
            stops.extend(
                normalized
                    .iter()
                    .copied()
                    .map(|stop| forward(stop, center, 0)),
            );
            // Skip only the generated identical join; authored hard stops remain.
            stops.extend(
                normalized
                    .iter()
                    .rev()
                    .copied()
                    .skip(1)
                    .map(|stop| reverse(stop, 100 - center, center * u64::from(SOURCE_ONE))),
            );
        }
        value => {
            let center = (100 + value) as u64;
            let tail = (-value) as u64;
            stops.extend(
                normalized
                    .iter()
                    .rev()
                    .copied()
                    .map(|stop| reverse(stop, center, 0)),
            );
            stops.extend(
                normalized
                    .iter()
                    .copied()
                    .skip(1)
                    .map(|stop| forward(stop, tail, center * u64::from(SOURCE_ONE))),
            );
        }
    }
    debug_assert_eq!(stops.len(), output_len);

    let full_turn = 360_i64 * i64::from(SOURCE_ONE);
    let angle = (90_i64 * i64::from(SOURCE_ONE) - i64::from(angle_16_16)).rem_euclid(full_turn);
    Ok(Projection {
        stops,
        angle: RationalAngle {
            numerator: angle,
            denominator: SOURCE_ONE,
        },
    })
}

#[cfg(test)]
mod tests {
    #![allow(const_item_mutation)]

    use super::*;

    const PURPLE: u32 = 1;
    const BLUE: u32 = 2;
    const CYAN: u32 = 3;
    const GREEN: u32 = 4;
    const SOURCE: [ShadeColor; 4] = [
        ShadeColor {
            color: PURPLE,
            position: 0,
        },
        ShadeColor {
            color: PURPLE,
            position: 0x2e14,
        },
        ShadeColor {
            color: BLUE,
            position: 0x63d7,
        },
        ShadeColor {
            color: CYAN,
            position: 0x9c29,
        },
    ];

    fn run(focus: i32, angle: i32) -> Projection {
        project(
            &SOURCE,
            PURPLE,
            GREEN,
            focus,
            angle,
            &mut usize::MAX,
            &mut usize::MAX,
            &mut usize::MAX,
        )
        .unwrap()
    }

    fn positions(focus: i32) -> Vec<u64> {
        run(focus, 270 << 16)
            .stops
            .into_iter()
            .map(|stop| stop.position_numerator)
            .collect()
    }

    fn colors(focus: i32) -> Vec<u32> {
        run(focus, 270 << 16)
            .stops
            .into_iter()
            .map(|stop| stop.color)
            .collect()
    }

    #[test]
    fn seven_office_measured_focus_cases_follow_exact_piecewise_rationals() {
        let (p1, p2, p3, one) = (0x2e14_u64, 0x63d7_u64, 0x9c29_u64, u64::from(SOURCE_ONE));
        assert_eq!(
            positions(100),
            vec![0, 100 * p1, 100 * p2, 100 * p3, 100 * one]
        );
        assert_eq!(positions(-100), positions(100));
        assert_eq!(
            positions(0),
            vec![
                0,
                100 * (one - p3),
                100 * (one - p2),
                100 * (one - p1),
                100 * one
            ]
        );
        assert_eq!(colors(100), vec![PURPLE, PURPLE, BLUE, CYAN, GREEN]);
        assert_eq!(colors(-100), colors(100));
        assert_eq!(colors(0), vec![GREEN, CYAN, BLUE, PURPLE, PURPLE]);
        for focus in [25_u64, 50] {
            assert_eq!(
                positions(focus as i32),
                vec![
                    0,
                    focus * p1,
                    focus * p2,
                    focus * p3,
                    focus * one,
                    focus * one + (100 - focus) * (one - p3),
                    focus * one + (100 - focus) * (one - p2),
                    focus * one + (100 - focus) * (one - p1),
                    100 * one
                ]
            );
            assert_eq!(
                colors(focus as i32),
                vec![PURPLE, PURPLE, BLUE, CYAN, GREEN, CYAN, BLUE, PURPLE, PURPLE]
            );
        }
        for focus in [-50_i32, -25] {
            let (center, tail) = ((100 + focus) as u64, (-focus) as u64);
            assert_eq!(
                positions(focus),
                vec![
                    0,
                    center * (one - p3),
                    center * (one - p2),
                    center * (one - p1),
                    center * one,
                    center * one + tail * p1,
                    center * one + tail * p2,
                    center * one + tail * p3,
                    100 * one
                ]
            );
            assert_eq!(
                colors(focus),
                vec![GREEN, CYAN, BLUE, PURPLE, PURPLE, PURPLE, BLUE, CYAN, GREEN]
            );
        }
    }

    #[test]
    fn explicit_array_endpoints_win_and_missing_last_uses_back_color() {
        let explicit = [
            ShadeColor {
                color: CYAN,
                position: 0,
            },
            ShadeColor {
                color: BLUE,
                position: SOURCE_ONE,
            },
        ];
        let result = project(
            &explicit,
            PURPLE,
            GREEN,
            100,
            0,
            &mut usize::MAX,
            &mut usize::MAX,
            &mut usize::MAX,
        )
        .unwrap();
        assert_eq!(
            result.stops.iter().map(|s| s.color).collect::<Vec<_>>(),
            vec![CYAN, BLUE]
        );
        let missing_last = [ShadeColor {
            color: CYAN,
            position: 0,
        }];
        let result = project(
            &missing_last,
            PURPLE,
            GREEN,
            100,
            0,
            &mut usize::MAX,
            &mut usize::MAX,
            &mut usize::MAX,
        )
        .unwrap();
        assert_eq!(result.stops.last().unwrap().color, GREEN);
    }

    #[test]
    // Missing-first and empty-array normalization are algebraic, not Office-verified.
    fn unverified_missing_first_and_empty_arrays_are_normalized() {
        let missing_first = [ShadeColor {
            color: CYAN,
            position: SOURCE_ONE,
        }];
        let result = project(
            &missing_first,
            PURPLE,
            GREEN,
            100,
            0,
            &mut usize::MAX,
            &mut usize::MAX,
            &mut usize::MAX,
        )
        .unwrap();
        assert_eq!(result.stops.first().unwrap().color, PURPLE);
        let result = project(
            &[],
            PURPLE,
            GREEN,
            100,
            0,
            &mut usize::MAX,
            &mut usize::MAX,
            &mut usize::MAX,
        )
        .unwrap();
        assert_eq!(
            result.stops.iter().map(|s| s.color).collect::<Vec<_>>(),
            vec![PURPLE, GREEN]
        );
    }

    #[test]
    fn authored_duplicate_endpoints_survive_both_focus_signs() {
        let source = [
            ShadeColor {
                color: PURPLE,
                position: 0,
            },
            ShadeColor {
                color: BLUE,
                position: 0,
            },
            ShadeColor {
                color: CYAN,
                position: SOURCE_ONE,
            },
            ShadeColor {
                color: GREEN,
                position: SOURCE_ONE,
            },
        ];
        for focus in [25, -25] {
            let result = project(
                &source,
                0,
                0,
                focus,
                0,
                &mut usize::MAX,
                &mut usize::MAX,
                &mut usize::MAX,
            )
            .unwrap();
            assert_eq!(result.stops.len(), source.len() * 2 - 1);
            assert!(result
                .stops
                .windows(2)
                .any(
                    |pair| pair[0].position_numerator == pair[1].position_numerator
                        && pair[0].color != pair[1].color
                ));
        }
    }

    #[test]
    fn angle_mapping_handles_measured_fractional_and_i32_extremes() {
        for (input, expected) in [
            (0, 90),
            (45 << 16, 45),
            (90 << 16, 0),
            (180 << 16, 270),
            (270 << 16, 180),
            (315 << 16, 135),
        ] {
            assert_eq!(
                run(100, input).angle.numerator,
                expected * i64::from(SOURCE_ONE)
            );
        }
        for input in [
            (90 << 16) + SOURCE_ONE as i32 / 2,
            -(SOURCE_ONE as i32 / 2),
            i32::MIN,
            i32::MAX,
        ] {
            let expected = (90_i64 * i64::from(SOURCE_ONE) - i64::from(input))
                .rem_euclid(360_i64 * i64::from(SOURCE_ONE));
            assert_eq!(run(100, input).angle.numerator, expected);
        }
    }

    #[test]
    fn exact_and_short_work_scratch_and_output_budgets() {
        let normalized = 5;
        let output = 9;
        let needs = (
            SOURCE.len() + normalized + output,
            normalized * size_of::<ShadeColor>(),
            output * size_of::<ProjectedShadeStop>(),
        );
        for (work, scratch, retained) in [
            (needs.0 - 1, needs.1, needs.2),
            (needs.0, needs.1 - 1, needs.2),
            (needs.0, needs.1, needs.2 - 1),
        ] {
            assert!(project(
                &SOURCE,
                PURPLE,
                GREEN,
                25,
                0,
                &mut work.clone(),
                &mut scratch.clone(),
                &mut retained.clone()
            )
            .is_err());
        }
        let (mut work, mut scratch, mut retained) = needs;
        assert!(project(
            &SOURCE,
            PURPLE,
            GREEN,
            25,
            0,
            &mut work,
            &mut scratch,
            &mut retained
        )
        .is_ok());
        assert_eq!((work, scratch, retained), (0, 0, 0));
    }

    #[test]
    fn rejects_invalid_focus_positions_and_order() {
        assert!(project(
            &SOURCE,
            0,
            0,
            101,
            0,
            &mut usize::MAX,
            &mut usize::MAX,
            &mut usize::MAX
        )
        .is_err());
        let unordered = [
            ShadeColor {
                color: 0,
                position: 2,
            },
            ShadeColor {
                color: 0,
                position: 1,
            },
        ];
        assert!(project(
            &unordered,
            0,
            0,
            0,
            0,
            &mut usize::MAX,
            &mut usize::MAX,
            &mut usize::MAX
        )
        .is_err());
        let out_of_range = [ShadeColor {
            color: 0,
            position: SOURCE_ONE + 1,
        }];
        assert!(project(
            &out_of_range,
            0,
            0,
            0,
            0,
            &mut usize::MAX,
            &mut usize::MAX,
            &mut usize::MAX
        )
        .is_err());
    }

    #[test]
    fn rational_native_values_and_compatibility_quantization_are_bounded() {
        for numerator in [0, 1, 32_768, POSITION_DENOMINATOR - 1, POSITION_DENOMINATOR] {
            let stop = ProjectedShadeStop {
                color: 0,
                position_numerator: numerator,
            };
            assert_eq!(
                stop.position(),
                numerator as f64 / POSITION_DENOMINATOR as f64
            );
            let exact = numerator as f64 * 100_000.0 / POSITION_DENOMINATOR as f64;
            assert!((f64::from(stop.position_units()) - exact).abs() <= 0.5);
        }
        for numerator in [
            0,
            1,
            32_768,
            90_i64 * i64::from(SOURCE_ONE),
            360_i64 * i64::from(SOURCE_ONE) - 1,
        ] {
            let angle = RationalAngle {
                numerator,
                denominator: SOURCE_ONE,
            };
            assert_eq!(angle.degrees(), numerator as f64 / f64::from(SOURCE_ONE));
            let exact = numerator as f64 * 60_000.0 / f64::from(SOURCE_ONE);
            assert!((f64::from(angle.angle_units()) - exact).abs() <= 0.5);
        }
        assert_eq!(
            ProjectedShadeStop {
                color: 0,
                position_numerator: POSITION_DENOMINATOR
            }
            .position_units(),
            100_000
        );
        assert_eq!(run(100, 90 << 16).angle.angle_units(), 0);
        assert_eq!(run(100, -(270 << 16)).angle.angle_units(), 0);
        let almost_full_turn = RationalAngle {
            numerator: 360_i64 * 131_072 - 1,
            denominator: 131_072,
        };
        assert_eq!(almost_full_turn.angle_units(), 0);
        let linear_error = (almost_full_turn.degrees() * 60_000.0
            - f64::from(almost_full_turn.angle_units()))
        .abs();
        assert!((21_600_000.0 - linear_error).abs() <= 0.5);
    }
}
