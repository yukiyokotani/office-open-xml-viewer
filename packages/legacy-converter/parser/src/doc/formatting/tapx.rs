//! Structural validation for table-style UpxTapx property arrays.
//!
//! [MS-DOC] 2.9.340 restricts UpxTapx to table SPRMs, ignores sprmTIstd,
//! excludes direct-row and sprmTIstd-preserved properties, and gives style
//! 0x000B its sole sprmTWidthBefore exception. Section 2.6.3 defines sprmTCnf
//! and the six cell-border-style exceptions in its bounded nested grpprl.

use super::{cnf, Budget, Sprms};
use crate::doc::unsupported;

const DEFAULT_TABLE_STYLE: usize = 0x000b;
const T_ISTD: u16 = 0x563a;
const T_WIDTH_BEFORE: u16 = 0xf617;
const T_CNF: u16 = 0xd66a;

// [MS-DOC] 2.9.340's explicit 41-member exclusion list. The first six
// conditional border codes are admitted only while validating sprmTCnf.
const EXPLICITLY_PROHIBITED: [u16; 41] = [
    0x9601, 0xd608, 0xd609, 0xd60c, 0xd612, 0xd616, 0xf618, 0x3619, 0xd61a, 0xd61b, 0xd61c, 0xd61d,
    0xd620, 0x7621, 0x5622, 0x7623, 0x5624, 0x5625, 0x7629, 0xd62b, 0xd62c, 0xd62f, 0xd632, 0xd635,
    0xf636, 0xd639, 0xd642, 0xd660, 0xd662, 0x5664, 0x3465, 0x7469, 0xd670, 0xd671, 0xd672, 0xd47f,
    0xd680, 0xd681, 0xd682, 0xd683, 0xd684,
];

// SPRM encodings for semantic properties that [MS-DOC] 2.6.3 says sprmTIstd
// preserves. The section describes the preserved properties and gives example
// SPRMs; this list is the validated encoding subset used to reject UpxTapx,
// not a claim that those examples exhaust every compatible encoding.
const PRESERVED_BY_T_ISTD: [u16; 16] = [
    0x3668, // sprmTWall
    0xd667, // sprmTPropRMark
    0x560b, // sprmTFBiDi
    0x7479, // sprmTRsid
    0x360d, // sprmTPc
    0x940e, // sprmTDxaAbs
    0x940f, // sprmTDyaAbs
    0x9410, // sprmTDxaFromText
    0x9411, // sprmTDyaFromText
    0x941e, // sprmTDxaFromTextRight
    0x941f, // sprmTDyaFromTextBottom
    0x9602, // sprmTDxaGapHalf
    0x9407, // sprmTDyaRowHeight
    0xf614, // sprmTTableWidth
    0x3615, // sprmTFAutofit
    0x740a, // sprmTTlp
];

const CONDITIONAL_BORDER_EXCEPTIONS: [u16; 6] = [0xd47f, 0xd680, 0xd681, 0xd682, 0xd683, 0xd684];

// These two diagonal counterparts are not in 2.9.340's six exceptions, but
// their individual 2.6.3 definitions independently require CNFOperand scope.
const OTHER_CNF_ONLY: [u16; 2] = [0xd685, 0xd686];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Scope {
    Unconditional,
    Conditional(u16),
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(super) struct Report {
    pub(super) ignored_tistd: usize,
    pub(super) unconditional_properties: usize,
    pub(super) conditional_properties: usize,
    /// True means the TAPX is valid, but the caller's current projection does
    /// not support at least one of its properties.
    pub(super) unsupported: bool,
}

/// Strictly validate and visit one borrowed UpxTapx in a single bounded pass.
///
/// `visit` describes implementation coverage; returning `false` records an
/// unsupported gate in [`Report`] rather than turning valid input into an
/// invalid-document error. Every top-level and conditional SPRM consumes the
/// shared formatting budget. Conditional groups are inspected once and nested
/// sprmTCnf is rejected rather than recursively expanded. Each record is
/// structurally validated before its callback; a caller that mutates state
/// MUST discard that state if a later record or callback returns an error.
pub(super) fn validate(
    style_id: usize,
    bytes: &[u8],
    budget: &mut Budget,
    mut visit: impl FnMut(Scope, u16, &[u8], &mut Budget) -> Result<bool, String>,
) -> Result<Report, String> {
    let mut report = Report::default();
    let mut has_default_width_before = false;
    let mut sprms = Sprms::new(bytes);
    while let Some((code, operand)) = sprms.next(budget)? {
        require_table_sprm(code)?;
        match code {
            T_ISTD => {
                report.ignored_tistd += 1;
            }
            T_WIDTH_BEFORE => {
                validate_width_before(style_id, operand)?;
                has_default_width_before = true;
                record(
                    &mut report,
                    Scope::Unconditional,
                    code,
                    operand,
                    budget,
                    &mut visit,
                )?;
            }
            T_CNF => {
                validate_conditional(style_id, operand, budget, &mut report, &mut visit)?;
            }
            _ => {
                validate_property(Scope::Unconditional, style_id, code, operand)?;
                record(
                    &mut report,
                    Scope::Unconditional,
                    code,
                    operand,
                    budget,
                    &mut visit,
                )?;
            }
        }
    }
    if style_id == DEFAULT_TABLE_STYLE && !has_default_width_before {
        return Err(unsupported(
            "default Word table-style TAPX lacks required zero dxa width-before",
        ));
    }
    Ok(report)
}

fn validate_conditional(
    style_id: usize,
    operand: &[u8],
    budget: &mut Budget,
    report: &mut Report,
    visit: &mut impl FnMut(Scope, u16, &[u8], &mut Budget) -> Result<bool, String>,
) -> Result<(), String> {
    let operand = cnf::parse(operand)?;
    let scope = Scope::Conditional(operand.condition);
    let mut nested = Sprms::new(operand.grpprl);
    while let Some((code, value)) = nested.next(budget)? {
        require_table_sprm(code)?;
        if code == T_ISTD {
            // Treat "any sprmTIstd ... contained in the array" in 2.9.340 as
            // covering the bounded grpprl carried by a top-level sprmTCnf too.
            // Ignoring this selector is narrower than assigning it conditional
            // table-style semantics that the specification never defines.
            report.ignored_tistd += 1;
            continue;
        }
        if code == T_CNF {
            return Err(unsupported(
                "nested sprmTCnf outside a Word table-style TAPX array",
            ));
        }
        validate_property(scope, style_id, code, value)?;
        record(report, scope, code, value, budget, visit)?;
    }
    Ok(())
}

fn validate_property(
    scope: Scope,
    style_id: usize,
    code: u16,
    operand: &[u8],
) -> Result<(), String> {
    if code == T_WIDTH_BEFORE {
        // The 0x000B exception is stated for its UpxTapx as a whole. An exact
        // value inside sprmTCnf is structurally valid but does not satisfy the
        // separate requirement for an unconditional default-style value.
        return validate_width_before(style_id, operand);
    }
    if PRESERVED_BY_T_ISTD.contains(&code) {
        return Err(prohibited(code));
    }
    if EXPLICITLY_PROHIBITED.contains(&code)
        && !(matches!(scope, Scope::Conditional(_))
            && CONDITIONAL_BORDER_EXCEPTIONS.contains(&code))
    {
        return Err(prohibited(code));
    }
    if matches!(scope, Scope::Unconditional)
        && (CONDITIONAL_BORDER_EXCEPTIONS.contains(&code) || OTHER_CNF_ONLY.contains(&code))
    {
        return Err(prohibited(code));
    }
    // sprmTCellNoWrapStyle is explicitly UpxTapx-only, not a property within
    // the grpprl owned by a CNFOperand.
    if matches!(scope, Scope::Conditional(_)) && code == 0x347d {
        return Err(prohibited(code));
    }
    Ok(())
}

fn validate_width_before(style_id: usize, operand: &[u8]) -> Result<(), String> {
    if style_id != DEFAULT_TABLE_STYLE || operand != [0x03, 0x00, 0x00] {
        return Err(unsupported(
            "invalid sprmTWidthBefore in Word table-style TAPX",
        ));
    }
    Ok(())
}

fn require_table_sprm(code: u16) -> Result<(), String> {
    if (code >> 10) & 7 != 5 {
        return Err(unsupported(format!(
            "non-table SPRM 0x{code:04X} in Word table-style TAPX"
        )));
    }
    Ok(())
}

fn prohibited(code: u16) -> String {
    unsupported(format!(
        "prohibited SPRM 0x{code:04X} in Word table-style TAPX"
    ))
}

fn record(
    report: &mut Report,
    scope: Scope,
    code: u16,
    operand: &[u8],
    budget: &mut Budget,
    visit: &mut impl FnMut(Scope, u16, &[u8], &mut Budget) -> Result<bool, String>,
) -> Result<(), String> {
    match scope {
        Scope::Unconditional => report.unconditional_properties += 1,
        Scope::Conditional(_) => report.conditional_properties += 1,
    }
    report.unsupported |= !visit(scope, code, operand, budget)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn prl(code: u16, operand: &[u8]) -> Vec<u8> {
        let mut bytes = code.to_le_bytes().to_vec();
        bytes.extend_from_slice(operand);
        bytes
    }

    fn framed(code: u16) -> Vec<u8> {
        let operand = match code >> 13 {
            0 | 1 => vec![0],
            2 | 4 | 5 => vec![0; 2],
            3 => vec![0; 4],
            7 => vec![0; 3],
            _ if code == 0xd608 => vec![1, 0],
            _ => vec![0],
        };
        prl(code, &operand)
    }

    fn cnf(nested: &[u8]) -> Vec<u8> {
        let mut operand = vec![(2 + nested.len()) as u8, 1, 0];
        operand.extend_from_slice(nested);
        prl(T_CNF, &operand)
    }

    fn validate_all(style_id: usize, bytes: &[u8]) -> Result<Report, String> {
        validate(style_id, bytes, &mut Budget::default(), |_, _, _, _| {
            Ok(true)
        })
    }

    #[test]
    fn ignores_tistd_and_reports_valid_unknown_properties_without_rejecting_them() {
        let bytes = [prl(T_ISTD, &[7, 0]), prl(0x3488, &[1]), prl(0x3489, &[2])].concat();
        let report = validate(7, &bytes, &mut Budget::default(), |_, code, operand, _| {
            Ok(code == 0x3488 && operand == [1])
        })
        .unwrap();
        assert_eq!(report.ignored_tistd, 1);
        assert_eq!(report.unconditional_properties, 2);
        assert!(report.unsupported);

        let report = validate(
            7,
            &cnf(&framed(0xd687)),
            &mut Budget::default(),
            |_, _, _, _| Ok(false),
        )
        .unwrap();
        assert_eq!(report.conditional_properties, 1);
        assert!(report.unsupported);

        let report = validate_all(7, &cnf(&prl(T_ISTD, &[9, 0]))).unwrap();
        assert_eq!(report.ignored_tistd, 1);
        assert_eq!(report.conditional_properties, 0);
    }

    #[test]
    fn callback_errors_propagate_and_invalid_properties_are_never_visited() {
        let mut visited = Vec::new();
        let error = validate(
            7,
            &prl(0xd608, &[1, 0]),
            &mut Budget::default(),
            |_, code, _, _| {
                visited.push(code);
                Ok(true)
            },
        )
        .unwrap_err();
        assert!(error.contains("prohibited SPRM"));
        assert!(visited.is_empty());

        let error = validate(
            7,
            &prl(0x3488, &[1]),
            &mut Budget::default(),
            |_, code, _, _| {
                visited.push(code);
                Err("projection failed".to_string())
            },
        )
        .unwrap_err();
        assert_eq!(error, "projection failed");
        assert_eq!(visited, [0x3488]);
    }

    #[test]
    fn rejects_non_table_preserved_and_explicitly_prohibited_sprms() {
        assert!(validate_all(7, &prl(0x0835, &[1]))
            .unwrap_err()
            .contains("non-table SPRM"));
        for code in PRESERVED_BY_T_ISTD {
            assert!(
                validate_all(7, &framed(code))
                    .unwrap_err()
                    .contains("prohibited SPRM"),
                "0x{code:04X}"
            );
        }
        for code in EXPLICITLY_PROHIBITED {
            assert!(
                validate_all(7, &framed(code))
                    .unwrap_err()
                    .contains("prohibited SPRM"),
                "0x{code:04X}"
            );
        }
    }

    #[test]
    fn default_width_before_is_exact_mandatory_and_forbidden_elsewhere() {
        assert!(validate_all(DEFAULT_TABLE_STYLE, &[])
            .unwrap_err()
            .contains("lacks required"));
        for operand in [[0, 0, 0], [3, 1, 0], [3, 0, 1]] {
            assert!(validate_all(DEFAULT_TABLE_STYLE, &prl(T_WIDTH_BEFORE, &operand)).is_err());
        }
        let report = validate_all(DEFAULT_TABLE_STYLE, &prl(T_WIDTH_BEFORE, &[3, 0, 0])).unwrap();
        assert_eq!(report.unconditional_properties, 1);
        assert!(validate_all(
            DEFAULT_TABLE_STYLE,
            &[
                cnf(&prl(T_WIDTH_BEFORE, &[3, 0, 0])),
                prl(T_WIDTH_BEFORE, &[3, 0, 0])
            ]
            .concat(),
        )
        .is_ok());
        assert!(
            validate_all(DEFAULT_TABLE_STYLE, &cnf(&prl(T_WIDTH_BEFORE, &[3, 0, 0])))
                .unwrap_err()
                .contains("lacks required")
        );
        assert!(validate_all(7, &prl(T_WIDTH_BEFORE, &[3, 0, 0])).is_err());
    }

    #[test]
    fn conditional_group_uses_shared_framing_and_six_border_exceptions() {
        let nested = CONDITIONAL_BORDER_EXCEPTIONS
            .into_iter()
            .flat_map(framed)
            .collect::<Vec<_>>();
        let report = validate_all(7, &cnf(&nested)).unwrap();
        assert_eq!(report.conditional_properties, 6);

        let mut malformed = cnf(&nested);
        malformed[2] -= 1;
        assert!(validate_all(7, &malformed).is_err());
        assert!(validate_all(7, &cnf(&framed(0xd608))).is_err());
        assert!(validate_all(7, &cnf(&prl(0x0835, &[1]))).is_err());
    }

    #[test]
    fn conditional_validation_is_single_level_and_context_specific() {
        assert!(validate_all(7, &cnf(&cnf(&[])))
            .unwrap_err()
            .contains("nested sprmTCnf"));
        assert!(validate_all(7, &framed(0xd685)).is_err());
        assert!(validate_all(7, &framed(0xd686)).is_err());
        assert!(validate_all(7, &cnf(&framed(0xd685))).is_ok());
        assert!(validate_all(7, &cnf(&framed(0xd686))).is_ok());
        assert!(validate_all(7, &cnf(&framed(0x347d))).is_err());

        let report = validate_all(7, &cnf(&prl(T_ISTD, &[9, 0]))).unwrap();
        assert_eq!(report.ignored_tistd, 1);
        assert_eq!(report.conditional_properties, 0);
    }

    #[test]
    fn nested_sprms_share_the_outer_formatting_budget() {
        let nested = [framed(0x3488), framed(0x3489)].concat();
        let bytes = cnf(&nested);
        let mut budget = Budget::default();
        let before = budget.remaining();
        validate(7, &bytes, &mut budget, |_, _, _, _| Ok(true)).unwrap();
        assert_eq!(before - budget.remaining(), 3);
    }
}
