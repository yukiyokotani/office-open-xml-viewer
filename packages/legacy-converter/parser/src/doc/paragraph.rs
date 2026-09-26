//! Paragraph layout from [MS-DOC] 2.6.2, LSPD and XAS to ECMA-376
//! 17.3.1 / CT_PPrBase. Values retain their units; no font/pixel fitting.

use super::{border::Border, u16_at, unsupported};
use std::collections::BTreeMap;

mod direct;
mod frame;
mod shading;
pub(super) use frame::FrameGap;
pub(in crate::doc) use frame::TableParagraphFrame;
pub(super) use shading::{fill as shading_fill, ShadingFill};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct AlignmentPatch {
    code: u16,
    value: u8,
}

impl AlignmentPatch {
    pub(super) fn from_sprm(code: u16, operand: &[u8]) -> Result<Option<Self>, String> {
        if !matches!(code, 0x2403 | 0x2461) {
            return Ok(None);
        }
        let value = *operand
            .first()
            .ok_or_else(|| unsupported("truncated Word paragraph alignment"))?;
        let maximum = if code == 0x2403 { 5 } else { 9 };
        if value > maximum {
            return Err(unsupported("invalid Word paragraph alignment"));
        }
        Ok(Some(Self { code, value }))
    }

    pub(super) fn apply(self, properties: &mut Properties) {
        properties.alignment = (self.value, self.code == 0x2403);
    }
}

#[derive(Clone)]
pub struct Properties {
    pub ilfo: i16,
    pub ilvl: u8,
    protected_list_indent: Option<(i32, i32)>,
    tabs: super::tabs::Stops,
    flags: BTreeMap<&'static str, bool>,
    line: (u16, &'static str),
    /// sprmPDyaLine was applied by a paragraph style or the direct PAPX.
    /// Without it the MS-DOC 2.6.2 default single spacing is unauthored.
    line_authored: bool,
    before: u16,
    after: u16,
    before_lines: Option<i16>,
    after_lines: Option<i16>,
    before_auto: bool,
    after_auto: bool,
    indents: [Option<(i32, u32)>; 4],
    indent_order: u32,
    first: i32,
    nest: Option<(i32, bool)>,
    chars: [Option<i16>; 3],
    alignment: (u8, bool),
    text_alignment: Option<&'static str>,
    borders: [Option<Border>; 5],
    /// MS-DOC 2.6.2 sprmPFContextualSpacing.
    contextual_spacing: bool,
    /// MS-DOC 2.6.2 sprmPOutLvl raw value (0..=9); see `outline_level`.
    outline_level: Option<u8>,
    /// MS-DOC 2.6.2 sprmPShd/sprmPShd80 projected to a fill.
    shading: Option<ShadingFill>,
    frame: frame::Frame,
}

impl Default for Properties {
    fn default() -> Self {
        Self {
            ilfo: 0,
            ilvl: 0,
            protected_list_indent: None,
            tabs: super::tabs::Stops::default(),
            flags: BTreeMap::from([
                ("widowControl", true),
                ("kinsoku", true),
                ("wordWrap", true),
                ("overflowPunct", true),
                ("autoSpaceDE", true),
                ("autoSpaceDN", true),
                ("snapToGrid", true),
                ("adjustRightInd", true),
            ]),
            line: (240, "auto"),
            line_authored: false,
            before: 0,
            after: 0,
            before_lines: None,
            after_lines: None,
            before_auto: false,
            after_auto: false,
            indents: [None; 4],
            indent_order: 0,
            first: 0,
            nest: None,
            chars: [None; 3],
            alignment: (0, false),
            text_alignment: None,
            borders: std::array::from_fn(|_| None),
            contextual_spacing: false,
            outline_level: None,
            shading: None,
            frame: frame::Frame::default(),
        }
    }
}

fn signed(bytes: &[u8]) -> Result<i32, String> {
    let value = u16_at(bytes, 0)? as i16 as i32;
    if !(-31680..=31680).contains(&value) {
        return Err(unsupported("invalid Word paragraph distance"));
    }
    Ok(value)
}

impl Properties {
    pub fn set_bidi(&mut self, value: bool) {
        self.flags.insert("bidi", value);
    }

    pub(super) fn is_bidi(&self) -> bool {
        self.flags.get("bidi") == Some(&true)
    }

    pub(super) fn clear_alignment(&mut self) {
        self.alignment = (0, false);
    }

    pub fn apply(&mut self, code: u16, operand: &[u8]) -> Result<bool, String> {
        let flag = match code {
            0x2405 => Some("keepLines"),
            0x2406 => Some("keepNext"),
            0x2407 => Some("pageBreakBefore"),
            0x240c => Some("suppressLineNumbers"),
            0x242a => Some("suppressAutoHyphens"),
            0x2431 => Some("widowControl"),
            0x2433 => Some("kinsoku"),
            0x2434 => Some("wordWrap"),
            0x2435 => Some("overflowPunct"),
            0x2436 => Some("topLinePunct"),
            0x2437 => Some("autoSpaceDE"),
            0x2438 => Some("autoSpaceDN"),
            0x2441 => Some("bidi"),
            0x2447 => Some("snapToGrid"),
            0x2448 => Some("adjustRightInd"),
            _ => None,
        };
        if let Some(flag) = flag {
            self.flags.insert(flag, bool8(operand[0])?);
            return Ok(true);
        }
        match code {
            0x6467 => {
                // MS-DOC 2.6.2 sprmPRsid is nonvisual revision-session
                // provenance, not a paragraph revision mark. Viewer policy:
                // validate its fixed-width ID and intentionally omit it.
                if operand.len() != 4 {
                    return Err(unsupported("invalid Word paragraph revision session ID"));
                }
                let _ = super::u32_at(operand, 0)?;
            }
            0xc66c => {
                // MS-DOC 2.6.2 sprmPTIstdInfo / 2.9.221
                // PTIstdInfoOperand: cb MUST be 16 and all 16 reserved bytes
                // MUST be ignored. Validate the complete framed operand before
                // treating this property as the required no-op.
                if operand.len() != 17 || operand[0] != 16 {
                    return Err(unsupported("invalid Word PTIstdInfo operand"));
                }
            }
            0x260a => self.ilvl = operand[0],
            0x460b => self.ilfo = u16_at(operand, 0)? as i16,
            0x6424..=0x6428 => {
                let side = usize::from(code - 0x6424);
                self.borders[side] = Some(Border::paragraph(operand, true, side)?);
            }
            0xc64e..=0xc652 => {
                // MS-DOC 2.9.21: BrcOperand.cb MUST be 8, excluding cb.
                if operand.len() != 9 || operand[0] != 8 {
                    return Err(unsupported("invalid Word paragraph border operand"));
                }
                let side = usize::from(code - 0xc64e);
                self.borders[side] = Some(Border::paragraph(&operand[1..], false, side)?);
            }
            0xc60d | 0xc615 => self.tabs.apply(operand, code == 0xc615)?,
            0x6412 => {
                let line = signed(operand)?;
                let multiple = u16_at(operand, 2)?;
                if multiple > 1 {
                    return Err(unsupported("invalid Word line spacing multiplier"));
                }
                self.line = if line < 0 {
                    ((-line) as u16, "exact")
                } else {
                    (line as u16, if multiple == 1 { "auto" } else { "atLeast" })
                };
                self.line_authored = true;
            }
            0xa413 | 0xa414 => {
                let value = u16_at(operand, 0)?;
                if value > 31680 {
                    return Err(unsupported("invalid Word paragraph spacing"));
                }
                if code == 0xa413 {
                    self.before = value;
                } else {
                    self.after = value;
                }
            }
            0x4458 | 0x4459 => {
                let value = u16_at(operand, 0)? as i16;
                if !(-20..=31680).contains(&value) {
                    return Err(unsupported("invalid Word line-unit paragraph spacing"));
                }
                if code == 0x4458 {
                    self.before_lines = Some(value);
                } else {
                    self.after_lines = Some(value);
                }
            }
            0x245b => self.before_auto = bool8(operand[0])?,
            0x245c => self.after_auto = bool8(operand[0])?,
            0x840e | 0x845d | 0x840f | 0x845e => {
                let index = match code {
                    0x840f => 0,
                    0x840e => 1,
                    0x845e => 2,
                    _ => 3,
                };
                self.indent_order += 1;
                self.indents[index] = Some((signed(operand)?, self.indent_order));
            }
            0x8411 | 0x8460 => self.first = signed(operand)?,
            0x4610 | 0x465f => {
                // PNest supersedes PNest80, irrespective of their order.
                if code == 0x465f || !self.nest.is_some_and(|(_, physical)| !physical) {
                    self.nest = Some((signed(operand)?, code == 0x4610));
                }
            }
            0x4455..=0x4457 => {
                let index = match code {
                    0x4455 => 1,
                    0x4456 => 0,
                    _ => 2,
                };
                self.chars[index] = Some(u16_at(operand, 0)? as i16);
            }
            0x2403 | 0x2461 => {
                AlignmentPatch::from_sprm(code, operand)?
                    .expect("alignment code")
                    .apply(self);
            }
            0x246d => {
                // MS-DOC 2.6.2 sprmPFContextualSpacing (Bool8): suppress
                // before/after spacing next to a paragraph of the same style,
                // exactly ECMA-376 17.3.1.9 contextualSpacing.
                self.contextual_spacing = bool8(operand[0])?;
            }
            0x2640 => {
                // MS-DOC 2.6.2 sprmPOutLvl: 0..=8 outline level, 9 body text.
                let level = operand[0];
                if level > 9 {
                    return Err(unsupported("invalid Word outline level"));
                }
                self.outline_level = Some(level);
            }
            0x6629 | 0xc653 => {
                // MS-DOC 2.6.2: sprmPBrcBar80 (Brc80) and sprmPBrcBar
                // (BrcOperand) are specified as "a value that has no effect".
                // Validate the framing and decode the border so malformed
                // input still fails, then apply nothing.
                if code == 0xc653 {
                    if operand.len() != 9 || operand[0] != 8 {
                        return Err(unsupported("invalid Word paragraph bar border"));
                    }
                    Border::read(&operand[1..], false)?;
                } else {
                    Border::read(operand, true)?;
                }
            }
            0xc64d | 0x442d => match shading::fill(operand, code == 0xc64d)? {
                Some(fill) => self.shading = Some(fill),
                None => return Ok(false),
            },
            0x8418 | 0x8419 | 0x841a | 0x442b | 0x261b | 0x2423 | 0x842f | 0x842e | 0x2430
            | 0x2462 | 0x442c | 0x443a => {
                self.frame.apply(code, operand)?;
            }
            0x4439 => {
                self.text_alignment = Some(match u16_at(operand, 0)? {
                    0 => "top",
                    1 => "center",
                    2 => "baseline",
                    3 => "bottom",
                    4 => "auto",
                    _ => return Err(unsupported("invalid Word paragraph text alignment")),
                });
            }
            _ => return Ok(false),
        }
        Ok(true)
    }

    fn normalized_alignment(&self) -> &'static str {
        let bidi = self.flags.get("bidi") == Some(&true);
        let (value, physical) = self.alignment;
        if physical {
            match value {
                0 if bidi => "right",
                2 if bidi => "left",
                0 => "left",
                1 => "center",
                2 => "right",
                3 => "both",
                4 => "mediumKashida",
                _ => "highKashida",
            }
        } else {
            [
                "left",
                "center",
                "right",
                "both",
                "distribute",
                "mediumKashida",
                "numTab",
                "highKashida",
                "lowKashida",
                "thaiDistribute",
            ][value as usize]
        }
    }

    fn normalized_model_alignment(&self) -> &'static str {
        match self.normalized_alignment() {
            // Match the DOCX parser's stable renderer-facing normalization.
            "both" => "justify",
            "numTab" => "left",
            value => value,
        }
    }

    /// MS-DOC sprmPIlfo: negative references retain the paragraph's logical
    /// left and first-line indent despite subsequent list/style formatting.
    pub fn preserve_list_indent(&mut self, source: &Self) {
        self.protected_list_indent = Some((source.logical_indents()[0], source.first));
    }

    fn logical_indents(&self) -> [i32; 2] {
        let bidi = self.flags.get("bidi") == Some(&true);
        // Resolve physical/logical coordinates only after bidi is known.
        let mut indents = [0, 0];
        let mut writes: Vec<_> = self
            .indents
            .iter()
            .enumerate()
            .filter_map(|(i, value)| value.map(|(v, order)| (order, i, v)))
            .collect();
        writes.sort_unstable_by_key(|(order, _, _)| *order);
        for (_, i, value) in writes {
            indents[(i % 2) ^ usize::from(i < 2 && bidi)] = value;
        }
        if let Some((value, physical)) = self.nest {
            indents[usize::from(physical && bidi)] += value;
        }
        if let Some((left, _)) = self.protected_list_indent {
            indents[0] = left;
        }
        indents
    }
}

fn bool8(value: u8) -> Result<bool, String> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(unsupported("invalid Word paragraph boolean")),
    }
}

pub fn prm0(prm: u16) -> Option<[u8; 3]> {
    let code: u16 = match (prm >> 1) & 127 {
        0x05 => 0x2461,
        0x07 => 0x2405,
        0x08 => 0x2406,
        0x09 => 0x2407,
        0x0c => 0x260a,
        0x0e => 0x240c,
        0x2c => 0x242a,
        0x33 => 0x2431,
        0x35 => 0x2433,
        0x36 => 0x2434,
        0x37 => 0x2435,
        0x38 => 0x2436,
        0x39 => 0x2437,
        0x3a => 0x2438,
        _ => return None,
    };
    let [a, b] = code.to_le_bytes();
    Some([a, b, (prm >> 8) as u8])
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::sprm::{Budget, Sprms};

    fn projected(properties: &Properties) -> serde_json::Value {
        serde_json::to_value(properties.direct_paragraph()).unwrap()
    }

    /// (style, color, width, space) of each physical side's border edge.
    fn border(value: &serde_json::Value, side: &str) -> serde_json::Value {
        value["borders"][side].clone()
    }

    fn raw_border(value: &serde_json::Value, side: &str) -> serde_json::Value {
        value["__paragraphTypographyAcquisition"]["borders"][side]["val"]["raw"].clone()
    }

    #[test]
    fn ptistdinfo_validates_and_ignores_exact_operand_without_consuming_neighbors() {
        for fill in [0x00, 0x55, 0xff] {
            let baseline = projected(&Properties::default());
            let mut bytes = vec![0x6c, 0xc6, 16];
            bytes.extend([fill; 16]);
            bytes.extend([0x07, 0x24, 1]);
            let mut sprms = Sprms::new(&bytes);
            let mut budget = Budget::default();
            let mut properties = Properties::default();
            let (code, operand) = sprms.next(&mut budget).unwrap().unwrap();
            assert_eq!((code, operand.len()), (0xc66c, 17));
            assert!(properties.apply(code, operand).unwrap());
            assert_eq!(projected(&properties), baseline);
            let (code, operand) = sprms.next(&mut budget).unwrap().unwrap();
            assert_eq!(code, 0x2407);
            assert!(properties.apply(code, operand).unwrap());
            assert!(properties.direct_paragraph().page_break_before);
            assert!(sprms.next(&mut budget).unwrap().is_none());
        }

        for cb in [0, 15, 17, 255] {
            let mut bytes = vec![0x6c, 0xc6, cb];
            bytes.resize(3 + usize::from(cb), 0);
            bytes.extend([0x07, 0x24, 1]);
            let mut sprms = Sprms::new(&bytes);
            let (code, operand) = sprms.next(&mut Budget::default()).unwrap().unwrap();
            assert_eq!(code, 0xc66c);
            assert!(Properties::default().apply(code, operand).is_err());
        }
        let truncated = [0x6c, 0xc6, 16, 0, 0];
        assert!(Sprms::new(&truncated).next(&mut Budget::default()).is_err());
    }

    #[test]
    fn paragraph_revision_session_id_is_validated_and_nonvisual() {
        let baseline = projected(&Properties::default());
        for value in [0, 0x7856_3412, u32::MAX] {
            let mut properties = Properties::default();
            assert!(properties.apply(0x6467, &value.to_le_bytes()).unwrap());
            assert_eq!(projected(&properties), baseline);
        }
        assert!(Properties::default().apply(0x6467, &[0; 3]).is_err());
        assert!(Properties::default().apply(0x6467, &[0; 5]).is_err());

        let mut bytes = vec![0x67, 0x64];
        bytes.extend(0x7856_3412u32.to_le_bytes());
        bytes.extend([0x07, 0x24, 1]);
        let mut sprms = Sprms::new(&bytes);
        let mut budget = Budget::default();
        let mut properties = Properties::default();
        while let Some((code, operand)) = sprms.next(&mut budget).unwrap() {
            assert!(properties.apply(code, operand).unwrap());
        }
        assert!(properties.direct_paragraph().page_break_before);

        assert!(Sprms::new(&[0x67, 0x64, 1, 2, 3])
            .next(&mut Budget::default())
            .is_err());
    }

    #[test]
    fn negative_list_reference_protects_logical_indents_in_rtl_without_freezing_the_other_side() {
        let mut original = Properties::default();
        original.apply(0x2441, &[1]).unwrap();
        original.apply(0x840e, &1000_i16.to_le_bytes()).unwrap(); // physical right = logical left
        original.apply(0x8460, &(-200_i16).to_le_bytes()).unwrap();
        let mut list = original.clone();
        list.apply(0x845e, &720_i16.to_le_bytes()).unwrap();
        list.apply(0x845d, &600_i16.to_le_bytes()).unwrap();
        list.apply(0x8460, &(-360_i16).to_le_bytes()).unwrap();
        list.preserve_list_indent(&original);
        let indents = |properties: &Properties| {
            let paragraph = properties.direct_paragraph();
            (
                paragraph.indent_left,
                paragraph.indent_right,
                paragraph.indent_first,
            )
        };
        assert_eq!(indents(&list), (50.0, 30.0, -10.0));
        assert_eq!(indents(&original), (50.0, 0.0, -10.0));
    }
    #[test]
    fn paragraph_borders_preserve_each_side() {
        let mut p = Properties::default();
        assert!(p.direct_paragraph().borders.is_none());
        for code in 0x6424..=0x6428 {
            assert!(p.apply(code, &[8, 1, 2, 3]).unwrap());
        }
        let value = projected(&p);
        for side in ["top", "left", "bottom", "right", "between"] {
            assert_eq!(
                border(&value, side),
                serde_json::json!({"style": "single", "color": "0000ff", "width": 1.0, "space": 3.0}),
                "{side}"
            );
        }
    }

    #[test]
    fn modern_border_replaces_only_its_side_and_explicit_none_clears_it() {
        let mut p = Properties::default();
        p.apply(0x6424, &[8, 1, 2, 0]).unwrap();
        p.apply(0x6426, &[8, 1, 2, 0]).unwrap();
        assert!(p
            .apply(0xc650, &[8, 0x12, 0x34, 0x56, 0, 16, 3, 7, 0])
            .unwrap());
        let value = projected(&p);
        assert_eq!(border(&value, "top")["style"], "single");
        assert_eq!(
            border(&value, "bottom"),
            serde_json::json!({"style": "double", "color": "123456", "width": 2.0, "space": 7.0})
        );
        p.apply(0xc650, &[8, 0, 0, 0, 0xff, 0, 0, 0, 0]).unwrap();
        let value = projected(&p);
        assert_eq!(raw_border(&value, "bottom"), "none");
        assert_eq!(border(&value, "bottom")["style"], "none");
        assert_eq!(border(&value, "top")["style"], "single");
        p.apply(0xc650, &[8, 0, 1, 2, 3, 255, 255, 255, 255])
            .unwrap();
        let value = projected(&p);
        assert_eq!(raw_border(&value, "bottom"), "nil");
        assert_eq!(border(&value, "top")["style"], "single");
        for bad in [&[0][..], &[7, 0, 0, 0, 0, 0, 0, 0], &[8, 0, 0]] {
            assert!(p.apply(0xc650, bad).is_err());
        }
    }

    #[test]
    fn paragraph_borders_resolve_logical_sides_after_the_final_bidi_setting() {
        let mut p = Properties::default();
        p.apply(0x6425, &[8, 1, 2, 0]).unwrap();
        p.apply(0xc651, &[8, 255, 0, 0, 0, 8, 3, 0, 0]).unwrap();
        p.apply(0x2441, &[1]).unwrap();
        let value = projected(&p);
        assert_eq!(border(&value, "left")["style"], "double");
        assert_eq!(border(&value, "right")["style"], "single");
        p.apply(0x2441, &[0]).unwrap();
        let value = projected(&p);
        assert_eq!(border(&value, "left")["style"], "single");
        assert_eq!(border(&value, "right")["style"], "double");
    }

    #[test]
    fn negative_line_spacing_is_exact_even_when_multiplier_flag_is_set() {
        let mut p = Properties::default();
        let spacing = |p: &Properties| {
            let spacing = p.direct_paragraph().line_spacing.unwrap();
            (spacing.value, spacing.rule)
        };
        p.apply(0x6412, &[0xd4, 0xfe, 1, 0]).unwrap(); // -300 twips
        assert_eq!(spacing(&p), (15.0, "exact".to_owned()));
        p.apply(0x6412, &[0x68, 1, 1, 0]).unwrap();
        assert_eq!(spacing(&p), (1.5, "auto".to_owned()));
        p.apply(0x6412, &[0x68, 1, 0, 0]).unwrap();
        assert_eq!(spacing(&p), (18.0, "atLeast".to_owned()));
        assert!(p.apply(0x6412, &[0, 0, 2, 0]).is_err());
    }

    #[test]
    fn preserves_signed_hanging_indents_and_explicit_spacing_resets() {
        let mut p = Properties::default();
        p.apply(0x845e, &720u16.to_le_bytes()).unwrap();
        p.apply(0x8460, &(-360i16).to_le_bytes()).unwrap();
        p.apply(0xa413, &240u16.to_le_bytes()).unwrap();
        p.apply(0xa413, &[0, 0]).unwrap();
        let paragraph = p.direct_paragraph();
        assert_eq!(
            (
                paragraph.indent_left,
                paragraph.indent_right,
                paragraph.indent_first
            ),
            (36.0, 0.0, -18.0)
        );
        assert_eq!(paragraph.space_before, 0.0);
    }

    #[test]
    fn late_bidi_property_does_not_turn_physical_indents_into_logical_ones() {
        let mut p = Properties::default();
        p.apply(0x840f, &720u16.to_le_bytes()).unwrap();
        p.apply(0x840e, &360u16.to_le_bytes()).unwrap();
        p.apply(0x2441, &[1]).unwrap();
        let sides = |p: &Properties| {
            let paragraph = p.direct_paragraph();
            (paragraph.indent_left, paragraph.indent_right)
        };
        assert_eq!(sides(&p), (18.0, 36.0));
        p.apply(0x845d, &240u16.to_le_bytes()).unwrap();
        assert_eq!(sides(&p), (18.0, 12.0));
        let mut single = Properties::default();
        single.apply(0x840f, &720u16.to_le_bytes()).unwrap();
        single.apply(0x2441, &[1]).unwrap();
        assert_eq!(sides(&single), (0.0, 36.0));
    }
}
