//! MS-DOC 2.9.179-183, 2.9.300-301, 2.9.310 -> ECMA-376 17.3.1.37-38.
use super::{u16_at, unsupported};
use std::collections::BTreeMap;

// Resource policy for the resolved set across arbitrarily many style/PRM edits.
// Each individual binary addition/deletion array has a normative 64-entry cap.
const MAX_RESOLVED_TABS: usize = 256;
#[derive(Clone, Default)]
pub struct Stops(BTreeMap<i16, (&'static str, &'static str)>);

fn byte(b: &[u8], p: usize) -> Result<u8, String> {
    b.get(p)
        .copied()
        .ok_or_else(|| unsupported("truncated Word custom tabs"))
}

fn count(b: &[u8], p: usize) -> Result<usize, String> {
    let value = byte(b, p)? as usize;
    if value > 64 {
        return Err(unsupported("invalid Word custom tab count"));
    }
    Ok(value)
}
fn positions(b: &[u8], p: usize, n: usize, close: bool) -> Result<Vec<i16>, String> {
    let mut result = Vec::with_capacity(n);
    for i in 0..n {
        let value = u16_at(b, p + i * 2)? as i16;
        // DelClose permits the complete negative i16 range (SHOULD, not MUST).
        if value > 31680 || (!close && value < -31680) {
            return Err(unsupported("invalid Word custom tab position"));
        }
        if result.last().is_some_and(|last| *last > value) {
            return Err(unsupported("unsorted Word custom tab positions"));
        }
        result.push(value);
    }
    Ok(result)
}
impl Stops {
    pub fn apply(&mut self, b: &[u8], close: bool) -> Result<(), String> {
        let cb = byte(b, 0)? as usize;
        if cb < 2 || (!(close && cb == 255) && cb + 1 != b.len()) {
            return Err(unsupported("invalid Word custom tab operand length"));
        }
        let deleted = count(b, 1)?;
        let add_at = 2 + deleted * if close { 4 } else { 2 };
        let added = count(b, add_at)?;
        if add_at + 1 + added * 3 != b.len() {
            return Err(unsupported("unexpected Word custom tab data"));
        }
        let removed = positions(b, 2, deleted, close)?;
        let added_positions = positions(b, add_at + 1, added, false)?;
        let mut radii = Vec::with_capacity(deleted);
        for i in 0..deleted {
            let radius = if close {
                let raw = u16_at(b, 2 + deleted * 2 + i * 2)? as i16 as i32;
                if !(-31679..=31681).contains(&raw) {
                    return Err(unsupported("invalid Word custom tab deletion distance"));
                }
                // XAS_plusOne is stored one larger than the actual twips.
                (raw - 1).max(25)
            } else {
                25
            };
            radii.push(radius);
        }
        let mut descriptors = Vec::with_capacity(added);
        for i in 0..added {
            let value = byte(b, add_at + 1 + added * 2 + i)?;
            let alignment = match value & 7 {
                0 => "left",
                1 => "center",
                2 => "right",
                3 => "decimal",
                4 => "bar",
                6 => "num",
                _ => return Err(unsupported("invalid Word custom tab alignment")),
            };
            // TBD explicitly ignores leaders for bar tabs, including reserved bits.
            let leader = if alignment == "bar" {
                "none"
            } else {
                match (value >> 3) & 7 {
                    0 | 7 => "none",
                    1 => "dot",
                    2 => "hyphen",
                    3 | 4 => "underscore",
                    5 => "middleDot",
                    _ => return Err(unsupported("invalid Word custom tab leader")),
                }
            };
            descriptors.push((alignment, leader));
        }
        for (&position, radius) in removed.iter().zip(radii) {
            let lo = (i32::from(position) - radius).max(i16::MIN as i32) as i16;
            let hi = (i32::from(position) + radius).min(i16::MAX as i32) as i16;
            // Each retained entry is removed at most once; avoid scanning the
            // full set for every deletion range.
            let keys: Vec<_> = self.0.range(lo..=hi).map(|(&p, _)| p).collect();
            for key in keys {
                self.0.remove(&key);
            }
        }
        for (position, descriptor) in added_positions.into_iter().zip(descriptors) {
            if self.0.len() >= MAX_RESOLVED_TABS && !self.0.contains_key(&position) {
                return Err(unsupported("Word resolved custom tab limit exceeded"));
            }
            self.0.insert(position, descriptor);
        }
        Ok(())
    }
    pub fn xml(&self) -> String {
        if self.0.is_empty() {
            return String::new();
        }
        let mut xml = String::from("<w:tabs>");
        for (position, (alignment, leader)) in &self.0 {
            xml.push_str(&format!(
                "<w:tab w:val=\"{alignment}\" w:pos=\"{position}\" w:leader=\"{leader}\"/>"
            ));
        }
        xml.push_str("</w:tabs>");
        xml
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn operand(del: &[(i16, i16)], add: &[(i16, u8)], close: bool) -> Vec<u8> {
        let mut b = vec![0, del.len() as u8];
        for (p, _) in del {
            b.extend(p.to_le_bytes());
        }
        if close {
            for (_, r) in del {
                b.extend(r.to_le_bytes());
            }
        }
        b.push(add.len() as u8);
        for (p, _) in add {
            b.extend(p.to_le_bytes());
        }
        for (_, d) in add {
            b.push(*d);
        }
        b[0] = (b.len() - 1).min(255) as u8;
        b
    }
    #[test]
    fn deletion_tolerance_and_plus_one_ranges_have_exact_boundaries() {
        let mut t = Stops::default();
        t.apply(
            &operand(
                &[],
                &[(974, 0), (975, 0), (1000, 0), (1025, 0), (1026, 0)],
                false,
            ),
            false,
        )
        .unwrap();
        t.apply(&operand(&[(1000, 0)], &[], false), false).unwrap();
        assert_eq!(t.0.keys().copied().collect::<Vec<_>>(), [974, 1026]);
        t.apply(&operand(&[(1000, 27)], &[], true), true).unwrap(); // Stored 27 = radius 26.
        assert!(t.0.is_empty());
        t.apply(&operand(&[], &[(975, 0), (1025, 0)], false), false)
            .unwrap();
        t.apply(&operand(&[(1000, 0)], &[], true), true).unwrap(); // Radius floors at 25.
        assert!(t.0.is_empty());
    }
    #[test]
    fn maps_descriptors_and_ignores_only_normatively_unused_bits() {
        let mut t = Stops::default();
        t.apply(
            &operand(
                &[],
                &[
                    (-100, 0xc0 | 8),
                    (100, 1 | 16),
                    (200, 2 | 24),
                    (300, 3 | 32),
                    (400, 4 | 48),
                    (500, 6 | 40),
                    (600, 56),
                ],
                false,
            ),
            false,
        )
        .unwrap();
        let xml = t.xml();
        for value in [
            "left",
            "center",
            "right",
            "decimal",
            "bar",
            "num",
            "dot",
            "hyphen",
            "underscore",
            "middleDot",
        ] {
            assert!(xml.contains(value));
        }
        assert_eq!(t.0[&400], ("bar", "none"));
        assert_eq!(t.0[&600], ("left", "none"));
        for d in [5, 7, 48] {
            assert!(t.apply(&operand(&[], &[(0, d)], false), false).is_err());
        }
    }
    #[test]
    fn extended_operand_counts_replace_the_255_length_marker() {
        let del: Vec<_> = (0..64).map(|i| (i * 100, 26)).collect();
        let mut t = Stops::default();
        let b = operand(&del, &[(100, 0)], true);
        assert_eq!(b[0], 255);
        t.apply(&b, true).unwrap();
        assert!(t.0.contains_key(&100));
        for n in 0..b.len() {
            assert!(Stops::default().apply(&b[..n], true).is_err());
        }
        let mut trailing = b.clone();
        trailing.push(0);
        assert!(t.apply(&trailing, true).is_err());
    }
    #[test]
    fn validates_counts_order_positions_and_aggregate_resource_limit() {
        let mut t = Stops::default();
        assert!(t.apply(&[2, 65, 0], false).is_err());
        assert!(t
            .apply(&operand(&[], &[(200, 0), (100, 0)], false), false)
            .is_err());
        assert!(t.apply(&operand(&[], &[(31681, 0)], false), false).is_err());
        assert!(t.apply(&operand(&[(0, 31682)], &[], true), true).is_err());
        for batch in 0..4 {
            let add: Vec<_> = (0..64).map(|i| ((batch * 64 + i) * 100, 0)).collect();
            t.apply(&operand(&[], &add, false), false).unwrap();
        }
        assert!(t
            .apply(&operand(&[], &[(26000, 0)], false), false)
            .unwrap_err()
            .contains("limit"));
        t.apply(&operand(&[], &[(0, 2)], false), false).unwrap(); // Replacement at the cap is valid.
    }
}
