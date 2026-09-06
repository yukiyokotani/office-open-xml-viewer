//! Passive color projection of MS-XLS Theme (2.4.326).
use super::{u16_at, u32_at, unsupported, Record};
use std::collections::BTreeSet;
use std::io::{Cursor, Read};
mod xml;

const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const R: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
// Resource policy, independent of the BIFF/OPC format limits.
const MAX_ZIP_BYTES: usize = 4 * 1024 * 1024;
const MAX_ENTRIES: usize = 64;
const MAX_PART_BYTES: usize = 256 * 1024;
const MAX_EXPANDED_BYTES: u64 = 2 * 1024 * 1024;

#[derive(Default)]
pub(super) struct Colors([Option<[u8; 3]>; 12]);

impl Colors {
    pub(super) fn parse(records: &[Record<'_>]) -> Result<Self, String> {
        let globals = records.iter().take_while(|r| r.kind != super::EOF).count();
        let mut matches = records[..globals]
            .iter()
            .enumerate()
            .filter(|(_, r)| r.kind == 0x0896);
        let Some((index, record)) = matches.next() else {
            return Ok(Self::default());
        };
        if matches.next().is_some() || record.data.len() < 16 || u16_at(record.data, 0)? != 0x0896 {
            return Err(unsupported("invalid or duplicate BIFF theme"));
        }
        if record.data.len() - 16 > MAX_ZIP_BYTES {
            return Err(unsupported("BIFF theme byte budget exceeded"));
        }
        let mut bytes = record.data[16..].to_vec();
        // MS-XLS 2.1.7.20.3 THEME = Theme *ContinueFrt12, 2.4.62.
        for continuation in records[index + 1..globals]
            .iter()
            .take_while(|r| r.kind == 0x087f)
        {
            if continuation.data.len() < 12
                || continuation.data.len() > 8224
                || u16_at(continuation.data, 0)? != 0x087f
            {
                return Err(unsupported("invalid BIFF theme continuation"));
            }
            let payload = &continuation.data[12..];
            if payload.len() > MAX_ZIP_BYTES.saturating_sub(bytes.len()) {
                return Err(unsupported("BIFF theme byte budget exceeded"));
            }
            bytes.extend_from_slice(payload);
        }
        if bytes.is_empty() {
            if u32_at(record.data, 12)? == 0 {
                return Err(unsupported("missing custom BIFF theme"));
            }
            // A version-only default theme is not an embedded color scheme.
            // Keep the palette fallback; do not guess a current Office theme.
            return Ok(Self::default());
        }
        Self::package(&bytes)
    }

    pub(super) fn rgb(&self, index: u32) -> Option<String> {
        let [r, g, b] = self.0.get(index as usize).copied().flatten()?;
        Some(format!("rgb=\"FF{r:02X}{g:02X}{b:02X}\""))
    }

    fn package(bytes: &[u8]) -> Result<Self, String> {
        let entries = preflight(bytes)?;
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
            .map_err(|_| unsupported("invalid BIFF theme ZIP"))?;
        let mut names = BTreeSet::new();
        let mut total = 0u64;
        if archive.len() != entries || archive.len() > MAX_ENTRIES {
            return Err(unsupported("ambiguous BIFF theme ZIP entry table"));
        }
        for i in 0..archive.len() {
            let entry = archive
                .by_index_raw(i)
                .map_err(|_| unsupported("invalid BIFF theme ZIP entry"))?;
            if entry.is_dir() {
                continue;
            }
            total = total
                .checked_add(entry.size())
                .ok_or_else(|| unsupported("BIFF theme size overflow"))?;
            if total > MAX_EXPANDED_BYTES
                || entry.size() > MAX_PART_BYTES as u64
                || entry.encrypted()
                || entry.is_symlink()
                || !names.insert(entry.name().to_owned())
                || normalize("", entry.name())? != entry.name()
            {
                return Err(unsupported("unsafe or oversized BIFF theme ZIP entry"));
            }
        }
        // Resolve the explicit package relationship chain. Filenames and ZIP
        // enumeration order do not identify the active theme.
        let root = xml::parse(&part(&mut archive, "_rels/.rels")?)?;
        let manager = relationship(&root, "", "officeDocument")?;
        let manager_xml = xml::parse(&part(&mut archive, &manager)?)?;
        if !matches!(manager_xml.ns.as_str(), A | STRICT_A) || manager_xml.name != "themeManager" {
            return Err(unsupported("invalid BIFF theme manager root"));
        }
        let (dir, name) = manager.rsplit_once('/').unwrap_or(("", &manager));
        let rels = if dir.is_empty() {
            format!("_rels/{name}.rels")
        } else {
            format!("{dir}/_rels/{name}.rels")
        };
        let relationships = xml::parse(&part(&mut archive, &rels)?)?;
        let target = relationship(&relationships, &manager, "theme")?;
        Self::document(&xml::parse(&part(&mut archive, &target)?)?)
    }

    fn document(root: &xml::Node) -> Result<Self, String> {
        if !matches!(root.ns.as_str(), A | STRICT_A) || root.name != "theme" {
            return Err(unsupported("invalid BIFF theme XML root"));
        }
        let ns = root.ns.as_str();
        let elements = root.only_child(ns, "themeElements")?;
        let scheme = elements.only_child(ns, "clrScheme")?;
        let mut colors = Self::default();
        let mut seen = BTreeSet::new();
        // MS-XLS ColorTheme 2.5.49 / ECMA-376 20.1.6.2: use names,
        // not the order of arbitrary input XML children.
        for slot in &scheme.children {
            if slot.ns != ns {
                continue;
            }
            let Some(index) = [
                "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
                "accent6", "hlink", "folHlink",
            ]
            .iter()
            .position(|name| *name == slot.name) else {
                continue;
            };
            if !seen.insert(index) || slot.children.len() != 1 {
                return Err(unsupported("ambiguous BIFF theme color slot"));
            }
            let color = &slot.children[0];
            // Never strip transforms and then claim the base color is the result.
            if color.ns != ns || !color.children.is_empty() {
                continue;
            }
            let value = match color.name.as_str() {
                "srgbClr" => color.attrs.get("val"),
                // ECMA-376 20.1.2.3.33: saved generating-application color.
                // No host OS query or invented system-color fallback.
                "sysClr" => color.attrs.get("lastClr"),
                _ => None,
            };
            if let Some(value) = value {
                let value = value.trim();
                if value.len() != 6 || !value.bytes().all(|b| b.is_ascii_hexdigit()) {
                    return Err(unsupported("invalid BIFF theme RGB value"));
                }
                colors.0[index] = Some([
                    u8::from_str_radix(&value[0..2], 16).unwrap(),
                    u8::from_str_radix(&value[2..4], 16).unwrap(),
                    u8::from_str_radix(&value[4..6], 16).unwrap(),
                ]);
            }
        }
        Ok(colors)
    }
}

fn part(archive: &mut zip::ZipArchive<Cursor<&[u8]>>, name: &str) -> Result<Vec<u8>, String> {
    let entry = archive
        .by_name(name)
        .map_err(|_| unsupported("missing BIFF theme part"))?;
    let declared = entry.size();
    if declared > MAX_PART_BYTES as u64 {
        return Err(unsupported("BIFF theme part byte budget exceeded"));
    }
    let mut bytes = Vec::with_capacity(declared as usize);
    entry
        .take(declared + 1)
        .read_to_end(&mut bytes)
        .map_err(|_| unsupported("invalid BIFF theme part data"))?;
    if bytes.len() as u64 != declared {
        return Err(unsupported("BIFF theme inflated size mismatch"));
    }
    Ok(bytes)
}

fn relationship(root: &xml::Node, source: &str, kind: &str) -> Result<String, String> {
    if root.ns != R || root.name != "Relationships" {
        return Err(unsupported("invalid BIFF theme relationships"));
    }
    let mut target = None;
    let mut ids = BTreeSet::new();
    for item in &root.children {
        if item.ns != R || item.name != "Relationship" {
            continue;
        }
        let id = item
            .attrs
            .get("Id")
            .ok_or_else(|| unsupported("missing BIFF theme relationship id"))?;
        if id.is_empty() || !ids.insert(id) {
            return Err(unsupported("duplicate BIFF theme relationship id"));
        }
        let typ = item.attrs.get("Type").map(String::as_str).unwrap_or("");
        if typ != format!("{REL}/{kind}") && typ != format!("{STRICT_REL}/{kind}") {
            continue;
        }
        if target.is_some()
            || item
                .attrs
                .get("TargetMode")
                .is_some_and(|m| m != "Internal")
        {
            return Err(unsupported("ambiguous or external BIFF theme relationship"));
        }
        target = Some(normalize(
            source,
            item.attrs
                .get("Target")
                .ok_or_else(|| unsupported("missing BIFF theme relationship target"))?,
        )?);
    }
    target.ok_or_else(|| unsupported("missing BIFF theme relationship"))
}

fn normalize(source: &str, target: &str) -> Result<String, String> {
    // These are OPC names only; no filesystem or network operation is performed.
    if target.is_empty()
        || target.starts_with("//")
        || target.contains(['\\', ':', '?', '#', '%'])
        || target.chars().any(|c| c.is_control() || c.is_whitespace())
    {
        return Err(unsupported("unsupported BIFF theme part URI"));
    }
    let mut components: Vec<&str> = if target.starts_with('/') {
        vec![]
    } else {
        source
            .rsplit_once('/')
            .map(|(p, _)| p.split('/').collect())
            .unwrap_or_default()
    };
    for component in target.trim_start_matches('/').split('/') {
        match component {
            "" => return Err(unsupported("invalid BIFF theme part URI")),
            "." => {}
            ".." => {
                components
                    .pop()
                    .ok_or_else(|| unsupported("BIFF theme target escapes package"))?;
            }
            _ => components.push(component),
        }
    }
    if components.is_empty() {
        return Err(unsupported("empty BIFF theme part URI"));
    }
    Ok(components.join("/"))
}

fn preflight(bytes: &[u8]) -> Result<usize, String> {
    if bytes.len() < 22 || bytes.len() > MAX_ZIP_BYTES {
        return Err(unsupported("BIFF theme ZIP byte budget exceeded"));
    }
    // Bound entry-table allocation before invoking the ZIP reader. ZIP64,
    // split archives and trailing data are outside this small embedded subset.
    let end = (bytes.len().saturating_sub(65557)..=bytes.len() - 22)
        .rev()
        .find(|&i| {
            bytes[i..].starts_with(b"PK\x05\x06")
                && u16_at(bytes, i + 20).is_ok_and(|n| i + 22 + usize::from(n) == bytes.len())
        })
        .ok_or_else(|| unsupported("missing BIFF theme ZIP directory"))?;
    let count = usize::from(u16_at(bytes, end + 10)?);
    if u16_at(bytes, end + 4)? != 0
        || u16_at(bytes, end + 6)? != 0
        || usize::from(u16_at(bytes, end + 8)?) != count
        || count > MAX_ENTRIES
        || u32_at(bytes, end + 12)? as u64 + u32_at(bytes, end + 16)? as u64 != end as u64
    {
        return Err(unsupported("unsupported BIFF theme ZIP directory"));
    }
    Ok(count)
}

#[cfg(test)]
mod tests {
    use super::*;
    const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const R: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn package(theme: &str) -> Vec<u8> {
        crate::ooxml::write_package(&[
            ("_rels/.rels".into(), format!("<Relationships xmlns=\"{R}\"><Relationship Id=\"main\" Type=\"{REL}/officeDocument\" Target=\"styles/manager.xml\"/></Relationships>")),
            ("styles/manager.xml".into(), format!("<a:themeManager xmlns:a=\"{A}\"/>")),
            ("styles/_rels/manager.xml.rels".into(), format!("<Relationships xmlns=\"{R}\"><Relationship Id=\"theme\" Type=\"{REL}/theme\" Target=\"../colors/custom.xml\"/></Relationships>")),
            ("colors/custom.xml".into(), theme.into()),
        ], 1024 * 1024).unwrap()
    }

    fn record_data(package: &[u8]) -> Vec<u8> {
        let mut data = vec![0; 16];
        data[..2].copy_from_slice(&0x0896u16.to_le_bytes());
        data.extend_from_slice(package);
        data
    }

    #[test]
    fn resolves_the_owned_theme_instead_of_guessing_part_names_or_color_order() {
        let xml = format!("<a:theme xmlns:a=\"{A}\"><a:themeElements><a:clrScheme name=\"Test\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"123456\"/></a:dk1><a:lt1><a:srgbClr val=\"ABCDEF\"/></a:lt1><a:accent1><a:srgbClr val=\"2468AC\"/></a:accent1></a:clrScheme></a:themeElements></a:theme>");
        let data = record_data(&package(&xml));
        let colors = Colors::parse(&[Record {
            kind: 0x0896,
            offset: 0,
            data: &data,
        }])
        .unwrap();
        assert_eq!(colors.0[0], Some([0x12, 0x34, 0x56]));
        assert_eq!(colors.0[1], Some([0xab, 0xcd, 0xef]));
        assert_eq!(colors.0[4], Some([0x24, 0x68, 0xac]));
        assert_eq!(colors.0[2], None);
    }

    fn document(colors: &str) -> String {
        format!("<a:theme xmlns:a=\"{A}\"><a:themeElements><a:clrScheme name=\"Test\">{colors}</a:clrScheme></a:themeElements></a:theme>")
    }

    #[test]
    fn missing_system_fallback_and_color_transforms_are_not_guessed() {
        let xml = document("<a:dk1><a:sysClr val=\"windowText\"/></a:dk1><a:accent1><a:srgbClr val=\"123456\"><a:tint val=\"50000\"/></a:srgbClr></a:accent1>");
        let colors = Colors::document(&xml::parse(xml.as_bytes()).unwrap()).unwrap();
        assert!(colors.0.iter().all(Option::is_none));
    }

    #[test]
    fn all_twelve_slots_are_name_bound_with_namespace_and_parent_ownership() {
        let names = [
            "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
            "accent6", "hlink", "folHlink",
        ];
        for ns in [A, STRICT_A] {
            let slots: String = names
                .iter()
                .enumerate()
                .rev()
                .map(|(i, name)| format!("<a:{name}><a:srgbClr val=\"0000{i:02X}\"/></a:{name}>"))
                .collect();
            let source = document(&slots).replace(A, ns);
            let colors = Colors::document(&xml::parse(source.as_bytes()).unwrap()).unwrap();
            for i in 0..12 {
                assert_eq!(colors.rgb(i), Some(format!("rgb=\"FF0000{i:02X}\"")));
            }
            assert!(colors.rgb(12).is_none());
        }
        let fake = document("<a:dk1 xmlns:a=\"urn:foreign\"><a:srgbClr val=\"123456\"/></a:dk1><a:extLst><a:dk1><a:srgbClr val=\"ABCDEF\"/></a:dk1></a:extLst>");
        assert!(Colors::document(&xml::parse(fake.as_bytes()).unwrap())
            .unwrap()
            .rgb(0)
            .is_none());
    }

    #[test]
    fn rejects_duplicate_slots_and_invalid_hex_without_clamping() {
        for slots in [
            "<a:dk1><a:srgbClr val=\"123456\"/></a:dk1><a:dk1><a:srgbClr val=\"FFFFFF\"/></a:dk1>",
            "<a:dk1><a:srgbClr val=\"123456\"/><a:srgbClr val=\"FFFFFF\"/></a:dk1>",
            "<a:dk1><a:srgbClr val=\"#12345\"/></a:dk1>",
            "<a:dk1><a:srgbClr val=\"12345678\"/></a:dk1>",
        ] {
            assert!(Colors::document(&xml::parse(document(slots).as_bytes()).unwrap()).is_err());
        }
    }

    #[test]
    fn reassembles_only_adjacent_theme_continuations() {
        let zip = package(&document("<a:dk1><a:srgbClr val=\"123456\"/></a:dk1>"));
        for split in [1, zip.len() / 2, zip.len() - 1] {
            let first = record_data(&zip[..split]);
            let mut second = vec![0; 12];
            second[..2].copy_from_slice(&0x087fu16.to_le_bytes());
            second.extend_from_slice(&zip[split..]);
            let r = [
                Record {
                    kind: 0x0896,
                    offset: 0,
                    data: &first,
                },
                Record {
                    kind: 0x087f,
                    offset: first.len() + 4,
                    data: &second,
                },
            ];
            assert_eq!(
                Colors::parse(&r).unwrap().rgb(0),
                Some("rgb=\"FF123456\"".into())
            );
            let interrupted = [
                r[0],
                Record {
                    kind: 0x089b,
                    offset: 0,
                    data: &[],
                },
                r[1],
            ];
            assert!(Colors::parse(&interrupted).is_err());
        }
    }

    #[test]
    fn validates_theme_header_and_global_ownership_but_does_not_guess_defaults() {
        let mut data = record_data(&[]);
        assert!(Colors::parse(&[Record {
            kind: 0x0896,
            offset: 0,
            data: &data
        }])
        .is_err());
        for version in [124226u32, 123820] {
            data[12..16].copy_from_slice(&version.to_le_bytes());
            let r = Record {
                kind: 0x0896,
                offset: 0,
                data: &data,
            };
            assert!(Colors::parse(&[r]).unwrap().rgb(0).is_none());
            assert!(Colors::parse(&[r, r]).is_err());
        }
        let hidden = [
            Record {
                kind: super::super::EOF,
                offset: 0,
                data: &[],
            },
            Record {
                kind: 0x0896,
                offset: 0,
                data: &[0],
            },
        ];
        assert!(Colors::parse(&hidden).unwrap().rgb(0).is_none());
    }

    #[test]
    fn relationship_resolution_rejects_external_ambiguous_and_escaping_targets() {
        for target in [
            "https://example.invalid/theme.xml",
            "//example.invalid/x",
            "../../escape.xml",
            "a%2fb.xml",
            "x.xml#fragment",
            "x\\y.xml",
        ] {
            assert!(normalize("folder/manager.xml", target).is_err());
        }
        assert_eq!(
            normalize("folder/manager.xml", "../colors/a.xml").unwrap(),
            "colors/a.xml"
        );
        assert_eq!(
            normalize("folder/manager.xml", "/colors/a.xml").unwrap(),
            "colors/a.xml"
        );
        for attrs in ["TargetMode=\"External\"", "TargetMode=\"unknown\""] {
            let source = format!("<Relationships xmlns=\"{R}\"><Relationship Id=\"a\" Type=\"{REL}/theme\" Target=\"a.xml\" {attrs}/></Relationships>");
            assert!(relationship(&xml::parse(source.as_bytes()).unwrap(), "", "theme").is_err());
        }
        for second in ["a", "b"] {
            let source = format!("<Relationships xmlns=\"{R}\"><Relationship Id=\"a\" Type=\"{REL}/theme\" Target=\"a.xml\"/><Relationship Id=\"{second}\" Type=\"{REL}/theme\" Target=\"b.xml\"/></Relationships>");
            assert!(relationship(&xml::parse(source.as_bytes()).unwrap(), "", "theme").is_err());
        }
    }

    #[test]
    fn preflight_bounds_zip_metadata_and_declared_inflation() {
        let small = package(&document(""));
        let end = small.len() - 22;
        for offset in [8, 10] {
            let mut bad = small.clone();
            bad[end + offset..end + offset + 2].copy_from_slice(&65u16.to_le_bytes());
            assert!(Colors::package(&bad).is_err());
        }
        let mut trailing = small.clone();
        trailing.push(0);
        assert!(Colors::package(&trailing).is_err());
        assert!(preflight(&vec![0; MAX_ZIP_BYTES + 1]).is_err());
        let large = package(&document(&format!("<!--{}-->", "x".repeat(MAX_PART_BYTES))));
        assert!(large.len() < MAX_PART_BYTES);
        assert!(Colors::package(&large).is_err());
    }

    #[test]
    fn rejects_duplicate_zip_names_and_mismatched_directory_counts() {
        let zip = crate::ooxml::write_package(
            &[("a.xml".into(), "a".into()), ("b.xml".into(), "b".into())],
            4096,
        )
        .unwrap();
        let mut duplicate = zip.clone();
        // Change both local and central names to create a true duplicate, not
        // merely a discrepancy between the two directory representations.
        for i in 0..duplicate.len() - 4 {
            if &duplicate[i..i + 5] == b"b.xml" {
                duplicate[i] = b'a';
            }
        }
        assert!(Colors::package(&duplicate)
            .err()
            .unwrap()
            .contains("entry table"));
        let mut wrong_count = zip;
        let end = wrong_count.len() - 22;
        for offset in [8, 10] {
            wrong_count[end + offset..end + offset + 2].copy_from_slice(&1u16.to_le_bytes());
        }
        assert!(Colors::package(&wrong_count).is_err());
    }
}
