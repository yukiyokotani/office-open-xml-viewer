//! Resolve document fonts, paragraph-style character defaults, character
//! styles and direct CHPX/PCD properties before serializing ordinary OOXML runs.
//! [MS-DOC] 2.4.6, STSH/STD, SttbfFfn and FFN. Table style conditions and
//! paragraph-property reconstruction are deliberately outside this subset.

use super::character::{self, Budget, Properties, Sprms};
use super::fkp::{self, Index, Kind};
use super::{u16_at, unsupported};
use std::collections::{BTreeMap, BTreeSet};

struct Style<'a> {
    base: usize,
    kind: u16,
    chpx: &'a [u8],
}

pub struct Formatting<'a> {
    pub characters: Index<'a>,
    paragraphs: Index<'a>,
    fonts: Vec<String>,
    defaults: Properties,
    styles: Vec<Option<Style<'a>>>,
    paragraph_cache: BTreeMap<usize, Properties>,
    budget: Budget,
    pub unsupported_character_properties: bool,
    pub missing_tables: bool,
}

impl<'a> Formatting<'a> {
    pub fn read(word: &'a [u8], table: &'a [u8]) -> Result<Self, String> {
        let fonts = read_fonts(fkp::table_part(word, table, 0x112)?)?;
        let (defaults, styles) = read_styles(fkp::table_part(word, table, 0xa2)?)?;
        let characters = Index::read(word, table, Kind::Character)?;
        let paragraphs = Index::read(word, table, Kind::Paragraph)?;
        let missing_tables =
            characters.is_empty() || paragraphs.is_empty() || styles.is_empty() || fonts.is_empty();
        Ok(Self {
            characters,
            paragraphs,
            fonts,
            defaults,
            styles,
            paragraph_cache: BTreeMap::new(),
            budget: Budget::default(),
            unsupported_character_properties: false,
            missing_tables,
        })
    }

    pub fn paragraph_style(&self, end_fc: usize) -> Result<usize, String> {
        if !self.paragraphs.is_empty() && self.paragraphs.at(end_fc).is_none() {
            return Err(unsupported("Word paragraph mark outside formatting ranges"));
        }
        fkp::paragraph_style(&self.paragraphs, end_fc)
    }

    fn chain(&mut self, mut id: usize, kind: u16) -> Result<Vec<usize>, String> {
        let mut result = Vec::new();
        let mut visited = BTreeSet::new();
        while id != 0xfff {
            self.budget.take()?;
            // Explicit resource policy. No recursive stack growth, cycles, or
            // exponentially expanded arrays of inherited property records.
            if result.len() >= 256 || !visited.insert(id) {
                return Err(unsupported("cyclic or excessive Word style inheritance"));
            }
            let Some(style) = self.styles.get(id).and_then(Option::as_ref) else {
                // Default Paragraph Font (10) is commonly a latent empty style.
                if (id == 10 && kind == 2) || (id == 0 && self.styles.is_empty()) {
                    break;
                }
                return Err(unsupported("Word style index references a missing style"));
            };
            if style.kind != kind {
                return Err(unsupported("Word style inheritance changes style kind"));
            }
            result.push(id);
            id = style.base;
        }
        result.reverse();
        Ok(result)
    }

    fn apply_style(&mut self, props: &mut Properties, id: usize, kind: u16) -> Result<(), String> {
        for id in self.chain(id, kind)? {
            let baseline = props.clone();
            let mut sprms = Sprms::new(self.styles[id].as_ref().expect("validated style").chpx);
            while let Some((code, operand)) = sprms.next(&mut self.budget)? {
                if !props.apply(code, operand, &baseline)? {
                    self.unsupported_character_properties = true;
                }
            }
        }
        Ok(())
    }

    fn paragraph_base(&mut self, id: usize) -> Result<Properties, String> {
        if let Some(value) = self.paragraph_cache.get(&id) {
            return Ok(value.clone());
        }
        let mut props = self.defaults.clone();
        self.apply_style(&mut props, id, 1)?;
        self.paragraph_cache.insert(id, props.clone());
        Ok(props)
    }

    /// A caller caches this result for a consecutive (paragraph style, CHPX,
    /// PCD) range. Properties are not decoded or allocated once per character.
    pub fn run_xml(
        &mut self,
        paragraph_style: usize,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<String, String> {
        let paragraph = self.paragraph_base(paragraph_style)?;
        let mut props = paragraph.clone();
        let mut style = paragraph.clone();
        let direct = match self.characters.at(fc) {
            Some((_, run)) => run.properties,
            None if self.characters.is_empty() => &[],
            None => return Err(unsupported("Word character outside formatting ranges")),
        };
        self.apply_direct(&mut props, &mut style, &paragraph, direct)?;
        if prm & 1 != 0 {
            let bytes = prcs
                .get((prm >> 1) as usize)
                .ok_or_else(|| unsupported("Word piece property index outside CLX"))?;
            self.apply_direct(&mut props, &mut style, &paragraph, bytes)?;
        } else if let Some(bytes) = character::prm0(prm) {
            self.apply_direct(&mut props, &mut style, &paragraph, &bytes)?;
        } else if prm != 0 {
            self.unsupported_character_properties = true;
        }
        props.xml(&self.fonts)
    }

    fn apply_direct(
        &mut self,
        props: &mut Properties,
        style: &mut Properties,
        paragraph: &Properties,
        bytes: &[u8],
    ) -> Result<(), String> {
        let mut sprms = Sprms::new(bytes);
        while let Some((code, operand)) = sprms.next(&mut self.budget)? {
            if (code >> 10) & 7 != 2 {
                continue;
            }
            match code {
                0x4a30 => {
                    let id = u16_at(operand, 0)? as usize;
                    let mut character_style = paragraph.clone();
                    self.apply_style(&mut character_style, id, 2)?;
                    // Reset exceptions survive both the reset and style application.
                    props.reset_to(&character_style);
                    *style = character_style;
                }
                0x2a33 => {
                    props.reset_to(paragraph);
                    *style = paragraph.clone();
                }
                _ => {
                    if !props.apply(code, operand, style)? {
                        self.unsupported_character_properties = true;
                    }
                }
            }
        }
        Ok(())
    }
}

fn read_fonts(bytes: &[u8]) -> Result<Vec<String>, String> {
    if bytes.is_empty() {
        return Ok(Vec::new());
    }
    let count = u16_at(bytes, 0)? as usize;
    if count > 0x7ff0 || u16_at(bytes, 2)? != 0 {
        return Err(unsupported("invalid Word font table header"));
    }
    let mut offset = 4;
    let mut fonts = Vec::with_capacity(count);
    for _ in 0..count {
        let size = *bytes
            .get(offset)
            .ok_or_else(|| unsupported("truncated Word font record"))? as usize;
        offset += 1;
        let font = bytes
            .get(offset..offset + size)
            .ok_or_else(|| unsupported("truncated Word font data"))?;
        let name = font
            .get(39..)
            .ok_or_else(|| unsupported("short Word font data"))?;
        let units: Vec<_> = name
            .chunks_exact(2)
            .map(|b| u16::from_le_bytes([b[0], b[1]]))
            .take_while(|u| *u != 0)
            .collect();
        if units.is_empty() || units.len() * 2 + 2 > name.len() {
            return Err(unsupported("unterminated or empty Word font name"));
        }
        fonts.push(
            String::from_utf16(&units)
                .map_err(|_| unsupported("invalid Unicode Word font name"))?,
        );
        offset += size;
    }
    Ok(fonts)
}

fn read_styles(bytes: &[u8]) -> Result<(Properties, Vec<Option<Style<'_>>>), String> {
    let mut defaults = Properties::default();
    if bytes.is_empty() {
        return Ok((defaults, Vec::new()));
    }
    let header_size = u16_at(bytes, 0)? as usize;
    let header = bytes
        .get(2..2 + header_size)
        .filter(|v| v.len() >= 18)
        .ok_or_else(|| unsupported("truncated Word stylesheet header"))?;
    let count = u16_at(header, 0)? as usize;
    let base_size = u16_at(header, 2)? as usize;
    if !(15..4094).contains(&count) || ![10, 18].contains(&base_size) {
        return Err(unsupported("invalid Word stylesheet header"));
    }
    for (slot, field) in [12, 14, 16, 18].iter().enumerate() {
        if *field + 2 <= header.len() {
            let font = u16_at(header, *field)?;
            if font > i16::MAX as u16 {
                return Err(unsupported("negative default Word font index"));
            }
            defaults.fonts[slot] = Some(font as usize);
        }
    }
    let mut offset = 2 + header_size;
    let mut styles = Vec::with_capacity(count);
    for _ in 0..count {
        let size = u16_at(bytes, offset)? as usize;
        if size > i16::MAX as usize {
            return Err(unsupported("negative Word style size"));
        }
        offset += 2;
        let std = bytes
            .get(offset..offset + size)
            .ok_or_else(|| unsupported("truncated Word style definition"))?;
        offset += size + size % 2;
        if size == 0 {
            styles.push(None);
            continue;
        }
        let kind_and_base = u16_at(std, 2)?;
        let kind = kind_and_base & 15;
        let count = (u16_at(std, 4)? & 15) as usize;
        let name_len = u16_at(std, base_size)? as usize;
        let mut p = base_size + 2 + name_len * 2;
        if u16_at(std, p)? != 0 {
            return Err(unsupported("unterminated Word style name"));
        }
        p += 2;
        let mut chpx = &[][..];
        for i in 0..count {
            let n = u16_at(std, p)? as usize;
            p += 2;
            let upx = std
                .get(p..p + n)
                .ok_or_else(|| unsupported("truncated Word style properties"))?;
            // Other sets include paragraph/table properties and old revision
            // formatting. Never apply those as current character properties.
            if (kind == 1 && i == 1) || (kind == 2 && i == 0) {
                chpx = upx;
            }
            p += n + n % 2;
        }
        styles.push(Some(Style {
            base: (kind_and_base >> 4) as usize,
            kind,
            chpx,
        }));
    }
    Ok((defaults, styles))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reads_font_names_after_ffn_metadata_not_as_latin1() {
        let mut bytes = vec![1, 0, 0, 0];
        let mut font = vec![0; 39];
        for unit in "日本語 Font\0".encode_utf16() {
            font.extend(unit.to_le_bytes());
        }
        bytes.push(font.len() as u8);
        bytes.extend(font);
        assert_eq!(read_fonts(&bytes).unwrap(), ["日本語 Font"]);
        bytes.pop();
        assert!(read_fonts(&bytes).is_err());
    }

    fn empty() -> Formatting<'static> {
        Formatting {
            characters: Index::default(),
            paragraphs: Index::default(),
            fonts: vec![],
            defaults: Properties::default(),
            styles: vec![],
            paragraph_cache: BTreeMap::new(),
            budget: Budget::default(),
            unsupported_character_properties: false,
            missing_tables: true,
        }
    }

    #[test]
    fn paragraph_and_character_styles_resolve_before_piece_overrides() {
        let mut f = empty();
        f.styles = vec![
            Some(Style {
                kind: 1,
                base: 0xfff,
                chpx: &[0x43, 0x4a, 24, 0, 0x35, 8, 1],
            }),
            Some(Style {
                kind: 1,
                base: 0,
                chpx: &[0x43, 0x4a, 32, 0],
            }),
            Some(Style {
                kind: 2,
                base: 0xfff,
                chpx: &[0x36, 8, 1],
            }),
        ];
        let xml = f
            .run_xml(1, 0, 1, &[&[0x30, 0x4a, 2, 0, 0x35, 8, 0x81]])
            .unwrap();
        assert!(xml.contains("w:sz w:val=\"32\""));
        assert!(xml.contains("w:i w:val=\"1\""));
        assert!(xml.contains("w:b w:val=\"0\""));
        let xml = f.run_xml(1, 0, 0x0100 | (0x55 << 1), &[]).unwrap();
        assert!(xml.contains("w:sz w:val=\"32\""));
    }

    #[test]
    fn rejects_style_cycles_and_invalid_complex_piece_references() {
        let mut f = empty();
        f.styles = vec![Some(Style {
            kind: 1,
            base: 0,
            chpx: &[],
        })];
        assert!(f.run_xml(0, 0, 0, &[]).unwrap_err().contains("cyclic"));
        f.styles.clear();
        assert!(f.run_xml(0, 0, 1, &[]).unwrap_err().contains("outside CLX"));
    }
}
