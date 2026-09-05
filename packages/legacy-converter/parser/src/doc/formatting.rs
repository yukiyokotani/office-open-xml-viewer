//! Resolve document fonts, paragraph-style character defaults, character
//! styles and direct CHPX/PCD properties before serializing ordinary OOXML runs.
//! [MS-DOC] 2.4.6, STSH/STD, SttbfFfn and FFN. Table style conditions and
//! advanced paragraph properties remain outside this subset.

use super::character::{self, Properties};
use super::fkp::{self, Index, Kind};
use super::sprm::{self, Budget, Sprms};
use super::{paragraph, table, u16_at, unsupported};
use std::collections::{BTreeMap, BTreeSet};

struct Style<'a> {
    base: usize,
    kind: u16,
    chpx: &'a [u8],
    papx: &'a [u8],
}

pub struct Formatting<'a> {
    pub characters: Index<'a>,
    paragraphs: Index<'a>,
    fonts: Vec<String>,
    defaults: Properties,
    styles: Vec<Option<Style<'a>>>,
    paragraph_cache: BTreeMap<usize, Properties>,
    paragraph_layout_cache: BTreeMap<usize, paragraph::Properties>,
    data: &'a [u8],
    budget: Budget,
    pub unsupported_character_properties: bool,
    pub unsupported_paragraph_properties: bool,
    pub unsupported_piece_properties: bool,
    pub missing_tables: bool,
    pub unsupported_table_properties: bool,
}

impl<'a> Formatting<'a> {
    pub fn read(word: &'a [u8], table: &'a [u8], data: &'a [u8]) -> Result<Self, String> {
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
            paragraph_layout_cache: BTreeMap::new(),
            data,
            budget: Budget::default(),
            unsupported_character_properties: false,
            unsupported_paragraph_properties: false,
            unsupported_piece_properties: false,
            missing_tables,
            unsupported_table_properties: false,
        })
    }

    pub fn paragraph_style(&self, end_fc: usize) -> Result<usize, String> {
        if !self.paragraphs.is_empty() && self.paragraphs.at(end_fc).is_none() {
            return Err(unsupported("Word paragraph mark outside formatting ranges"));
        }
        fkp::paragraph_style(&self.paragraphs, end_fc)
    }

    pub fn paragraph_xml(
        &mut self,
        style: usize,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<String, String> {
        let mut props = if let Some(props) = self.paragraph_layout_cache.get(&style) {
            props.clone()
        } else {
            let mut props = paragraph::Properties::default();
            for id in self.chain(style, 1)? {
                let bytes = self.styles[id].as_ref().expect("validated style").papx;
                self.apply_paragraph(&mut props, bytes)?;
            }
            self.paragraph_layout_cache.insert(style, props.clone());
            props
        };
        let direct = self
            .paragraphs
            .at(fc)
            .map_or(&[][..], |(_, run)| run.properties);
        if !direct.is_empty() {
            self.apply_paragraph(&mut props, &direct[2..])?;
        }
        if prm & 1 != 0 {
            let bytes = prcs
                .get((prm >> 1) as usize)
                .ok_or_else(|| unsupported("Word piece property index outside CLX"))?;
            self.apply_paragraph(&mut props, bytes)?;
        } else if let Some(bytes) = paragraph::prm0(prm) {
            self.apply_paragraph(&mut props, &bytes)?;
        }
        Ok(props.xml())
    }

    fn apply_paragraph<'b>(
        &mut self,
        props: &mut paragraph::Properties,
        bytes: &'b [u8],
    ) -> Result<(), String>
    where
        'a: 'b,
    {
        sprm::paragraph_properties(bytes, self.data, &mut self.budget, |code, operand| {
            if (code >> 10) & 7 == 1
                && !matches!(code, 0x2416 | 0x2417 | 0x6649 | 0x664a | 0x244b | 0x244c)
                && !props.apply(code, operand)?
            {
                self.unsupported_paragraph_properties = true;
            }
            Ok(())
        })
    }

    pub fn table_properties(
        &mut self,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<table::Properties, String> {
        let mut properties = table::Properties::default();
        // MS-DOC 2.4.3: structural flags are direct, never inherited from STSH.
        let direct = self
            .paragraphs
            .at(fc)
            .map_or(&[][..], |(_, r)| r.properties);
        let direct = if direct.is_empty() {
            direct
        } else {
            &direct[2..]
        };
        let piece = if prm & 1 != 0 {
            *prcs
                .get((prm >> 1) as usize)
                .ok_or_else(|| unsupported("Word piece property index outside CLX"))?
        } else {
            &[]
        };
        for bytes in [direct, piece] {
            sprm::paragraph_properties(bytes, self.data, &mut self.budget, |code, operand| {
                if !properties.apply(code, operand)? && (code >> 10) & 7 == 5 {
                    self.unsupported_table_properties = true;
                }
                Ok(())
            })?;
        }
        if prm & 1 == 0 {
            if let Some([a, b, value]) = table::prm0(prm) {
                properties.apply(u16::from_le_bytes([a, b]), &[value])?;
            }
        }
        Ok(properties)
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
        let props = self.run_properties(paragraph_style, fc, prm, prcs)?;
        props.xml(&self.fonts)
    }

    pub fn inline_picture_location(
        &mut self,
        style: usize,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<Option<usize>, String> {
        self.run_properties(style, fc, prm, prcs)?
            .picture
            .inline_location()
    }

    pub fn floating_picture_allowed(
        &mut self,
        style: usize,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<bool, String> {
        Ok(self
            .run_properties(style, fc, prm, prcs)?
            .picture
            .passive_special())
    }

    fn run_properties(
        &mut self,
        paragraph_style: usize,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<Properties, String> {
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
        } else if prm != 0 && paragraph::prm0(prm).is_none() && table::prm0(prm).is_none() {
            self.unsupported_piece_properties = true;
        }
        Ok(props)
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
                    props.reset_to(&character_style, true);
                    *style = character_style;
                }
                0x2a33 => {
                    props.reset_to(paragraph, false);
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
        let mut papx = &[][..];
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
            if kind == 1 && i == 0 {
                papx = upx
                    .get(2..)
                    .ok_or_else(|| unsupported("missing Word style paragraph index"))?;
            }
            p += n + n % 2;
        }
        styles.push(Some(Style {
            base: (kind_and_base >> 4) as usize,
            kind,
            chpx,
            papx,
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
            paragraph_layout_cache: BTreeMap::new(),
            data: &[],
            budget: Budget::default(),
            unsupported_character_properties: false,
            unsupported_paragraph_properties: false,
            unsupported_piece_properties: false,
            missing_tables: true,
            unsupported_table_properties: false,
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
                papx: &[],
            }),
            Some(Style {
                kind: 1,
                base: 0,
                chpx: &[0x43, 0x4a, 32, 0],
                papx: &[],
            }),
            Some(Style {
                kind: 2,
                base: 0xfff,
                chpx: &[0x36, 8, 1],
                papx: &[],
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
            papx: &[],
        })];
        assert!(f.run_xml(0, 0, 0, &[]).unwrap_err().contains("cyclic"));
        f.styles.clear();
        assert!(f.run_xml(0, 0, 1, &[]).unwrap_err().contains("outside CLX"));
    }

    #[test]
    fn custom_tabs_resolve_style_additions_then_direct_deletions_and_replacements() {
        let mut f = empty();
        f.styles = vec![Some(Style {
            kind: 1,
            base: 0xfff,
            chpx: &[],
            // Two tabs: 720 left/dotted, 1440 right/no leader.
            papx: &[0x0d, 0xc6, 8, 0, 2, 0xd0, 2, 0xa0, 5, 8, 2],
        })];
        let original = f.paragraph_xml(0, 0, 0, &[]).unwrap();
        assert!(original.contains("<w:tab w:val=\"left\" w:pos=\"720\" w:leader=\"dot\"/>"));
        // Delete at 740 (+20 twips from inherited 720), replace 1440 with center.
        let modified = f
            .paragraph_xml(0, 0, 1, &[&[0x0d, 0xc6, 7, 1, 0xe4, 2, 1, 0xa0, 5, 1]])
            .unwrap();
        assert!(!modified.contains("w:pos=\"720\""));
        assert!(modified.contains("<w:tab w:val=\"center\" w:pos=\"1440\" w:leader=\"none\"/>"));
        assert_eq!(f.paragraph_xml(0, 0, 0, &[]).unwrap(), original);
        assert!(!f.unsupported_paragraph_properties);
    }
    #[test]
    fn inherited_paragraph_layout_is_overridden_by_piece_properties() {
        let mut f = empty();
        f.styles = vec![Some(Style {
            kind: 1,
            base: 0xfff,
            chpx: &[],
            papx: &[0x12, 0x64, 0xd4, 0xfe, 0, 0, 0x13, 0xa4, 240, 0],
        })];
        let xml = f
            .paragraph_xml(0, 0, 1, &[&[0x13, 0xa4, 0, 0, 0x07, 0x24, 1]])
            .unwrap();
        assert!(xml.contains("w:line=\"300\" w:lineRule=\"exact\""));
        assert!(xml.contains("w:before=\"0\""));
        assert!(xml.contains("<w:pageBreakBefore w:val=\"1\"/>"));
        assert!(f
            .paragraph_xml(0, 0, 0, &[])
            .unwrap()
            .contains("w:before=\"240\""));
    }

    #[test]
    fn paragraph_mark_physical_offset_selects_direct_layout_before_piece_override() {
        let mut word = vec![0u8; 1024];
        // A single FKP can contain both paragraph runs; its BTE covers both.
        let table: Vec<u8> = [100u32, 120, 1]
            .into_iter()
            .flat_map(u32::to_le_bytes)
            .collect();
        word[0x106..0x10a].copy_from_slice(&12u32.to_le_bytes());
        let page = &mut word[512..1024];
        for (i, fc) in [100u32, 110, 120].into_iter().enumerate() {
            page[i * 4..i * 4 + 4].copy_from_slice(&fc.to_le_bytes());
        }
        page[12] = 32;
        page[25] = 48;
        page[64..74].copy_from_slice(&[5, 0, 0, 0x13, 0xa4, 120, 0, 0x61, 0x24, 1]);
        page[96..106].copy_from_slice(&[5, 0, 0, 0x13, 0xa4, 240, 0, 0x61, 0x24, 2]);
        page[511] = 2;
        let mut f = Formatting::read(&word, &table, &[]).unwrap();
        let first = f.paragraph_xml(0, 109, 0, &[]).unwrap();
        assert!(first.contains("w:before=\"120\""));
        assert!(first.contains("<w:jc w:val=\"center\"/>"));
        let second = f.paragraph_xml(0, 110, 1, &[&[0x13, 0xa4, 0, 0]]).unwrap();
        assert!(second.contains("w:before=\"0\""));
        assert!(second.contains("<w:jc w:val=\"right\"/>"));
        assert!(f
            .paragraph_xml(0, 120, 0, &[])
            .unwrap()
            .contains("w:before=\"0\""));
    }

    #[test]
    fn paragraph_data_indirection_replaces_tail_and_rejects_cycles() {
        let data = [10, 0, 0x12, 0x64, 0xd4, 0xfe, 0, 0, 0x13, 0xa4, 120, 0];
        let mut f = empty();
        f.data = &data;
        let mut props = paragraph::Properties::default();
        f.apply_paragraph(&mut props, &[0x46, 0x66, 0, 0, 0, 0, 0x13, 0xa4, 240, 0])
            .unwrap();
        assert!(props.xml().contains("w:before=\"120\""));
        assert!(props.xml().contains("w:line=\"300\""));
        // A non-first PHugePapx must be ignored, including an invalid pointer.
        f.apply_paragraph(&mut props, &[0x07, 0x24, 1, 0x46, 0x66, 255, 255, 255, 255])
            .unwrap();
        assert!(f
            .apply_paragraph(&mut props, &[0x46, 0x66, 255, 255, 255, 255])
            .is_err());
        let cyclic = [10, 0, 0x46, 0x66, 0, 0, 0, 0, 0x13, 0xa4, 0, 0];
        f.data = &cyclic;
        assert!(f
            .apply_paragraph(&mut props, &[0x46, 0x66, 0, 0, 0, 0])
            .unwrap_err()
            .contains("cyclic"));
    }

    #[test]
    fn supported_paragraph_prm_does_not_report_character_loss() {
        let mut f = empty();
        for code in [0x09, 0x18, 0x19] {
            f.run_xml(0, 0, 0x0100 | (code << 1), &[]).unwrap();
        }
        assert!(!f.unsupported_piece_properties && !f.unsupported_character_properties);
    }
}
