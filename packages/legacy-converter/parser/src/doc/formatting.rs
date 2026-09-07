//! Resolve document fonts, paragraph-style character defaults, character
//! styles and direct CHPX/PCD properties before serializing ordinary OOXML runs.
//! [MS-DOC] 2.4.6, STSH/STD, SttbfFfn and FFN. Table style conditions and
//! advanced paragraph properties remain outside this subset.

use super::character::{self, Properties};
use super::fkp::{self, Index, Kind};
use super::sprm::{self, Budget, Sprms};
use super::{numbering, paragraph, table, u16_at, unsupported};
use std::collections::{BTreeMap, BTreeSet};

struct Style<'a> {
    base: usize,
    kind: u16,
    chpx: &'a [u8],
    papx: &'a [u8],
}

#[derive(Clone, Copy, Debug, Default)]
struct DirectParagraphProperties {
    bidi: Option<bool>,
    alignment: Option<(u16, u8)>,
    absolute_indents: [Option<(u16, i16)>; 6],
}

impl DirectParagraphProperties {
    fn overlay(&mut self, later: Self) {
        if later.bidi.is_some() {
            self.bidi = later.bidi;
        }
        if later.alignment.is_some() {
            self.alignment = later.alignment;
        }
        for value in later.absolute_indents.into_iter().flatten() {
            self.push_absolute_indent(value);
        }
    }

    fn push_absolute_indent(&mut self, value: (u16, i16)) {
        if let Some(index) = self
            .absolute_indents
            .iter()
            .position(|entry| entry.is_some_and(|entry| entry.0 == value.0))
        {
            self.absolute_indents[index..].rotate_left(1);
            self.absolute_indents[5] = None;
        }
        if let Some(slot) = self.absolute_indents.iter_mut().find(|slot| slot.is_none()) {
            *slot = Some(value);
        } else {
            // There are exactly six recognized absolute-indent codes. A full
            // array with no matching code therefore cannot accept another one.
            debug_assert!(false, "absolute-indent code set exceeded its fixed bound");
        }
    }
}

pub struct Formatting<'a> {
    pub characters: Index<'a>,
    paragraphs: Index<'a>,
    fonts: Vec<String>,
    defaults: Properties,
    styles: Vec<Option<Style<'a>>>,
    paragraph_cache: BTreeMap<usize, Properties>,
    paragraph_marker_styles: BTreeMap<usize, Properties>,
    paragraph_layout_cache: BTreeMap<usize, paragraph::Properties>,
    data: &'a [u8],
    budget: Budget,
    numbering: numbering::Tables<'a>,
    pub numbering_output: numbering::output::Store,
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
            paragraph_marker_styles: BTreeMap::new(),
            paragraph_layout_cache: BTreeMap::new(),
            data,
            budget: Budget::default(),
            numbering: numbering::Tables::read(word, table)?,
            numbering_output: numbering::output::Store::default(),
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
                let _ = self.apply_paragraph(&mut props, bytes)?;
            }
            self.paragraph_layout_cache.insert(style, props.clone());
            props
        };
        let direct = self
            .paragraphs
            .at(fc)
            .map_or(&[][..], |(_, run)| run.properties);
        let mut direct_properties = if !direct.is_empty() {
            self.apply_paragraph(&mut props, &direct[2..])?
        } else {
            DirectParagraphProperties::default()
        };
        if prm & 1 != 0 {
            let bytes = prcs
                .get((prm >> 1) as usize)
                .ok_or_else(|| unsupported("Word piece property index outside CLX"))?;
            direct_properties.overlay(self.apply_paragraph(&mut props, bytes)?);
        } else if let Some(bytes) = paragraph::prm0(prm) {
            direct_properties.overlay(self.apply_paragraph(&mut props, &bytes)?);
        }
        if let Some(reference) = numbering::Reference::new(props.ilfo, props.ilvl)? {
            let selected = self.numbering.resolve(reference)?;
            let level = *selected.level;
            let linked = selected.list.styles[usize::from(reference.level)];
            let original = props.clone();
            // MS-DOC 2.4.6.3 part 3 / 2.4.6.6 part 2: list paragraph
            // properties follow style/direct/PCD properties. Body character
            // runs continue to resolve their own original style and CHPX.
            if linked != 0xfff {
                for id in self.chain(usize::from(linked), 1)? {
                    let bytes = self.styles[id].as_ref().expect("validated style").papx;
                    let _ = self.apply_paragraph(&mut props, bytes)?;
                }
            }
            let _ = self.apply_paragraph(&mut props, level.papx)?;
            // MS-DOC 2.4.6.3 describes applying list properties after the
            // paragraph's style/direct properties. Word-produced evidence
            // establishes narrower precedence exceptions for explicitly
            // authored direct bidi and alignment: retain them after list
            // application. This does not alter other paragraph properties or
            // the same properties inherited only through a style.
            if let Some(value) = direct_properties.bidi {
                props.set_bidi(value);
            }
            if let Some((code, value)) = direct_properties.alignment {
                // Reuse the paragraph property's enum/range validation.
                props.apply(code, &[value])?;
            }
            // Replay only explicitly direct absolute twip indents, after the
            // list and after final direct bidi restoration. The fixed-size
            // capture preserves PAPX→PCD chronology without retaining style,
            // default, relative-nest, or character-unit properties.
            // Word-produced positive-iLfo controls establish this for signed
            // absolute values, explicit zero, and both paragraph directions;
            // this is observed Office precedence, not the literal list-last
            // ordering in MS-DOC 2.4.6.3. No coordinate correction is applied.
            for (code, value) in direct_properties.absolute_indents.into_iter().flatten() {
                props.apply(code, &value.to_le_bytes())?;
            }
            if reference.preserve_indent {
                props.preserve_list_indent(&original);
            }

            let mut marker = self.run_properties(style, fc, prm, prcs)?;
            let mut baseline = self.paragraph_base(style)?;
            if linked != 0xfff {
                // Resolve the linked style's toggles against its own base
                // chain once, then overlay its explicit visible properties.
                // Reapplying that chain to an already styled mark would
                // invert relative toggles twice and inject a default size.
                let id = usize::from(linked);
                let patch = if let Some(patch) = self.paragraph_marker_styles.get(&id) {
                    patch.clone()
                } else {
                    let mut patch = Properties::sparse();
                    self.apply_style(&mut patch, id, 1)?;
                    self.paragraph_marker_styles.insert(id, patch.clone());
                    patch
                };
                marker.overlay_visible(&patch);
                baseline.overlay_visible(&patch);
            }
            let mut character_style = baseline.clone();
            self.apply_direct(&mut marker, &mut character_style, &baseline, level.chpx)?;
            let ppr = props.xml();
            let rpr = marker.xml(&self.fonts)?;
            let id = self.numbering_output.activate(
                &self.numbering,
                reference,
                ppr,
                rpr,
                super::MAX_DOCUMENT_XML_BYTES,
            )?;
            props.numbering = id.map(|id| (id, reference.level));
        }
        Ok(props.xml())
    }

    fn apply_paragraph<'b>(
        &mut self,
        props: &mut paragraph::Properties,
        bytes: &'b [u8],
    ) -> Result<DirectParagraphProperties, String>
    where
        'a: 'b,
    {
        let mut direct = DirectParagraphProperties::default();
        sprm::paragraph_properties(bytes, self.data, &mut self.budget, |code, operand| {
            if (code >> 10) & 7 == 1
                && !matches!(code, 0x2416 | 0x2417 | 0x6649 | 0x664a | 0x244b | 0x244c)
                && !props.apply(code, operand)?
            {
                self.unsupported_paragraph_properties = true;
            }
            if code == 0x2441 {
                // Properties::apply validated this Bool8 operand above.
                direct.bidi = Some(operand[0] == 1);
            } else if matches!(code, 0x2403 | 0x2461) {
                // Properties::apply validated the alignment enum above.
                direct.alignment = Some((code, operand[0]));
            } else if matches!(code, 0x840e | 0x840f | 0x845d | 0x845e | 0x8411 | 0x8460) {
                // Properties::apply validated the signed XAS operand above.
                direct.push_absolute_indent((code, u16_at(operand, 0)? as i16));
            }
            Ok(())
        })?;
        Ok(direct)
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

    pub fn passive_special_character(
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
            paragraph_marker_styles: BTreeMap::new(),
            paragraph_layout_cache: BTreeMap::new(),
            data: &[],
            budget: Budget::default(),
            numbering: numbering::Tables::default(),
            numbering_output: numbering::output::Store::default(),
            unsupported_character_properties: false,
            unsupported_paragraph_properties: false,
            unsupported_piece_properties: false,
            missing_tables: true,
            unsupported_table_properties: false,
        }
    }

    fn with_direct_paragraph(properties: &[u8]) -> Formatting<'_> {
        let mut word = vec![0u8; 1024];
        let table: Vec<u8> = [100u32, 110, 1]
            .into_iter()
            .flat_map(u32::to_le_bytes)
            .collect();
        word[0x106..0x10a].copy_from_slice(&12u32.to_le_bytes());
        let page = &mut word[512..1024];
        page[0..4].copy_from_slice(&100u32.to_le_bytes());
        page[4..8].copy_from_slice(&110u32.to_le_bytes());
        if !properties.is_empty() {
            page[8] = 32;
            if properties.len() % 2 == 0 {
                page[64] = 0;
                page[65] = (properties.len() / 2) as u8;
                page[66..66 + properties.len()].copy_from_slice(properties);
            } else {
                page[64] = ((properties.len() + 1) / 2) as u8;
                page[65..65 + properties.len()].copy_from_slice(properties);
            }
        }
        page[511] = 1;
        // The index borrows its inputs, so leak this small test fixture.
        Formatting::read(
            Box::leak(word.into_boxed_slice()),
            Box::leak(table.into_boxed_slice()),
            &[],
        )
        .unwrap()
    }

    fn level_bidi_formatting(value: u8) -> numbering::Tables<'static> {
        numbering::Tables {
            lists: vec![numbering::List {
                id: 42,
                styles: [0xfff; 9],
                simple: true,
                hybrid: false,
                auto_number: false,
                levels: vec![numbering::Level {
                    start: Some(1),
                    format: 0,
                    justification: 0,
                    legal: false,
                    restart: Some(0),
                    follow: 0,
                    tentative: false,
                    papx: Box::leak(vec![0x41, 0x24, value, 0x03, 0x24, 2].into_boxed_slice()),
                    chpx: &[],
                    text: &[0, 0, b'.', 0],
                    placeholders: [Some((1, 0)), None, None, None, None, None, None, None, None],
                }],
            }],
            overrides: vec![numbering::Override {
                list_index: 0,
                first_cp: None,
                auto_number_field: None,
                levels: vec![],
            }],
        }
    }

    fn level_paragraph_formatting(papx: Vec<u8>) -> numbering::Tables<'static> {
        let mut tables = level_bidi_formatting(0);
        tables.lists[0].levels[0].papx = Box::leak(papx.into_boxed_slice());
        tables
    }

    fn list_piece(ilfo: i16) -> Vec<u8> {
        [vec![0x0b, 0x46], ilfo.to_le_bytes().to_vec()].concat()
    }

    #[test]
    fn direct_absolute_indent_axes_replay_after_list_in_last_write_order() {
        let mut f = with_direct_paragraph(&[
            0, 0, 0x41, 0x24, 1, // direct RTL
            0x0e, 0x84, 0xd0, 0x02, // physical right 720 => RTL logical left
            0x5e, 0x84, 0x68, 0x01, // later logical left 360 wins
            0x0f, 0x84, 0xf0, 0x00, // physical left 240 => RTL logical right
            0x5d, 0x84, 0xe0, 0x01, // later logical right 480 wins
        ]);
        f.numbering =
            level_paragraph_formatting(vec![0x5e, 0x84, 0xd0, 0x02, 0x5d, 0x84, 0xd0, 0x02]);
        let piece = list_piece(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        assert!(
            xml.contains("<w:ind w:left=\"360\" w:right=\"480\""),
            "{xml}"
        );
    }

    #[test]
    fn direct_zero_and_first_line_replay_but_absent_axis_inherits_list() {
        let mut f = with_direct_paragraph(&[
            0, 0, 0x0f, 0x84, 0, 0, // explicit physical-left zero
            0x60, 0x84, 0x68, 0x01, // explicit first line +360
        ]);
        f.numbering = level_paragraph_formatting(vec![
            0x5e, 0x84, 0xd0, 0x02, // list left 720
            0x5d, 0x84, 0x68, 0x01, // list right 360 (must remain: absent direct)
            0x60, 0x84, 0x98, 0xfe, // list hanging 360
        ]);
        let piece = list_piece(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        assert!(
            xml.contains("<w:ind w:left=\"0\" w:right=\"360\" w:firstLine=\"360\""),
            "{xml}"
        );
    }

    #[test]
    fn ltr_physical_logical_and_first_line_use_latest_direct_writes() {
        let mut f = with_direct_paragraph(&[
            0, 0, 0x0f, 0x84, 0xd0, 0x02, // physical left 720
            0x5e, 0x84, 0xa0, 0x05, // later logical left 1440
            0x11, 0x84, 0x68, 0x01, // first80 +360
            0x60, 0x84, 0x98, 0xfe, // later first -360
        ]);
        f.numbering = level_paragraph_formatting(vec![0x5e, 0x84, 0x68, 0x01, 0x60, 0x84, 0, 0]);
        let piece = list_piece(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        assert!(xml.contains("<w:ind w:left=\"1440\""), "{xml}");
        assert!(xml.contains("w:hanging=\"360\""), "{xml}");
    }

    #[test]
    fn pcd_absolute_indent_is_the_later_direct_layer() {
        let mut f = with_direct_paragraph(&[0, 0, 0x5e, 0x84, 0xd0, 0x02]);
        f.numbering = level_paragraph_formatting(vec![0x5e, 0x84, 0x68, 0x01]);
        let piece = [list_piece(1), vec![0x5e, 0x84, 0xa0, 0x05]].concat();
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        assert!(xml.contains("<w:ind w:left=\"1440\""), "{xml}");
    }

    #[test]
    fn repeated_same_axis_writes_never_evict_an_unchanged_opposite_axis() {
        let mut bytes = vec![
            0, 0, 0x41, 0x24, 1, // RTL
            0x0e, 0x84, 0, 0, // physical right zero => RTL logical left
        ];
        for value in 1_i16..=7 {
            bytes.extend([0x0f, 0x84]); // repeatedly replace physical left
            bytes.extend(value.to_le_bytes());
        }
        let mut f = with_direct_paragraph(&bytes);
        f.numbering =
            level_paragraph_formatting(vec![0x5e, 0x84, 0xd0, 0x02, 0x5d, 0x84, 0xd0, 0x02]);
        let piece = list_piece(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        assert!(xml.contains("<w:ind w:left=\"0\" w:right=\"7\""), "{xml}");
    }

    #[test]
    fn all_six_codes_survive_when_pcd_repeats_one_papx_code() {
        let mut f = with_direct_paragraph(&[
            0, 0, 0x0e, 0x84, 10, 0, 0x0f, 0x84, 20, 0, 0x5d, 0x84, 30, 0, 0x5e, 0x84, 40, 0, 0x11,
            0x84, 50, 0, 0x60, 0x84, 60, 0,
        ]);
        f.numbering = level_paragraph_formatting(vec![0x5e, 0x84, 0xd0, 0x02]);
        let piece = [list_piece(1), vec![0x0e, 0x84, 70, 0]].concat();
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        // PCD's repeated physical-right code moves last without dropping the
        // other five distinct codes; logical-left remains later for LTR left.
        assert!(xml.contains("<w:ind w:left=\"40\" w:right=\"70\""), "{xml}");
        assert!(xml.contains("w:firstLine=\"60\""), "{xml}");
    }

    #[test]
    fn negative_list_reference_preservation_stays_authoritative() {
        let mut f = with_direct_paragraph(&[0, 0, 0x5e, 0x84, 0xd0, 0x02, 0x60, 0x84, 0x98, 0xfe]);
        f.numbering =
            level_paragraph_formatting(vec![0x5e, 0x84, 0xa0, 0x05, 0x60, 0x84, 0x68, 0x01]);
        let piece = list_piece(-1);
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        assert!(xml.contains("<w:ind w:left=\"720\""), "{xml}");
        assert!(xml.contains("w:hanging=\"360\""), "{xml}");
    }

    #[test]
    fn explicit_direct_bidi_survives_list_level_formatting() {
        let mut f = with_direct_paragraph(&[0, 0, 0x41, 0x24, 0]);
        f.numbering = level_bidi_formatting(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]]).unwrap();
        assert!(xml.contains("<w:bidi w:val=\"0\"/>"));
    }

    #[test]
    fn explicit_direct_bidi_and_physical_alignment_survive_list_level_formatting() {
        let mut f = with_direct_paragraph(&[0, 0, 0x41, 0x24, 0, 0x03, 0x24, 0]);
        f.numbering = level_bidi_formatting(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]]).unwrap();
        assert!(xml.contains("<w:bidi w:val=\"0\"/>"));
        assert!(xml.contains("<w:jc w:val=\"left\"/>"));
    }

    #[test]
    fn direct_bidi_does_not_protect_absent_alignment_from_the_list_level() {
        let mut f = with_direct_paragraph(&[0, 0, 0x41, 0x24, 0]);
        f.numbering = level_bidi_formatting(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]]).unwrap();
        assert!(xml.contains("<w:bidi w:val=\"0\"/>"));
        assert!(xml.contains("<w:jc w:val=\"right\"/>"));
    }

    #[test]
    fn direct_alignment_is_independent_and_uses_its_last_explicit_write() {
        let mut f = with_direct_paragraph(&[0, 0, 0x03, 0x24, 2, 0x61, 0x24, 0]);
        f.numbering = level_bidi_formatting(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]]).unwrap();
        assert!(xml.contains("<w:bidi w:val=\"1\"/>"));
        // The later logical PJc left replaces the earlier physical PJc80 right.
        assert!(xml.contains("<w:jc w:val=\"left\"/>"));
    }

    #[test]
    fn direct_bidi_controls_are_explicit_last_write_and_piece_override() {
        let mut explicit_true = with_direct_paragraph(&[0, 0, 0x41, 0x24, 1]);
        explicit_true.numbering = level_bidi_formatting(0);
        assert!(explicit_true
            .paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]])
            .unwrap()
            .contains("<w:bidi w:val=\"1\"/>"));

        let mut absent = with_direct_paragraph(&[]);
        absent.numbering = level_bidi_formatting(1);
        assert!(absent
            .paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]])
            .unwrap()
            .contains("<w:bidi w:val=\"1\"/>"));

        let mut sequential = with_direct_paragraph(&[0, 0, 0x41, 0x24, 1, 0x41, 0x24, 0]);
        sequential.numbering = level_bidi_formatting(1);
        assert!(sequential
            .paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]])
            .unwrap()
            .contains("<w:bidi w:val=\"0\"/>"));

        let mut piece_override = with_direct_paragraph(&[0, 0, 0x41, 0x24, 0]);
        piece_override.numbering = level_bidi_formatting(0);
        assert!(piece_override
            .paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0, 0x41, 0x24, 1]])
            .unwrap()
            .contains("<w:bidi w:val=\"1\"/>"));
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
    fn list_marker_style_patches_do_not_toggle_twice_or_inject_default_sizes() {
        for (chpx, bold) in [(&[][..], "1"), (&[0x35, 0x08, 0x81][..], "0")] {
            let mut f = empty();
            f.styles = vec![Some(Style {
                kind: 1,
                base: 0xfff,
                chpx: &[0x35, 0x08, 0x81],
                papx: &[],
            })];
            f.numbering = numbering::Tables {
                lists: vec![numbering::List {
                    id: 42,
                    styles: [0; 9],
                    simple: true,
                    hybrid: false,
                    auto_number: false,
                    levels: vec![numbering::Level {
                        start: Some(1),
                        format: 0,
                        justification: 0,
                        legal: false,
                        restart: Some(0),
                        follow: 0,
                        tentative: false,
                        papx: &[],
                        chpx,
                        text: &[0, 0, b'.', 0],
                        placeholders: [
                            Some((1, 0)),
                            None,
                            None,
                            None,
                            None,
                            None,
                            None,
                            None,
                            None,
                        ],
                    }],
                }],
                overrides: vec![numbering::Override {
                    list_index: 0,
                    first_cp: None,
                    auto_number_field: None,
                    levels: vec![],
                }],
            };
            let piece: &[u8] = &[0x0b, 0x46, 1, 0, 0x43, 0x4a, 40, 0];
            for _ in 0..2 {
                f.paragraph_xml(0, 0, 1, &[piece]).unwrap();
            }
            let xml = f.numbering_output.xml(10000).unwrap().unwrap();
            assert!(xml.contains(&format!("<w:b w:val=\"{bold}\"/>")));
            assert!(xml.contains("<w:sz w:val=\"40\"/>"));
            assert!(!xml.contains("<w:sz w:val=\"20\"/>"));
            assert_eq!(xml.matches("<w:num w:numId=").count(), 1);
            assert!(f
                .run_xml(0, 0, 1, &[piece])
                .unwrap()
                .contains("<w:b w:val=\"1\"/>"));
        }
    }

    #[test]
    fn paragraph_border_style_cascade_and_piece_resets_do_not_mutate_the_cache() {
        let mut f = empty();
        f.styles = vec![
            Some(Style {
                kind: 1,
                base: 0xfff,
                chpx: &[],
                papx: &[0x24, 0x64, 8, 1, 2, 0, 0x26, 0x64, 8, 1, 2, 0],
            }),
            Some(Style {
                kind: 1,
                base: 0,
                chpx: &[],
                papx: &[0x50, 0xc6, 8, 0xff, 0, 0, 0, 16, 3, 0, 0],
            }),
        ];
        let before = f.paragraph_xml(1, 0, 0, &[]).unwrap();
        assert!(before.contains("<w:top w:val=\"single\""));
        assert!(before.contains("<w:bottom w:val=\"double\""));
        let cleared = f
            .paragraph_xml(1, 0, 1, &[&[0x50, 0xc6, 8, 0, 0, 0, 0xff, 0, 0, 0, 0]])
            .unwrap();
        assert!(cleared.contains("<w:top w:val=\"single\""));
        assert!(cleared.contains("<w:bottom w:val=\"none\""));
        assert_eq!(before, f.paragraph_xml(1, 0, 0, &[]).unwrap());
        assert!(!f.unsupported_paragraph_properties);
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
