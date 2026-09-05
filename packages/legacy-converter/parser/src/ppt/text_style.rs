//! Direct PowerPoint text runs: [MS-PPT] 2.9.14/20/41/44–46.
use super::*;
mod bullet;

#[derive(Default, Clone, Copy)]
pub(super) struct Context<'a> {
    pub fonts: &'a [String],
    pub scheme: Option<&'a scheme::Scheme>,
    pub levels: Option<&'a [Level]>,
}
pub(super) fn text_type(atom: Record<'_>) -> Result<u16, String> {
    if atom.version != 0 || atom.payload.len() != 4 {
        return Err(unsupported("invalid PowerPoint text header"));
    }
    let kind = u32_at(atom.payload, 0)?;
    if !matches!(kind, 0..=2 | 4..=8) {
        return Err(unsupported("invalid PowerPoint text type"));
    }
    Ok(kind as u16)
}

struct Runs {
    paragraphs: Vec<(usize, Paragraph)>,
    characters: Vec<(usize, Character)>,
}

fn read_runs(text: &str, style: &[u8], work_budget: &mut usize) -> Result<Runs, String> {
    // TextHeaderAtom adds an implicit CR; run counts include that character.
    let length = text.encode_utf16().count() + 1;
    let mut reader = Reader {
        bytes: style,
        pos: 0,
        budget: work_budget,
    };
    let mut pf = Vec::new();
    let mut cf = Vec::new();
    let mut end = 0;
    while end < length {
        end = reader.run_end(end, length)?;
        let level = reader.u16()?;
        if level > 4 {
            return Err(unsupported("invalid PowerPoint paragraph level"));
        }
        pf.push((end, Paragraph::read(&mut reader, level)?));
    }
    end = 0;
    while end < length {
        end = reader.run_end(end, length)?;
        cf.push((end, Character::read(&mut reader)?));
    }
    if reader.pos != style.len() {
        return Err(unsupported("unexpected PowerPoint text style tail"));
    }
    // Validate shared boundaries once, including master exemplars that are not
    // themselves emitted. Never accept half of a UTF-16 surrogate pair.
    let (mut cp, mut run) = (0, 0);
    for character in text.chars() {
        cp += character.len_utf16();
        if cf[run].0 < cp {
            return Err(unsupported(
                "PowerPoint character style splits a surrogate pair",
            ));
        }
        if cf[run].0 == cp {
            run += 1;
        }
    }
    Ok(Runs {
        paragraphs: pf,
        characters: cf,
    })
}

pub(super) fn write(
    text: &str,
    style: &[u8],
    context: Context<'_>,
    output: &mut String,
    xml_budget: &mut usize,
    work_budget: &mut usize,
) -> Result<(), String> {
    let Runs {
        paragraphs: pf,
        characters: cf,
    } = read_runs(text, style, work_budget)?;
    let (mut pi, mut ci, mut cp) = (0, 0, 0);
    for paragraph in text.split('\r') {
        while pf[pi].0 <= cp {
            pi += 1;
        }
        let para_end = cp + paragraph.encode_utf16().count() + 1;
        if pf[pi].0 < para_end {
            return Err(unsupported("PowerPoint paragraph style splits a paragraph"));
        }
        drawing::append(output, xml_budget, "<a:p>")?;
        let base = context
            .levels
            .and_then(|levels| levels.get(usize::from(pf[pi].1.level)));
        drawing::append(
            output,
            xml_budget,
            &pf[pi].1.inherit(base.map(|v| &v.paragraph)).xml(context)?,
        )?;
        let mut start = 0;
        let mut iter = paragraph.char_indices().peekable();
        while let Some((offset, c)) = iter.next() {
            while cf[ci].0 <= cp {
                ci += 1;
            }
            let run = ci;
            cp += c.len_utf16();
            if cp > cf[run].0 {
                return Err(unsupported(
                    "PowerPoint character style splits a surrogate pair",
                ));
            }
            if cp == cf[run].0 || iter.peek().is_none() {
                let end = offset + c.len_utf8();
                write_run(
                    &paragraph[start..end],
                    &cf[run].1.inherit(base.map(|v| &v.character)).xml(
                        "rPr",
                        context.fonts,
                        context.scheme,
                    )?,
                    output,
                    xml_budget,
                )?;
                start = end;
            }
        }
        while cf[ci].0 <= cp {
            ci += 1;
        }
        drawing::append(
            output,
            xml_budget,
            &cf[ci].1.inherit(base.map(|v| &v.character)).xml(
                "endParaRPr",
                context.fonts,
                context.scheme,
            )?,
        )?;
        drawing::append(output, xml_budget, "</a:p>")?;
        cp += 1;
    }
    Ok(())
}

pub(super) fn write_run(
    text: &str,
    properties: &str,
    output: &mut String,
    budget: &mut usize,
) -> Result<(), String> {
    // Unicode UAX #14 BK/LF: VT, LF and LINE SEPARATOR force line breaks,
    // not new paragraphs. DrawingML CT_TextLineBreak preserves that distinction.
    // MS-PPT TextHeaderAtom assigns CR the separate paragraph-mark role.
    for (index, part) in text.split(['\u{b}', '\n', '\u{2028}']).enumerate() {
        if index != 0 {
            drawing::append(output, budget, "<a:br>")?;
            drawing::append(output, budget, properties)?;
            drawing::append(output, budget, "</a:br>")?;
        }
        if !part.is_empty() {
            drawing::append(output, budget, "<a:r>")?;
            drawing::append(output, budget, properties)?;
            drawing::append(output, budget, "<a:t>")?;
            escaped(part, output, budget)?;
            drawing::append(output, budget, "</a:t></a:r>")?;
        }
    }
    Ok(())
}

pub(super) fn escaped(text: &str, output: &mut String, budget: &mut usize) -> Result<(), String> {
    let mut remaining = text;
    while !remaining.is_empty() {
        let mut end = remaining.len().min(1024);
        while !remaining.is_char_boundary(end) {
            end -= 1;
        }
        drawing::append(output, budget, &xml_text(&remaining[..end]))?;
        remaining = &remaining[end..];
    }
    Ok(())
}

struct Reader<'a, 'b> {
    bytes: &'a [u8],
    pos: usize,
    budget: &'b mut usize,
}
impl Reader<'_, '_> {
    fn u16(&mut self) -> Result<u16, String> {
        let n = u16_at(self.bytes, self.pos)?;
        self.pos += 2;
        Ok(n)
    }
    fn u32(&mut self) -> Result<u32, String> {
        let n = u32_at(self.bytes, self.pos)?;
        self.pos += 4;
        Ok(n)
    }
    fn run_end(&mut self, start: usize, length: usize) -> Result<usize, String> {
        *self.budget = self
            .budget
            .checked_sub(1)
            .ok_or_else(|| unsupported("PowerPoint text style work budget exceeded"))?;
        let count = self.u32()? as usize;
        start
            .checked_add(count)
            .filter(|end| count != 0 && *end <= length)
            .ok_or_else(|| unsupported("invalid PowerPoint text run count"))
    }
    fn optional16(&mut self, mask: u32, bit: u32) -> Result<Option<u16>, String> {
        if mask & bit != 0 {
            Ok(Some(self.u16()?))
        } else {
            Ok(None)
        }
    }
}

#[derive(Clone, PartialEq)]
struct Character {
    mask: u32,
    style: u16,
    size: u16,
    font: Option<u16>,
    ea: Option<u16>,
    symbol: Option<u16>,
    color: Option<u32>,
}
impl Character {
    fn inherit(&self, base: Option<&Self>) -> Self {
        let Some(base) = base else {
            return self.clone();
        };
        Self {
            mask: self.mask | base.mask,
            style: (base.style & !(self.mask as u16)) | (self.style & self.mask as u16),
            size: if self.mask & 0x20000 != 0 {
                self.size
            } else {
                base.size
            },
            font: self.font.or(base.font),
            ea: self.ea.or(base.ea),
            symbol: self.symbol.or(base.symbol),
            color: self.color.or(base.color),
        }
    }
    fn read(r: &mut Reader<'_, '_>) -> Result<Self, String> {
        let mask = r.u32()?;
        if mask & 0x07100000 != 0 {
            return Err(unsupported(
                "extended PowerPoint character style in base run",
            ));
        }
        let style = r.optional16(mask, 0x3eb7)?.unwrap_or(0);
        let font = r.optional16(mask, 0x10000)?;
        let ea = r.optional16(mask, 0x200000)?;
        let _ansi = r.optional16(mask, 0x400000)?;
        let symbol = r.optional16(mask, 0x800000)?;
        let size = r.optional16(mask, 0x20000)?.unwrap_or(18);
        if !(1..=4000).contains(&size) {
            return Err(unsupported("invalid PowerPoint font size"));
        }
        let color = if mask & 0x40000 != 0 {
            Some(r.u32()?)
        } else {
            None
        };
        let _position = r.optional16(mask, 0x80000)?;
        Ok(Self {
            mask,
            style,
            size,
            font,
            ea,
            symbol,
            color,
        })
    }
    fn xml(
        &self,
        tag: &str,
        fonts: &[String],
        scheme: Option<&scheme::Scheme>,
    ) -> Result<String, String> {
        let mut xml = format!("<a:{tag} sz=\"{}\"", u32::from(self.size) * 100);
        for (bit, name) in [(1, "b"), (2, "i")] {
            if self.mask & bit != 0 {
                xml.push_str(&format!(
                    " {name}=\"{}\"",
                    u8::from(self.style & bit as u16 != 0)
                ));
            }
        }
        if self.mask & 4 != 0 {
            xml.push_str(if self.style & 4 != 0 {
                " u=\"sng\""
            } else {
                " u=\"none\""
            });
        }
        let mut children = String::new();
        if let Some(color) = self.color.and_then(|c| scheme::text(c, scheme)) {
            children.push_str(&format!(
                "<a:solidFill><a:srgbClr val=\"{:02X}{:02X}{:02X}\"/></a:solidFill>",
                color & 255,
                (color >> 8) & 255,
                (color >> 16) & 255
            ));
        }
        for (id, name) in [(self.font, "latin"), (self.ea, "ea"), (self.symbol, "sym")] {
            if let Some(id) = id {
                // An unsupported font reference must not invent a font or
                // discard otherwise usable inherited text properties.
                let Some(font) = fonts.get(usize::from(id)) else {
                    continue;
                };
                children.push_str(&format!(
                    "<a:{name} typeface=\"{}\"/>",
                    crate::ooxml::xml_attr(font)
                ));
            }
        }
        if children.is_empty() {
            xml.push_str("/>");
        } else {
            xml.push('>');
            xml.push_str(&children);
            xml.push_str(&format!("</a:{tag}>"));
        }
        Ok(xml)
    }
}

#[derive(Clone, PartialEq)]
struct Paragraph {
    level: u16,
    align: Option<u16>,
    spacing: [Option<i16>; 3],
    bullet: bullet::Bullet,
    margin: Option<i16>,
    indent: Option<i16>,
}
impl Paragraph {
    fn inherit(&self, base: Option<&Self>) -> Self {
        let Some(base) = base else {
            return self.clone();
        };
        Self {
            level: self.level,
            align: self.align.or(base.align),
            spacing: std::array::from_fn(|i| self.spacing[i].or(base.spacing[i])),
            bullet: self.bullet.inherit(&base.bullet),
            margin: self.margin.or(base.margin),
            indent: self.indent.or(base.indent),
        }
    }
    fn read(r: &mut Reader<'_, '_>, level: u16) -> Result<Self, String> {
        let mask = r.u32()?;
        if mask & 0x03800000 != 0 {
            return Err(unsupported(
                "extended PowerPoint paragraph style in base run",
            ));
        }
        let bullet = bullet::Bullet::read(r, mask)?;
        let align = r.optional16(mask, 0x800)?;
        let mut spacing = [None; 3];
        for (value, flag) in spacing.iter_mut().zip([0x1000, 0x2000, 0x4000]) {
            *value = r.optional16(mask, flag)?.map(|n| n as i16);
        }
        let margin = r.optional16(mask, 0x100)?.map(|v| v as i16);
        let indent = r.optional16(mask, 0x400)?.map(|v| v as i16);
        r.optional16(mask, 0x8000)?;
        if mask & 0x100000 != 0 {
            let count = usize::from(r.u16()?);
            *r.budget = r
                .budget
                .checked_sub(count)
                .ok_or_else(|| unsupported("PowerPoint tab work budget exceeded"))?;
            for _ in 0..count {
                r.u32()?;
            }
        }
        r.optional16(mask, 0x10000)?;
        r.optional16(mask, 0xe0000)?;
        r.optional16(mask, 0x200000)?;
        Ok(Self {
            level,
            align,
            spacing,
            bullet,
            margin,
            indent,
        })
    }
    fn xml(&self, context: Context<'_>) -> Result<String, String> {
        let mut xml = format!("<a:pPr lvl=\"{}\"", self.level);
        // Binary text/bullet offsets share a text-body origin. DrawingML indent
        // is relative to marL, so retain their difference (including hanging).
        // A negative binary margin cannot be expressed by ST_TextMargin; omit
        // it and its dependent first-line offset rather than clamp the layout.
        if let Some(margin) = self
            .margin
            .map(|m| master_to_emu(i64::from(m)))
            .filter(|m| (0..=51_206_400).contains(m))
        {
            xml.push_str(&format!(" marL=\"{margin}\""));
            if let Some(indent) = self.indent {
                let indent = master_to_emu(i64::from(indent)) - margin;
                if (-51_206_400..=51_206_400).contains(&indent) {
                    xml.push_str(&format!(" indent=\"{indent}\""));
                }
            }
        }
        if let Some(align) = self.align {
            let value = ["l", "ctr", "r", "just", "dist", "thaiDist", "justLow"]
                .get(usize::from(align))
                .ok_or_else(|| unsupported("invalid PowerPoint text alignment"))?;
            xml.push_str(&format!(" algn=\"{value}\""));
        }
        xml.push('>');
        for (value, tag) in self.spacing.iter().zip(["lnSpc", "spcBef", "spcAft"]) {
            if let Some(n) = value {
                // MS-PPT ParaSpacing: >=0 percent; <0 master units (1/8 pt).
                let (kind, value) = if *n >= 0 {
                    ("spcPct", i32::from(*n) * 1000)
                } else {
                    ("spcPts", (-i32::from(*n) * 100 + 4) / 8)
                };
                if (kind == "spcPct" && value > 13200000) || (kind == "spcPts" && value > 158400) {
                    return Err(unsupported("PowerPoint spacing exceeds DrawingML range"));
                }
                xml.push_str(&format!("<a:{tag}><a:{kind} val=\"{value}\"/></a:{tag}>"));
            }
        }
        xml.push_str(&self.bullet.xml(context));
        xml.push_str("</a:pPr>");
        Ok(xml)
    }
}

#[derive(Clone, PartialEq)]
pub(super) struct Level {
    paragraph: Paragraph,
    character: Character,
}
impl Level {
    pub fn inherit(&self, base: Option<&Self>) -> Self {
        Self {
            paragraph: self.paragraph.inherit(base.map(|b| &b.paragraph)),
            character: self.character.inherit(base.map(|b| &b.character)),
        }
    }
    pub fn empty(level: u16) -> Self {
        Self {
            paragraph: Paragraph {
                level,
                align: None,
                spacing: [None; 3],
                bullet: bullet::Bullet::default(),
                margin: None,
                indent: None,
            },
            character: Character {
                mask: 0,
                style: 0,
                size: 18,
                font: None,
                ea: None,
                symbol: None,
                color: None,
            },
        }
    }
}

/// Supported master-shape subset: uniform formatting within each indent level.
/// Conflicting exemplar runs do not justify choosing an arbitrary first run.
pub(super) fn shape_levels(
    text: &str,
    style: &[u8],
    budget: &mut usize,
) -> Result<Vec<Option<Level>>, String> {
    let Runs {
        paragraphs: pf,
        characters: cf,
    } = read_runs(text, style, budget)?;
    let mut levels = vec![None; 5];
    let mut conflict = [false; 5];
    let (mut pi, mut ci, mut cp) = (0, 0, 0);
    for paragraph in text.split('\r') {
        while pf[pi].0 <= cp {
            pi += 1;
        }
        let end = cp + paragraph.encode_utf16().count() + 1;
        if pf[pi].0 < end {
            return Err(unsupported("PowerPoint master style splits a paragraph"));
        }
        let level = usize::from(pf[pi].1.level);
        while cp < end {
            while cf[ci].0 <= cp {
                ci += 1;
            }
            let value = Level {
                paragraph: pf[pi].1.clone(),
                character: cf[ci].1.clone(),
            };
            if levels[level].as_ref().is_some_and(|old| old != &value) {
                conflict[level] = true;
            }
            levels[level] = Some(value);
            cp = cf[ci].0.min(end);
        }
    }
    for (value, conflicting) in levels.iter_mut().zip(conflict) {
        if conflicting {
            *value = None;
        }
    }
    Ok(levels)
}
pub(super) struct Master {
    types: std::collections::BTreeMap<u16, Vec<Level>>,
    defaults: Vec<Level>,
}
impl Master {
    pub fn parse(
        records: &[Record<'_>],
        defaults: &[Level],
        budget: &mut usize,
    ) -> Result<Self, String> {
        let mut types = std::collections::BTreeMap::new();
        for atom in records.iter().filter(|r| r.kind == 4003) {
            if types.contains_key(&atom.instance) {
                return Err(unsupported("duplicate PowerPoint master text type"));
            }
            let mut levels = read_levels(*atom, budget)?;
            for (i, level) in levels.iter_mut().enumerate() {
                level.paragraph = level
                    .paragraph
                    .inherit(defaults.get(i).map(|v| &v.paragraph));
                level.character = level
                    .character
                    .inherit(defaults.get(i).map(|v| &v.character));
            }
            levels.extend(defaults.iter().skip(levels.len()).cloned());
            types.insert(atom.instance, levels);
        }
        Ok(Self {
            types,
            defaults: defaults.to_vec(),
        })
    }
    pub fn levels(&self, kind: u16) -> Option<&[Level]> {
        self.types
            .get(&kind)
            .map(Vec::as_slice)
            .or_else(|| (!self.defaults.is_empty()).then_some(self.defaults.as_slice()))
    }
}
fn read_levels(atom: Record<'_>, budget: &mut usize) -> Result<Vec<Level>, String> {
    if atom.version != 0 || !matches!(atom.instance, 0..=2 | 4..=8) {
        return Err(unsupported("invalid PowerPoint master text style"));
    }
    let mut r = Reader {
        bytes: atom.payload,
        pos: 0,
        budget,
    };
    let count = usize::from(r.u16()?);
    if count > 5 {
        return Err(unsupported("too many PowerPoint master text levels"));
    }
    *r.budget = r
        .budget
        .checked_sub(count)
        .ok_or_else(|| unsupported("PowerPoint master text work budget exceeded"))?;
    let mut levels: Vec<Option<Level>> = vec![None; count];
    for index in 0..count {
        // MS-PPT 2.9.36: only text types >=5 include an explicit level field.
        let level = if atom.instance >= 5 {
            usize::from(r.u16()?)
        } else {
            index
        };
        if level >= count || levels[level].is_some() {
            return Err(unsupported("invalid PowerPoint master text level"));
        }
        levels[level] = Some(Level {
            paragraph: Paragraph::read(&mut r, level as u16)?,
            character: Character::read(&mut r)?,
        });
    }
    if r.pos != atom.payload.len() {
        return Err(unsupported("unexpected PowerPoint master text style tail"));
    }
    Ok(levels
        .into_iter()
        .map(|v| v.expect("every level assigned"))
        .collect())
}
pub(super) fn document_defaults(
    children: &[Record<'_>],
    budget: &mut usize,
) -> Result<Vec<Level>, String> {
    let mut defaults = None;
    for env in children
        .iter()
        .filter(|r| r.kind == 1010 && r.version == 15)
    {
        for atom in parse_records(env.payload, budget)?
            .iter()
            .filter(|r| r.kind == 4003)
        {
            if defaults.is_some() {
                return Err(unsupported("duplicate PowerPoint document text defaults"));
            }
            defaults = Some(read_levels(*atom, budget)?);
        }
    }
    Ok(defaults.unwrap_or_default())
}

pub(super) fn default_style(text: &str) -> Vec<u8> {
    let count = (text.encode_utf16().count() + 1) as u32;
    [
        count.to_le_bytes().to_vec(),
        vec![0; 6],
        count.to_le_bytes().to_vec(),
        vec![0; 4],
    ]
    .concat()
}

pub(super) fn fonts(children: &[Record<'_>], budget: &mut usize) -> Result<Vec<String>, String> {
    let mut fonts = Vec::new();
    for env in children
        .iter()
        .filter(|r| r.kind == 1010 && r.version == 15)
    {
        for collection in parse_records(env.payload, budget)?
            .iter()
            .filter(|r| r.kind == 2005 && r.version == 15)
        {
            for entity in parse_records(collection.payload, budget)?
                .iter()
                .filter(|r| r.kind == 4023)
            {
                if entity.version != 0 || entity.payload.len() != 68 || fonts.len() >= 65536 {
                    return Err(unsupported("invalid PowerPoint font entity"));
                }
                let units: Vec<u16> = entity.payload[..64]
                    .chunks_exact(2)
                    .map(|v| u16::from_le_bytes([v[0], v[1]]))
                    .take_while(|n| *n != 0)
                    .collect();
                if units.len() >= 32 {
                    return Err(unsupported("unterminated PowerPoint font name"));
                }
                fonts.push(String::from_utf16_lossy(&units));
            }
        }
    }
    Ok(fonts)
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn preserves_character_bullet_properties_and_hanging_indent() {
        // MS-PPT TextPFException: explicit bullet flags, glyph, font, point size,
        // RGB color, text offset and absolute first-line/bullet offset.
        let data = [
            u32s(2),
            u16s(0),
            u32s(0x5ff),
            u16s(15),
            u16s('&' as u16),
            u16s(0),
            u16s((-12i16) as u16),
            u32s(0xfe332211),
            u16s(144),
            u16s(0),
            u32s(2),
            u32s(0),
        ]
        .concat();
        let output = xml("X", &data, &["Bullet & Font".into()]).unwrap();
        assert!(output.contains("marL=\"228600\" indent=\"-228600\""));
        assert!(output.contains("<a:buClr><a:srgbClr val=\"112233\"/></a:buClr>"));
        assert!(output.contains("<a:buSzPts val=\"1200\"/>"));
        assert!(output.contains("<a:buFont typeface=\"Bullet &amp; Font\"/>"));
        assert!(output.contains("<a:buChar char=\"&amp;\"/>"));
    }
    #[test]
    fn paragraph_offsets_merge_independently_and_obey_ooxml_ranges() {
        let mut base = Level::empty(0).paragraph;
        base.margin = Some(144);
        base.indent = Some(0);
        let mut direct = Level::empty(0).paragraph;
        direct.margin = Some(288);
        let output = direct.inherit(Some(&base)).xml(Context::default()).unwrap();
        assert!(output.contains("marL=\"457200\" indent=\"-457200\""));
        direct.indent = Some(432);
        assert!(direct
            .xml(Context::default())
            .unwrap()
            .contains("indent=\"228600\""));
        for invalid in [-1, 32767] {
            direct.margin = Some(invalid);
            let output = direct.xml(Context::default()).unwrap();
            assert!(!output.contains("marL="));
            assert!(!output.contains("indent="));
        }
        direct.margin = Some(30000);
        direct.indent = Some(-32768);
        let output = direct.xml(Context::default()).unwrap();
        assert!(output.contains("marL="));
        assert!(!output.contains("indent="));
        direct.margin = None;
        assert!(!direct.xml(Context::default()).unwrap().contains("indent="));
    }
    #[test]
    fn master_shape_levels_keep_uniform_styles_without_selecting_arbitrary_runs() {
        let style = [
            u32s(4),
            u16s(0),
            u32s(0),
            u32s(4),
            u32s(0x60001),
            u16s(1),
            u16s(36),
            u32s(0xfeffffff),
        ]
        .concat();
        let levels = shape_levels("one", &style, &mut 100).unwrap();
        let level = levels[0].as_ref().unwrap();
        assert_eq!(level.character.color, Some(0xfeffffff));
        assert_eq!(level.character.size, 36);
        let mixed = [
            u32s(4),
            u16s(0),
            u32s(0),
            u32s(2),
            u32s(0x40000),
            u32s(0xfeffffff),
            u32s(2),
            u32s(0x40000),
            u32s(0xfe000000),
        ]
        .concat();
        assert!(shape_levels("one", &mixed, &mut 100).unwrap()[0].is_none());
        assert!(shape_levels("one", &style, &mut 1).is_err());
        // Two paragraphs at separate levels retain separate colors.
        let distinct = [
            u32s(2),
            u16s(0),
            u32s(0),
            u32s(2),
            u16s(1),
            u32s(0),
            u32s(2),
            u32s(0x40000),
            u32s(0xfeffffff),
            u32s(2),
            u32s(0x40000),
            u32s(0xfe000000),
        ]
        .concat();
        let levels = shape_levels("a\rb", &distinct, &mut 100).unwrap();
        assert_eq!(
            levels[0].as_ref().unwrap().character.color,
            Some(0xfeffffff)
        );
        assert_eq!(
            levels[1].as_ref().unwrap().character.color,
            Some(0xfe000000)
        );
    }
    #[test]
    fn master_defaults_merge_by_level_and_direct_false_overrides_true() {
        // One title level: centered, bold, 48pt. Direct formatting turns bold
        // off and changes only alignment; absent size must not become 18pt.
        let data = [
            u16s(1),
            u32s(0x800),
            u16s(1),
            u32s(0x20001),
            u16s(1),
            u16s(48),
        ]
        .concat();
        let record = Record {
            version: 0,
            instance: 0,
            kind: 4003,
            payload: &data,
        };
        let master = Master::parse(&[record], &[], &mut 100).unwrap();
        let direct = [
            u32s(2),
            u16s(0),
            u32s(0x800),
            u16s(2),
            u32s(2),
            u32s(1),
            u16s(0),
        ]
        .concat();
        let mut output = String::new();
        write(
            "X",
            &direct,
            Context {
                levels: master.levels(0),
                ..Context::default()
            },
            &mut output,
            &mut 4096,
            &mut 100,
        )
        .unwrap();
        assert!(output.contains("sz=\"4800\" b=\"0\""));
        assert!(output.contains("algn=\"r\""));
        assert!(!output.contains("1800"));
        let local = [u16s(1), u32s(0), u32s(0x20000), u16s(32)].concat();
        let merged = Master::parse(
            &[Record {
                payload: &local,
                ..record
            }],
            master.levels(0).unwrap(),
            &mut 100,
        )
        .unwrap();
        let base = &merged.levels(0).unwrap()[0];
        assert_eq!(base.character.size, 32);
        assert_eq!(base.character.style & 1, 1);
        assert_eq!(base.paragraph.align, Some(1));
    }

    #[test]
    fn master_style_level_ids_are_present_only_for_extended_text_types() {
        let data = [
            u16s(2),
            u16s(1),
            u32s(0),
            u32s(0x20000),
            u16s(28),
            u16s(0),
            u32s(0),
            u32s(0x20000),
            u16s(44),
        ]
        .concat();
        let mut atom = Record {
            version: 0,
            instance: 6,
            kind: 4003,
            payload: &data,
        };
        let levels = read_levels(atom, &mut 100).unwrap();
        assert_eq!(levels[0].character.size, 44);
        assert_eq!(levels[1].character.size, 28);
        assert!(read_levels(atom, &mut 1).is_err());
        atom.instance = 0;
        assert!(read_levels(atom, &mut 100).is_err());
        for bad in [
            vec![6, 0],
            vec![1, 0],
            [u16s(1), u16s(1), u32s(0), u32s(0)].concat(),
        ] {
            assert!(read_levels(
                Record {
                    instance: 6,
                    payload: &bad,
                    ..atom
                },
                &mut 100
            )
            .is_err());
        }
    }
    fn u16s(n: u16) -> Vec<u8> {
        n.to_le_bytes().to_vec()
    }
    fn u32s(n: u32) -> Vec<u8> {
        n.to_le_bytes().to_vec()
    }
    fn style(count: u32, cf: Vec<u8>) -> Vec<u8> {
        [u32s(count), u16s(0), u32s(0), u32s(count), cf].concat()
    }
    fn xml(text: &str, data: &[u8], fonts: &[String]) -> Result<String, String> {
        let mut output = String::new();
        write(
            text,
            data,
            Context {
                fonts,
                ..Context::default()
            },
            &mut output,
            &mut (1024 * 1024),
            &mut MAX_RECORDS.clone(),
        )?;
        Ok(output)
    }
    #[test]
    fn direct_font_size_styles_color_and_name_are_not_replaced_with_defaults() {
        let data = style(
            3,
            [u32s(0x70007), u16s(3), u16s(0), u16s(36), u32s(0xfe563412)].concat(),
        );
        let out = xml("ab", &data, &["A & B\"".into()]).unwrap();
        assert!(out.contains("sz=\"3600\" b=\"1\" i=\"1\" u=\"none\""));
        assert!(out.contains("val=\"123456\""));
        assert!(out.contains("typeface=\"A &amp; B&quot;\""));
    }
    #[test]
    fn counts_utf16_and_retains_implicit_paragraph_mark_style() {
        let data = [
            u32s(4),
            u16s(0),
            u32s(0),
            u32s(2),
            u32s(0x20000),
            u16s(40),
            u32s(2),
            u32s(0x20000),
            u16s(20),
        ]
        .concat();
        let out = xml("😀x", &data, &[]).unwrap();
        assert!(out.contains("sz=\"4000\"/><a:t>😀</a:t>"));
        assert!(out.contains("sz=\"2000\"/><a:t>x</a:t>"));
        assert!(out.contains("<a:endParaRPr sz=\"2000\"/>"));
    }
    #[test]
    fn paragraph_alignment_and_signed_spacing_use_their_own_units() {
        let data = [
            u32s(2),
            u16s(1),
            u32s(0x7800),
            u16s(1),
            u16s(120),
            u16s((-96i16) as u16),
            u16s(50),
            u32s(2),
            u32s(0),
        ]
        .concat();
        let out = xml("x", &data, &[]).unwrap();
        assert!(out.contains("lvl=\"1\" algn=\"ctr\""));
        assert!(out.contains("<a:lnSpc><a:spcPct val=\"120000\"/></a:lnSpc>"));
        assert!(out.contains("<a:spcBef><a:spcPts val=\"1200\"/></a:spcBef>"));
        assert!(out.contains("<a:spcAft><a:spcPct val=\"50000\"/></a:spcAft>"));
    }
    #[test]
    fn rejects_zero_overrun_truncated_and_surrogate_splitting_runs() {
        assert!(xml("a", &style(0, u32s(0)), &[]).is_err());
        assert!(xml("a", &style(3, u32s(0)), &[]).is_err());
        assert!(xml("a", &style(2, u32s(0x20000)), &[]).is_err());
        let split = [
            u32s(3),
            u16s(0),
            u32s(0),
            u32s(1),
            u32s(0),
            u32s(2),
            u32s(0),
        ]
        .concat();
        assert!(xml("😀", &split, &[]).is_err());
    }

    #[test]
    fn line_breaks_remain_inside_the_paragraph_and_keep_character_style() {
        for text in ["a\u{b}b", "\u{b}ab", "ab\u{b}", "a\nb", "a\u{2028}b"] {
            let data = style(4, [u32s(0x20000), u16s(36)].concat());
            let out = xml(text, &data, &[]).unwrap();
            assert_eq!(out.matches("<a:p>").count(), 1);
            assert_eq!(out.matches("<a:br><a:rPr sz=\"3600\"/></a:br>").count(), 1);
            assert!(!out.contains('\u{fffd}'));
        }
        let data = style(5, u32s(0));
        let out = xml("a\r\u{b}b", &data, &[]).unwrap();
        assert_eq!(out.matches("<a:p>").count(), 2);
        assert_eq!(out.matches("<a:br>").count(), 1);
    }

    #[test]
    fn style_work_and_expanded_xml_have_independent_budgets() {
        let data = style(3, u32s(0));
        let mut output = String::new();
        assert!(write(
            "ab",
            &data,
            Context::default(),
            &mut output,
            &mut 1024,
            &mut 1
        )
        .unwrap_err()
        .contains("work budget"));
        assert!(write(
            "ab",
            &data,
            Context::default(),
            &mut output,
            &mut 8,
            &mut 10
        )
        .unwrap_err()
        .contains("OUTPUT_TOO_LARGE"));
    }

    #[test]
    fn font_collection_indexes_entities_without_embedding_font_data() {
        use super::super::persist::tests::record;
        let mut font = vec![0; 68];
        font[..8].copy_from_slice(&[b'N', 0, b'a', 0, b'm', 0, b'e', 0]);
        let env = record(
            15,
            1010,
            &record(
                15,
                2005,
                &[
                    record(0, 4023, &font),
                    record(0, 4024, b"opaque embedded bytes"),
                ]
                .concat(),
            ),
        );
        let children = parse_records(&env, &mut 100).unwrap();
        assert_eq!(fonts(&children, &mut 100).unwrap(), ["Name"]);
        let data = style(2, [u32s(0x10000), u16s(1)].concat());
        let output = xml("x", &data, &["Name".into()]).unwrap();
        assert!(output.contains("<a:t>x</a:t>"));
        assert!(!output.contains("typeface=")); // Omit invalid references; never guess an index.
    }
}
