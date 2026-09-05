//! Direct PowerPoint text runs: [MS-PPT] 2.9.14/20/41/44–46.
use super::*;

pub(super) fn write(
    text: &str,
    style: &[u8],
    fonts: &[String],
    output: &mut String,
    xml_budget: &mut usize,
    work_budget: &mut usize,
) -> Result<(), String> {
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
        drawing::append(output, xml_budget, &pf[pi].1.xml()?)?;
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
                    &cf[run].1.xml("rPr", fonts)?,
                    output,
                    xml_budget,
                )?;
                start = end;
            }
        }
        while cf[ci].0 <= cp {
            ci += 1;
        }
        drawing::append(output, xml_budget, &cf[ci].1.xml("endParaRPr", fonts)?)?;
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
    fn xml(&self, tag: &str, fonts: &[String]) -> Result<String, String> {
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
        if let Some(color) = self.color.filter(|c| c >> 24 == 0xfe) {
            children.push_str(&format!(
                "<a:solidFill><a:srgbClr val=\"{:02X}{:02X}{:02X}\"/></a:solidFill>",
                color & 255,
                (color >> 8) & 255,
                (color >> 16) & 255
            ));
        }
        for (id, name) in [(self.font, "latin"), (self.ea, "ea"), (self.symbol, "sym")] {
            if let Some(id) = id {
                let font = fonts
                    .get(usize::from(id))
                    .ok_or_else(|| unsupported("PowerPoint font index out of range"))?;
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

struct Paragraph {
    level: u16,
    align: Option<u16>,
    spacing: [Option<i16>; 3],
    no_bullet: bool,
}
impl Paragraph {
    fn read(r: &mut Reader<'_, '_>, level: u16) -> Result<Self, String> {
        let mask = r.u32()?;
        if mask & 0x03800000 != 0 {
            return Err(unsupported(
                "extended PowerPoint paragraph style in base run",
            ));
        }
        let bullet_flags = r.optional16(mask, 15)?;
        r.optional16(mask, 0x80)?; // bulletChar
        r.optional16(mask, 0x10)?; // bulletFontRef
        r.optional16(mask, 0x40)?; // bulletSize
        if mask & 0x20 != 0 {
            r.u32()?;
        } // bulletColor
        let align = r.optional16(mask, 0x800)?;
        let mut spacing = [None; 3];
        for (value, flag) in spacing.iter_mut().zip([0x1000, 0x2000, 0x4000]) {
            *value = r.optional16(mask, flag)?.map(|n| n as i16);
        }
        for flag in [0x100, 0x400, 0x8000] {
            r.optional16(mask, flag)?;
        }
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
            no_bullet: mask & 1 != 0 && bullet_flags.is_some_and(|v| v & 1 == 0),
        })
    }
    fn xml(&self) -> Result<String, String> {
        let mut xml = format!("<a:pPr lvl=\"{}\"", self.level);
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
        if self.no_bullet {
            xml.push_str("<a:buNone/>");
        }
        xml.push_str("</a:pPr>");
        Ok(xml)
    }
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
            fonts,
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
        assert!(write("ab", &data, &[], &mut output, &mut 1024, &mut 1)
            .unwrap_err()
            .contains("work budget"));
        assert!(write("ab", &data, &[], &mut output, &mut 8, &mut 10)
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
        assert!(xml("x", &data, &["Name".into()])
            .unwrap_err()
            .contains("font index"));
    }
}
