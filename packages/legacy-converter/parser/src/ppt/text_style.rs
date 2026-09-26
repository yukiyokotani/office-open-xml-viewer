//! Direct PowerPoint text runs: [MS-PPT] 2.9.14/20/41/44–46.
use super::*;
pub(super) mod auto_number;
mod bullet;
pub(super) mod direct_model;
pub(super) mod master_chain;

#[derive(Default, Clone, Copy)]
pub(super) struct Context<'a> {
    pub fonts: &'a [String],
    pub scheme: Option<&'a scheme::Scheme>,
    pub levels: Option<&'a [Level]>,
    pub slide_numbers: &'a [u32],
    pub slide_number: u32,
    pub ruler_tabs: Option<ruler::Tabs<'a>>,
    pub style9: Option<&'a [u8]>,
    /// Direct model only: record a glyph effect the model cannot express
    /// here instead of rejecting, for a caller with an alternative source.
    pub deferred_effect: Option<&'a std::cell::Cell<Option<&'static str>>>,
}

#[derive(Clone, Copy, Default, PartialEq, Eq, Debug)]
pub(super) struct ParagraphAxes {
    pub margin: Option<i16>,
    pub indent: Option<i16>,
}

/// Per-level origins of the document `Tx_TYPE_OTHER` master style, indexed
/// by IndentLevel. A level absent from the record keeps both fields absent.
pub(super) type DocumentAxes = [ParagraphAxes; 5];

pub(super) struct DocumentDefaults {
    pub levels: Vec<Level>,
    pub type4_axes: Option<DocumentAxes>,
}

/// MS-PPT 2.9.47 / 2.2.30: this is a passive positional substitution, not
/// evaluation of arbitrary fields or replacement of literal asterisks.
pub(super) fn slide_number_position(atom: Record<'_>) -> Result<u32, String> {
    if atom.version != 0 || atom.instance != 0 || atom.payload.len() != 4 {
        return Err(unsupported("invalid PowerPoint slide-number atom"));
    }
    let position = u32_at(atom.payload, 0)?;
    if position > i32::MAX as u32 {
        return Err(unsupported("negative PowerPoint slide-number position"));
    }
    Ok(position)
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
        // MS-PPT 2.9.14: signed percentage of line height, in [-100, 100].
        // Validate even though projection remains unsupported: DrawingML
        // CT_TextCharacterProperties/@baseline is relative to font size, so
        // multiplying by 1000 alone is not a justified conversion rule.
        if r.optional16(mask, 0x80000)?
            .is_some_and(|value| !(-100..=100).contains(&(value as i16)))
        {
            return Err(unsupported(
                "invalid PowerPoint character baseline position",
            ));
        }
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
}

#[derive(Clone, PartialEq)]
struct Paragraph {
    level: u16,
    rtl: Option<bool>,
    align: Option<u16>,
    spacing: [Option<i16>; 3],
    bullet: bullet::Bullet,
    margin: Option<i16>,
    indent: Option<i16>,
    default_tab: Option<i16>,
}
impl Paragraph {
    fn inherit(&self, base: Option<&Self>) -> Self {
        let Some(base) = base else {
            return self.clone();
        };
        Self {
            level: self.level,
            rtl: self.rtl.or(base.rtl),
            align: self.align.or(base.align),
            spacing: std::array::from_fn(|i| self.spacing[i].or(base.spacing[i])),
            bullet: self.bullet.inherit(&base.bullet),
            margin: self.margin.or(base.margin),
            indent: self.indent.or(base.indent),
            default_tab: self.default_tab.or(base.default_tab),
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
        let default_tab = r.optional16(mask, 0x8000)?.map(|v| v as i16);
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
        // MS-PPT 2.9.20 / 2.13.30 TextDirectionEnum. Absence inherits;
        // explicit LeftToRight must clear a right-to-left master value.
        let rtl = match r.optional16(mask, 0x200000)? {
            None => None,
            Some(0) => Some(false),
            Some(1) => Some(true),
            Some(_) => return Err(unsupported("invalid PowerPoint text direction")),
        };
        Ok(Self {
            level,
            rtl,
            align,
            spacing,
            bullet,
            margin,
            indent,
            default_tab,
        })
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
                rtl: None,
                align: None,
                spacing: [None; 3],
                bullet: bullet::Bullet::default(),
                margin: None,
                indent: None,
                default_tab: None,
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
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct AuthoredFontSizes {
    level_count: u8,
    values: [Option<u16>; 5],
}
impl AuthoredFontSizes {
    #[cfg(test)]
    pub fn level_count(self) -> usize {
        usize::from(self.level_count)
    }
    #[cfg(test)]
    pub fn get(self, level: usize) -> Option<u16> {
        (level < self.level_count())
            .then(|| self.values[level])
            .flatten()
    }
}
pub(super) type AuthoredFontSizeTable = [Option<AuthoredFontSizes>; 9];
pub(super) struct Master {
    // One fixed nine-slot table per Master covers the eight admitted
    // TextTypeEnum values and five levels, with no second Level vector or
    // per-paragraph metadata allocation.
    types: std::collections::BTreeMap<u16, Vec<Level>>,
    authored_font_sizes: std::rc::Rc<AuthoredFontSizeTable>,
    defaults: Vec<Level>,
    /// Unmerged atom levels for the direct model's level-chain resolution.
    raw: std::collections::BTreeMap<u16, Vec<Level>>,
}
impl Master {
    pub fn parse(
        records: &[Record<'_>],
        defaults: &[Level],
        budget: &mut usize,
    ) -> Result<Self, String> {
        let mut types = std::collections::BTreeMap::new();
        let mut raw = std::collections::BTreeMap::new();
        let mut authored_types: AuthoredFontSizeTable = [None; 9];
        for atom in records.iter().filter(|r| r.kind == 4003) {
            if types.contains_key(&atom.instance) {
                return Err(unsupported("duplicate PowerPoint master text type"));
            }
            let mut levels = read_levels(*atom, budget)?;
            let authored_font_sizes = AuthoredFontSizes {
                level_count: levels.len() as u8,
                values: std::array::from_fn(|index| {
                    levels.get(index).and_then(|level| {
                        (level.character.mask & 0x20000 != 0).then_some(level.character.size)
                    })
                }),
            };
            authored_types[usize::from(atom.instance)] = Some(authored_font_sizes);
            raw.insert(atom.instance, levels.clone());
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
            authored_font_sizes: std::rc::Rc::new(authored_types),
            defaults: defaults.to_vec(),
            raw,
        })
    }
    /// Direct-model levels for text of `kind` (see [`master_chain`]).
    pub fn direct_levels(&self, kind: u16) -> Option<master_chain::DirectLevels> {
        let own = self.raw.get(&kind);
        let base = master_chain::base_type(kind).and_then(|b| self.raw.get(&b));
        if own.is_none() && base.is_none() && self.defaults.is_empty() {
            return None;
        }
        Some(master_chain::resolve(
            own.map_or(&[], Vec::as_slice),
            base.map_or(&[], Vec::as_slice),
            &self.defaults,
        ))
    }
    /// Direct-model levels of the document Tx_TYPE_OTHER atom alone.
    pub fn document_levels(&self) -> Option<master_chain::DirectLevels> {
        (!self.defaults.is_empty()).then(|| master_chain::resolve(&[], &[], &self.defaults))
    }
    pub fn levels(&self, kind: u16) -> Option<&[Level]> {
        self.types
            .get(&kind)
            .map(Vec::as_slice)
            .or_else(|| (!self.defaults.is_empty()).then_some(self.defaults.as_slice()))
    }
    #[cfg(test)]
    pub fn authored_font_sizes(&self, kind: u16) -> Option<AuthoredFontSizes> {
        self.authored_font_sizes
            .get(usize::from(kind))
            .copied()
            .flatten()
    }
    pub fn authored_font_size_table(&self) -> std::rc::Rc<AuthoredFontSizeTable> {
        self.authored_font_sizes.clone()
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
) -> Result<DocumentDefaults, String> {
    let mut defaults = None;
    let mut type4_axes = None;
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
            let levels = read_levels(*atom, budget)?;
            if atom.instance == 4 && !levels.is_empty() {
                type4_axes = Some(std::array::from_fn(|index| {
                    levels
                        .get(index)
                        .map(|level| ParagraphAxes {
                            margin: level.paragraph.margin,
                            indent: level.paragraph.indent,
                        })
                        .unwrap_or_default()
                }));
            }
            defaults = Some(levels);
        }
    }
    Ok(DocumentDefaults {
        levels: defaults.unwrap_or_default(),
        type4_axes,
    })
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
    use pptx_model::{Bullet as ModelBullet, Paragraph as ModelParagraph, TextRun};

    /// Project through the direct text model. Paragraph origins the tested
    /// level does not supply come from zero document axes, which only fill
    /// absent fields.
    fn project(
        text: &str,
        style: &[u8],
        context: Context<'_>,
    ) -> Result<Vec<ModelParagraph>, String> {
        let zero = ParagraphAxes {
            margin: Some(0),
            indent: Some(0),
        };
        direct_model::paragraphs_with_axes(
            text,
            style,
            context,
            direct_model::DirectAxes {
                document: Some([zero; 5]),
                ..Default::default()
            },
            &mut MAX_RECORDS.clone(),
            &mut (1024 * 1024),
        )
    }

    /// The text runs of one paragraph as (text, font size).
    fn runs(paragraph: &ModelParagraph) -> Vec<(&str, Option<f64>)> {
        paragraph
            .runs
            .iter()
            .filter_map(|run| match run {
                TextRun::Text(run) => Some((run.text.as_str(), run.font_size)),
                _ => None,
            })
            .collect()
    }

    fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let options = (instance << 4) | version;
        [
            options.to_le_bytes().as_slice(),
            &kind.to_le_bytes(),
            &(payload.len() as u32).to_le_bytes(),
            payload,
        ]
        .concat()
    }

    #[test]
    fn document_defaults_retain_type4_paragraph_axes_by_level() {
        let absent = document_defaults(&[], &mut 100).unwrap();
        assert!(absent.levels.is_empty());
        assert_eq!(absent.type4_axes, None);

        let axes = |margin, indent| ParagraphAxes { margin, indent };
        // Level 0 carries both fields, level 1 only its indent and level 2
        // only its margin; levels 3-4 are absent from the record.
        let levels = [
            u16s(3),
            u32s(0x500),
            u16s(180),
            u16s(90),
            u32s(0),
            u32s(0x400),
            u16s(300),
            u32s(0),
            u32s(0x100),
            u16s(576),
            u32s(0),
        ]
        .concat();
        let master = record_bytes(0, 4, 4003, &levels);
        let environment = Record {
            version: 15,
            instance: 0,
            kind: 1010,
            payload: &master,
        };
        let defaults = document_defaults(&[environment], &mut 100).unwrap();
        assert_eq!(defaults.levels.len(), 3);
        assert_eq!(
            defaults.type4_axes,
            Some([
                axes(Some(180), Some(90)),
                axes(None, Some(300)),
                axes(Some(576), None),
                axes(None, None),
                axes(None, None),
            ])
        );

        let title = record_bytes(0, 0, 4003, &levels);
        let environment = Record {
            payload: &title,
            ..environment
        };
        assert_eq!(
            document_defaults(&[environment], &mut 100)
                .unwrap()
                .type4_axes,
            None
        );

        let zero_level = [u16s(1), u32s(0x500), u16s(0), u16s(0), u32s(0)].concat();
        let zero_master = record_bytes(0, 4, 4003, &zero_level);
        let zero_environment = Record {
            payload: &zero_master,
            ..environment
        };
        assert_eq!(
            document_defaults(&[zero_environment], &mut 100)
                .unwrap()
                .type4_axes
                .map(|levels| levels[0]),
            Some(axes(Some(0), Some(0)))
        );
        let empty = record_bytes(0, 4, 4003, &u16s(0));
        let empty_environment = Record {
            payload: &empty,
            ..environment
        };
        assert_eq!(
            document_defaults(&[empty_environment], &mut 100)
                .unwrap()
                .type4_axes,
            None
        );
    }

    fn paragraph_master(properties: &[u8], defaults: &[Level]) -> Master {
        // MS-PPT 2.9.35-36: one title master level, followed by its empty CF.
        let payload = [u16s(1), properties.to_vec(), u32s(0)].concat();
        Master::parse(
            &[Record {
                version: 0,
                instance: 0,
                kind: 4003,
                payload: &payload,
            }],
            defaults,
            &mut 100,
        )
        .unwrap()
    }

    fn inherited_paragraph(master: &Master) -> ModelParagraph {
        project(
            "X",
            &default_style("X"),
            Context {
                levels: master.levels(0),
                ..Context::default()
            },
        )
        .unwrap()
        .remove(0)
    }

    #[test]
    fn text_direction_preserves_inheritance_and_explicit_left_to_right() {
        let read = |value: u16| {
            let bytes = [u32s(0x200000), u16s(value)].concat();
            Paragraph::read(
                &mut Reader {
                    bytes: &bytes,
                    pos: 0,
                    budget: &mut 100,
                },
                0,
            )
        };
        let rtl = Level {
            paragraph: read(1).unwrap(),
            ..Level::empty(0)
        };
        let direction = |value: Option<u16>, levels: Option<&[Level]>| {
            let pf = value.map_or_else(
                || [u32s(0)].concat(),
                |v| [u32s(0x200000), u16s(v)].concat(),
            );
            let style = [u32s(2), u16s(0), pf, u32s(2), u32s(0)].concat();
            let paragraph = project(
                "X",
                &style,
                Context {
                    levels,
                    ..Context::default()
                },
            )
            .unwrap()
            .remove(0);
            (paragraph.rtl, paragraph.alignment)
        };
        let master = Some(std::slice::from_ref(&rtl));
        // Right-to-left text without an alignment aligns right (ECMA-376
        // 21.1.2.2.7 keeps direction independent of an explicit alignment).
        assert_eq!(direction(Some(1), None), (true, "r".to_owned()));
        assert_eq!(direction(None, master), (true, "r".to_owned()));
        assert_eq!(direction(Some(0), master), (false, "l".to_owned()));
        assert_eq!(direction(None, None), (false, "l".to_owned()));
        for value in [2, 255, 32767, 65535] {
            assert!(read(value).is_err());
        }
    }

    #[test]
    fn preserves_default_tab_size_units_and_explicit_zero_overrides() {
        // These fields belong in TextMasterStyleLevel, not TextPFRun (2.9.45).
        for value in [i16::MIN, -1, 0, 1, 288, 575, 576, i16::MAX] {
            let bytes = [u32s(0x8000), value.to_le_bytes().to_vec()].concat();
            let master = paragraph_master(&bytes, &[]);
            assert_eq!(
                inherited_paragraph(&master).def_tab_sz,
                Some(master_to_emu(i64::from(value)))
            );
        }
        let base_bytes = [u32s(0x8000), u16s(288)].concat();
        let base = paragraph_master(&base_bytes, &[]);
        let zero_bytes = [u32s(0x8000), u16s(0)].concat();
        let zero = paragraph_master(&zero_bytes, base.levels(0).unwrap());
        let absent = paragraph_master(&u32s(0), base.levels(0).unwrap());
        assert_eq!(inherited_paragraph(&zero).def_tab_sz, Some(0));
        assert_eq!(inherited_paragraph(&absent).def_tab_sz, Some(457200));
        assert_eq!(
            inherited_paragraph(&paragraph_master(&u32s(0), &[])).def_tab_sz,
            None
        );
    }

    #[test]
    fn preserves_character_bullet_properties_and_hanging_indent() {
        // MS-PPT 2.9.44-45: direct bullets can inherit offsets from a master,
        // but the offsets themselves cannot be stored in a TextPFRun.
        let master = paragraph_master(&[u32s(0x500), u16s(144), u16s(0)].concat(), &[]);
        let data = [
            u32s(2),
            u16s(0),
            u32s(0xff),
            u16s(15),
            u16s('&' as u16),
            u16s(0),
            u16s((-12i16) as u16),
            u32s(0xfe332211),
            u32s(2),
            u32s(0),
        ]
        .concat();
        // MS-PPT 2.9.45 forbids ruler fields inside a TextPFRun.
        assert_eq!(u32_at(&data, 6).unwrap() & 0x108500, 0);
        let paragraph = project(
            "X",
            &data,
            Context {
                fonts: &["Bullet & Font".into()],
                levels: master.levels(0),
                ..Context::default()
            },
        )
        .unwrap()
        .remove(0);
        assert_eq!((paragraph.mar_l, paragraph.indent), (228600, -228600));
        assert!(matches!(
            paragraph.bullet,
            ModelBullet::Char { ref ch, ref color, size_pct: None, size_pts: Some(12.0), ref font_family }
                if ch == "&" && color.as_deref() == Some("112233")
                    && font_family.as_deref() == Some("Bullet & Font")
        ));
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
        let paragraph = project(
            "X",
            &direct,
            Context {
                levels: master.levels(0),
                ..Context::default()
            },
        )
        .unwrap()
        .remove(0);
        assert_eq!(paragraph.alignment, "r");
        let TextRun::Text(run) = &paragraph.runs[0] else {
            panic!("text run")
        };
        assert_eq!((run.font_size, run.bold), (Some(48.0), Some(false)));
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
    fn master_retains_authored_type_level_count_and_explicit_sizes_before_defaults_merge() {
        let default_bytes = [
            u16s(2),
            u32s(0),
            u32s(0x20000),
            u16s(18),
            u32s(0),
            u32s(0x20000),
            u16s(16),
        ]
        .concat();
        let defaults = read_levels(
            Record {
                version: 0,
                instance: 4,
                kind: 4003,
                payload: &default_bytes,
            },
            &mut 100,
        )
        .unwrap();
        let authored_bytes = [
            u16s(2),
            u16s(0),
            u32s(0),
            u32s(0x20000),
            u16s(20),
            u16s(1),
            u32s(0),
            u32s(0),
        ]
        .concat();
        let master = Master::parse(
            &[Record {
                version: 0,
                instance: 8,
                kind: 4003,
                payload: &authored_bytes,
            }],
            &defaults,
            &mut 100,
        )
        .unwrap();
        let authored = master.authored_font_sizes(8).unwrap();
        assert_eq!(authored.level_count(), 2);
        assert_eq!(authored.get(0), Some(20));
        assert_eq!(authored.get(1), None);
        assert_eq!(authored.get(2), None);
        assert_eq!(authored.get(usize::MAX), None);
        assert!(master.authored_font_sizes(7).is_none());
        assert_eq!(master.levels(8).unwrap()[1].character.size, 16);
        assert_eq!(master.levels(7).unwrap()[0].character.size, 18);
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
    #[test]
    fn slide_numbers_keep_original_utf16_style_boundaries_and_literal_text() {
        let text = "😀*\r*Z";
        let data = [
            u32s(4),
            u16s(0),
            u32s(0),
            u32s(3),
            u16s(0),
            u32s(0),
            u32s(3),
            u32s(0x20000),
            u16s(40),
            u32s(4),
            u32s(0x20000),
            u16s(20),
        ]
        .concat();
        let context = |slide_numbers| Context {
            slide_numbers,
            slide_number: 42,
            ..Context::default()
        };
        let model = project(text, &data, context(&[2, 4])).unwrap();
        assert_eq!(model.len(), 2);
        assert_eq!(runs(&model[0]), [("😀", Some(40.0)), ("42", Some(40.0))]);
        assert_eq!(runs(&model[1]), [("42", Some(20.0)), ("Z", Some(20.0))]);
        for positions in [&[1][..], &[3], &[6], &[7], &[2, 2], &[4, 2]] {
            assert!(
                project(text, &data, context(positions)).is_err(),
                "{positions:?}"
            );
        }
    }

    #[test]
    fn slide_number_atoms_reject_malformed_headers_lengths_and_negative_positions() {
        let valid = [0u8; 4];
        let atom = Record {
            kind: 4056,
            version: 0,
            instance: 0,
            payload: &valid,
        };
        assert_eq!(slide_number_position(atom).unwrap(), 0);
        for invalid in [
            Record { version: 1, ..atom },
            Record {
                instance: 1,
                ..atom
            },
            Record {
                payload: &valid[..3],
                ..atom
            },
            Record {
                payload: &[255; 4],
                ..atom
            },
        ] {
            assert!(slide_number_position(invalid).is_err());
        }
    }
    fn u32s(n: u32) -> Vec<u8> {
        n.to_le_bytes().to_vec()
    }
    fn style(count: u32, cf: Vec<u8>) -> Vec<u8> {
        [u32s(count), u16s(0), u32s(0), u32s(count), cf].concat()
    }
    fn model(text: &str, data: &[u8], fonts: &[String]) -> Result<Vec<ModelParagraph>, String> {
        project(
            text,
            data,
            Context {
                fonts,
                ..Context::default()
            },
        )
    }
    fn first_run(paragraphs: &[ModelParagraph]) -> &pptx_model::TextRunData {
        match &paragraphs[0].runs[0] {
            TextRun::Text(run) => run,
            _ => panic!("text run"),
        }
    }
    #[test]
    fn direct_font_size_styles_color_and_name_are_not_replaced_with_defaults() {
        let data = style(
            3,
            [u32s(0x70007), u16s(3), u16s(0), u16s(36), u32s(0xfe563412)].concat(),
        );
        let paragraphs = model("ab", &data, &["A & B\"".into()]).unwrap();
        let run = first_run(&paragraphs);
        assert_eq!(run.font_size, Some(36.0));
        assert_eq!(
            (run.bold, run.italic, run.underline),
            (Some(true), Some(true), false)
        );
        assert_eq!(run.color.as_deref(), Some("123456"));
        assert_eq!(run.font_family.as_deref(), Some("A & B\""));
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
        let paragraphs = model("😀x", &data, &[]).unwrap();
        assert_eq!(
            runs(&paragraphs[0]),
            [("😀", Some(40.0)), ("x", Some(20.0))]
        );
        // The implicit terminal CR's style sizes an empty final paragraph.
        let data = [
            u32s(3),
            u16s(0),
            u32s(0),
            u32s(2),
            u32s(0x20000),
            u16s(40),
            u32s(1),
            u32s(0x20000),
            u16s(20),
        ]
        .concat();
        let paragraphs = model("a\r", &data, &[]).unwrap();
        assert_eq!(paragraphs[0].def_font_size, None);
        assert_eq!(paragraphs[1].def_font_size, Some(20.0));
    }

    #[test]
    fn character_position_accepts_only_the_signed_percentage_domain() {
        // MS-PPT 2.9.14: position is signed and MUST be within [-100, 100].
        // Its line-height reference does not establish an OOXML font-size
        // percentage mapping. This test validates input, not superscript paint.
        for position in i16::MIN..=i16::MAX {
            let data = style(2, [u32s(0x80000), u16s(position as u16)].concat());
            let result = model("x", &data, &[]);
            assert_eq!(
                result.is_ok(),
                (-100..=100).contains(&position),
                "{position}"
            );
            if let Ok(paragraphs) = result {
                assert_eq!(first_run(&paragraphs).baseline, None);
            }
        }
        assert!(model("x", &style(2, u32s(0x80000)), &[]).is_err());
        assert!(model("x", &style(2, [u32s(0x80000), vec![0]].concat()), &[]).is_err());
    }

    #[test]
    fn master_character_position_uses_the_same_validation_as_direct_runs() {
        for position in [-101i16, -100, 0, 100, 101] {
            // TextMasterStyleAtom: one level, empty PF, CF with position only.
            let data = [u16s(1), u32s(0), u32s(0x80000), u16s(position as u16)].concat();
            let result = Master::parse(
                &[Record {
                    version: 0,
                    instance: 0,
                    kind: 4003,
                    payload: &data,
                }],
                &[],
                &mut 100,
            );
            assert_eq!(
                result.is_ok(),
                (-100..=100).contains(&position),
                "{position}"
            );
        }
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
        let paragraph = model("x", &data, &[]).unwrap().remove(0);
        assert_eq!((paragraph.lvl, paragraph.alignment.as_str()), (1, "ctr"));
        assert!(matches!(
            paragraph.space_line,
            Some(ooxml_common::text::SpaceLine::Pct { val }) if val == 120_000.0
        ));
        assert_eq!(
            (paragraph.space_before, paragraph.space_before_pct),
            (Some(1200), None)
        );
        assert_eq!(
            (paragraph.space_after, paragraph.space_after_pct),
            (None, Some(50_000.0))
        );
    }
    #[test]
    fn rejects_zero_overrun_truncated_and_surrogate_splitting_runs() {
        assert!(model("a", &style(0, u32s(0)), &[]).is_err());
        assert!(model("a", &style(3, u32s(0)), &[]).is_err());
        assert!(model("a", &style(2, u32s(0x20000)), &[]).is_err());
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
        assert!(model("😀", &split, &[]).is_err());
    }

    #[test]
    fn line_breaks_remain_inside_the_paragraph_and_keep_character_style() {
        let breaks = |paragraph: &ModelParagraph| {
            paragraph
                .runs
                .iter()
                .filter(|run| matches!(run, TextRun::Break))
                .count()
        };
        for text in ["a\u{b}b", "\u{b}ab", "ab\u{b}", "a\nb", "a\u{2028}b"] {
            let data = style(4, [u32s(0x20000), u16s(36)].concat());
            let paragraphs = model(text, &data, &[]).unwrap();
            assert_eq!(paragraphs.len(), 1);
            assert_eq!(breaks(&paragraphs[0]), 1);
            let runs = runs(&paragraphs[0]);
            assert!(runs
                .iter()
                .all(|(text, size)| !text.contains('\u{fffd}') && *size == Some(36.0)));
            assert_eq!(runs.iter().map(|(text, _)| *text).collect::<String>(), "ab");
        }
        let data = style(5, u32s(0));
        let paragraphs = model("a\r\u{b}b", &data, &[]).unwrap();
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(breaks(&paragraphs[0]) + breaks(&paragraphs[1]), 1);
    }

    #[test]
    fn style_work_and_model_output_have_independent_budgets() {
        let data = style(3, u32s(0));
        let origin = [Level {
            paragraph: Paragraph {
                margin: Some(0),
                indent: Some(0),
                ..Level::empty(0).paragraph
            },
            ..Level::empty(0)
        }];
        let context = Context {
            levels: Some(&origin),
            ..Context::default()
        };
        assert!(
            direct_model::paragraphs("ab", &data, context, &mut 1, &mut 1024)
                .unwrap_err()
                .contains("work budget")
        );
        assert!(
            direct_model::paragraphs("ab", &data, context, &mut 10, &mut 8)
                .unwrap_err()
                .contains("model budget")
        );
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
        let paragraphs = model("x", &data, &["Name".into()]).unwrap();
        let run = first_run(&paragraphs);
        assert_eq!(run.text, "x");
        assert_eq!(run.font_family, None); // Omit invalid references; never guess an index.
    }
}
