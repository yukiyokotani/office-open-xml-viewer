//! Direct projection of binary PowerPoint text into the shared PPTX model.
//!
//! The binary semantics come from MS-PPT 2.9.14, 2.9.20, 2.9.41 and
//! 2.9.44-46. This deliberately shares the existing binary run decoding and
//! inheritance; it does not serialize or parse DrawingML as an intermediate.
//!
//! Internal producer groundwork, not a complete or publicly routed converter.
//! Unresolved paragraph origins and percentage before/after spacing need further
//! model admission support. Baseline information discarded by the native decoder
//! cannot be recovered here. Keep these limits explicit before wiring a producer.

use super::*;
use ooxml_common::text::SpaceLine;
use pptx_model::{
    Bullet as ModelBullet, Paragraph as ModelParagraph, TabStop, TextRun, TextRunData,
};

/// Build the receiving presentation model directly from a TextCharsAtom (or
/// decoded TextBytesAtom), StyleTextPropAtom and its resolved master context.
/// `model_budget` bounds requested output backing storage and owned string bytes;
/// allocator bookkeeping, input bytes and decoder temporaries are not included.
pub(in crate::ppt) fn paragraphs(
    text: &str,
    style: &[u8],
    context: Context<'_>,
    work_budget: &mut usize,
    model_budget: &mut usize,
) -> Result<Vec<ModelParagraph>, String> {
    paragraphs_with_axes(
        text,
        style,
        context,
        DirectAxes::default(),
        work_budget,
        model_budget,
    )
}

#[derive(Clone, Copy, Default)]
pub(in crate::ppt) struct DirectAxes<'a> {
    pub ruler: Option<ruler::Ruler<'a>>,
    pub document: Option<ParagraphAxes>,
}

pub(in crate::ppt) fn paragraphs_with_axes(
    text: &str,
    style: &[u8],
    context: Context<'_>,
    axes: DirectAxes<'_>,
    work_budget: &mut usize,
    model_budget: &mut usize,
) -> Result<Vec<ModelParagraph>, String> {
    charge_units(
        work_budget,
        text.len(),
        "PowerPoint direct text work budget exceeded",
    )?;
    if context.slide_numbers.windows(2).any(|v| v[0] >= v[1]) {
        return Err(unsupported(
            "duplicate or unordered PowerPoint slide-number positions",
        ));
    }
    let Runs {
        paragraphs: pf,
        characters: cf,
    } = read_runs(text, style, work_budget)?;
    let groups = context
        .style9
        .map(|bytes| auto_number::bind(bytes, &cf, work_budget))
        .transpose()?
        .unwrap_or_default();
    let mut numbers = context.slide_numbers.iter().peekable();
    let slide_number = context.slide_number.to_string();
    let mut number_group = 0;
    let (mut pi, mut ci, mut cp) = (0, 0, 0);
    let paragraph_count = text.matches('\r').count().saturating_add(1);
    model_charge::<ModelParagraph>(model_budget, paragraph_count)?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(paragraph_count)
        .map_err(|_| unsupported("PowerPoint direct text model allocation failed"))?;

    for paragraph in text.split('\r') {
        charge(work_budget, "PowerPoint direct text work budget exceeded")?;
        while pf[pi].0 <= cp {
            pi += 1;
        }
        let para_end = cp + paragraph.encode_utf16().count() + 1;
        if pf[pi].0 < para_end {
            return Err(unsupported("PowerPoint paragraph style splits a paragraph"));
        }
        let base = context
            .levels
            .and_then(|levels| levels.get(usize::from(pf[pi].1.level)));
        let mut properties = pf[pi].1.inherit(base.map(|v| &v.paragraph));
        // MS-PPT 2.9.30 supplies independent local ruler origins for the
        // paragraph's active level. Office-rendered binary counterfactuals
        // establish the document type-4 origin for ordinary level-0 freeform
        // text, including independent fallback beneath partial local axes.
        // Callers gate that narrower fallback context.
        if let Some(ruler) = axes.ruler {
            let level = usize::from(properties.level);
            if let Some(margin) = ruler.margins[level] {
                properties.margin = Some(margin);
            }
            if let Some(indent) = ruler.indents[level] {
                properties.indent = Some(indent);
            }
        }
        if properties.level == 0 {
            if let Some(document) = axes.document {
                properties.margin = properties.margin.or(document.margin);
                properties.indent = properties.indent.or(document.indent);
            }
        }
        let number = auto_number::paragraph(&groups, &mut number_group, cp, para_end);
        let mut runs = Vec::new();
        let mut start = 0;
        let mut iter = paragraph.char_indices().peekable();
        while let Some((offset, c)) = iter.next() {
            while cf[ci].0 <= cp {
                ci += 1;
            }
            let run = ci;
            if numbers.peek().is_some_and(|&&p| (p as usize) < cp) {
                return Err(unsupported(
                    "invalid PowerPoint slide-number character boundary",
                ));
            }
            if numbers.peek().is_some_and(|&&p| p as usize == cp) {
                push_text_runs(
                    &paragraph[start..offset],
                    &cf[run].1.inherit(base.map(|v| &v.character)),
                    context,
                    &mut runs,
                    work_budget,
                    model_budget,
                )?;
                push_text_runs(
                    &slide_number,
                    &cf[run].1.inherit(base.map(|v| &v.character)),
                    context,
                    &mut runs,
                    work_budget,
                    model_budget,
                )?;
                start = offset + c.len_utf8();
                numbers.next();
            }
            cp += c.len_utf16();
            if cp > cf[run].0 {
                return Err(unsupported(
                    "PowerPoint character style splits a surrogate pair",
                ));
            }
            if cp == cf[run].0 || iter.peek().is_none() {
                let end = offset + c.len_utf8();
                push_text_runs(
                    &paragraph[start..end],
                    &cf[run].1.inherit(base.map(|v| &v.character)),
                    context,
                    &mut runs,
                    work_budget,
                    model_budget,
                )?;
                start = end;
            }
        }
        while cf[ci].0 <= cp {
            ci += 1;
        }
        let end_character = cf[ci].1.inherit(base.map(|v| &v.character));
        result.push(model_paragraph(
            &properties,
            number,
            context,
            runs,
            &end_character,
            work_budget,
            model_budget,
        )?);
        cp += 1;
    }
    if numbers.next().is_some() {
        return Err(unsupported("PowerPoint slide-number position outside text"));
    }
    Ok(result)
}

fn charge(budget: &mut usize, message: &str) -> Result<(), String> {
    *budget = budget.checked_sub(1).ok_or_else(|| unsupported(message))?;
    Ok(())
}

fn charge_units(budget: &mut usize, units: usize, message: &str) -> Result<(), String> {
    *budget = budget
        .checked_sub(units)
        .ok_or_else(|| unsupported(message))?;
    Ok(())
}

fn model_charge<T>(budget: &mut usize, count: usize) -> Result<(), String> {
    let bytes = std::mem::size_of::<T>()
        .checked_mul(count)
        .ok_or_else(|| unsupported("PowerPoint direct text model size overflow"))?;
    charge_units(
        budget,
        bytes,
        "PowerPoint direct text model budget exceeded",
    )
}

fn push_text_runs(
    text: &str,
    character: &Character,
    context: Context<'_>,
    output: &mut Vec<TextRun>,
    budget: &mut usize,
    model_budget: &mut usize,
) -> Result<(), String> {
    for (index, part) in text.split(['\u{b}', '\n', '\u{2028}']).enumerate() {
        if index != 0 {
            charge(budget, "PowerPoint direct text work budget exceeded")?;
            reserve_run_slot(output, model_budget)?;
            output.push(TextRun::Break);
        }
        if !part.is_empty() {
            charge(budget, "PowerPoint direct text work budget exceeded")?;
            reserve_run_slot(output, model_budget)?;
            output.push(TextRun::Text(model_run(
                part,
                character,
                context,
                model_budget,
            )?));
        }
    }
    Ok(())
}

/// Charge backing storage before growing it. Capacity follows actual emitted
/// runs, not character count; uniform text stays a single run. Geometric
/// growth avoids quadratic copying for heavily styled input.
fn reserve_run_slot(output: &mut Vec<TextRun>, budget: &mut usize) -> Result<(), String> {
    if output.len() < output.capacity() {
        return Ok(());
    }
    let capacity = output
        .capacity()
        .max(1)
        .checked_mul(2)
        .ok_or_else(|| unsupported("PowerPoint direct text model size overflow"))?;
    model_charge::<TextRun>(budget, capacity - output.capacity())?;
    output
        .try_reserve_exact(capacity - output.len())
        .map_err(|_| unsupported("PowerPoint direct text model allocation failed"))
}

fn model_run(
    text: &str,
    character: &Character,
    context: Context<'_>,
    model_budget: &mut usize,
) -> Result<TextRunData, String> {
    charge_units(
        model_budget,
        text.len(),
        "PowerPoint direct text model budget exceeded",
    )?;
    let color = model_color(character.color, context.scheme, model_budget)?;
    let mut font = |id: Option<u16>| -> Result<Option<String>, String> {
        let value = id.and_then(|n| context.fonts.get(usize::from(n)));
        if let Some(value) = value {
            charge_units(
                model_budget,
                value.len(),
                "PowerPoint direct text model budget exceeded",
            )?;
        }
        Ok(value.cloned())
    };
    Ok(TextRunData {
        text: text.to_owned(),
        bold: (character.mask & 1 != 0).then_some(character.style & 1 != 0),
        italic: (character.mask & 2 != 0).then_some(character.style & 2 != 0),
        underline: character.mask & 4 != 0 && character.style & 4 != 0,
        underline_style: None,
        underline_color: None,
        strikethrough: false,
        strike_double: false,
        font_size: (character.mask & 0x20000 != 0).then_some(f64::from(character.size)),
        color,
        font_family: font(character.font)?,
        font_family_ea: font(character.ea)?,
        font_family_sym: font(character.symbol)?,
        // MS-PPT baseline is a percentage of line height, whereas this model's
        // baseline is thousandths of a point. Character::read validates it, but
        // no dimensionally invalid conversion is invented here.
        baseline: None,
        caps: None,
        letter_spacing: None,
        field_type: None,
        hyperlink: None,
        hyperlink_action: None,
        shadow: None,
        reflection: None,
        outline: None,
        highlight: None,
    })
}

fn model_paragraph(
    paragraph: &Paragraph,
    number: Option<auto_number::Number>,
    context: Context<'_>,
    runs: Vec<TextRun>,
    end_character: &Character,
    budget: &mut usize,
    model_budget: &mut usize,
) -> Result<ModelParagraph, String> {
    let rtl = paragraph.rtl.unwrap_or(false);
    let alignment = match paragraph.align {
        Some(value) => *["l", "ctr", "r", "just", "dist", "thaiDist", "justLow"]
            .get(usize::from(value))
            .ok_or_else(|| unsupported("invalid PowerPoint text alignment"))?,
        None if rtl => "r",
        None => "l",
    };
    charge_units(
        model_budget,
        alignment.len(),
        "PowerPoint direct text model budget exceeded",
    )?;
    let bullet = model_bullet(&paragraph.bullet, number, context, model_budget)?;
    let mar_l = paragraph
        .margin
        .map(|v| master_to_emu(i64::from(v)))
        .ok_or_else(|| {
            unsupported(
                "implicit binary PowerPoint paragraph margin requires model admission context",
            )
        })?;
    let indent = paragraph
        .indent
        .map(|v| master_to_emu(i64::from(v)) - mar_l)
        .ok_or_else(|| {
            unsupported(
                "implicit binary PowerPoint paragraph indent requires model admission context",
            )
        })?;
    let spacing = |value: Option<i16>| -> Result<Option<i64>, String> {
        match value {
            None => Ok(None),
            Some(0) => Ok(Some(0)),
            Some(1..) => Err(unsupported(
                "percentage PowerPoint before/after spacing requires presentation model support",
            )),
            Some(v) => {
                let raw = (-i32::from(v) * 100 + 4) / 8;
                Ok(Some(i64::from(raw)))
            }
        }
    };
    let space_line = match paragraph.spacing[0] {
        None => None,
        Some(v @ 0..=13200) => Some(SpaceLine::Pct {
            val: f64::from(v) * 1000.0,
        }),
        Some(13201..) => return Err(unsupported("invalid PowerPoint percentage line spacing")),
        // MS-PPT 2.2.20: absolute master units (1/8 point). The shared
        // floating-point model preserves these exactly; do not quantize via
        // DrawingML's hundredths-of-a-point serialization unit.
        Some(v) => Some(SpaceLine::Pts {
            val: -f64::from(v) / 8.0,
        }),
    };
    let tab_stops = if let Some(tabs) = context.ruler_tabs {
        let positions = tabs.positions();
        let count = positions.len();
        *budget = budget
            .checked_sub(positions.len())
            .ok_or_else(|| unsupported("PowerPoint ruler tab work budget exceeded"))?;
        model_charge::<TabStop>(model_budget, count)?;
        let mut result = Vec::new();
        result
            .try_reserve_exact(count)
            .map_err(|_| unsupported("PowerPoint direct text model allocation failed"))?;
        for (pos, algn) in positions {
            charge_units(
                model_budget,
                algn.len(),
                "PowerPoint direct text model budget exceeded",
            )?;
            result.push(TabStop {
                pos,
                algn: algn.to_owned(),
            });
        }
        result
    } else {
        Vec::new()
    };
    let empty = !runs.iter().any(|run| matches!(run, TextRun::Text(_)));
    Ok(ModelParagraph {
        alignment: alignment.to_owned(),
        mar_l,
        mar_r: 0,
        indent,
        space_before: spacing(paragraph.spacing[1])?,
        space_after: spacing(paragraph.spacing[2])?,
        space_line,
        lvl: u32::from(paragraph.level),
        bullet,
        def_font_size: (empty && end_character.mask & 0x20000 != 0)
            .then_some(f64::from(end_character.size)),
        def_color: None,
        def_bold: None,
        def_italic: None,
        def_font_family: None,
        tab_stops,
        def_tab_sz: paragraph.default_tab.map(|v| master_to_emu(i64::from(v))),
        rtl,
        ea_ln_brk: true,
        runs,
    })
}

fn model_bullet(
    bullet: &bullet::Bullet,
    number: Option<auto_number::Number>,
    context: Context<'_>,
    model_budget: &mut usize,
) -> Result<ModelBullet, String> {
    if bullet.enabled == Some(false) {
        return Ok(ModelBullet::None);
    }
    if bullet.enabled != Some(true) {
        return Ok(ModelBullet::Inherit);
    }
    let character = bullet
        .character
        .and_then(|v| char::from_u32(u32::from(v)))
        .filter(|c| !c.is_control() && !matches!(*c, '\u{fffe}' | '\u{ffff}'));
    if character.is_none() && number.is_none() {
        return Ok(ModelBullet::None);
    }
    let color = if bullet.has_color == Some(true) {
        model_color(bullet.color, context.scheme, model_budget)?
    } else {
        None
    };
    let (size_pct, size_pts) = if bullet.has_size == Some(true) {
        match bullet.size {
            Some(v @ 25..=400) => (Some(f64::from(v)), None),
            Some(v @ -4000..=-1) => (None, Some(f64::from(-v))),
            _ => (None, None),
        }
    } else {
        (None, None)
    };
    let font_family = if bullet.has_font == Some(true) {
        let value = bullet
            .font
            .and_then(|id| context.fonts.get(usize::from(id)));
        if let Some(value) = value {
            charge_units(
                model_budget,
                value.len(),
                "PowerPoint direct text model budget exceeded",
            )?;
        }
        value.cloned()
    } else {
        None
    };
    if let Some(number) = number {
        charge_units(
            model_budget,
            number.scheme.len(),
            "PowerPoint direct text model budget exceeded",
        )?;
        Ok(ModelBullet::AutoNum {
            num_type: number.scheme.to_owned(),
            start_at: Some(u32::from(number.start)),
            color,
            size_pct,
            size_pts,
            font_family,
        })
    } else {
        let ch = character.expect("validated above");
        charge_units(
            model_budget,
            ch.len_utf8(),
            "PowerPoint direct text model budget exceeded",
        )?;
        Ok(ModelBullet::Char {
            ch: ch.to_string(),
            color,
            size_pct,
            size_pts,
            font_family,
        })
    }
}

fn model_color(
    value: Option<u32>,
    scheme: Option<&scheme::Scheme>,
    model_budget: &mut usize,
) -> Result<Option<String>, String> {
    let Some(rgb) = value.and_then(|value| scheme::text(value, scheme)) else {
        return Ok(None);
    };
    charge_units(
        model_budget,
        6,
        "PowerPoint direct text model budget exceeded",
    )?;
    Ok(Some(format!(
        "{:02X}{:02X}{:02X}",
        rgb & 255,
        (rgb >> 8) & 255,
        (rgb >> 16) & 255
    )))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn plain_style(level: u16) -> Vec<u8> {
        [u32s(2), u16s(level), u32s(0), u32s(2), u32s(0)].concat()
    }

    fn ruler_with_axes(
        margins: [Option<i16>; 5],
        indents: [Option<i16>; 5],
    ) -> ruler::Ruler<'static> {
        ruler::Ruler {
            c_levels: None,
            default_tab_size: None,
            tabs: None,
            margins,
            indents,
        }
    }

    #[test]
    fn direct_origins_resolve_local_level_then_document_type4_level0_by_field() {
        let document = ParagraphAxes {
            margin: Some(180),
            indent: Some(90),
        };
        let project = |level, ruler, levels: Option<&[Level]>| {
            paragraphs_with_axes(
                "X",
                &plain_style(level),
                Context {
                    levels,
                    ..Context::default()
                },
                DirectAxes {
                    ruler,
                    document: Some(document),
                },
                &mut 100,
                &mut 100_000,
            )
        };

        let model = project(0, None, None).unwrap();
        assert_eq!(model[0].mar_l, master_to_emu(180));
        assert_eq!(model[0].indent, master_to_emu(-90));

        let ruler = ruler_with_axes(
            [Some(144), None, None, None, None],
            [Some(144), None, None, None, None],
        );
        let model = project(0, Some(ruler), None).unwrap();
        assert_eq!(model[0].mar_l, master_to_emu(144));
        assert_eq!(model[0].indent, 0);

        let ruler = ruler_with_axes([Some(288), None, None, None, None], [None; 5]);
        let model = project(0, Some(ruler), None).unwrap();
        assert_eq!(model[0].mar_l, master_to_emu(288));
        assert_eq!(model[0].indent, master_to_emu(90 - 288));

        for local_indent in [-144, 144] {
            let ruler = ruler_with_axes([None; 5], [Some(local_indent), None, None, None, None]);
            let model = project(0, Some(ruler), None).unwrap();
            assert_eq!(model[0].mar_l, master_to_emu(180));
            assert_eq!(
                model[0].indent,
                master_to_emu(i64::from(local_indent) - 180)
            );
        }

        let mut inherited = Level::empty(0);
        inherited.paragraph.margin = Some(500);
        inherited.paragraph.indent = Some(400);
        let ruler = ruler_with_axes(
            [Some(0), None, None, None, None],
            [Some(0), None, None, None, None],
        );
        let model = project(0, Some(ruler), Some(std::slice::from_ref(&inherited))).unwrap();
        assert_eq!(model[0].mar_l, 0);
        assert_eq!(model[0].indent, 0);

        let vt_style = [u32s(4), u16s(0), u32s(0), u32s(4), u32s(0)].concat();
        let model = paragraphs_with_axes(
            "A\u{b}B",
            &vt_style,
            Context::default(),
            DirectAxes {
                ruler: None,
                document: Some(document),
            },
            &mut 100,
            &mut 100_000,
        )
        .unwrap();
        assert_eq!(model.len(), 1);
        assert_eq!(model[0].mar_l, master_to_emu(180));
        assert_eq!(model[0].indent, master_to_emu(-90));
    }

    #[test]
    fn document_origin_does_not_fill_higher_levels_but_local_ruler_does() {
        let axes = DirectAxes {
            ruler: None,
            document: Some(ParagraphAxes {
                margin: Some(180),
                indent: Some(90),
            }),
        };
        let error = paragraphs_with_axes(
            "X",
            &plain_style(1),
            Context::default(),
            axes,
            &mut 100,
            &mut 100_000,
        )
        .unwrap_err();
        assert!(error.contains("margin requires model admission context"));

        let ruler = ruler_with_axes(
            [None, Some(0), None, None, None],
            [None, Some(0), None, None, None],
        );
        let model = paragraphs_with_axes(
            "X",
            &plain_style(1),
            Context::default(),
            DirectAxes {
                ruler: Some(ruler),
                document: axes.document,
            },
            &mut 100,
            &mut 100_000,
        )
        .unwrap();
        assert_eq!((model[0].mar_l, model[0].indent), (0, 0));

        assert!(paragraphs(
            "X",
            &plain_style(0),
            Context::default(),
            &mut 100,
            &mut 100_000
        )
        .unwrap_err()
        .contains("margin requires model admission context"));
    }

    #[test]
    fn missing_font_size_stays_absent_instead_of_using_decoder_storage_default() {
        let mut base = Level::empty(0);
        base.paragraph.margin = Some(0);
        base.paragraph.indent = Some(0);
        let run = model_run("A", &base.character, Context::default(), &mut 100).unwrap();
        assert_eq!(run.font_size, None);
        let paragraph = model_paragraph(
            &base.paragraph,
            None,
            Context::default(),
            vec![],
            &base.character,
            &mut 100,
            &mut 100,
        )
        .unwrap();
        assert_eq!(paragraph.def_font_size, None);
        base.character.mask |= 0x20000;
        base.character.size = 24;
        let run = model_run("A", &base.character, Context::default(), &mut 100).unwrap();
        assert_eq!(run.font_size, Some(24.0));
    }

    #[test]
    fn uniform_text_storage_tracks_runs_not_character_count() {
        let text = "A".repeat(2048);
        let length = u32::try_from(text.len() + 1).unwrap();
        let style = [u32s(length), u16s(0), u32s(0), u32s(length), u32s(0)].concat();
        let mut base = Level::empty(0);
        base.paragraph.margin = Some(0);
        base.paragraph.indent = Some(0);
        let context = Context {
            levels: Some(std::slice::from_ref(&base)),
            ..Context::default()
        };
        let required = std::mem::size_of::<ModelParagraph>()
            + 2 * std::mem::size_of::<TextRun>()
            + text.len()
            + 1;
        let mut available = required;
        let result = paragraphs(&text, &style, context, &mut 10_000, &mut available).unwrap();
        assert_eq!(result[0].runs.len(), 1);
        assert_eq!(result[0].runs.capacity(), 2);
        assert_eq!(available, 0);
        assert!(
            paragraphs(&text, &style, context, &mut 10_000, &mut (required - 1))
                .unwrap_err()
                .contains("model budget")
        );
    }

    #[test]
    fn run_capacity_growth_is_charged_before_reserving() {
        let mut runs = Vec::new();
        let mut budget = 4 * std::mem::size_of::<TextRun>();
        for _ in 0..4 {
            reserve_run_slot(&mut runs, &mut budget).unwrap();
            runs.push(TextRun::Break);
        }
        assert_eq!(budget, 0);
        assert_eq!(runs.capacity(), 4);
        assert!(reserve_run_slot(&mut runs, &mut budget).is_err());
        assert_eq!(runs.len(), 4);
        assert_eq!(runs.capacity(), 4);
    }

    #[test]
    fn tab_storage_counts_structs_and_owned_alignment_strings() {
        let data = [
            4u32.to_le_bytes().as_slice(),
            &2u16.to_le_bytes(),
            &0i16.to_le_bytes(),
            &0u16.to_le_bytes(),
            &576i16.to_le_bytes(),
            &3u16.to_le_bytes(),
        ]
        .concat();
        let tabs = ruler::read(
            Record {
                kind: 4006,
                version: 0,
                instance: 0,
                payload: &data,
            },
            &mut 100,
        )
        .unwrap()
        .unwrap();
        let mut base = Level::empty(0);
        base.paragraph.margin = Some(0);
        base.paragraph.indent = Some(0);
        let context = Context {
            ruler_tabs: Some(tabs),
            ..Context::default()
        };
        let required = 1 + 2 * std::mem::size_of::<TabStop>() + 1 + 3;
        let mut budget = required;
        let result = model_paragraph(
            &base.paragraph,
            None,
            context,
            vec![],
            &base.character,
            &mut 100,
            &mut budget,
        )
        .unwrap();
        assert_eq!(budget, 0);
        assert_eq!(result.tab_stops.len(), 2);
        assert_eq!(result.tab_stops[1].pos, 914400);
        assert!(model_paragraph(
            &base.paragraph,
            None,
            context,
            vec![],
            &base.character,
            &mut 100,
            &mut (required - 1)
        )
        .is_err());
    }

    fn u16s(value: u16) -> Vec<u8> {
        value.to_le_bytes().to_vec()
    }
    fn u32s(value: u32) -> Vec<u8> {
        value.to_le_bytes().to_vec()
    }

    fn explicit_origin() -> Level {
        let mut level = Level::empty(0);
        level.paragraph.margin = Some(0);
        level.paragraph.indent = Some(0);
        level
    }

    fn topology_style(total: u32, character_counts: &[u32]) -> Vec<u8> {
        let mut style = [u32s(total), u16s(0), u32s(0)].concat();
        for &count in character_counts {
            style.extend(u32s(count));
            style.extend(u32s(0));
        }
        style
    }

    #[test]
    fn preserves_empty_leading_and_trailing_classic_paragraphs() {
        // MS-PPT 2.9.41 adds one implicit terminal CR outside TextBytesAtom.
        // A controlled Office roundtrip retained CR + "AB" + CR as three
        // paragraphs while one PF run covered all five UTF-16 positions and
        // the CF runs covered 1/2/2 positions.
        let base = explicit_origin();
        let style = topology_style(5, &[1, 2, 2]);
        let model = paragraphs(
            "\rAB\r",
            &style,
            Context {
                levels: Some(std::slice::from_ref(&base)),
                ..Context::default()
            },
            &mut 100,
            &mut 100_000,
        )
        .unwrap();
        assert_eq!(model.len(), 3);
        assert!(model[0].runs.is_empty());
        assert!(model[2].runs.is_empty());
        assert!(matches!(
            model[1].runs.as_slice(),
            [TextRun::Text(run)] if run.text == "AB"
        ));
    }

    #[test]
    fn preserves_classic_vertical_tab_as_an_explicit_line_break() {
        // The controlled Office protocol retained U+000B inside one classic
        // paragraph and roundtripped it as DrawingML a:br.
        let base = explicit_origin();
        let model = paragraphs(
            "AB\u{b}CD",
            &topology_style(6, &[6]),
            Context {
                levels: Some(std::slice::from_ref(&base)),
                ..Context::default()
            },
            &mut 100,
            &mut 100_000,
        )
        .unwrap();
        assert_eq!(model.len(), 1);
        assert!(matches!(
            model[0].runs.as_slice(),
            [TextRun::Text(first), TextRun::Break, TextRun::Text(last)]
                if first.text == "AB" && last.text == "CD"
        ));
    }

    #[test]
    fn accepts_a_valid_character_run_spanning_a_supplementary_scalar() {
        // Style run counts use UTF-16 positions. The scalar occupies two units;
        // the controlled classic file retained one run across all four text
        // units without splitting its surrogate pair.
        let base = explicit_origin();
        let model = paragraphs(
            "A😀B",
            &topology_style(5, &[5]),
            Context {
                levels: Some(std::slice::from_ref(&base)),
                ..Context::default()
            },
            &mut 100,
            &mut 100_000,
        )
        .unwrap();
        assert_eq!(model.len(), 1);
        assert!(matches!(
            model[0].runs.as_slice(),
            [TextRun::Text(run)] if run.text == "A😀B"
        ));
    }

    #[test]
    fn applies_one_body_ruler_tab_to_each_classic_paragraph() {
        // The controlled Office protocol retained one TextRulerAtom for a text
        // body containing multiple CR-delimited paragraphs, not paragraph-local
        // ruler arrays.
        let ruler_data = [
            4u32.to_le_bytes().as_slice(),
            &1u16.to_le_bytes(),
            &1152i16.to_le_bytes(),
            &0u16.to_le_bytes(),
        ]
        .concat();
        let tabs = ruler::read(
            Record {
                kind: 4006,
                version: 0,
                instance: 0,
                payload: &ruler_data,
            },
            &mut 100,
        )
        .unwrap()
        .unwrap();
        let base = explicit_origin();
        let model = paragraphs(
            "A\rB",
            &topology_style(4, &[4]),
            Context {
                levels: Some(std::slice::from_ref(&base)),
                ruler_tabs: Some(tabs),
                ..Context::default()
            },
            &mut 100,
            &mut 100_000,
        )
        .unwrap();
        assert_eq!(model.len(), 2);
        for paragraph in model {
            assert!(matches!(
                paragraph.tab_stops.as_slice(),
                [TabStop { pos: 1_828_800, algn }] if algn == "l"
            ));
        }
    }

    #[test]
    fn projects_inherited_binary_styles_breaks_and_signed_geometry() {
        let base = Level {
            paragraph: Paragraph {
                level: 0,
                rtl: Some(true),
                align: None,
                spacing: [Some(150), Some(-16), Some(-24)],
                bullet: bullet::Bullet {
                    enabled: Some(true),
                    character: Some(0x2022),
                    ..Default::default()
                },
                margin: Some(-10),
                indent: Some(20),
                default_tab: Some(576),
            },
            character: Character {
                mask: 0x20003,
                style: 1,
                size: 24,
                font: Some(0),
                ea: None,
                symbol: None,
                color: Some(0xfe33_2211),
            },
        };
        // One paragraph run and one character run, both covering "A\nB" + CR.
        let style = [u32s(4), u16s(0), u32s(0), u32s(4), u32s(0)].concat();
        let fonts = ["Example".to_string()];
        let model = paragraphs(
            "A\nB",
            &style,
            Context {
                fonts: &fonts,
                levels: Some(&[base]),
                ..Context::default()
            },
            &mut 100,
            &mut 100_000,
        )
        .unwrap();
        assert_eq!(model[0].alignment, "r");
        assert!(model[0].mar_l < 0);
        assert_eq!(model[0].indent, master_to_emu(30));
        assert_eq!(model[0].space_before, Some(200));
        assert_eq!(model[0].space_after, Some(300));
        assert!(matches!(model[0].space_line, Some(SpaceLine::Pct { val }) if val == 150_000.0));
        assert_eq!(model[0].def_tab_sz, Some(914_400));
        assert!(matches!(
            model[0].runs.as_slice(),
            [TextRun::Text(_), TextRun::Break, TextRun::Text(_)]
        ));
        let TextRun::Text(run) = &model[0].runs[0] else {
            panic!()
        };
        assert_eq!(run.font_family.as_deref(), Some("Example"));
        assert_eq!(run.color.as_deref(), Some("112233"));
        assert_eq!(run.bold, Some(true));
        assert_eq!(run.italic, Some(false));
        assert_eq!(run.font_size, Some(24.0));
    }

    #[test]
    fn substitutes_slide_numbers_at_utf16_positions_and_bounds_output_work() {
        let style = [u32s(4), u16s(0), u32s(0), u32s(4), u32s(0)].concat();
        let mut base = Level::empty(0);
        base.paragraph.margin = Some(0);
        base.paragraph.indent = Some(0);
        let context = Context {
            slide_numbers: &[1],
            slide_number: 27,
            levels: Some(std::slice::from_ref(&base)),
            ..Context::default()
        };
        let model = paragraphs("A*B", &style, context, &mut 100, &mut 100_000).unwrap();
        let texts: Vec<_> = model[0]
            .runs
            .iter()
            .filter_map(|run| match run {
                TextRun::Text(run) => Some(run.text.as_str()),
                _ => None,
            })
            .collect();
        assert_eq!(texts, ["A", "27", "B"]);
        assert!(paragraphs("A*B", &style, context, &mut 2, &mut 100_000).is_err());
        assert!(paragraphs("A*B", &style, context, &mut 100, &mut 1).is_err());
    }

    #[test]
    fn line_spacing_preserves_every_valid_native_value_without_xml_quantization() {
        let mut base = Level::empty(0);
        base.paragraph.margin = Some(0);
        base.paragraph.indent = Some(0);
        // MS-PPT 2.2.20: negative master units or percentages in 0..=13200.
        // One master unit is 1/576 inch = 1/8 point, exactly representable here.
        for value in i16::MIN..=i16::MAX {
            base.paragraph.spacing[0] = Some(value);
            let result = model_paragraph(
                &base.paragraph,
                None,
                Context::default(),
                vec![],
                &base.character,
                &mut 100,
                &mut 100,
            );
            if value > 13200 {
                assert!(result.unwrap_err().contains("percentage line spacing"));
            } else {
                let expected = if value < 0 {
                    SpaceLine::Pts {
                        val: -f64::from(value) / 8.0,
                    }
                } else {
                    SpaceLine::Pct {
                        val: f64::from(value) * 1000.0,
                    }
                };
                assert_eq!(result.unwrap().space_line, Some(expected), "{value}");
            }
        }
    }

    #[test]
    fn spacing_boundaries_and_unresolved_geometry_fail_closed() {
        let character = Level::empty(0).character;
        for (spacing, expected) in [
            (
                [Some(13_200), Some(1), None],
                "percentage PowerPoint before/after",
            ),
            (
                [Some(13_201), None, None],
                "invalid PowerPoint percentage line",
            ),
        ] {
            let paragraph = Paragraph {
                spacing,
                margin: Some(0),
                indent: Some(0),
                ..Level::empty(0).paragraph
            };
            assert!(model_paragraph(
                &paragraph,
                None,
                Context::default(),
                vec![],
                &character,
                &mut 10,
                &mut 10_000,
            )
            .unwrap_err()
            .contains(expected));
        }
        let unresolved = Level::empty(0);
        assert!(model_paragraph(
            &unresolved.paragraph,
            None,
            Context::default(),
            vec![],
            &unresolved.character,
            &mut 10,
            &mut 10_000,
        )
        .unwrap_err()
        .contains("margin requires model admission context"));
    }

    #[test]
    fn rejects_surrogate_splits_before_allocating_model_runs() {
        // Paragraph covers the scalar plus CR; character runs illegally split
        // the UTF-16 surrogate pair after its first code unit.
        let style = [
            u32s(3),
            u16s(0),
            u32s(0),
            u32s(1),
            u32s(0),
            u32s(2),
            u32s(0),
        ]
        .concat();
        assert!(
            paragraphs("😀", &style, Context::default(), &mut 100, &mut 10_000)
                .unwrap_err()
                .contains("surrogate pair")
        );
    }
}
