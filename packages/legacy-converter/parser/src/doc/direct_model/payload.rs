//! Heap-payload accounting for the currently supported direct DOC projection.
//!
//! These totals exclude the outer Rust struct (charged by the caller) and are
//! not an RSS estimate. String and vector *capacities* are charged because that
//! is the retained allocation. The explicitly enumerated unsupported owning
//! fields fail closed until included in the direct projection and accounting.

use docx_model::{
    CtBorderTypographyWire, DocParagraph, ParagraphBorders, ParagraphTypographyWire,
    RunFontAxisValues, RunFontFacts, RunFontSlots, RunTypographyWire, SectionProps, TextRun,
    TypographyValueWire,
};

pub(super) fn table(value: &docx_model::DocTable) -> Result<usize, String> {
    let mut total = Total::default();
    for string in [&value.jc, &value.table_layout.logical_sequence_id] {
        total.string(string)?;
    }
    total.option_string(&value.layout)?;
    total.option_string(&value.overlap)?;
    total.option_string(&value.table_layout.effective_style_id)?;
    if let Some(position) = &value.tblp_pr {
        total.string(&position.horz_anchor)?;
        total.string(&position.vert_anchor)?;
        total.option_string(&position.tblp_x_spec)?;
        total.option_string(&position.tblp_y_spec)?;
    }
    for column in &value.table_layout.grid.columns {
        total.option_string(&column.width)?;
    }
    if let Some(width) = &value.table_layout.preferred_width {
        total.table_width(width)?;
    }
    if let Some(layout) = &value.table_layout.layout {
        total.option_string(&layout.kind)?;
    }
    if let Some(margins) = &value.table_layout.cell_margins {
        total.table_margins(margins)?;
    }
    for row in &value.rows {
        total.string(&row.row_height_rule)?;
        if let Some(height) = &row.table_row_layout.height {
            total.option_string(&height.value)?;
            total.string(&height.rule)?;
        }
        for width in [
            &row.table_row_layout.before_width,
            &row.table_row_layout.after_width,
        ]
        .into_iter()
        .flatten()
        {
            total.option_string(&width.kind)?;
            total.option_string(&width.value)?;
        }
        if let Some(exception) = &row.table_row_layout.exception {
            if let Some(width) = &exception.preferred_width {
                total.table_width(width)?;
            }
            if exception.layout.is_some()
                || exception.justification.is_some()
                || exception.indent.is_some()
                || exception.borders.is_some()
                || exception.cell_margins.is_some()
                || exception.cell_spacing.is_some()
            {
                return Err(unsupported(
                    "unaccounted direct DOC table exception payload",
                ));
            }
        }
        for cell in &row.cells {
            total.string(&cell.v_align)?;
            total.option_string(&cell.background)?;
            if let Some(width) = &cell.table_cell_layout.preferred_width {
                total.table_width(width)?;
            }
            if let Some(margins) = &cell.table_cell_layout.margins {
                total.table_margins(margins)?;
            }
            for border in [
                &cell.borders.top,
                &cell.borders.left,
                &cell.borders.bottom,
                &cell.borders.right,
                &cell.borders.inside_h,
                &cell.borders.inside_v,
            ]
            .into_iter()
            .flatten()
            {
                total.string(&border.style)?;
                total.option_string(&border.color)?;
            }
        }
    }
    Ok(total.0)
}

pub(super) fn text_run(run: &TextRun) -> Result<usize, String> {
    if run.no_break_hyphen_offsets.capacity() != 0
        || run.border.is_some()
        || run.ruby.is_some()
        || run.revision.is_some()
        || run.note_ref.is_some()
    {
        return Err(unsupported("unaccounted direct DOC text-run payload"));
    }
    let mut total = Total::default();
    total.string(&run.text)?;
    for value in [
        &run.underline_style,
        &run.underline_color,
        &run.color,
        &run.font_family,
        &run.font_family_high_ansi,
        &run.font_family_east_asia,
        &run.font_hint,
        &run.background,
        &run.vert_align,
        &run.hyperlink,
        &run.hyperlink_anchor,
        &run.highlight,
        &run.emphasis_mark,
        &run.font_family_cs,
        &run.lang_bidi,
        &run.lang_east_asia,
        &run.lang_default,
        &run.fit_text_id,
        &run.east_asian_combine_brackets,
    ] {
        total.option_string(value)?;
    }
    if let Some(slots) = &run.font_slots {
        total.font_slots(slots)?;
    }
    if let Some(wire) = &run.typography_acquisition {
        total.run_typography(wire)?;
    }
    Ok(total.0)
}

pub(super) fn paragraph(value: &DocParagraph) -> Result<usize, String> {
    if value.runs.capacity() != 0 {
        return Err(unsupported("unaccounted direct DOC paragraph run payload"));
    }
    paragraph_metadata(value)
}

/// Metadata only: run payloads are owned and charged by the run producer.
pub(super) fn paragraph_metadata(value: &DocParagraph) -> Result<usize, String> {
    if value.run_revisions.capacity() != 0
        || value.complex_field_boundaries.capacity() != 0
        || value.bookmarks.capacity() != 0
        || value.comment_marks.capacity() != 0
        || value.frame_pr.is_some()
    {
        return Err(unsupported("unaccounted direct DOC paragraph payload"));
    }
    let mut total = Total::default();
    if let Some(numbering) = &value.numbering {
        total.add(std::mem::size_of::<docx_model::NumberingInfo>())?;
        for string in [
            &numbering.format,
            &numbering.text,
            &numbering.suff,
            &numbering.jc,
        ] {
            total.string(string)?;
        }
        for string in [
            &numbering.font_family,
            &numbering.font_family_east_asia,
            &numbering.color,
            &numbering.pic_bullet_image_path,
            &numbering.pic_bullet_mime_type,
        ] {
            total.option_string(string)?;
        }
        if let Some(facts) = &numbering.font_facts {
            total.font_facts(facts)?;
        }
    }
    for string in [
        &value.paragraph_id,
        &value.shading,
        &value.style_id,
        &value.default_font_family,
        &value.default_font_family_east_asia,
        &value.paragraph_mark_color,
    ] {
        total.option_string(string)?;
    }
    total.string(&value.alignment)?;
    if let Some(line_spacing) = &value.line_spacing {
        total.string(&line_spacing.rule)?;
    }
    total.vec::<docx_model::TabStop>(value.tab_stops.capacity())?;
    for tab in &value.tab_stops {
        total.string(&tab.alignment)?;
        total.string(&tab.leader)?;
    }
    if let Some(borders) = &value.borders {
        total.paragraph_borders(borders)?;
    }
    if let Some(facts) = &value.paragraph_mark_font_facts {
        total.font_facts(facts)?;
    }
    if let Some(wire) = &value.paragraph_typography_acquisition {
        total.paragraph_typography(wire)?;
    }
    Ok(total.0)
}

pub(super) fn section(value: &SectionProps) -> Result<usize, String> {
    if value.page_borders.is_some() || value.line_numbering.is_some() {
        return Err(unsupported("unaccounted direct DOC section payload"));
    }
    let mut total = Total::default();
    for string in [
        &value.section_start,
        &value.text_direction,
        &value.doc_grid_type,
        &value.v_align,
    ] {
        total.option_string(string)?;
    }
    if let Some(columns) = &value.columns {
        total.vec::<docx_model::ColSpec>(columns.cols.capacity())?;
    }
    if let Some(numbering) = &value.page_num_type {
        total.option_string(&numbering.fmt)?;
    }
    if let Some(placement) = &value.section_placement {
        if placement.line_numbering.is_some() || placement.page_borders.is_some() {
            return Err(unsupported(
                "unaccounted direct DOC section placement payload",
            ));
        }
        total.add(std::mem::size_of_val(placement.as_ref()))?;
        total.string(&placement.section_id)?;
        total.option_string(&placement.v_align)?;
        total.option_string(&placement.doc_grid_type)?;
        if let Some(geometry) = &placement.page_geometry {
            total.add(std::mem::size_of_val(geometry.as_ref()))?;
        }
    }
    Ok(total.0)
}

pub(super) fn ending_section(
    kind: &String,
    columns: Option<&docx_model::ColumnsSpec>,
    page_num_type: Option<&docx_model::PageNumType>,
    text_direction: &Option<String>,
    geom: &docx_model::SectionGeom,
    placement: &docx_model::SectionPlacementWire,
) -> Result<usize, String> {
    let mut total = Total::default();
    total.string(kind)?;
    total.columns(columns)?;
    if let Some(numbering) = page_num_type {
        total.option_string(&numbering.fmt)?;
    }
    total.option_string(text_direction)?;
    total.add(std::mem::size_of_val(geom))?;
    total.add(2 * std::mem::size_of::<docx_model::HeadersFooters>())?;
    total.section_placement(placement)?;
    Ok(total.0)
}

#[derive(Default)]
struct Total(usize);

impl Total {
    fn add(&mut self, bytes: usize) -> Result<(), String> {
        self.0 = self.0.checked_add(bytes).ok_or("OUTPUT_TOO_LARGE")?;
        Ok(())
    }

    fn string(&mut self, value: &String) -> Result<(), String> {
        self.add(value.capacity())
    }

    fn option_string(&mut self, value: &Option<String>) -> Result<(), String> {
        if let Some(value) = value {
            self.string(value)?;
        }
        Ok(())
    }

    fn vec<T>(&mut self, capacity: usize) -> Result<(), String> {
        self.add(
            capacity
                .checked_mul(std::mem::size_of::<T>())
                .ok_or("OUTPUT_TOO_LARGE")?,
        )
    }

    fn table_width(&mut self, value: &docx_model::TableWidthAcquisitionWire) -> Result<(), String> {
        self.option_string(&value.kind)?;
        self.option_string(&value.value)
    }

    fn table_margins(
        &mut self,
        value: &docx_model::TableMarginAcquisitionWire,
    ) -> Result<(), String> {
        for width in [
            &value.top,
            &value.bottom,
            &value.start,
            &value.end,
            &value.left,
            &value.right,
        ]
        .into_iter()
        .flatten()
        {
            self.table_width(width)?;
        }
        Ok(())
    }

    fn axes(&mut self, axes: &RunFontAxisValues) -> Result<(), String> {
        for value in [
            &axes.ascii,
            &axes.east_asia,
            &axes.high_ansi,
            &axes.complex_script,
        ] {
            self.option_string(value)?;
        }
        Ok(())
    }

    fn font_slots(&mut self, slots: &RunFontSlots) -> Result<(), String> {
        self.axes(&slots.direct)?;
        self.axes(&slots.theme)
    }

    fn font_facts(&mut self, facts: &RunFontFacts) -> Result<(), String> {
        for value in [
            &facts.font_family,
            &facts.font_family_high_ansi,
            &facts.font_family_east_asia,
            &facts.font_hint,
            &facts.font_family_cs,
            &facts.lang_bidi,
            &facts.lang_east_asia,
            &facts.lang_default,
        ] {
            self.option_string(value)?;
        }
        if let Some(slots) = &facts.font_slots {
            self.font_slots(slots)?;
        }
        Ok(())
    }

    fn columns(&mut self, columns: Option<&docx_model::ColumnsSpec>) -> Result<(), String> {
        if let Some(columns) = columns {
            self.vec::<docx_model::ColSpec>(columns.cols.capacity())?;
        }
        Ok(())
    }

    fn section_placement(
        &mut self,
        placement: &docx_model::SectionPlacementWire,
    ) -> Result<(), String> {
        if placement.line_numbering.is_some() || placement.page_borders.is_some() {
            return Err(unsupported(
                "unaccounted direct DOC section placement payload",
            ));
        }
        self.add(std::mem::size_of_val(placement))?;
        self.string(&placement.section_id)?;
        self.option_string(&placement.v_align)?;
        self.option_string(&placement.doc_grid_type)?;
        if let Some(geometry) = &placement.page_geometry {
            self.add(std::mem::size_of_val(geometry.as_ref()))?;
        }
        Ok(())
    }

    fn typography_string(&mut self, value: &TypographyValueWire<String>) -> Result<(), String> {
        self.option_string(&value.raw)?;
        self.option_string(&value.value)
    }

    fn typography_raw<T>(&mut self, value: &TypographyValueWire<T>) -> Result<(), String> {
        self.option_string(&value.raw)
    }

    fn border_typography(&mut self, value: &CtBorderTypographyWire) -> Result<(), String> {
        for item in [
            &value.val,
            &value.color,
            &value.theme_color,
            &value.theme_tint,
            &value.theme_shade,
        ] {
            self.typography_string(item)?;
        }
        self.typography_raw(&value.size_pt)?;
        self.typography_raw(&value.space_pt)?;
        self.typography_raw(&value.shadow)?;
        self.typography_raw(&value.frame)
    }

    fn run_typography(&mut self, value: &RunTypographyWire) -> Result<(), String> {
        if value.fit_text.is_some()
            || value.border.is_some()
            || value.ruby.is_some()
            || value.revision.is_some()
        {
            return Err(unsupported("unaccounted direct DOC run typography payload"));
        }
        if let Some(underline) = &value.underline {
            for item in [
                &underline.val,
                &underline.color,
                &underline.theme_color,
                &underline.theme_tint,
                &underline.theme_shade,
            ] {
                self.typography_string(item)?;
            }
        }
        self.typography_string(&value.vertical_align)?;
        self.typography_raw(&value.position_pt)?;
        self.typography_string(&value.emphasis)?;
        self.option_string(&value.languages.east_asia)?;
        self.option_string(&value.languages.bidi)?;
        self.option_string(&value.languages.default)?;
        self.typography_string(&value.east_asian_layout.combine_brackets)
    }

    fn paragraph_borders(&mut self, value: &ParagraphBorders) -> Result<(), String> {
        for edge in [
            &value.top,
            &value.right,
            &value.bottom,
            &value.left,
            &value.between,
        ]
        .into_iter()
        .flatten()
        {
            self.string(&edge.style)?;
            self.option_string(&edge.color)?;
        }
        Ok(())
    }

    fn paragraph_typography(&mut self, value: &ParagraphTypographyWire) -> Result<(), String> {
        for edge in [
            &value.borders.top,
            &value.borders.right,
            &value.borders.bottom,
            &value.borders.left,
            &value.borders.between,
            &value.borders.bar,
        ]
        .into_iter()
        .flatten()
        {
            self.border_typography(edge)?;
        }
        Ok(())
    }
}

fn unsupported(message: &str) -> String {
    format!("UNSUPPORTED:{message}")
}

#[cfg(test)]
mod tests {
    use super::*;

    fn allocated(value: &str, capacity: usize) -> String {
        let mut output = String::with_capacity(capacity);
        output.push_str(value);
        output
    }

    #[test]
    fn counts_string_capacities_and_vector_backing_not_lengths() {
        let mut run = TextRun::default();
        run.text = allocated("x", 32);
        run.font_family = Some(allocated("font", 64));
        assert_eq!(text_run(&run).unwrap(), 96);

        let mut paragraph_value = DocParagraph::default();
        paragraph_value.alignment = allocated("left", 16);
        paragraph_value.line_spacing = Some(docx_model::LineSpacing {
            value: 1.0,
            rule: allocated("auto", 10),
            explicit: true,
        });
        paragraph_value.tab_stops = Vec::with_capacity(3);
        paragraph_value.tab_stops.push(docx_model::TabStop {
            pos: -1.0,
            alignment: allocated("left", 8),
            leader: allocated("dot", 12),
        });
        assert_eq!(
            paragraph(&paragraph_value).unwrap(),
            16 + 10 + 3 * std::mem::size_of::<docx_model::TabStop>() + 8 + 12
        );
    }

    #[test]
    fn counts_western_language_capacities_at_each_reachable_run_seam() {
        let mut run = TextRun::default();
        run.lang_default = Some(allocated("en-US", 31));
        run.typography_acquisition = Some(RunTypographyWire {
            languages: docx_model::TypographyLanguagesWire {
                default: Some(allocated("en-US", 47)),
                ..docx_model::TypographyLanguagesWire::default()
            },
            ..RunTypographyWire::default()
        });
        assert_eq!(text_run(&run).unwrap(), 31 + 47);

        let facts = RunFontFacts {
            lang_default: Some(allocated("fr-FR", 53)),
            ..RunFontFacts::default()
        };
        let mut total = Total::default();
        total.font_facts(&facts).unwrap();
        assert_eq!(total.0, 53);
    }

    #[test]
    fn counts_row_preferred_width_exception_strings() {
        let mut table_value = docx_model::DocTable::default();
        let before = table(&table_value).unwrap();
        table_value.rows.push(docx_model::DocTableRow {
            table_row_layout: docx_model::TableRowLayoutAcquisitionWire {
                exception: Some(docx_model::TablePropertyExceptionAcquisitionWire {
                    preferred_width: Some(docx_model::TableWidthAcquisitionWire {
                        kind: Some(allocated("auto", 17)),
                        value: Some(allocated("0", 23)),
                    }),
                    ..Default::default()
                }),
                ..Default::default()
            },
            ..Default::default()
        });
        assert_eq!(table(&table_value).unwrap() - before, 40);
        table_value.rows[0]
            .table_row_layout
            .exception
            .as_mut()
            .unwrap()
            .justification = Some("left".into());
        assert!(table(&table_value).unwrap_err().starts_with("UNSUPPORTED:"));
    }

    #[test]
    fn western_language_payload_is_rejected_one_byte_over_budget_before_retention() {
        let mut run = TextRun {
            lang_default: Some(allocated("en-US", 31)),
            typography_acquisition: Some(RunTypographyWire {
                languages: docx_model::TypographyLanguagesWire {
                    default: Some(allocated("en-US", 47)),
                    ..docx_model::TypographyLanguagesWire::default()
                },
                ..RunTypographyWire::default()
            }),
            ..TextRun::default()
        };
        let required = std::mem::size_of::<TextRun>() + 31 + 47;
        let mut runs = Vec::with_capacity(1);
        let mut budget = super::super::ModelBudget::new(required - 1);
        assert_eq!(
            budget.text(&mut runs, &mut run, "").unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
        assert!(runs.is_empty());
        assert_eq!(run.lang_default.as_deref(), Some("en-US"));

        let mut budget = super::super::ModelBudget::new(required);
        budget.text(&mut runs, &mut run, "").unwrap();
        assert_eq!(runs.len(), 1);
        assert_eq!(budget.remaining_bytes, 0);
    }

    #[test]
    fn counts_boxed_section_payload_and_rejects_unowned_fields() {
        let mut section_value = SectionProps::default();
        section_value.section_start = Some(allocated("nextPage", 24));
        section_value.section_placement = Some(Box::new(docx_model::SectionPlacementWire {
            section_id: allocated("section:0", 32),
            section_bidi: false,
            v_align: None,
            line_numbering: None,
            doc_grid_type: None,
            doc_grid_line_pitch: None,
            doc_grid_char_space: None,
            gutter_pt: None,
            rtl_gutter: None,
            page_borders_authored: None,
            page_borders: None,
            page_geometry: None,
        }));
        assert_eq!(
            section(&section_value).unwrap(),
            24 + std::mem::size_of::<docx_model::SectionPlacementWire>() + 32
        );

        let mut run = TextRun::default();
        run.no_break_hyphen_offsets = Vec::with_capacity(1);
        assert!(text_run(&run).unwrap_err().starts_with("UNSUPPORTED:"));
        let mut paragraph_value = DocParagraph::default();
        paragraph_value.runs.push(docx_model::DocRun::Break {
            break_type: docx_model::BreakType::Line,
        });
        assert!(paragraph(&paragraph_value)
            .unwrap_err()
            .starts_with("UNSUPPORTED:"));
    }
}
