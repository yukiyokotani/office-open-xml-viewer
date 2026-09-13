//! Per-cell projection of the bounded native table-style margin cascade.

use super::*;

pub(super) fn resolve(
    prepared: &mut [PreparedParagraph],
    index: &table_context::Index,
    formatting: &mut formatting::Formatting<'_>,
) -> Result<(), String> {
    for (table_id, table_context) in index.tables().iter().enumerate() {
        let first = table_context
            .rows
            .first()
            .and_then(|context| prepared.get(context.ttp_id))
            .map(|paragraph| &paragraph.table_properties.row);
        let unsupported_shape = table_context
            .rows
            .iter()
            .try_fold(false, |shaped, context| {
                let row = &prepared
                    .get(context.ttp_id)
                    .ok_or_else(|| unsupported("Word table margin TTP outside story"))?
                    .table_properties
                    .row;
                let geometry_varies = first.is_some_and(|first| {
                    row.origin() != first.origin()
                        || row.cells.len() != first.cells.len()
                        || row
                            .cells
                            .iter()
                            .zip(&first.cells)
                            .any(|(cell, first)| cell.width != first.width)
                });
                Ok::<_, String>(
                    shaped
                        || geometry_varies
                        || row.bidi
                        || row
                            .cells
                            .iter()
                            .any(|cell| cell.flags & (0x0003 | 0x0060) != 0),
                )
            })?;
        for (row_index, row_context) in table_context.rows.iter().enumerate() {
            let options = if let (Some(table_style), Some(flags)) =
                (row_context.table_style, row_context.table_style_options)
            {
                let (horizontal, vertical, presence) =
                    formatting.table_style_selector_profile(Some(table_style))?;
                Some(table_style_condition::Options::new(
                    flags, horizontal, vertical, presence,
                )?)
            } else {
                None
            };
            let conditional_margins =
                formatting.has_conditional_first_row_margins(row_context.table_style)?;
            let mut defaults = None;
            if row_context.source_cell_count > 63 {
                return Err(unsupported("Word table margin cell budget exceeded"));
            }
            let mut style_cells = [table::MarginPatch::default(); 63];
            for ordinal in 0..row_context.source_cell_count {
                let mut matches = if let Some(options) = options {
                    table_style_condition::select(
                        index,
                        table_id,
                        row_index,
                        table_style_condition::LogicalColumn {
                            ordinal,
                            count: row_context.source_cell_count,
                        },
                        options,
                    )?
                } else {
                    [None; 5]
                };
                if matches
                    .into_iter()
                    .flatten()
                    .any(|condition| condition == table_style_condition::FIRST_ROW)
                    && conditional_margins
                    && unsupported_shape
                {
                    formatting.unsupported_table_properties = true;
                    matches = [None; 5];
                }
                let key = formatting.table_formatting_key(
                    row_context.table_style,
                    row_context.table_style_options,
                    matches,
                )?;
                let (row_defaults, cell_patch) = formatting.table_cell_margins_for_key(key)?;
                if let Some(previous) = defaults {
                    if previous != row_defaults {
                        return Err(unsupported("Word table margin defaults vary within row"));
                    }
                } else {
                    defaults = Some(row_defaults);
                }
                style_cells[ordinal] = cell_patch;
            }
            let row = &mut prepared
                .get_mut(row_context.ttp_id)
                .ok_or_else(|| unsupported("Word table margin TTP outside story"))?
                .table_properties
                .row;
            if row.cells.len() != row_context.source_cell_count {
                return Err(unsupported("Word table margin cell count mismatch"));
            }
            row.resolve_style_aware_margins_by_cell(
                defaults.unwrap_or_default(),
                &style_cells[..row_context.source_cell_count],
            )?;
        }
    }
    Ok(())
}
