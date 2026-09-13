//! Native DOC table-border style resolution after row acquisition.

use super::{ModelBudget, PreparedParagraph};
use crate::doc::{formatting, table, table_context, table_style_condition, unsupported};

pub(super) fn resolve(
    prepared: &mut [PreparedParagraph],
    index: &table_context::Index,
    formatting: &mut formatting::Formatting<'_>,
    budget: &mut ModelBudget,
) -> Result<(), String> {
    for context in index.tables() {
        let conditional = active_conditional(context, formatting)?;
        let conditional = if let Some((style, options, condition, sides)) = conditional {
            if supported_conditional_shape(prepared, context, style, options)? {
                Some((condition, sides))
            } else {
                formatting.unsupported_table_properties = true;
                None
            }
        } else {
            None
        };

        for row_context in &context.rows {
            resolve_row(prepared, row_context, formatting, budget)?;
        }
        if let Some((condition, sides)) = conditional {
            apply_conditional_region(prepared, context, condition, sides, budget)?;
        }
    }
    Ok(())
}

fn active_conditional(
    context: &table_context::TableContext,
    formatting: &mut formatting::Formatting<'_>,
) -> Result<Option<(usize, u16, u16, [Option<table::PreparedBorder>; 6])>, String> {
    for row in &context.rows {
        let Some(style) = row.table_style else {
            continue;
        };
        let Some((condition, sides, presence)) =
            formatting.conditional_table_borders(Some(style))?
        else {
            continue;
        };
        let Some(options) = row.table_style_options else {
            continue;
        };
        let enabled = match condition {
            table_style_condition::FIRST_ROW => options & (1 << 5) != 0,
            table_style_condition::LAST_ROW => options & (1 << 6) != 0,
            table_style_condition::FIRST_COLUMN => options & (1 << 7) != 0,
            table_style_condition::LAST_COLUMN => options & (1 << 8) != 0,
            _ => false,
        };
        if enabled {
            let active_edges = [
                (table_style_condition::FIRST_ROW, 1 << 5),
                (table_style_condition::LAST_ROW, 1 << 6),
                (table_style_condition::FIRST_COLUMN, 1 << 7),
                (table_style_condition::LAST_COLUMN, 1 << 8),
            ]
            .into_iter()
            .filter(|(candidate, flag)| presence & candidate != 0 && options & flag != 0)
            .fold(0, |mask, (candidate, _)| mask | candidate);
            if active_edges != condition {
                // Multiple active edge conditions can exclude one another on a
                // singleton and have an ordered overlap elsewhere. That cascade
                // is resolved only by the shared condition selector.
                formatting.unsupported_table_properties = true;
                return Ok(None);
            }
            return Ok(Some((style, options, condition, sides)));
        }
    }
    Ok(None)
}

fn supported_conditional_shape(
    prepared: &[PreparedParagraph],
    context: &table_context::TableContext,
    style: usize,
    options: u16,
) -> Result<bool, String> {
    let Some(first_context) = context.rows.first() else {
        return Ok(false);
    };
    let first = row(prepared, first_context)?;
    if first.cells.is_empty() {
        return Ok(false);
    }
    let count = first.cells.len();
    let origin = first.origin();
    let gap = first.gap;

    for (row_index, row_context) in context.rows.iter().enumerate() {
        if row_context.table_style != Some(style)
            || row_context.table_style_options != Some(options)
            || (row_index > 0 && row_context.header)
        {
            return Ok(false);
        }
        let row = row(prepared, row_context)?;
        if row.bidi
            || row.border_tistd_count > 1
            || row.origin() != origin
            || row.gap != gap
            || row.cells.len() != count
            || row.prepared_borders.iter().any(Option::is_some)
            || row.borders.iter().any(Option::is_some)
        {
            return Ok(false);
        }
        for (ordinal, cell) in row.cells.iter().enumerate() {
            if cell.width != first.cells[ordinal].width
                || cell.flags & 3 != 0
                || (cell.flags >> 5) & 3 != 0
                || cell.prepared_borders.iter().any(Option::is_some)
                || cell.borders.iter().any(Option::is_some)
            {
                return Ok(false);
            }
        }
    }
    Ok(true)
}

fn resolve_row(
    prepared: &mut [PreparedParagraph],
    context: &table_context::RowContext,
    formatting: &mut formatting::Formatting<'_>,
    budget: &mut ModelBudget,
) -> Result<(), String> {
    let style = formatting.table_borders(context.table_style)?;
    let row = row_mut(prepared, context)?;
    if row.cells.len() != context.source_cell_count {
        return Err(unsupported("Word table border cell count mismatch"));
    }

    let has_style = style.iter().any(Option::is_some);
    let has_direct = row.prepared_borders.iter().any(Option::is_some)
        || row
            .cells
            .iter()
            .any(|cell| cell.prepared_borders.iter().any(Option::is_some));
    let direct = row
        .prepared_borders
        .iter()
        .chain(
            row.cells
                .iter()
                .flat_map(|cell| cell.prepared_borders.iter()),
        )
        .flatten()
        .copied();
    let has_old_direct = direct.clone().any(table::PreparedBorder::is_old);
    let has_nil_direct = direct.clone().any(table::PreparedBorder::is_nil);
    let has_diagonal_direct = row
        .cells
        .iter()
        .any(|cell| cell.prepared_borders[4].is_some() || cell.prepared_borders[5].is_some());
    let has_tc80 = row
        .cells
        .iter()
        .any(|cell| cell.borders.iter().any(Option::is_some));
    let selected = context.table_style.is_some();
    let border_style_interaction = has_style || row.border_tistd_count > 0;
    if border_style_interaction
        && (has_tc80
            || has_old_direct
            || has_nil_direct
            || has_diagonal_direct
            || ((has_style || has_direct) && row.bidi)
            || ((has_style || has_direct) && row.border_tistd_count > 1))
    {
        formatting.unsupported_table_properties = true;
        return Ok(());
    }

    for side in 0..6 {
        if let Some(value) = row.prepared_borders[side].or(style[side]) {
            row.borders[side] = Some(materialize_border(value, budget)?);
        } else if selected {
            row.borders[side] = None;
        }
    }
    row.prepared_borders = [None; 6];
    for cell in &mut row.cells {
        for side in 0..6 {
            if let Some(value) = cell.prepared_borders[side] {
                cell.borders[side] = Some(materialize_border(value, budget)?);
            }
        }
        cell.prepared_borders = [None; 6];
    }
    Ok(())
}

fn apply_conditional_region(
    prepared: &mut [PreparedParagraph],
    context: &table_context::TableContext,
    condition: u16,
    sides: [Option<table::PreparedBorder>; 6],
    budget: &mut ModelBudget,
) -> Result<(), String> {
    // MS-DOC 2.6.3 defines the six border properties allowed in a TCnf. Word
    // 16.112.4 controls establish their region boundaries here: first/last-row
    // left/right and insideV, first/last-column top/bottom and insideH, with no
    // inside edge for a singleton. This mapper is limited to the rectangular,
    // unmerged LTR shape checked above; it is not a general per-cell rule.
    let rows = context.rows.len();
    let columns = context
        .rows
        .first()
        .ok_or_else(|| unsupported("Word conditional border table has no rows"))?
        .source_cell_count;
    match condition {
        table_style_condition::FIRST_ROW | table_style_condition::LAST_ROW => {
            let row = if condition == table_style_condition::FIRST_ROW {
                0
            } else {
                rows - 1
            };
            for column in 0..columns {
                set(prepared, context, row, column, 0, sides[0], budget)?;
                set(prepared, context, row, column, 2, sides[2], budget)?;
            }
            set(prepared, context, row, 0, 1, sides[1], budget)?;
            set(prepared, context, row, columns - 1, 3, sides[3], budget)?;
            for column in 0..columns.saturating_sub(1) {
                set(prepared, context, row, column, 3, sides[5], budget)?;
                set(prepared, context, row, column + 1, 1, sides[5], budget)?;
            }
        }
        table_style_condition::FIRST_COLUMN | table_style_condition::LAST_COLUMN => {
            let column = if condition == table_style_condition::FIRST_COLUMN {
                0
            } else {
                columns - 1
            };
            for row in 0..rows {
                set(prepared, context, row, column, 1, sides[1], budget)?;
                set(prepared, context, row, column, 3, sides[3], budget)?;
            }
            set(prepared, context, 0, column, 0, sides[0], budget)?;
            set(prepared, context, rows - 1, column, 2, sides[2], budget)?;
            for row in 0..rows.saturating_sub(1) {
                set(prepared, context, row, column, 2, sides[4], budget)?;
                set(prepared, context, row + 1, column, 0, sides[4], budget)?;
            }
        }
        _ => return Err(unsupported("unsupported Word conditional border region")),
    }
    Ok(())
}

fn set(
    prepared: &mut [PreparedParagraph],
    context: &table_context::TableContext,
    row_index: usize,
    column: usize,
    side: usize,
    value: Option<table::PreparedBorder>,
    budget: &mut ModelBudget,
) -> Result<(), String> {
    let Some(value) = value else {
        return Ok(());
    };
    let row_context = context
        .rows
        .get(row_index)
        .ok_or_else(|| unsupported("Word conditional border row outside table"))?;
    let cell = row_mut(prepared, row_context)?
        .cells
        .get_mut(column)
        .ok_or_else(|| unsupported("Word conditional border cell outside row"))?;
    cell.borders[side] = Some(materialize_border(value, budget)?);
    Ok(())
}

fn row<'a>(
    prepared: &'a [PreparedParagraph],
    context: &table_context::RowContext,
) -> Result<&'a table::Row, String> {
    Ok(&prepared
        .get(context.ttp_id)
        .ok_or_else(|| unsupported("Word table border TTP outside story"))?
        .table_properties
        .row)
}

fn row_mut<'a>(
    prepared: &'a mut [PreparedParagraph],
    context: &table_context::RowContext,
) -> Result<&'a mut table::Row, String> {
    Ok(&mut prepared
        .get_mut(context.ttp_id)
        .ok_or_else(|| unsupported("Word table border TTP outside story"))?
        .table_properties
        .row)
}

pub(super) fn materialize_border(
    value: table::PreparedBorder,
    budget: &mut ModelBudget,
) -> Result<crate::doc::border::Border, String> {
    let border = value.decode()?;
    // This allocation happens after Properties::retained_heap_bytes. Charge
    // its actual retained String capacities before it enters the prepared
    // story; an error drops this local value and then the whole model.
    budget.charge(border.retained_bytes()?)?;
    Ok(border)
}
