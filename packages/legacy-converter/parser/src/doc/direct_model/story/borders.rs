//! Native DOC table-border style resolution after row acquisition.

use super::{ModelBudget, PreparedParagraph};
use crate::doc::{formatting, table, table_context, table_style_condition, unsupported};

type BorderSides = [Option<table::PreparedBorder>; 6];

struct ActiveConditional {
    style: usize,
    options: u16,
    /// Column then row is the edge-condition portion of the required
    /// [MS-DOC] 2.4.6.6 order. The bounded admission below permits at most one
    /// of each family.
    patches: [Option<(u16, BorderSides)>; 2],
    presence: u16,
}

pub(super) fn resolve(
    prepared: &mut [PreparedParagraph],
    index: &table_context::Index,
    formatting: &mut formatting::Formatting<'_>,
    budget: &mut ModelBudget,
) -> Result<(), String> {
    for (table_id, context) in index.tables().iter().enumerate() {
        let conditional = active_conditional(context, formatting)?;
        let conditional = if let Some(active) = conditional {
            if supported_conditional_shape(prepared, context, active.style, active.options)? {
                Some(active)
            } else {
                formatting.unsupported_table_properties = true;
                None
            }
        } else {
            None
        };

        if let Some(active) = conditional {
            apply_conditional_regions(prepared, index, table_id, context, active, formatting)?;
        }
        for row_context in &context.rows {
            resolve_row(prepared, row_context, formatting, budget)?;
        }
    }
    Ok(())
}

fn active_conditional(
    context: &table_context::TableContext,
    formatting: &mut formatting::Formatting<'_>,
) -> Result<Option<ActiveConditional>, String> {
    for row in &context.rows {
        let Some(style) = row.table_style else {
            continue;
        };
        let Some((borders, presence)) = formatting.conditional_table_borders(Some(style))? else {
            continue;
        };
        let Some(options) = row.table_style_options else {
            continue;
        };
        let eligible_corners = [
            (table_style_condition::TOP_LEFT, (1 << 5) | (1 << 7)),
            (table_style_condition::TOP_RIGHT, (1 << 5) | (1 << 8)),
            (table_style_condition::BOTTOM_LEFT, (1 << 6) | (1 << 7)),
            (table_style_condition::BOTTOM_RIGHT, (1 << 6) | (1 << 8)),
        ]
        .into_iter()
        .any(|(condition, flags)| presence & condition != 0 && options & flags == flags);
        if eligible_corners {
            // A corner condition changes adjacent edge eligibility in the
            // shared selector. Native border controls do not yet cover that
            // cascade, so no partial edge patch may escape it.
            formatting.unsupported_table_properties = true;
            return Ok(None);
        }
        let conditions = [
            (table_style_condition::FIRST_COLUMN, 1 << 7),
            (table_style_condition::LAST_COLUMN, 1 << 8),
            (table_style_condition::FIRST_ROW, 1 << 5),
            (table_style_condition::LAST_ROW, 1 << 6),
        ];
        let active_edges = conditions
            .into_iter()
            .filter(|(condition, flag)| presence & condition != 0 && options & flag != 0)
            .fold(0, |mask, (condition, _)| mask | condition);
        let active_borders = conditions
            .into_iter()
            .enumerate()
            .filter(|(slot, (_, flag))| borders.present & (1 << slot) != 0 && options & flag != 0)
            .fold(0, |mask, (_, (condition, _))| mask | condition);
        if active_edges != active_borders {
            // DOC85: a selected condition from another property family must
            // not be silently dropped by the border-only path.
            formatting.unsupported_table_properties = true;
            return Ok(None);
        }
        if active_borders == 0 {
            continue;
        }
        if active_borders.count_ones() > 1
            && active_borders
                != table_style_condition::FIRST_COLUMN | table_style_condition::FIRST_ROW
        {
            // Native multi-condition controls currently cover one first-column
            // plus one first-row condition. Other family combinations remain
            // behind the admission gate.
            formatting.unsupported_table_properties = true;
            return Ok(None);
        }
        let patches = [
            (borders.present & 1 != 0 && options & (1 << 7) != 0)
                .then_some((table_style_condition::FIRST_COLUMN, borders.sides[0])),
            (borders.present & (1 << 2) != 0 && options & (1 << 5) != 0)
                .then_some((table_style_condition::FIRST_ROW, borders.sides[2])),
        ];
        if active_borders.count_ones() == 1 {
            let patch = conditions
                .into_iter()
                .enumerate()
                .find(|(slot, (_, flag))| borders.present & (1 << slot) != 0 && options & flag != 0)
                .map(|(slot, (condition, _))| (condition, borders.sides[slot]));
            return Ok(Some(ActiveConditional {
                style,
                options,
                patches: [patch, None],
                presence,
            }));
        }
        return Ok(Some(ActiveConditional {
            style,
            options,
            patches,
            presence,
        }));
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

fn apply_conditional_regions(
    prepared: &mut [PreparedParagraph],
    index: &table_context::Index,
    table_id: usize,
    context: &table_context::TableContext,
    active: ActiveConditional,
    formatting: &mut formatting::Formatting<'_>,
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
    let (horizontal, vertical, presence) =
        formatting.table_style_selector_profile(Some(active.style))?;
    debug_assert_eq!(presence, active.presence);
    let options =
        table_style_condition::Options::new(active.options, horizontal, vertical, presence)?;
    for (condition, sides) in active.patches.into_iter().flatten() {
        match condition {
            table_style_condition::FIRST_ROW | table_style_condition::LAST_ROW => {
                let row = if condition == table_style_condition::FIRST_ROW {
                    0
                } else {
                    rows - 1
                };
                for column in 0..columns {
                    if matches_condition(index, table_id, row, column, columns, options, condition)?
                    {
                        set(prepared, context, row, column, 0, sides[0])?;
                        set(prepared, context, row, column, 2, sides[2])?;
                    }
                }
                if matches_condition(index, table_id, row, 0, columns, options, condition)? {
                    set(prepared, context, row, 0, 1, sides[1])?;
                }
                if matches_condition(
                    index,
                    table_id,
                    row,
                    columns - 1,
                    columns,
                    options,
                    condition,
                )? {
                    set(prepared, context, row, columns - 1, 3, sides[3])?;
                }
                for column in 0..columns.saturating_sub(1) {
                    if matches_condition(index, table_id, row, column, columns, options, condition)?
                        && matches_condition(
                            index,
                            table_id,
                            row,
                            column + 1,
                            columns,
                            options,
                            condition,
                        )?
                    {
                        set(prepared, context, row, column, 3, sides[5])?;
                        set(prepared, context, row, column + 1, 1, sides[5])?;
                    }
                }
            }
            table_style_condition::FIRST_COLUMN | table_style_condition::LAST_COLUMN => {
                let column = if condition == table_style_condition::FIRST_COLUMN {
                    0
                } else {
                    columns - 1
                };
                for row in 0..rows {
                    if matches_condition(index, table_id, row, column, columns, options, condition)?
                    {
                        set(prepared, context, row, column, 1, sides[1])?;
                        set(prepared, context, row, column, 3, sides[3])?;
                    }
                }
                if matches_condition(index, table_id, 0, column, columns, options, condition)? {
                    set(prepared, context, 0, column, 0, sides[0])?;
                }
                if matches_condition(
                    index,
                    table_id,
                    rows - 1,
                    column,
                    columns,
                    options,
                    condition,
                )? {
                    set(prepared, context, rows - 1, column, 2, sides[2])?;
                }
                for row in 0..rows.saturating_sub(1) {
                    if matches_condition(index, table_id, row, column, columns, options, condition)?
                        && matches_condition(
                            index,
                            table_id,
                            row + 1,
                            column,
                            columns,
                            options,
                            condition,
                        )?
                    {
                        set(prepared, context, row, column, 2, sides[4])?;
                        set(prepared, context, row + 1, column, 0, sides[4])?;
                    }
                }
            }
            _ => return Err(unsupported("unsupported Word conditional border region")),
        }
    }
    Ok(())
}

fn matches_condition(
    index: &table_context::Index,
    table_id: usize,
    row: usize,
    column: usize,
    count: usize,
    options: table_style_condition::Options,
    condition: u16,
) -> Result<bool, String> {
    Ok(table_style_condition::select(
        index,
        table_id,
        row,
        table_style_condition::LogicalColumn {
            ordinal: column,
            count,
        },
        options,
    )?
    .contains(&Some(condition)))
}

fn set(
    prepared: &mut [PreparedParagraph],
    context: &table_context::TableContext,
    row_index: usize,
    column: usize,
    side: usize,
    value: Option<table::PreparedBorder>,
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
    // The ordered conditional cascade retains its final raw winner here.
    // resolve_row decodes and charges each retained border only once.
    cell.prepared_borders[side] = Some(value);
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
