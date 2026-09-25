//! Native DOC table-border style resolution after row acquisition.

use super::{ModelBudget, PreparedParagraph};
use crate::doc::{formatting, table, table_context, table_style_condition, unsupported};

type BorderSides = [Option<table::PreparedBorder>; 6];

const EDGE_CONDITIONS: [(u16, u16); 4] = [
    (table_style_condition::FIRST_COLUMN, 1 << 7),
    (table_style_condition::LAST_COLUMN, 1 << 8),
    (table_style_condition::FIRST_ROW, 1 << 5),
    (table_style_condition::LAST_ROW, 1 << 6),
];

struct ActiveConditional {
    style: usize,
    options: u16,
    /// Column then row is the edge-condition portion of the required
    /// [MS-DOC] 2.4.6.6 order. The bounded admission below permits at most one
    /// of each family.
    patches: [Option<(u16, BorderSides)>; 2],
    presence: u16,
}

#[derive(Clone, Copy)]
struct SourceSpans {
    values: [(usize, usize); 63],
    len: usize,
}

impl SourceSpans {
    fn parse(row: &table::Row) -> Option<Self> {
        if row.cells.is_empty() || row.cells.len() > 63 {
            return None;
        }
        let mut values = [(0, 0); 63];
        let mut len = 0;
        let mut start = 0;
        while start < row.cells.len() {
            let flags = row.cells[start].flags & 3;
            if flags == 1 {
                return None;
            }
            let mut end = start + 1;
            if flags >= 2 {
                while end < row.cells.len() && row.cells[end].flags & 3 == 1 {
                    end += 1;
                }
                if end == start + 1 {
                    return None;
                }
            }
            values[len] = (start, end);
            len += 1;
            start = end;
        }
        Some(Self { values, len })
    }

    fn as_slice(&self) -> &[(usize, usize)] {
        &self.values[..self.len]
    }
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
            let conditions = active
                .patches
                .iter()
                .flatten()
                .fold(0, |mask, (condition, _)| mask | condition);
            if supported_conditional_shape(
                prepared,
                context,
                active.style,
                active.options,
                conditions,
            )? {
                Some(active)
            } else {
                formatting.unsupported_table_properties = true;
                None
            }
        } else {
            None
        };

        if let Some(active) = conditional {
            let merged_geometry = has_horizontal_merge(prepared, context)?;
            apply_conditional_regions(prepared, index, table_id, context, active, formatting)?;
            if merged_geometry {
                // DOC150 shows that these newly projected horizontal merged
                // shapes require native grid redistribution. Keep the border
                // ownership result for validation while gating this conditional
                // projection independently of TIstd/TTlp admission.
                formatting.unsupported_table_properties = true;
            }
        }
        for row_context in &context.rows {
            resolve_row(prepared, row_context, formatting, budget)?;
        }
    }
    Ok(())
}

fn has_horizontal_merge(
    prepared: &[PreparedParagraph],
    context: &table_context::TableContext,
) -> Result<bool, String> {
    for row_context in &context.rows {
        if source_row(prepared, row_context)?
            .cells
            .iter()
            .any(|cell| cell.flags & 3 != 0)
        {
            return Ok(true);
        }
    }
    Ok(false)
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
        let active_edges = EDGE_CONDITIONS
            .into_iter()
            .filter(|(condition, flag)| presence & condition != 0 && options & flag != 0)
            .fold(0, |mask, (condition, _)| mask | condition);
        let active_borders = EDGE_CONDITIONS
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
        if !supported_edge_combination(active_borders) {
            // The bounded mapper composes at most one column condition and one
            // row condition. Same-family and three-or-more-condition cascades
            // remain behind the admission gate.
            formatting.unsupported_table_properties = true;
            return Ok(None);
        }
        let mut patches = [None; 2];
        for (target, condition) in selected_edge_conditions(active_borders)
            .into_iter()
            .flatten()
            .enumerate()
        {
            let slot = EDGE_CONDITIONS
                .iter()
                .position(|(candidate, _)| *candidate == condition)
                .expect("selected edge condition");
            patches[target] = Some((condition, borders.sides[slot]));
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

fn supported_edge_combination(mask: u16) -> bool {
    let columns = mask & (table_style_condition::FIRST_COLUMN | table_style_condition::LAST_COLUMN);
    let rows = mask & (table_style_condition::FIRST_ROW | table_style_condition::LAST_ROW);
    mask.count_ones() <= 1
        || (mask.count_ones() == 2 && columns.count_ones() == 1 && rows.count_ones() == 1)
}

fn selected_edge_conditions(mask: u16) -> [Option<u16>; 2] {
    let mut selected = [None; 2];
    let mut target = 0;
    // [MS-DOC] 2.4.6.6 applies first/last column before first/last row,
    // independent of the serialized order of their sprmTCnf records.
    for (condition, _) in EDGE_CONDITIONS {
        if mask & condition != 0 {
            selected[target] = Some(condition);
            target += 1;
        }
    }
    selected
}

fn supported_conditional_shape(
    prepared: &[PreparedParagraph],
    context: &table_context::TableContext,
    style: usize,
    options: u16,
    conditions: u16,
) -> Result<bool, String> {
    let Some(first_context) = context.rows.first() else {
        return Ok(false);
    };
    let first = source_row(prepared, first_context)?;
    if first.cells.is_empty() {
        return Ok(false);
    }
    let count = first.cells.len();
    let Some(first_spans) = SourceSpans::parse(first) else {
        return Ok(false);
    };
    // The native combined controls establish source-span ownership only for
    // FIRST_ROW. Column conditions and other row scopes remain gated when a
    // horizontal merge is present; their unmerged behavior is unchanged.
    if first_spans.len < count && conditions != table_style_condition::FIRST_ROW {
        return Ok(false);
    }
    let origin = first.origin();
    let gap = first.gap;

    for (row_index, row_context) in context.rows.iter().enumerate() {
        if row_context.table_style != Some(style)
            || row_context.table_style_options != Some(options)
            || (row_index > 0 && row_context.header)
        {
            return Ok(false);
        }
        let row = source_row(prepared, row_context)?;
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
        let Some(spans) = SourceSpans::parse(row) else {
            return Ok(false);
        };
        if spans.as_slice() != first_spans.as_slice() {
            return Ok(false);
        }
        for (ordinal, cell) in row.cells.iter().enumerate() {
            if cell.width != first.cells[ordinal].width
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

#[allow(clippy::needless_range_loop)] // One side index selects three parallel arrays.
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
    // [MS-DOC] 2.9.20/2.9.157: a direct NilBrc states that the cells have no
    // border on that side. It is an ordinary direct value above the table
    // style (projected as an explicit "nil" edge), and a Nil diagonal is the
    // same as the default absence of a diagonal. Cell diagonals (sprmTSetBrc
    // 0x10/0x20) are direct cell values that no table style supplies, so
    // they project as ECMA-376 tl2br/tr2bl over the style like any other
    // direct cell value.
    let has_tc80 = row
        .cells
        .iter()
        .any(|cell| cell.borders.iter().any(Option::is_some));
    let selected = context.table_style.is_some();
    let border_style_interaction = has_style || row.border_tistd_count > 0;
    // A selected style without border values leaves right-to-left rows with
    // the same direct-border cascade as an unstyled row; only style borders
    // under RTL remain unverified.
    if border_style_interaction
        && (has_tc80
            || has_old_direct
            || (has_style && row.bidi)
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
    // inside edge for a singleton. Horizontal merges retain source-slot
    // ownership: the start owns top/left/bottom, the end owns right, and
    // insideV is emitted only between visible spans. The admission check keeps
    // this limited to rectangular, equal-grid LTR rows.
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
                let spans = SourceSpans::parse(source_row(prepared, &context.rows[row])?)
                    .ok_or_else(|| unsupported("invalid Word horizontal merge spans"))?;
                for &(start, _) in spans.as_slice() {
                    if matches_condition(index, table_id, row, start, columns, options, condition)?
                    {
                        set(prepared, context, row, start, 0, sides[0])?;
                        set(prepared, context, row, start, 2, sides[2])?;
                    }
                }
                let &(first_start, _) = spans.as_slice().first().expect("nonempty spans");
                if matches_condition(
                    index,
                    table_id,
                    row,
                    first_start,
                    columns,
                    options,
                    condition,
                )? {
                    set(prepared, context, row, first_start, 1, sides[1])?;
                }
                let &(last_start, last_end) = spans.as_slice().last().expect("nonempty spans");
                if matches_condition(
                    index,
                    table_id,
                    row,
                    last_end - 1,
                    columns,
                    options,
                    condition,
                )? {
                    set(prepared, context, row, last_start, 3, sides[3])?;
                }
                for pair in spans.as_slice().windows(2) {
                    let (left_start, left_end) = pair[0];
                    let (right_start, _) = pair[1];
                    if matches_condition(
                        index,
                        table_id,
                        row,
                        left_end - 1,
                        columns,
                        options,
                        condition,
                    )? && matches_condition(
                        index,
                        table_id,
                        row,
                        right_start,
                        columns,
                        options,
                        condition,
                    )? {
                        set(prepared, context, row, left_start, 3, sides[5])?;
                        set(prepared, context, row, right_start, 1, sides[5])?;
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

fn source_row<'a>(
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn source_span_mapping_is_fixed_size_and_rejects_malformed_chains() {
        let mut row = table::Row::default();
        row.cells = vec![table::Cell::default(); 63];
        assert_eq!(SourceSpans::parse(&row).unwrap().as_slice().len(), 63);

        row.cells.push(table::Cell::default());
        assert!(SourceSpans::parse(&row).is_none());

        row.cells.truncate(3);
        row.cells[0].flags = 2;
        row.cells[1].flags = 1;
        assert_eq!(
            SourceSpans::parse(&row).unwrap().as_slice(),
            [(0, 2), (2, 3)]
        );
        row.cells[0].flags = 0;
        assert!(SourceSpans::parse(&row).is_none());
        row.cells[0].flags = 2;
        row.cells[1].flags = 0;
        assert!(SourceSpans::parse(&row).is_none());
    }
}
