//! DOC table-style conditional selection, [MS-DOC] 2.4.6.6, 2.9.41, 2.9.69,
//! and 2.9.326. This is intentionally DOC-local: its specified application order
//! differs from the OOXML conditional-style order. It neither resolves complete
//! styles nor replays historical table auto-formatting.

pub(super) use super::table_context::LogicalColumn;
use super::{table_context::Index, unsupported};

pub(super) const HORIZONTAL_ODD: u16 = 0x0040;
pub(super) const HORIZONTAL_EVEN: u16 = 0x0080;
pub(super) const VERTICAL_ODD: u16 = 0x0010;
pub(super) const VERTICAL_EVEN: u16 = 0x0020;
pub(super) const FIRST_COLUMN: u16 = 0x0004;
pub(super) const LAST_COLUMN: u16 = 0x0008;
pub(super) const FIRST_ROW: u16 = 0x0001;
pub(super) const LAST_ROW: u16 = 0x0002;
pub(super) const TOP_RIGHT: u16 = 0x0100;
pub(super) const TOP_LEFT: u16 = 0x0200;
pub(super) const BOTTOM_RIGHT: u16 = 0x0400;
pub(super) const BOTTOM_LEFT: u16 = 0x0800;

#[derive(Clone, Copy)]
pub(super) struct Options {
    first_row: bool,
    last_row: bool,
    first_column: bool,
    last_column: bool,
    horizontal_band: Option<u8>,
    vertical_band: Option<u8>,
}

impl Options {
    /// Band operands follow [MS-DOC] 2.6.3 (0x3488/0x3489): 1..=3,
    /// with no banding when absent. Historical auto-format bits and Fatl
    /// padding are not live table-style conditions.
    pub(super) fn new(
        grfatl: u16,
        horizontal_band: Option<u8>,
        vertical_band: Option<u8>,
    ) -> Result<Self, String> {
        validate_band(horizontal_band)?;
        validate_band(vertical_band)?;
        Ok(Self {
            first_row: grfatl & (1 << 5) != 0,
            last_row: grfatl & (1 << 6) != 0,
            first_column: grfatl & (1 << 7) != 0,
            last_column: grfatl & (1 << 8) != 0,
            horizontal_band: (grfatl & (1 << 9) == 0)
                .then_some(horizontal_band)
                .flatten(),
            vertical_band: (grfatl & (1 << 10) == 0).then_some(vertical_band).flatten(),
        })
    }
}

fn validate_band(value: Option<u8>) -> Result<(), String> {
    if value.is_some_and(|size| !(1..=3).contains(&size)) {
        return Err(unsupported("invalid Word table style band size"));
    }
    Ok(())
}

/// Matching CNFC values in MS-DOC's required application order: horizontal
/// band, vertical band, column, row, then corner. The checked table-context
/// lookup supplies source-cell ordinals for supported LTR, unmerged rows,
/// including ragged rows and rows whose cell widths differ. Selection allocates
/// nothing.
pub(super) fn select(
    index: &Index,
    table_id: usize,
    row_index: usize,
    column: LogicalColumn,
    options: Options,
) -> Result<[Option<u16>; 5], String> {
    let table = index
        .tables()
        .get(table_id)
        .ok_or_else(|| unsupported("Word table style table outside context index"))?;
    let row = table
        .rows
        .get(row_index)
        .ok_or_else(|| unsupported("Word table style row outside context index"))?;
    if column.count == 0 || column.ordinal >= column.count {
        return Err(unsupported("invalid Word table logical column position"));
    }

    let actual_top = row_index == 0;
    let actual_bottom = row_index + 1 == table.rows.len();
    let logical_left = column.ordinal == 0;
    let logical_right = column.ordinal + 1 == column.count;
    let first_row = options.first_row && (actual_top || row.header);
    let last_row = options.last_row && actual_bottom && !first_row;
    let first_column = options.first_column && logical_left;
    let last_column = options.last_column && logical_right && !first_column;

    let horizontal = if first_row || last_row {
        None
    } else {
        let ordinal = if options.first_row {
            row.preceding_body_rows
        } else {
            row_index
        };
        band(
            ordinal,
            options.horizontal_band,
            HORIZONTAL_ODD,
            HORIZONTAL_EVEN,
        )
    };
    let vertical = if first_column || last_column {
        None
    } else {
        let ordinal = column.ordinal - usize::from(options.first_column);
        band(ordinal, options.vertical_band, VERTICAL_ODD, VERTICAL_EVEN)
    };
    let column_match = first_column
        .then_some(FIRST_COLUMN)
        .or_else(|| last_column.then_some(LAST_COLUMN));
    let row_match = first_row
        .then_some(FIRST_ROW)
        .or_else(|| last_row.then_some(LAST_ROW));

    let corner = if actual_top && logical_left && options.first_row && options.first_column {
        Some(TOP_LEFT)
    } else if actual_top && logical_right && options.first_row && options.last_column {
        Some(TOP_RIGHT)
    } else if actual_bottom && logical_left && options.last_row && options.first_column {
        Some(BOTTOM_LEFT)
    } else if actual_bottom && logical_right && options.last_row && options.last_column {
        Some(BOTTOM_RIGHT)
    } else {
        None
    };
    Ok([horizontal, vertical, column_match, row_match, corner])
}

fn band(ordinal: usize, size: Option<u8>, odd: u16, even: u16) -> Option<u16> {
    let band = ordinal / usize::from(size?);
    Some(if band % 2 == 0 { odd } else { even })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::table::{Cell, Properties};

    const ALL: u16 = 0b1111 << 5;

    fn properties(row_end: bool, header: bool) -> Properties {
        let mut value = Properties::default();
        value.apply(0x6649, &1u32.to_le_bytes()).unwrap();
        value.row_end = row_end;
        value.row.header = header;
        if row_end {
            value.row.cells = vec![Cell {
                width: 10,
                ..Cell::default()
            }];
        }
        value
    }

    fn index(headers: &[bool]) -> Index {
        let mut paragraphs = Vec::new();
        for header in headers {
            paragraphs.push((properties(false, false), '\u{7}'));
            paragraphs.push((properties(true, *header), '\u{7}'));
        }
        Index::build(paragraphs.len(), paragraphs, &mut |_| Ok(())).unwrap()
    }

    fn column(ordinal: usize) -> LogicalColumn {
        LogicalColumn { ordinal, count: 3 }
    }

    #[test]
    fn matches_in_doc_application_order_with_row_column_and_corner_precedence() {
        let index = index(&[false, false, true, false, true]);
        let options = Options::new(ALL, Some(1), Some(1)).unwrap();
        assert_eq!(
            select(&index, 0, 0, column(0), options).unwrap(),
            [
                None,
                None,
                Some(FIRST_COLUMN),
                Some(FIRST_ROW),
                Some(TOP_LEFT)
            ]
        );
        assert_eq!(
            select(&index, 0, 0, column(2), options).unwrap(),
            [
                None,
                None,
                Some(LAST_COLUMN),
                Some(FIRST_ROW),
                Some(TOP_RIGHT)
            ]
        );
        assert_eq!(
            select(&index, 0, 1, column(1), options).unwrap(),
            [Some(HORIZONTAL_ODD), Some(VERTICAL_ODD), None, None, None]
        );
        assert_eq!(
            select(&index, 0, 2, column(1), options).unwrap(),
            [None, Some(VERTICAL_ODD), None, Some(FIRST_ROW), None]
        );
        assert_eq!(
            select(&index, 0, 3, column(1), options).unwrap(),
            [Some(HORIZONTAL_EVEN), Some(VERTICAL_ODD), None, None, None]
        );
        // A final header matches first-row formatting, while corner geometry is
        // still the actual bottom-right position specified for CNFC 0x0400.
        assert_eq!(
            select(&index, 0, 4, column(2), options).unwrap(),
            [
                None,
                None,
                Some(LAST_COLUMN),
                Some(FIRST_ROW),
                Some(BOTTOM_RIGHT)
            ]
        );
    }

    #[test]
    fn band_sizes_one_two_and_three_use_one_based_odd_then_even_bands() {
        let index = index(&[false, false, false, false, false, false, false]);
        for (size, row, expected) in [
            (1, 1, HORIZONTAL_EVEN),
            (2, 2, HORIZONTAL_EVEN),
            (3, 4, HORIZONTAL_EVEN),
        ] {
            let options = Options::new(0, Some(size), Some(size)).unwrap();
            assert_eq!(
                select(&index, 0, row, column(1), options).unwrap()[0],
                Some(expected)
            );
        }
        for (size, ordinal, expected) in [
            (1, 1, VERTICAL_EVEN),
            (2, 1, VERTICAL_ODD),
            (3, 2, VERTICAL_ODD),
        ] {
            let options = Options::new(0, Some(size), Some(size)).unwrap();
            assert_eq!(
                select(&index, 0, 1, column(ordinal), options).unwrap()[1],
                Some(expected)
            );
        }
    }

    #[test]
    fn singleton_dimensions_follow_exclusions_and_corner_precedence() {
        let index = index(&[false]);
        let options = Options::new(ALL, Some(1), Some(1)).unwrap();
        assert_eq!(
            select(
                &index,
                0,
                0,
                LogicalColumn {
                    ordinal: 0,
                    count: 1
                },
                options
            )
            .unwrap(),
            [
                None,
                None,
                Some(FIRST_COLUMN),
                Some(FIRST_ROW),
                Some(TOP_LEFT)
            ]
        );
    }

    #[test]
    fn corner_matrix_matches_spec_exclusions_for_all_boundary_flag_combinations() {
        fn expected(top: bool, bottom: bool, left: bool, right: bool, flags: u16) -> Option<u16> {
            let first_row = flags & (1 << 5) != 0;
            let last_row = flags & (1 << 6) != 0;
            let first_column = flags & (1 << 7) != 0;
            let last_column = flags & (1 << 8) != 0;
            let top_left = top && left && first_row && first_column;
            let top_right = top && right && first_row && last_column && !top_left;
            let bottom_left = bottom && left && last_row && first_column && !top_right && !top_left;
            let bottom_right = bottom
                && right
                && last_row
                && last_column
                && !top_right
                && !top_left
                && !bottom_left;
            if top_left {
                Some(TOP_LEFT)
            } else if top_right {
                Some(TOP_RIGHT)
            } else if bottom_left {
                Some(BOTTOM_LEFT)
            } else if bottom_right {
                Some(BOTTOM_RIGHT)
            } else {
                None
            }
        }

        for (rows, columns) in [(1, 1), (1, 3), (3, 1), (3, 3)] {
            let index = index(&vec![false; rows]);
            for flag_bits in 0u16..16 {
                let flags = flag_bits << 5;
                let options = Options::new(flags, None, None).unwrap();
                for row in 0..rows {
                    for column in 0..columns {
                        let actual = select(
                            &index,
                            0,
                            row,
                            LogicalColumn {
                                ordinal: column,
                                count: columns,
                            },
                            options,
                        )
                        .unwrap()[4];
                        assert_eq!(
                            actual,
                            expected(
                                row == 0,
                                row + 1 == rows,
                                column == 0,
                                column + 1 == columns,
                                flags,
                            ),
                            "rows={rows} columns={columns} row={row} column={column} flags={flags:#x}"
                        );
                    }
                }
            }
        }
    }

    #[test]
    fn disabled_or_absent_bands_do_not_match_and_padding_is_ignored() {
        let index = index(&[false, false, false]);
        let disabled = Options::new((1 << 9) | (1 << 10), Some(1), Some(1)).unwrap();
        assert_eq!(
            select(&index, 0, 1, column(1), disabled).unwrap(),
            [None; 5]
        );
        let absent = Options::new(0, None, None).unwrap();
        assert_eq!(select(&index, 0, 1, column(1), absent).unwrap(), [None; 5]);
        let padded = Options::new(0xf800, Some(1), Some(1)).unwrap();
        let plain = Options::new(0, Some(1), Some(1)).unwrap();
        assert_eq!(
            select(&index, 0, 1, column(1), padded).unwrap(),
            select(&index, 0, 1, column(1), plain).unwrap()
        );
    }

    #[test]
    fn header_rows_count_normally_when_first_row_formatting_is_disabled() {
        let index = index(&[true, false, true, false, true]);
        let options = Options::new(0x001f, Some(1), None).unwrap();
        let expected = [
            HORIZONTAL_ODD,
            HORIZONTAL_EVEN,
            HORIZONTAL_ODD,
            HORIZONTAL_EVEN,
            HORIZONTAL_ODD,
        ];
        for (row, condition) in expected.into_iter().enumerate() {
            assert_eq!(
                select(&index, 0, row, column(0), options).unwrap(),
                [Some(condition), None, None, None, None]
            );
        }
    }

    #[test]
    fn rejects_invalid_band_sizes_and_context_positions() {
        for size in [0, 4, 255] {
            assert!(Options::new(0, Some(size), None).is_err());
            assert!(Options::new(0, None, Some(size)).is_err());
        }
        let index = index(&[false]);
        let options = Options::new(0, None, None).unwrap();
        assert!(select(&index, 1, 0, column(0), options).is_err());
        assert!(select(&index, 0, 1, column(0), options).is_err());
        for invalid in [
            LogicalColumn {
                ordinal: 0,
                count: 0,
            },
            LogicalColumn {
                ordinal: 3,
                count: 3,
            },
        ] {
            assert!(select(&index, 0, 0, invalid, options).is_err());
        }
    }
}
