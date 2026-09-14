//! Gate regressions for conditional table-style margins.
//!
//! Word 16.112.4 DOC93-DOC95 controls establish isolated serialization, while
//! New138-new145 bordered controls isolate D63E from D634 and establish one
//! bounded conditional path: a single non-inherited FIRST_ROW D63E record.

use super::*;

fn empty<'a>() -> Formatting<'a> {
    Formatting {
        characters: Default::default(),
        paragraphs: Default::default(),
        fonts: vec![],
        defaults: Properties::default(),
        styles: vec![],
        paragraph_cache: BTreeMap::new(),
        paragraph_marker_styles: BTreeMap::new(),
        paragraph_layout_cache: BTreeMap::new(),
        table_style_cache: BTreeMap::new(),
        effective_nfib: 0x0112,
        interpret_table_styles: true,
        data: &[],
        budget: Default::default(),
        numbering: Default::default(),
        numbering_output: Default::default(),
        unsupported_character_properties: false,
        unsupported_paragraph_properties: false,
        unsupported_piece_properties: false,
        missing_tables: true,
        unsupported_table_properties: false,
    }
}

fn formatting(tapx: &[u8]) -> Formatting<'_> {
    let mut value = empty();
    value.styles.push(Some(super::super::Style {
        base: 0x0fff,
        kind: 3,
        chpx: &[],
        papx: &[],
        table: Some(super::super::TableStylePropertySets {
            tapx,
            papx: &[0, 0],
            chpx: &[],
        }),
        language_compatibility: Default::default(),
    }));
    value
}

fn margin(code: u16, sides: u8, width: u16) -> Vec<u8> {
    let mut value = Vec::from(code.to_le_bytes());
    value.extend([6, 0, 1, sides, 3]);
    value.extend(width.to_le_bytes());
    value
}

fn first_row(nested: &[u8]) -> Vec<u8> {
    let mut value = Vec::from(0xd66au16.to_le_bytes());
    value.push(u8::try_from(2 + nested.len()).unwrap());
    value.extend(crate::doc::table_style_condition::FIRST_ROW.to_le_bytes());
    value.extend(nested);
    value
}

#[test]
fn conditional_d634_stays_gated_in_both_record_orders() {
    let unconditional = margin(0xd634, 0x02, 72);
    let conditional = first_row(&margin(0xd634, 0x02, 288));

    let mut results = Vec::new();
    for tapx in [
        [unconditional.as_slice(), conditional.as_slice()].concat(),
        [conditional.as_slice(), unconditional.as_slice()].concat(),
    ] {
        let mut value = formatting(&tapx);
        let margins = value.table_cell_margins(Some(0)).unwrap();
        results.push((margins, value.unsupported_table_properties));
    }

    assert_eq!(results[0], results[1]);
    let mut expected = table::MarginPatch::default();
    expected.apply_style(0xd634, &unconditional[2..]).unwrap();
    assert_eq!(results[0].0 .0, expected);
    assert_eq!(results[0].0 .1, table::MarginPatch::default());
    assert!(results[0].1);
}

#[test]
fn unconditional_d63e_overrides_d634_in_both_record_orders() {
    for tapx in [
        [
            margin(0xd634, 0x02, 288).as_slice(),
            margin(0xd63e, 0x02, 72).as_slice(),
        ]
        .concat(),
        [
            margin(0xd63e, 0x02, 72).as_slice(),
            margin(0xd634, 0x02, 288).as_slice(),
        ]
        .concat(),
    ] {
        let mut value = formatting(&tapx);
        let (defaults, cells) = value.table_cell_margins(Some(0)).unwrap();
        let mut row = table::Row::default();
        row.apply(0x7621, &[0, 1, 0xe8, 3]).unwrap();
        row.resolve_style_aware_margins(defaults, cells);

        assert_eq!(row.cells[0].margins[1], Some(72));
        assert!(!value.unsupported_table_properties);
    }
}

#[test]
fn conditional_d63e_with_overlapping_d634_stays_gated() {
    for (side, width) in [(0x01, 216), (0x02, 288), (0x04, 360), (0x08, 432)] {
        let unconditional = margin(0xd634, 0x0f, 72);
        let conditional = first_row(&margin(0xd63e, side, width));
        let tapx = [unconditional.as_slice(), conditional.as_slice()].concat();
        let mut value = formatting(&tapx);

        let (defaults, cells) = value.table_cell_margins(Some(0)).unwrap();

        let mut expected = table::MarginPatch::default();
        expected.apply_style(0xd634, &unconditional[2..]).unwrap();
        assert_eq!(defaults, expected);
        assert_eq!(cells, table::MarginPatch::default());
        assert!(value.unsupported_table_properties);
    }
}

#[test]
fn conditional_d63e_accepts_each_physical_side_over_d63e_baseline() {
    for (side, width) in [(0x01, 216), (0x02, 288), (0x04, 360), (0x08, 432)] {
        let unconditional = margin(0xd63e, 0x0f, 72);
        let conditional = first_row(&margin(0xd63e, side, width));
        let tapx = [unconditional.as_slice(), conditional.as_slice()].concat();
        let mut value = formatting(&tapx);
        let key = TableFormattingKey {
            selected_style: 0,
            matches: [
                Some(crate::doc::table_style_condition::FIRST_ROW),
                None,
                None,
                None,
                None,
            ],
        };

        let (_, cells) = value.table_cell_margins_for_key(Some(key)).unwrap();

        let index = side.trailing_zeros() as usize;
        assert_eq!(cells.get(index).map(|value| value.resolved()), Some(width));
        assert!(!value.unsupported_table_properties);
    }
}

#[test]
fn single_first_row_d63e_overlays_unconditional_d63e_per_side() {
    let unconditional = margin(0xd63e, 0x0f, 72);
    let conditional = first_row(&margin(0xd63e, 0x05, 288));
    let tapx = [unconditional.as_slice(), conditional.as_slice()].concat();
    let mut value = formatting(&tapx);
    let key = TableFormattingKey {
        selected_style: 0,
        matches: [
            Some(crate::doc::table_style_condition::FIRST_ROW),
            None,
            None,
            None,
            None,
        ],
    };

    let (_, matched) = value.table_cell_margins_for_key(Some(key)).unwrap();
    assert_eq!(matched.get(0).map(|value| value.resolved()), Some(288));
    assert_eq!(matched.get(1).map(|value| value.resolved()), Some(72));
    assert_eq!(matched.get(2).map(|value| value.resolved()), Some(288));
    assert_eq!(matched.get(3).map(|value| value.resolved()), Some(72));
    assert!(!value.unsupported_table_properties);

    let key = TableFormattingKey::unconditional(0);
    let (_, unmatched) = value.table_cell_margins_for_key(Some(key)).unwrap();
    for side in 0..4 {
        assert_eq!(unmatched.get(side).map(|value| value.resolved()), Some(72));
    }
}

#[test]
fn repeated_first_row_d63e_remains_gated() {
    let tapx = [
        first_row(&margin(0xd63e, 0x01, 216)),
        first_row(&margin(0xd63e, 0x02, 288)),
    ]
    .concat();
    let mut value = formatting(&tapx);

    let _ = value.table_cell_margins(Some(0)).unwrap();

    assert!(value.unsupported_table_properties);
}
