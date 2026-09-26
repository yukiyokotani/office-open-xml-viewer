//! Full-container acceptance fixtures for direct binary table projection.

use super::tests::{source_with_typography, with_numbering, with_picture_data};
use crate::cfb::{test_support::build_scoped_cfb, CompoundFile};
use crate::doc::table_structure::Payload;
use docx_model::{BodyElement, CellElement, DocRun};

fn sprm(code: u16, operand: &[u8], variable: bool) -> Vec<u8> {
    let mut value = code.to_le_bytes().to_vec();
    if variable {
        value.push(operand.len() as u8);
    }
    value.extend_from_slice(operand);
    value
}

fn cell() -> Vec<u8> {
    sprm(0x2416, &[1], false)
}
fn row(width: u16) -> Vec<u8> {
    [
        sprm(0x2416, &[1], false),
        sprm(0x2417, &[1], false),
        sprm(0x7621, &[0, 1, width as u8, (width >> 8) as u8], false),
        sprm(0x2416, &[1], false),
    ]
    .concat()
}
fn large_row() -> Vec<u8> {
    [
        sprm(0x2416, &[1], false),
        sprm(0x2417, &[1], false),
        sprm(0x7621, &[0, 63, 1, 0], false),
        sprm(0x5622, &[1, 63], false),
        sprm(0x2416, &[1], false),
    ]
    .concat()
}
fn depth(value: u32) -> Vec<u8> {
    sprm(0x6649, &value.to_le_bytes(), false)
}
fn nested_cell() -> Vec<u8> {
    [depth(2), sprm(0x244b, &[1], false)].concat()
}
fn nested_row(width: u16) -> Vec<u8> {
    [
        depth(2),
        sprm(0x244c, &[1], false),
        sprm(0x7621, &[0, 1, width as u8, (width >> 8) as u8], false),
    ]
    .concat()
}

/// Replace the fixture's default PAP FKP with exact CP-addressed PAPX runs.
pub(super) fn with_papx(source: &[u8], runs: &[(usize, usize, Vec<u8>)]) -> Vec<u8> {
    let cfb = CompoundFile::open(source).unwrap();
    let mut word = cfb.stream("WordDocument").unwrap();
    let table = cfb.stream("0Table").unwrap();
    let bte = u32::from_le_bytes(word[0x102..0x106].try_into().unwrap()) as usize;
    let pn = u32::from_le_bytes(table[bte + 8..bte + 12].try_into().unwrap()) as usize;
    let page = &mut word[pn * 512..(pn + 1) * 512];
    page.fill(0);
    let n = runs.len();
    assert!(n <= 29);
    for (index, (start, _, _)) in runs.iter().enumerate() {
        page[index * 4..index * 4 + 4]
            .copy_from_slice(&(0x400u32 + (*start as u32) * 2).to_le_bytes());
    }
    page[n * 4..n * 4 + 4]
        .copy_from_slice(&(0x400u32 + (runs.last().unwrap().1 as u32) * 2).to_le_bytes());
    let bx = (n + 1) * 4;
    let mut payload = 510usize;
    for (index, (_, _, sprms)) in runs.iter().enumerate().rev() {
        if sprms.is_empty() {
            continue;
        }
        let mut papx = vec![0, 0];
        papx.extend_from_slice(sprms);
        assert_eq!(papx.len() % 2, 1);
        let cb = papx.len().div_ceil(2);
        payload -= 1 + papx.len();
        payload &= !1;
        page[payload] = cb as u8;
        page[payload + 1..payload + 1 + papx.len()].copy_from_slice(&papx);
        page[bx + index * 13] = (payload / 2) as u8;
    }
    page[511] = n as u8;
    build_scoped_cfb(&[("WordDocument", word), ("0Table", table)])
}

fn body_table_source(text: &str) -> Vec<u8> {
    let units = text.encode_utf16().count();
    let source = source_with_typography(
        text,
        &[(units, 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    with_papx(
        &source,
        &[(0, 2, cell()), (2, 3, row(1000)), (3, units, Vec::new())],
    )
}

fn with_piece_prc(source: &[u8], prc: &[u8], data: &[u8]) -> Vec<u8> {
    with_piece_prc_and_prm(source, prc, data, 1)
}

fn with_piece_prc_and_prm(source: &[u8], prc: &[u8], data: &[u8], prm: u16) -> Vec<u8> {
    let cfb = CompoundFile::open(source).unwrap();
    let mut word = cfb.stream("WordDocument").unwrap();
    let table = cfb.stream("0Table").unwrap();
    let clx_offset = u32::from_le_bytes(word[0x1a2..0x1a6].try_into().unwrap()) as usize;
    let clx_size = u32::from_le_bytes(word[0x1a6..0x1aa].try_into().unwrap()) as usize;
    let mut replacement = table.clone();
    let replacement_offset = replacement.len();
    replacement.push(1);
    replacement.extend(u16::try_from(prc.len()).unwrap().to_le_bytes());
    replacement.extend(prc);
    let prefix = 3 + prc.len();
    replacement.extend_from_slice(&table[clx_offset..clx_offset + clx_size]);
    replacement[replacement_offset + prefix + 19..replacement_offset + prefix + 21]
        .copy_from_slice(&prm.to_le_bytes());
    word[0x1a2..0x1a6].copy_from_slice(&(replacement_offset as u32).to_le_bytes());
    word[0x1a6..0x1aa].copy_from_slice(&((prefix + clx_size) as u32).to_le_bytes());
    build_scoped_cfb(&[
        ("WordDocument", word),
        ("0Table", replacement),
        ("Data", data.to_vec()),
    ])
}

#[test]
fn full_cfb_piece_alignment_applies_after_direct_table_props_data() {
    let text = "a\u{7}\u{7}\r";
    let units = text.encode_utf16().count();
    let base = source_with_typography(
        text,
        &[(units, 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let table_data = [row(1000), sprm(0x2403, &[2], false)].concat();
    let cell_offset = 2 + table_data.len();
    let cell_data = [
        cell(),
        sprm(0x2403, &[2], false),
        sprm(0x2407, &[0], false),
        sprm(0x2407, &[0], false),
    ]
    .concat();
    let mut data = Vec::new();
    data.extend(u16::try_from(table_data.len()).unwrap().to_le_bytes());
    data.extend(table_data);
    data.extend(u16::try_from(cell_data.len()).unwrap().to_le_bytes());
    data.extend(cell_data);
    let pointer = sprm(0x646b, &0u32.to_le_bytes(), false);
    let physical = with_papx(
        &base,
        &[
            (
                0,
                2,
                [
                    sprm(0x646b, &(cell_offset as u32).to_le_bytes(), false),
                    sprm(0x2407, &[0], false),
                ]
                .concat(),
            ),
            (2, 3, [pointer, row(1000)].concat()),
            (3, units, Vec::new()),
        ],
    );
    let simple_center_prm = (1u16 << 8) | (0x05 << 1);
    let complex_center = sprm(0x2403, &[1], false);
    for (prc, prm) in [(&[][..], simple_center_prm), (&complex_center[..], 1)] {
        let bytes = with_piece_prc_and_prm(&physical, prc, &data, prm);
        let document = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
            .unwrap()
            .document;
        let BodyElement::Table(table) = &document.body[0] else {
            panic!("table")
        };
        let CellElement::Paragraph(paragraph) = &table.rows[0].cells[0].content[0] else {
            panic!("cell paragraph")
        };
        assert_eq!(paragraph.alignment, "center");
    }

    let inline = with_papx(
        &base,
        &[
            (
                0,
                2,
                [cell(), sprm(0x2403, &[2], false), sprm(0x2407, &[0], false)].concat(),
            ),
            (
                2,
                3,
                [
                    row(1000),
                    sprm(0x2403, &[2], false),
                    sprm(0x2407, &[0], false),
                ]
                .concat(),
            ),
            (3, units, Vec::new()),
        ],
    );
    let bytes = with_piece_prc(&inline, &complex_center, &[]);
    let document = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
        .unwrap()
        .document;
    let BodyElement::Table(table) = &document.body[0] else {
        panic!("table")
    };
    let CellElement::Paragraph(paragraph) = &table.rows[0].cells[0].content[0] else {
        panic!("cell paragraph")
    };
    assert_eq!(paragraph.alignment, "center");
}

#[test]
fn full_cfb_complex_piece_applies_table_sprms_but_not_its_redirects() {
    let source = body_table_source("a\u{7}\u{7}\r");
    let definition = sprm(0xd608, &[6, 0, 1, 0, 0, 0xe8, 3], false);
    let width = sprm(0x7623, &[0, 1, 0xdc, 5], false);
    let column_widths = |bytes: &[u8]| {
        let document = super::super::direct_model(&CompoundFile::open(bytes).unwrap(), 1_000_000)
            .unwrap()
            .document;
        let Some(BodyElement::Table(table)) = document.body.first() else {
            panic!("table")
        };
        let CellElement::Paragraph(cell) = &table.rows[0].cells[0].content[0] else {
            panic!("cell paragraph")
        };
        assert!(matches!(&cell.runs[0], DocRun::Text(text) if text.text == "a"));
        assert!(matches!(
            document.body.get(1),
            Some(BodyElement::Paragraph(_))
        ));
        table.col_widths.clone()
    };

    // Top-level piece table SPRMs apply to the row mark after its direct PAPX.
    let raw = with_piece_prc(&source, &[definition.clone(), width.clone()].concat(), &[]);
    assert_eq!(column_widths(&raw), [75.0]);

    // Piece PTableProps/PHugePapx are ignored without reading their Data.
    let mut data = Vec::new();
    let redirected = [definition, width].concat();
    data.extend(u16::try_from(redirected.len()).unwrap().to_le_bytes());
    data.extend(redirected);
    for redirect in [0x646b, 0x6646] {
        let prc = sprm(redirect, &0u32.to_le_bytes(), false);
        let wrapped = with_piece_prc(&source, &prc, &data);
        assert_eq!(column_widths(&wrapped), [50.0]);
    }
}

#[test]
fn full_cfb_table_keeps_tdxacol_ranges_across_a_later_same_count_definition() {
    fn definition(boundaries: &[i16]) -> Vec<u8> {
        let count = boundaries.len() - 1;
        let mut operand = vec![0, 0, count as u8];
        for boundary in boundaries {
            operand.extend_from_slice(&boundary.to_le_bytes());
        }
        operand.resize(operand.len() + count * 20, 0);
        let cb = (operand.len() - 1) as u16;
        operand[..2].copy_from_slice(&cb.to_le_bytes());
        sprm(0xd608, &operand, false)
    }

    let text = "a\u{7}b\u{7}c\u{7}\u{7}\r";
    let units = text.encode_utf16().count();
    let source = source_with_typography(
        text,
        &[(units, 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let first = definition(&[0, 1500, 6000, 9000]);
    let second = definition(&[0, 2500, 4000, 9000]);
    let middle = sprm(0x7623, &[1, 2, 0xb8, 0x0b], false);

    for before_second_definition in [true, false] {
        let geometry = if before_second_definition {
            [first.clone(), middle.clone(), second.clone()].concat()
        } else {
            [first.clone(), second.clone(), middle.clone()].concat()
        };
        let row = [
            cell(),
            sprm(0x2417, &[1], false),
            geometry,
            sprm(0x3615, &[0], false),
        ]
        .concat();
        assert_eq!(row.len() % 2, 1, "with_papx requires an odd grpprl");
        let bytes = with_papx(
            &source,
            &[
                (0, 2, cell()),
                (2, 4, cell()),
                (4, 6, cell()),
                (6, 7, row),
                (7, units, Vec::new()),
            ],
        );
        let document = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
            .unwrap()
            .document;
        let BodyElement::Table(table) = &document.body[0] else {
            panic!("table")
        };
        assert_eq!(table.col_widths, [125.0, 150.0, 250.0]);
        assert_eq!(table.rows.len(), 1);
        assert_eq!(table.rows[0].cells.len(), 3);
        let cell_text = table.rows[0]
            .cells
            .iter()
            .map(|cell| {
                cell.content
                    .iter()
                    .filter_map(|element| match element {
                        CellElement::Paragraph(paragraph) => Some(paragraph),
                        CellElement::Table(_) => None,
                    })
                    .flat_map(|paragraph| &paragraph.runs)
                    .filter_map(|run| match run {
                        DocRun::Text(text) => Some(text.text.as_str()),
                        _ => None,
                    })
                    .collect::<String>()
            })
            .collect::<Vec<_>>();
        assert_eq!(cell_text, ["a", "b", "c"]);
        let Some(BodyElement::Paragraph(trailing)) = document.body.get(1) else {
            panic!("trailing paragraph")
        };
        assert!(trailing.runs.is_empty());
    }

    let late_autofit = [
        cell(),
        sprm(0x2417, &[1], false),
        first.clone(),
        middle.clone(),
        second.clone(),
        sprm(0x3615, &[1], false),
    ]
    .concat();
    let bytes = with_papx(
        &source,
        &[
            (0, 2, cell()),
            (2, 4, cell()),
            (4, 6, cell()),
            (6, 7, late_autofit),
            (7, units, Vec::new()),
        ],
    );
    let error =
        super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000).unwrap_err();
    assert!(error.contains("unsupported formatting"), "{error}");

    let alignment_before_definition = [
        cell(),
        sprm(0x2417, &[1], false),
        first,
        sprm(0xd62c, &[1, 2, 1], true),
        middle,
        second,
        sprm(0x3615, &[0], false),
    ]
    .concat();
    let bytes = with_papx(
        &source,
        &[
            (0, 2, cell()),
            (2, 4, cell()),
            (4, 6, cell()),
            (6, 7, alignment_before_definition),
            (7, units, Vec::new()),
        ],
    );
    let document = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
        .unwrap()
        .document;
    let BodyElement::Table(table) = &document.body[0] else {
        panic!("table")
    };
    assert_eq!(
        table.rows[0]
            .cells
            .iter()
            .map(|cell| cell.v_align.as_str())
            .collect::<Vec<_>>(),
        ["top", "center", "top"]
    );
}

#[test]
fn full_cfb_alignment_persists_before_appended_piece_paragraph_properties() {
    fn definition(boundaries: &[i16]) -> Vec<u8> {
        let count = boundaries.len() - 1;
        let mut operand = vec![0, 0, count as u8];
        for boundary in boundaries {
            operand.extend(boundary.to_le_bytes());
        }
        operand.resize(operand.len() + count * 20, 0);
        let cb = (operand.len() - 1) as u16;
        operand[..2].copy_from_slice(&cb.to_le_bytes());
        sprm(0xd608, &operand, false)
    }

    let text = "a\u{7}\u{7}\r";
    let units = text.encode_utf16().count();
    let source = source_with_typography(
        text,
        &[(units, 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let row = [
        cell(),
        sprm(0x2417, &[1], false),
        definition(&[0, 1000]),
        sprm(0xd62c, &[0, 1, 1], true),
        definition(&[0, 1500]),
        sprm(0x3615, &[0], false),
    ]
    .concat();
    let source = with_papx(
        &source,
        &[(0, 2, cell()), (2, 3, row), (3, units, Vec::new())],
    );
    let piece_alignment = sprm(0x2403, &[1], false);
    let bytes = with_piece_prc(&source, &piece_alignment, &[]);
    let document = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
        .unwrap()
        .document;
    let BodyElement::Table(table) = &document.body[0] else {
        panic!("table")
    };
    assert_eq!(table.rows[0].cells[0].v_align, "center");
    let CellElement::Paragraph(paragraph) = &table.rows[0].cells[0].content[0] else {
        panic!("cell paragraph")
    };
    assert_eq!(paragraph.alignment, "center");
    assert!(matches!(
        document.body.get(1),
        Some(BodyElement::Paragraph(_))
    ));
}

#[test]
fn full_cfb_variable_vertical_merge_length_keeps_the_native_gate_closed() {
    let text = "a\u{7}\u{7}\r";
    let units = text.encode_utf16().count();
    let source = source_with_typography(
        text,
        &[(units, 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let build = |vertical_merge: Vec<u8>| {
        let mut row = [row(1000), vertical_merge].concat();
        if row.len() % 2 == 0 {
            row.extend(sprm(0x2416, &[1], false));
        }
        with_papx(
            &source,
            &[(0, 2, cell()), (2, 3, row), (3, units, Vec::new())],
        )
    };

    let valid = build(sprm(0xd62b, &[0, 3], true));
    let document = super::super::direct_model(&CompoundFile::open(&valid).unwrap(), 1_000_000)
        .unwrap()
        .document;
    assert!(matches!(document.body.first(), Some(BodyElement::Table(_))));

    let malformed = build(sprm(0xd62b, &[0], true));
    let error = super::super::direct_model(&CompoundFile::open(&malformed).unwrap(), 1_000_000)
        .unwrap_err();
    assert!(error.contains("unsupported formatting"), "{error}");
}

#[test]
fn full_cfb_body_table_retains_cell_local_page_and_column_break_runs() {
    let bytes = body_table_source("\u{c}\u{7}\u{7}\r");
    let document = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
        .unwrap()
        .document;
    let BodyElement::Table(table) = &document.body[0] else {
        panic!("table")
    };
    let CellElement::Paragraph(paragraph) = &table.rows[0].cells[0].content[0] else {
        panic!("paragraph")
    };
    assert!(matches!(
        paragraph.runs.as_slice(),
        [DocRun::Break {
            break_type: docx_model::BreakType::Page
        }]
    ));

    let bytes = body_table_source("\u{e}\u{7}\u{7}\r");
    let document = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
        .unwrap()
        .document;
    let BodyElement::Table(table) = &document.body[0] else {
        panic!("table")
    };
    let CellElement::Paragraph(paragraph) = &table.rows[0].cells[0].content[0] else {
        panic!("paragraph")
    };
    assert!(matches!(
        paragraph.runs.as_slice(),
        [DocRun::Break {
            break_type: docx_model::BreakType::Column
        }]
    ));
}

#[test]
fn retained_tistd_does_not_bypass_the_existing_style_projection_gate() {
    let text = "x\u{7}\u{7}\r";
    let source = source_with_typography(
        text,
        &[(text.encode_utf16().count(), 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let mut row = row(1000);
    row.extend(sprm(0x563a, &7u16.to_le_bytes(), false));
    let bytes = with_papx(&source, &[(0, 2, cell()), (2, 3, row), (3, 4, Vec::new())]);
    let error =
        super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000).unwrap_err();
    assert!(error.contains("unsupported formatting"), "{error}");
}

#[test]
fn full_cfb_table_budget_is_atomic_and_image_resource_outlives_input() {
    let bytes = with_picture_data(&body_table_source("\u{1}\u{7}\u{7}\r"), false);
    let result = {
        let cfb = CompoundFile::open(&bytes).unwrap();
        super::super::direct_model(&cfb, 1_000_000).unwrap()
    };
    drop(bytes);
    let BodyElement::Table(table) = &result.document.body[0] else {
        panic!("table")
    };
    let CellElement::Paragraph(paragraph) = &table.rows[0].cells[0].content[0] else {
        panic!("paragraph")
    };
    assert!(paragraph
        .runs
        .iter()
        .any(|run| matches!(run, DocRun::Image(_))));
    assert_eq!(result.resources.len(), 1);
    assert!(result.resources[0].bytes.starts_with(b"\x89PNG"));
    let image = paragraph
        .runs
        .iter()
        .find_map(|run| match run {
            DocRun::Image(image) => Some(image),
            _ => None,
        })
        .unwrap();
    assert_eq!(image.image_path, result.resources[0].key);
    assert_eq!(
        super::super::direct_model(
            &CompoundFile::open(&body_table_source("a\u{7}\u{7}\r")).unwrap(),
            1
        )
        .unwrap_err(),
        "OUTPUT_TOO_LARGE"
    );
}

#[test]
fn horizontal_merge_continuation_does_not_retain_an_orphan_picture_resource() {
    let text = "a\u{7}\u{1}\u{7}c\u{7}\u{7}\r";
    let source = source_with_typography(
        text,
        &[(text.encode_utf16().count(), 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let merged_row = [
        sprm(0x2416, &[1], false),
        sprm(0x2417, &[1], false),
        sprm(0x7621, &[0, 3, 0xe8, 3], false),
        sprm(0x5624, &[0, 2], false),
        sprm(0x2416, &[1], false),
    ]
    .concat();
    let bytes = with_picture_data(
        &with_papx(
            &source,
            &[
                (0, 2, cell()),
                (2, 4, cell()),
                (4, 6, cell()),
                (6, 7, merged_row),
                (7, 8, Vec::new()),
            ],
        ),
        false,
    );
    let result =
        super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000).unwrap();
    let BodyElement::Table(table) = &result.document.body[0] else {
        panic!("table")
    };
    assert_eq!(table.rows[0].cells.len(), 2);
    assert!(table.rows[0].cells.iter().all(|cell| cell
        .content
        .iter()
        .filter_map(|block| match block {
            CellElement::Paragraph(paragraph) => Some(paragraph),
            CellElement::Table(_) => None,
        })
        .flat_map(|paragraph| &paragraph.runs)
        .all(|run| !matches!(run, DocRun::Image(_)))));
    assert!(result.resources.is_empty());
}

#[test]
fn vertical_merge_continuation_does_not_retain_an_orphan_picture_resource() {
    let text = "a\u{7}\u{7}\u{1}\u{7}\u{7}\r";
    let source = source_with_typography(
        text,
        &[(text.encode_utf16().count(), 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let vertical_row = |flag| {
        [
            row(1000),
            sprm(0xd62b, &[0, flag], true),
            sprm(0x2416, &[1], false),
        ]
        .concat()
    };
    let bytes = with_picture_data(
        &with_papx(
            &source,
            &[
                (0, 2, cell()),
                (2, 3, vertical_row(3)),
                (3, 5, cell()),
                (5, 6, vertical_row(1)),
                (6, 7, Vec::new()),
            ],
        ),
        false,
    );
    let result =
        super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000).unwrap();
    let BodyElement::Table(table) = &result.document.body[0] else {
        panic!("table")
    };
    assert_eq!(table.rows.len(), 2);
    assert_eq!(table.rows[1].cells[0].v_merge, Some(false));
    assert!(table.rows[1].cells[0]
        .content
        .iter()
        .all(|block| match block {
            CellElement::Paragraph(paragraph) => paragraph
                .runs
                .iter()
                .all(|run| !matches!(run, DocRun::Image(_))),
            CellElement::Table(_) => true,
        }));
    assert!(result.resources.is_empty());
}

#[test]
fn merged_continuations_advance_numbering_before_their_content_is_suppressed() {
    fn numbered_cell() -> Vec<u8> {
        [
            cell(),
            sprm(0x2417, &[0], false),
            sprm(0x260a, &[0], false),
            sprm(0x460b, &1u16.to_le_bytes(), false),
        ]
        .concat()
    }
    fn numbered_body() -> Vec<u8> {
        [
            sprm(0x260a, &[0], false),
            sprm(0x460b, &1u16.to_le_bytes(), false),
        ]
        .concat()
    }
    fn row_with_cells(count: u8, merge: Option<(u16, Vec<u8>, bool)>) -> Vec<u8> {
        let mut properties = [
            sprm(0x2416, &[1], false),
            sprm(0x2417, &[1], false),
            sprm(0x7621, &[0, count, 0xe8, 3], false),
            sprm(0x2416, &[1], false),
        ]
        .concat();
        if let Some((code, operand, variable)) = merge {
            properties.extend(sprm(code, &operand, variable));
            // Keep the PAPX word-sized after adding a variable-length vertical
            // merge record; this repeated in-table value is semantically inert.
            if variable {
                properties.extend(sprm(0x2416, &[1], false));
            }
        }
        properties
    }
    fn marker(cell: &docx_model::DocTableCell) -> Option<&str> {
        cell.content.iter().find_map(|element| match element {
            CellElement::Paragraph(paragraph) => paragraph
                .numbering
                .as_ref()
                .map(|value| value.text.as_str()),
            CellElement::Table(_) => None,
        })
    }
    fn trailing_marker(document: &docx_model::Document) -> &str {
        let BodyElement::Paragraph(paragraph) = &document.body[1] else {
            panic!("trailing numbered paragraph")
        };
        paragraph.numbering.as_ref().unwrap().text.as_str()
    }

    let horizontal_text = "a\u{7}b\u{7}c\u{7}\u{7}d\r";
    for (range, expected) in [([0, 2], ["1.", "3."]), ([1, 3], ["1.", "2."])] {
        let source = with_numbering(&source_with_typography(
            horizontal_text,
            &[(
                horizontal_text.encode_utf16().count(),
                2,
                12240,
                15840,
                1,
                720,
            )],
            None,
            None,
            None,
            None,
        ));
        let row = row_with_cells(3, Some((0x5624, range.to_vec(), false)));
        let bytes = with_papx(
            &source,
            &[
                (0, 2, numbered_cell()),
                (2, 4, numbered_cell()),
                (4, 6, numbered_cell()),
                (6, 7, row),
                (7, 9, numbered_body()),
            ],
        );
        let document = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
            .unwrap()
            .document;
        let BodyElement::Table(table) = &document.body[0] else {
            panic!("horizontal table")
        };
        assert_eq!(table.rows[0].cells.len(), 2);
        assert_eq!(
            table.rows[0]
                .cells
                .iter()
                .filter_map(marker)
                .collect::<Vec<_>>(),
            expected
        );
        assert_eq!(trailing_marker(&document), "4.");
    }

    let vertical_text = "a\u{7}b\u{7}\u{7}c\u{7}d\u{7}\u{7}e\r";
    let source = with_numbering(&source_with_typography(
        vertical_text,
        &[(
            vertical_text.encode_utf16().count(),
            2,
            12240,
            15840,
            1,
            720,
        )],
        None,
        None,
        None,
        None,
    ));
    let bytes = with_papx(
        &source,
        &[
            (0, 2, numbered_cell()),
            (2, 4, numbered_cell()),
            (4, 5, row_with_cells(2, Some((0xd62b, vec![0, 3], true)))),
            (5, 7, numbered_cell()),
            (7, 9, numbered_cell()),
            (9, 10, row_with_cells(2, Some((0xd62b, vec![0, 1], true)))),
            (10, 12, numbered_body()),
        ],
    );
    let document = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
        .unwrap()
        .document;
    let BodyElement::Table(table) = &document.body[0] else {
        panic!("vertical table")
    };
    assert_eq!(table.rows.len(), 2);
    assert_eq!(
        table
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .filter_map(marker)
            .collect::<Vec<_>>(),
        ["1.", "2.", "4."]
    );
    assert_eq!(trailing_marker(&document), "5.");
}

#[test]
fn retained_nested_table_picture_keeps_its_single_deduplicated_resource() {
    let text = "\u{1}\r\r\u{7}\u{7}\r";
    let source = source_with_typography(
        text,
        &[(text.encode_utf16().count(), 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let bytes = with_picture_data(
        &with_papx(
            &source,
            &[
                (0, 2, nested_cell()),
                (2, 3, nested_row(500)),
                (3, 4, cell()),
                (4, 5, row(1000)),
                (5, 6, Vec::new()),
            ],
        ),
        false,
    );
    let result =
        super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000).unwrap();
    let BodyElement::Table(outer) = &result.document.body[0] else {
        panic!("outer table")
    };
    let nested = outer.rows[0].cells[0]
        .content
        .iter()
        .find_map(|block| match block {
            CellElement::Table(table) => Some(table),
            CellElement::Paragraph(_) => None,
        })
        .expect("nested table");
    let image = nested.rows[0].cells[0]
        .content
        .iter()
        .filter_map(|block| match block {
            CellElement::Paragraph(paragraph) => Some(paragraph),
            CellElement::Table(_) => None,
        })
        .flat_map(|paragraph| &paragraph.runs)
        .find_map(|run| match run {
            DocRun::Image(image) => Some(image),
            _ => None,
        })
        .expect("nested image");
    assert_eq!(result.resources.len(), 1);
    assert_eq!(image.image_path, result.resources[0].key);
}

#[test]
fn large_prepared_table_properties_are_admitted_before_paragraph_and_resource_projection() {
    let text = "\u{1}\u{7}\u{7}\r";
    let compact_bytes = with_picture_data(&body_table_source(text), false);
    let compact =
        super::super::direct_model(&CompoundFile::open(&compact_bytes).unwrap(), 32 * 1024)
            .unwrap();
    assert_eq!(compact.resources.len(), 1);

    let source = source_with_typography(
        text,
        &[(text.encode_utf16().count(), 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let bytes = with_picture_data(
        &with_papx(
            &source,
            &[(0, 2, cell()), (2, 3, large_row()), (3, 4, Vec::new())],
        ),
        false,
    );

    assert_eq!(
        super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 32 * 1024).unwrap_err(),
        "OUTPUT_TOO_LARGE"
    );

    let result =
        super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000).unwrap();
    let BodyElement::Table(table) = &result.document.body[0] else {
        panic!("table")
    };
    assert_eq!(table.rows[0].cells.len(), 1);
    assert_eq!(result.resources.len(), 1);
}

#[test]
fn full_cfb_nested_and_header_tables_share_the_story_producer() {
    let text = "n\r\r\u{7}\u{7}\r";
    let source = source_with_typography(
        text,
        &[(text.encode_utf16().count(), 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let bytes = with_papx(
        &source,
        &[
            (0, 2, nested_cell()),
            (2, 3, nested_row(500)),
            (3, 4, cell()),
            (4, 5, row(1000)),
            (5, 6, Vec::new()),
        ],
    );
    let document = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
        .unwrap()
        .document;
    let BodyElement::Table(outer) = &document.body[0] else {
        panic!("outer")
    };
    let nested = outer.rows[0].cells[0]
        .content
        .iter()
        .find_map(|value| match value {
            CellElement::Table(table) => Some(table),
            _ => None,
        })
        .expect("nested table");
    assert_ne!(
        outer.table_layout.logical_sequence_id,
        nested.table_layout.logical_sequence_id
    );
    assert_eq!(outer.table_layout.logical_total_rows, 1);
    assert_eq!(nested.table_layout.logical_total_rows, 1);

    let main = "body\r";
    let header = "h\u{7}\u{7}\r";
    let slots = [None, Some(header), None, None, None, None];
    let source = source_with_typography(
        main,
        &[(main.encode_utf16().count(), 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        Some(&slots),
    );
    let base = main.encode_utf16().count() + 1;
    let total = main.encode_utf16().count() + 1 + header.encode_utf16().count() + 2;
    let bytes = with_papx(
        &source,
        &[
            (0, base, Vec::new()),
            (base, base + 2, cell()),
            (base + 2, base + 3, row(1000)),
            (base + 3, total, Vec::new()),
        ],
    );
    let document = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
        .unwrap()
        .document;
    let header = document.headers.default.unwrap();
    assert!(header
        .body
        .iter()
        .any(|value| matches!(value, BodyElement::Table(_))));
}

#[test]
fn table_budget_fails_at_row_retention_and_finish_grid_projection() {
    use super::tables::{Block, Blocks, Writer};
    use crate::doc::table::Properties;
    let paragraph = || {
        let mut value = docx_model::DocParagraph::default();
        value.runs.push(DocRun::Text(Box::new(docx_model::TextRun {
            text: "cell".into(),
            ..Default::default()
        })));
        Blocks(vec![Block::Paragraph(Box::new(value))])
    };
    let cell_props = || {
        let mut value = Properties::default();
        value.apply(0x6649, &1u32.to_le_bytes()).unwrap();
        value
    };
    let row_props = || {
        let mut value = cell_props();
        value.row_end = true;
        value.inner_row = true;
        value.row.apply(0x7621, &[0, 1, 0xe8, 3]).unwrap();
        value
    };

    let mut sequence = 0;
    let mut writer = Writer::new(&mut sequence);
    let mut budget = super::ModelBudget::new(64 * 1024);
    writer
        .push(cell_props(), '\u{7}', paragraph(), &mut budget)
        .unwrap();
    budget.remaining_bytes = 0;
    assert_eq!(
        writer
            .push(row_props(), '\u{7}', Blocks::default(), &mut budget)
            .unwrap_err(),
        "OUTPUT_TOO_LARGE"
    );

    let mut sequence = 0;
    let mut writer = Writer::new(&mut sequence);
    let mut budget = super::ModelBudget::new(64 * 1024);
    writer
        .push(cell_props(), '\u{7}', paragraph(), &mut budget)
        .unwrap();
    writer
        .push(row_props(), '\u{7}', Blocks::default(), &mut budget)
        .unwrap();
    budget.remaining_bytes = 0;
    assert_eq!(
        writer.finish(&mut budget).err().unwrap(),
        "OUTPUT_TOO_LARGE"
    );
}

#[test]
fn clear_and_solid_cell_backgrounds_use_their_exact_colors() {
    use super::tables::{Block, Blocks, Writer};
    use crate::doc::table::Properties;
    let paragraph = || Blocks(vec![Block::Paragraph(Box::default())]);
    let build = |pattern: u16, foreground_auto: bool| {
        let mut cell = Properties::default();
        cell.apply(0x6649, &1u32.to_le_bytes()).unwrap();
        let mut end = Properties::default();
        end.apply(0x6649, &1u32.to_le_bytes()).unwrap();
        end.row_end = true;
        end.inner_row = true;
        end.row.apply(0x7621, &[0, 1, 0xe8, 3]).unwrap();
        let [lo, hi] = pattern.to_le_bytes();
        let foreground = if foreground_auto {
            [0, 0, 0, 255]
        } else {
            [0x11, 0x22, 0x33, 0]
        };
        let operand = [
            10,
            foreground[0],
            foreground[1],
            foreground[2],
            foreground[3],
            0x44,
            0x55,
            0x66,
            0,
            lo,
            hi,
        ];
        end.row.apply(0xd612, &operand).unwrap();
        let mut sequence = 0;
        let mut writer = Writer::new(&mut sequence);
        let mut budget = super::ModelBudget::new(64 * 1024);
        writer
            .push(cell, '\u{7}', paragraph(), &mut budget)
            .unwrap();
        writer
            .push(end, '\u{7}', Blocks::default(), &mut budget)
            .unwrap();
        writer.finish(&mut budget)
    };
    let clear = build(0, false).unwrap();
    let Block::Table(table) = &clear.0[0] else {
        panic!()
    };
    assert_eq!(table.rows[0].cells[0].background.as_deref(), Some("445566"));
    let solid = build(1, false).unwrap();
    let Block::Table(table) = &solid.0[0] else {
        panic!()
    };
    assert_eq!(table.rows[0].cells[0].background.as_deref(), Some("112233"));
    assert!(build(1, true)
        .err()
        .unwrap()
        .contains("automatic solid cell shading"));
    assert!(build(2, false)
        .err()
        .unwrap()
        .contains("patterned cell shading"));
}

#[test]
fn multiappend_reserves_for_len_plus_add_before_mutating() {
    use super::tables::{Block, Blocks};
    let paragraph = || Block::Paragraph(Box::default());
    let mut target = Blocks(Vec::with_capacity(4));
    target.0.push(paragraph());
    target.0.push(paragraph());
    let other = Blocks(vec![paragraph(), paragraph(), paragraph()]);
    let before = target.0.capacity();
    let mut charged = 0usize;
    target
        .append(other, &mut |bytes| {
            charged += bytes;
            Ok(())
        })
        .unwrap();
    assert_eq!(target.0.len(), 5);
    assert!(target.0.capacity() >= 5);
    assert!(charged >= (5 - before) * std::mem::size_of::<Block>());

    let mut target = Blocks(Vec::with_capacity(4));
    target.0.push(paragraph());
    target.0.push(paragraph());
    let original_capacity = target.0.capacity();
    assert_eq!(
        target
            .append(
                Blocks(vec![paragraph(), paragraph(), paragraph()]),
                &mut |_| Err("OUTPUT_TOO_LARGE".into())
            )
            .unwrap_err(),
        "OUTPUT_TOO_LARGE"
    );
    assert_eq!(target.0.len(), 2);
    assert_eq!(target.0.capacity(), original_capacity);
}
