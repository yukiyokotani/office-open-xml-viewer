//! Full-container acceptance fixtures for direct binary table projection.

use super::tests::{source_with_typography, with_numbering, with_picture_data};
use crate::cfb::{test_support::build_scoped_cfb, CompoundFile};
use crate::doc::table_structure::{FrameKey, Payload};
use docx_model::{BodyElement, CellElement, DocRun};

fn sprm(code: u16, operand: &[u8], variable: bool) -> Vec<u8> {
    let mut value = code.to_le_bytes().to_vec();
    if variable {
        value.push(operand.len() as u8);
    }
    value.extend_from_slice(operand);
    value
}

fn unframed() -> Result<FrameKey, String> {
    Ok(FrameKey::default())
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

pub(super) fn with_piece_prc(source: &[u8], prc: &[u8], data: &[u8]) -> Vec<u8> {
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

fn cell_border_assignment(sides: u8, border: [u8; 8]) -> Vec<u8> {
    let mut operand = vec![11, 0, 1, sides];
    operand.extend(border);
    sprm(0xd62f, &operand, false)
}

fn cell_border_tc80_row(border: [u8; 4]) -> Vec<u8> {
    let mut definition = vec![26, 0, 1, 0, 0, 0xe8, 3, 0, 0, 0, 0];
    for _ in 0..4 {
        definition.extend(border);
    }
    [
        cell(),
        sprm(0x2417, &[1], false),
        sprm(0xd608, &definition, false),
    ]
    .concat()
}

/// Fill the existing fixture's empty style slot with a red table border.
fn with_cell_border_style_border(source: &[u8], border: [u8; 8]) -> Vec<u8> {
    let cfb = CompoundFile::open(source).unwrap();
    let mut word = cfb.stream("WordDocument").unwrap();
    let mut table = cfb.stream("0Table").unwrap();
    let offset = u32::from_le_bytes(word[0xa2..0xa6].try_into().unwrap()) as usize;
    let length = u32::from_le_bytes(word[0xa6..0xaa].try_into().unwrap()) as usize;
    let mut stylesheet = table[offset..offset + length].to_vec();
    let normal = 2 + u16::from_le_bytes(stylesheet[..2].try_into().unwrap()) as usize;
    let slot = normal
        + 2
        + u16::from_le_bytes(stylesheet[normal..normal + 2].try_into().unwrap()) as usize;
    assert_eq!(&stylesheet[slot..slot + 2], &[0, 0]);
    let mut borders = vec![48];
    for _ in 0..6 {
        borders.extend(border);
    }
    let tapx = sprm(0xd613, &borders, false);
    let mut style = vec![0; 14];
    style[2..4].copy_from_slice(&0xfff3u16.to_le_bytes());
    style[4..6].copy_from_slice(&3u16.to_le_bytes());
    for set in [tapx.as_slice(), &[][..], &[][..]] {
        style.extend((set.len() as u16).to_le_bytes());
        style.extend(set);
        if set.len() % 2 != 0 {
            style.push(0);
        }
    }
    let size = style.len() as u16;
    style[6..8].copy_from_slice(&size.to_le_bytes());
    stylesheet.splice(
        slot..slot + 2,
        [size.to_le_bytes().as_slice(), &style].concat(),
    );
    word[0xa2..0xa6].copy_from_slice(&(table.len() as u32).to_le_bytes());
    word[0xa6..0xaa].copy_from_slice(&(stylesheet.len() as u32).to_le_bytes());
    table.extend(stylesheet);
    build_scoped_cfb(&[("WordDocument", word), ("0Table", table)])
}

fn cell_border_source(mut row_properties: Vec<u8>, styled: bool) -> Vec<u8> {
    let source = body_table_source("x\u{7}\u{7}\r");
    if row_properties.len().is_multiple_of(2) {
        row_properties.extend(cell());
    }
    let source = with_papx(
        &source,
        &[(0, 2, cell()), (2, 3, row_properties), (3, 4, Vec::new())],
    );
    if styled {
        with_cell_border_style_border(&source, [0xff, 0, 0, 0, 8, 1, 0, 0])
    } else {
        source
    }
}

fn cell_border_document(
    row_properties: Vec<u8>,
    styled: bool,
) -> Result<docx_model::Document, String> {
    let source = cell_border_source(row_properties, styled);
    super::super::direct_model(&CompoundFile::open(&source).unwrap(), 1_000_000)
        .map(|result| result.document)
}

#[test]
fn cell_border_superseded_complete_assignment_matches_drawn_and_nil_controls() {
    // MS-DOC 2.4.6 and 2.9.305: a complete later direct assignment owns
    // every field on the selected edges. This does not decode a winning FF.
    for (replacement, expected_style, unresolved) in [
        (
            [0, 0, 0xff, 0, 24, 1, 0, 0],
            "single",
            [0, 0, 0, 0, 8, 0xff, 0, 0],
        ),
        ([0xff; 8], "nil", [0, 0, 0, 0, 8, 0xff, 0, 0]),
        (
            [0, 0, 0xff, 0, 24, 1, 0, 0],
            "single",
            [0, 0, 0, 0, 31, 0xff, 0, 0],
        ),
        // cvAuto ignores its RGB bytes; FF width zero remains in its valid domain.
        (
            [0, 0, 0xff, 0, 24, 1, 0, 0],
            "single",
            [1, 2, 3, 0xff, 0, 0xff, 0, 0],
        ),
        // Exact Nil must precede ordinary COLORREF and width validation.
        (
            [0, 0, 0, 1, 0xff, 0xff, 0xff, 0xff],
            "nil",
            [0, 0, 0, 0, 8, 0xff, 0, 0],
        ),
    ] {
        // MS-DOC 2.9.16 requires width < 32 for types >= 0x40.
        let unresolved = cell_border_assignment(0x0f, unresolved);
        let later = cell_border_assignment(0x0f, replacement);
        let control = cell_border_document([row(1000), later.clone()].concat(), false).unwrap();
        let BodyElement::Table(table) = &control.body[0] else {
            panic!("table")
        };
        let top = table.rows[0].cells[0].borders.top.as_ref().unwrap();
        assert_eq!(top.style, expected_style);
        if expected_style == "single" {
            assert_eq!(top.color.as_deref(), Some("0000ff"));
        }
        let actual =
            cell_border_document([row(1000), unresolved.clone(), later].concat(), false).unwrap();
        assert_eq!(
            serde_json::to_value(actual).unwrap(),
            serde_json::to_value(control).unwrap()
        );
    }
}

#[test]
fn cell_border_superseded_leaves_winning_and_partially_replaced_owners_gated() {
    let unresolved = cell_border_assignment(0x03, [0, 0, 0, 0, 8, 0xff, 0, 0]);
    let complete = cell_border_assignment(0x03, [0xff, 0, 0, 0, 8, 1, 0, 0]);
    let top_only = cell_border_assignment(0x01, [0xff, 0, 0, 0, 8, 1, 0, 0]);
    for properties in [
        [row(1000), unresolved.clone()].concat(),
        [row(1000), complete, unresolved.clone()].concat(),
        [row(1000), unresolved.clone(), unresolved.clone()].concat(),
        [row(1000), unresolved, top_only].concat(),
    ] {
        assert!(cell_border_document(properties, false).is_err());
    }

    // Replacing one cell must not discharge the neighbouring cell's owner.
    let source = body_table_source("x\u{7}y\u{7}\u{7}\r");
    let project = |properties: Vec<u8>| {
        let mut properties = properties;
        if properties.len().is_multiple_of(2) {
            properties.extend(cell());
        }
        let bytes = with_papx(
            &source,
            &[
                (0, 2, cell()),
                (2, 4, cell()),
                (4, 5, properties),
                (5, 6, Vec::new()),
            ],
        );
        super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
            .map(|result| result.document)
    };
    let two_cells = [
        cell(),
        sprm(0x2417, &[1], false),
        sprm(0x7621, &[0, 2, 0xe8, 3], false),
    ]
    .concat();
    let assignment = |limit, border: [u8; 8]| {
        let mut operand = vec![11, 0, limit, 1];
        operand.extend(border);
        sprm(0xd62f, &operand, false)
    };
    let both_ff = assignment(2, [0, 0, 0, 0, 8, 0xff, 0, 0]);
    let one_red = assignment(1, [0xff, 0, 0, 0, 8, 1, 0, 0]);
    assert!(project([two_cells.clone(), both_ff.clone(), one_red].concat()).is_err());
    let both_red = assignment(2, [0xff, 0, 0, 0, 8, 1, 0, 0]);
    let control = project([two_cells.clone(), both_red.clone()].concat()).unwrap();
    let BodyElement::Table(table) = &control.body[0] else {
        panic!("table")
    };
    assert_eq!(table.rows[0].cells.len(), 2);
    assert!(table.rows[0].cells.iter().all(|cell| cell
        .borders
        .top
        .as_ref()
        .unwrap()
        .color
        .as_deref()
        == Some("ff0000")));
    let actual = project([two_cells, both_ff, both_red].concat()).unwrap();
    assert_eq!(
        serde_json::to_value(actual).unwrap(),
        serde_json::to_value(control).unwrap()
    );
}

#[test]
fn cell_border_superseded_tistd_resets_only_the_prepared_direct_layer() {
    let unresolved = cell_border_assignment(0x0f, [0, 0, 0, 0, 8, 0xff, 0, 0]);
    let reset = sprm(0x563a, &1u16.to_le_bytes(), false);
    let control = cell_border_document([row(1000), reset.clone()].concat(), true).unwrap();
    let BodyElement::Table(table) = &control.body[0] else {
        panic!("table")
    };
    assert_eq!(
        table.rows[0].cells[0]
            .borders
            .top
            .as_ref()
            .unwrap()
            .color
            .as_deref(),
        Some("ff0000")
    );
    let actual = cell_border_document(
        [row(1000), unresolved.clone(), reset.clone()].concat(),
        true,
    )
    .unwrap();
    assert_eq!(
        serde_json::to_value(actual).unwrap(),
        serde_json::to_value(control).unwrap()
    );

    // TC80 is outside that resettable layer; the existing styled-TC80 gate
    // and unrelated malformed property refusals cannot be cleared by TIstd.
    assert!(cell_border_document(
        [
            cell_border_tc80_row([8, 1, 6, 0]),
            unresolved.clone(),
            reset.clone(),
        ]
        .concat(),
        true
    )
    .is_err());
    let mut undefined_row_border = vec![48];
    for _ in 0..6 {
        undefined_row_border.extend([0, 0, 0, 0, 8, 2, 0, 0]);
    }
    assert!(cell_border_document(
        [
            row(1000),
            sprm(0xd613, &undefined_row_border, false),
            unresolved,
            reset,
        ]
        .concat(),
        true
    )
    .is_err());
}

#[test]
fn cell_border_superseded_does_not_excuse_invalid_earlier_operands() {
    let valid = vec![11, 0, 1, 1, 0, 0, 0, 0, 8, 0xff, 0, 0];
    let mut range = valid.clone();
    range[2] = 2;
    let mut reversed = valid.clone();
    reversed[1] = 1;
    reversed[2] = 0;
    let mut sides = valid.clone();
    sides[3] = 0x40;
    let mut cb = valid.clone();
    cb[0] = 10;
    let mut color = valid.clone();
    color[7] = 1;
    let mut kind = valid.clone();
    kind[9] = 2;
    let mut width = valid.clone();
    width[8] = 32;
    let later = cell_border_assignment(1, [0xff, 0, 0, 0, 8, 1, 0, 0]);
    cell_border_document([row(1000), later.clone()].concat(), false).unwrap();
    for (operand, reason) in [
        (range, "cell range outside row"),
        (reversed, "cell range outside row"),
        (sides, "unsupported formatting"),
        (cb.clone(), "truncated Word integer"),
        (color, "invalid Word COLORREF"),
        (kind, "undefined Word border type 0x02"),
        (width, "invalid Word ignored border width"),
    ] {
        let error = cell_border_document(
            [row(1000), sprm(0xd62f, &operand, false), later.clone()].concat(),
            false,
        )
        .expect_err("invalid earlier assignment must refuse");
        assert!(error.contains(reason), "{reason}: {error}");
    }
    // A truncated operand needs its own acquisition boundary: concatenating
    // another Prl could supply its missing bytes from that Prl's opcode.
    let mut direct = crate::doc::table::Row::default();
    direct.apply(0x7621, &[0, 1, 0xe8, 3]).unwrap();
    for operand in [&valid[..10], cb.as_slice()] {
        let error = direct
            .apply_style_aware_borders(0xd62f, operand)
            .unwrap_err();
        assert!(
            error.contains("invalid Word cell border operand length"),
            "{error}"
        );
    }
    // Exact Nil is independently valid, not an ordinary FF with invalid
    // COLORREF. The full-container Nil replacement positive above exercises it.
}

#[test]
fn cell_border_superseded_keeps_row_old_tc80_and_style_carriers_gated() {
    let ff = [0, 0, 0, 0, 8, 0xff, 0, 0];
    let mut array = vec![48];
    for _ in 0..6 {
        array.extend(ff);
    }
    let later = cell_border_assignment(0x0f, [0xff, 0, 0, 0, 8, 1, 0, 0]);
    let mut valid_array = vec![48];
    for _ in 0..6 {
        valid_array.extend([0xff, 0, 0, 0, 8, 1, 0, 0]);
    }
    for (control, earlier) in [
        (
            [row(1000), sprm(0xd613, &valid_array, false)].concat(),
            [row(1000), sprm(0xd613, &array, false)].concat(),
        ),
        (
            cell_border_tc80_row([8, 1, 6, 0]),
            cell_border_tc80_row([8, 0xff, 6, 0]),
        ),
        (
            [row(1000), sprm(0xd620, &[7, 0, 1, 1, 8, 1, 6, 0], false)].concat(),
            [row(1000), sprm(0xd620, &[7, 0, 1, 1, 8, 0xff, 6, 0], false)].concat(),
        ),
    ] {
        cell_border_document([control, later.clone()].concat(), false).unwrap();
        let error = cell_border_document([earlier, later.clone()].concat(), false)
            .expect_err("unsupported carrier must refuse");
        assert!(
            error.contains("unsupported Word border type 0xFF"),
            "{error}"
        );
    }
    // A later row border belongs to a different layer from direct cell FF.
    let mut red_array = vec![48];
    for _ in 0..6 {
        red_array.extend([0xff, 0, 0, 0, 8, 1, 0, 0]);
    }
    assert!(cell_border_document(
        [
            row(1000),
            cell_border_assignment(1, ff),
            sprm(0xd613, &red_array, false),
        ]
        .concat(),
        false
    )
    .is_err());
    let source = cell_border_source(
        [row(1000), sprm(0x563a, &1u16.to_le_bytes(), false), later].concat(),
        false,
    );
    let control = with_cell_border_style_border(&source, [0xff, 0, 0, 0, 8, 1, 0, 0]);
    super::super::direct_model(&CompoundFile::open(&control).unwrap(), 1_000_000).unwrap();
    let source = with_cell_border_style_border(&source, ff);
    let error = super::super::direct_model(&CompoundFile::open(&source).unwrap(), 1_000_000)
        .expect_err("table-style FF must refuse");
    assert!(
        error.contains("unsupported Word border type 0xFF"),
        "{error}"
    );
}

#[test]
fn cell_border_superseded_textbox_overwrite_matches_control_but_winner_refuses() {
    let source = super::tests::drawing_shape_source(Some("x\u{7}\u{7}\r"), false);
    let unresolved = cell_border_assignment(0x0f, [0, 0, 0, 0, 8, 0xff, 0, 0]);
    let red = cell_border_assignment(0x0f, [0xff, 0, 0, 0, 8, 1, 0, 0]);
    let project = |mut properties: Vec<u8>| {
        if properties.len().is_multiple_of(2) {
            properties.extend(cell());
        }
        // The real fixture's main story precedes its anchored textbox story.
        let bytes = with_papx(
            &source,
            &[
                (0, 3, Vec::new()),
                (3, 5, cell()),
                (5, 6, properties),
                (6, 7, Vec::new()),
            ],
        );
        super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
            .map(|result| result.document)
    };
    let control = project([row(1000), red.clone()].concat()).unwrap();
    let BodyElement::Paragraph(paragraph) = &control.body[0] else {
        panic!("main paragraph")
    };
    let shape = paragraph
        .runs
        .iter()
        .find_map(|run| match run {
            DocRun::Shape(shape) => Some(shape),
            _ => None,
        })
        .expect("anchored textbox");
    let table = shape
        .text_box_content
        .iter()
        .find_map(|block| match block {
            docx_model::TextBoxBlockWire::Body(BodyElement::Table(table)) => Some(table),
            _ => None,
        })
        .expect("textbox table");
    assert_eq!(
        table.rows[0].cells[0]
            .borders
            .top
            .as_ref()
            .unwrap()
            .color
            .as_deref(),
        Some("ff0000")
    );
    let actual = project([row(1000), unresolved.clone(), red.clone()].concat()).unwrap();
    assert_eq!(
        serde_json::to_value(actual).unwrap(),
        serde_json::to_value(control).unwrap()
    );
    assert!(project([row(1000), unresolved.clone()].concat()).is_err());
    assert!(project([row(1000), red, unresolved].concat()).is_err());
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
        .push(
            cell_props(),
            '\u{7}',
            paragraph(),
            &mut unframed,
            &mut budget,
        )
        .unwrap();
    budget.remaining_bytes = 0;
    assert_eq!(
        writer
            .push(
                row_props(),
                '\u{7}',
                Blocks::default(),
                &mut unframed,
                &mut budget
            )
            .unwrap_err(),
        "OUTPUT_TOO_LARGE"
    );

    let mut sequence = 0;
    let mut writer = Writer::new(&mut sequence);
    let mut budget = super::ModelBudget::new(64 * 1024);
    writer
        .push(
            cell_props(),
            '\u{7}',
            paragraph(),
            &mut unframed,
            &mut budget,
        )
        .unwrap();
    writer
        .push(
            row_props(),
            '\u{7}',
            Blocks::default(),
            &mut unframed,
            &mut budget,
        )
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
            .push(cell, '\u{7}', paragraph(), &mut unframed, &mut budget)
            .unwrap();
        writer
            .push(end, '\u{7}', Blocks::default(), &mut unframed, &mut budget)
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

#[test]
fn positioned_header_table_stays_outside_main_story_admission() {
    let slots = [None, Some("a\u{7}\u{7}\r"), None, None, None, None];
    let bytes = source_with_typography(
        "b\r",
        &[(2, 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        Some(&slots),
    );
    // Physical CPs: main 0..2, separator 2..3, header table 3..6,
    // final header paragraph and guards 6..9.
    let positioned_row = [
        row(1000),
        sprm(0x360d, &[0x60], false),
        sprm(0x940f, &159i16.to_le_bytes(), false),
        cell(),
    ]
    .concat();
    let bytes = with_papx(
        &bytes,
        &[
            (0, 3, Vec::new()),
            (3, 5, cell()),
            (5, 6, positioned_row),
            (6, 9, Vec::new()),
        ],
    );
    let result = crate::doc::direct_model(&CompoundFile::open(&bytes).unwrap(), 1024 * 1024);
    match result {
        Ok(value) => panic!(
            "positioned header table was admitted: {:?}",
            value.document.headers.default
        ),
        Err(error) => assert!(error.contains("outside the main story"), "{error}"),
    }
}

/// Main-story projection through the production story producer (context
/// index plus emitted tables), observed before the final model admission
/// gates. Framed cell paragraphs remain refused; the flag is returned so each
/// case also proves that gate is unchanged.
fn main_story_before_admission(bytes: &[u8]) -> (Vec<BodyElement>, bool) {
    let cfb = CompoundFile::open(bytes).unwrap();
    crate::doc::with_acquired_doc(&cfb, |mut facts| {
        let body = super::project_story_for_test(&facts.story, &mut facts.formatting)?;
        Ok((body, facts.formatting.unsupported_paragraph_properties))
    })
    .unwrap()
}

fn table_rows(body: &[BodyElement]) -> Vec<usize> {
    body.iter()
        .filter_map(|element| match element {
            BodyElement::Table(table) => Some(table.rows.len()),
            _ => None,
        })
        .collect()
}

/// sprmPPc text/column anchors with an absolute Y and the given X. The
/// explicit unlocked anchor keeps the fixture's PAPX grpprl length odd.
fn frame(x: i16) -> Vec<u8> {
    padded_frame(0x20, x)
}

fn padded_frame(position_code: u8, x: i16) -> Vec<u8> {
    [
        sprm(0x261b, &[position_code], false),
        sprm(0x8419, &241i16.to_le_bytes(), false),
        sprm(0x8418, &x.to_le_bytes(), false),
        sprm(0x2430, &[0], false),
    ]
    .concat()
}

fn two_row_source(first: Vec<u8>, second: Vec<u8>) -> Vec<u8> {
    // CPs: row 1 cell 0..2, TTP 2..3; row 2 cell 3..5, TTP 5..6.
    let text = "a\u{7}\u{7}b\u{7}\u{7}\r";
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
        &[
            (0, 2, [cell(), first].concat()),
            (2, 3, row(1000)),
            (3, 5, [cell(), second].concat()),
            (5, 6, row(1000)),
            (6, units, Vec::new()),
        ],
    )
}

#[test]
fn first_cell_frame_properties_split_adjacent_unpositioned_rows() {
    // [MS-DOC] 2.4.3: rows without nondefault table position/wrapping are
    // different tables when their first cells' first paragraphs differ in
    // a frame property.
    let (body, framed_gate) = main_story_before_admission(&two_row_source(frame(-4), frame(1441)));
    assert_eq!(table_rows(&body), vec![1, 1]);
    assert!(framed_gate, "framed cell paragraphs stay refused");
    let [BodyElement::Table(first), BodyElement::Table(second)] = &body[..2] else {
        panic!("two tables")
    };
    assert_ne!(
        first.table_layout.logical_sequence_id,
        second.table_layout.logical_sequence_id
    );
    for table in [first, second] {
        assert_eq!(
            (
                table.table_layout.logical_row_offset,
                table.table_layout.logical_total_rows
            ),
            (0, 1)
        );
    }

    // Equal effective values keep one table, including explicit
    // default-valued distances against their omission.
    let equal = [
        frame(-4),
        sprm(0x842f, &0u16.to_le_bytes(), false),
        sprm(0x842e, &0u16.to_le_bytes(), false),
    ]
    .concat();
    let (body, _) = main_story_before_admission(&two_row_source(frame(-4), equal));
    assert_eq!(table_rows(&body), vec![2]);
    // PositionCodeOperand padding bits MUST be ignored ([MS-DOC] 2.9.208).
    let (body, _) = main_story_before_admission(&two_row_source(frame(-4), padded_frame(0x2f, -4)));
    assert_eq!(table_rows(&body), vec![2]);
    // An unframed row against a framed one is a frame-property difference.
    let (body, _) = main_story_before_admission(&two_row_source(Vec::new(), frame(-4)));
    assert_eq!(table_rows(&body), vec![1, 1]);
}

#[test]
fn only_the_first_cell_first_paragraph_frame_is_row_identity() {
    // Row 1: an empty first paragraph (its cell mark only) carries frame A;
    // its later paragraph and second cell carry other frames. Row 2's first
    // paragraph repeats frame A, so both rows remain one table.
    // CPs: row 1 cell 1 0..1, cell 2 1..3, TTP 3..4; row 2 cell 4..6, cell
    // 6..8, TTP 8..9.
    let text = "\u{7}b\u{7}\u{7}c\u{7}d\u{7}\u{7}\r";
    let units = text.encode_utf16().count();
    let source = source_with_typography(
        text,
        &[(units, 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let two_cells = || {
        [
            sprm(0x2416, &[1], false),
            sprm(0x2417, &[1], false),
            sprm(0x7621, &[0, 2, 0xe8, 0x03], false),
            sprm(0x2416, &[1], false),
        ]
        .concat()
    };
    let document = |row_two_first: Vec<u8>| {
        with_papx(
            &source,
            &[
                (0, 1, [cell(), frame(-4)].concat()),
                (1, 3, [cell(), frame(-8)].concat()),
                (3, 4, two_cells()),
                (4, 6, [cell(), row_two_first].concat()),
                (6, 8, [cell(), frame(-12)].concat()),
                (8, 9, two_cells()),
                (9, units, Vec::new()),
            ],
        )
    };
    let (body, framed_gate) = main_story_before_admission(&document(frame(-4)));
    assert_eq!(table_rows(&body), vec![2]);
    assert!(framed_gate);
    // The empty first paragraph is still the identity source.
    let (body, _) = main_story_before_admission(&document(frame(-8)));
    assert_eq!(table_rows(&body), vec![1, 1]);
}

#[test]
fn positioned_rows_keep_their_own_identity_over_cell_frames() {
    // sprmTFNoAllowOverlap is a nondefault table wrapping property, so the
    // first-paragraph frames are not consulted ([MS-DOC] 2.4.3).
    // The repeated (idempotent) operand keeps the PAPX grpprl length odd.
    let no_overlap = || {
        [
            row(1000),
            sprm(0x3465, &[1], false),
            sprm(0x3465, &[1], false),
        ]
        .concat()
    };
    let text = "a\u{7}\u{7}b\u{7}\u{7}\r";
    let units = text.encode_utf16().count();
    let source = source_with_typography(
        text,
        &[(units, 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let bytes = with_papx(
        &source,
        &[
            (0, 2, [cell(), frame(-4)].concat()),
            (2, 3, no_overlap()),
            (3, 5, [cell(), frame(1441)].concat()),
            (5, 6, no_overlap()),
            (6, units, Vec::new()),
        ],
    );
    let (body, _) = main_story_before_admission(&bytes);
    assert_eq!(table_rows(&body), vec![2]);
}

#[test]
fn nested_rows_use_their_own_depth_first_paragraph_frames() {
    // Outer row 1's first cell starts with a two-row nested table whose rows
    // carry different frames; the outer cell's first depth-1 paragraph (after
    // the nested table) and outer row 2's first paragraph share frame A.
    // CPs: inner row 1 cell 0..2, TTP 2..3; inner row 2 cell 3..5, TTP 5..6;
    // outer cell 6..8, outer TTP 8..9; outer row 2 cell 9..11, TTP 11..12.
    let text = "n\r\rm\r\rx\u{7}\u{7}y\u{7}\u{7}\r";
    let units = text.encode_utf16().count();
    let source = source_with_typography(
        text,
        &[(units, 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let bytes = with_papx(
        &source,
        &[
            (0, 2, [nested_cell(), frame(-8)].concat()),
            (2, 3, nested_row(500)),
            (3, 5, [nested_cell(), frame(-12)].concat()),
            (5, 6, nested_row(500)),
            (6, 8, [cell(), frame(-4)].concat()),
            (8, 9, row(1000)),
            (9, 11, [cell(), frame(-4)].concat()),
            (11, 12, row(1000)),
            (12, units, Vec::new()),
        ],
    );
    let (body, _) = main_story_before_admission(&bytes);
    assert_eq!(table_rows(&body), vec![2]);
    let BodyElement::Table(outer) = &body[0] else {
        panic!("outer table")
    };
    let nested: Vec<_> = outer.rows[0].cells[0]
        .content
        .iter()
        .filter_map(|value| match value {
            CellElement::Table(table) => Some(table.rows.len()),
            _ => None,
        })
        .collect();
    assert_eq!(nested, vec![1, 1]);
    assert!(outer.rows[1].cells[0]
        .content
        .iter()
        .all(|value| !matches!(value, CellElement::Table(_))));
}

#[test]
fn whole_model_keeps_refusing_framed_unpositioned_cell_paragraphs() {
    for bytes in [
        two_row_source(frame(-4), frame(1441)),
        two_row_source(frame(-4), frame(-4)),
    ] {
        let error = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
            .expect_err("framed cell paragraphs are refused");
        assert!(error.contains("unsupported formatting"), "{error}");
    }
}

/// A depth-2 cell paragraph PAPX with the given frame SPRMs. An explicit
/// unlocked anchor keeps the fixture's PAPX grpprl length odd.
fn nested_frame_cell(frame: Vec<u8>) -> Vec<u8> {
    let mut properties = [nested_cell(), frame].concat();
    if properties.len() % 2 == 0 {
        properties.extend(sprm(0x2430, &[0], false));
    }
    properties
}

fn nested_frame_source(first: Vec<u8>, second: Vec<u8>) -> Vec<u8> {
    nested_frame_source_with_rows(first, second, nested_row(500), nested_row(500))
}

fn nested_frame_source_with_rows(
    first: Vec<u8>,
    second: Vec<u8>,
    first_row: Vec<u8>,
    second_row: Vec<u8>,
) -> Vec<u8> {
    let text = "n\r\rm\r\rx\u{7}\u{7}\r";
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
        &[
            (0, 2, nested_frame_cell(first)),
            (2, 3, first_row),
            (3, 5, nested_frame_cell(second)),
            (5, 6, second_row),
            (6, 8, cell()),
            (8, 9, row(1000)),
            (9, units, Vec::new()),
        ],
    )
}

fn nested_owner_frame(x: i16) -> Vec<u8> {
    [padded_frame(0x60, x), sprm(0x2423, &[2], false)].concat()
}

#[test]
fn nested_cell_frame_leaves_unproved_grid_policies_in_flow_with_warning() {
    for extra in [
        sprm(0x3615, &[1], false),
        sprm(0x560b, &1u16.to_le_bytes(), false),
        sprm(0x7629, &[0, 1, 1, 0], false),
    ] {
        let mut row = [nested_row(500), extra].concat();
        if row.len() % 2 == 0 {
            row.extend(sprm(0x2416, &[1], false));
        }
        let bytes = nested_frame_source_with_rows(
            nested_owner_frame(-4),
            nested_owner_frame(-4),
            row.clone(),
            row,
        );
        let result = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
            .expect("previously admitted nested grid policy");
        let BodyElement::Table(outer) = &result.document.body[0] else {
            panic!("outer")
        };
        let CellElement::Table(nested) = &outer.rows[0].cells[0].content[0] else {
            panic!("nested")
        };
        let wire = serde_json::to_value(nested).unwrap();
        assert!(wire["__tableLayout"]["cellFrame"].is_null());
        assert_eq!(wire["__tableLayout"]["ordinaryFlow"], true);
        assert!(!result.document.diagnostics.is_empty());
    }
}

#[test]
fn nested_cell_frame_does_not_override_nondefault_tap_without_active_anchors() {
    for tap in [sprm(0x360d, &[0xf0], false), sprm(0x3465, &[1], false)] {
        let source = nested_frame_source(nested_owner_frame(-4), nested_owner_frame(-4));
        let text_units = "n\r\rm\r\rx\u{7}\u{7}\r".encode_utf16().count();
        let positioned_row = |extra: &[u8]| {
            let mut properties = [nested_row(500), extra.to_vec()].concat();
            if properties.len() % 2 == 0 {
                properties.extend(sprm(0x2416, &[1], false));
            }
            properties
        };
        for row_index in [0, 1] {
            let bytes = with_papx(
                &source,
                &[
                    (0, 2, nested_frame_cell(nested_owner_frame(-4))),
                    (
                        2,
                        3,
                        positioned_row(if row_index == 0 { &tap } else { &[] }),
                    ),
                    (3, 5, nested_frame_cell(nested_owner_frame(-4))),
                    (
                        5,
                        6,
                        positioned_row(if row_index == 1 { &tap } else { &[] }),
                    ),
                    (6, 8, cell()),
                    (8, 9, row(1000)),
                    (9, text_units, Vec::new()),
                ],
            );
            let model = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
                .expect("existing bounded native admission");
            let BodyElement::Table(outer) = &model.document.body[0] else {
                panic!("outer table")
            };
            // TAP row identity can split the normal first row from a later
            // positioned row. Check the group owning the authored TAP, not
            // the independently valid preceding group.
            let wanted = if row_index == 0 { "n" } else { "m" };
            let nested = outer.rows[0].cells[0]
                .content
                .iter()
                .find_map(|block| {
                    let CellElement::Table(table) = block else {
                        return None;
                    };
                    table
                        .rows
                        .iter()
                        .flat_map(|row| &row.cells)
                        .flat_map(|cell| &cell.content)
                        .any(|block| {
                            matches!(block, CellElement::Paragraph(p) if p.runs.iter()
                        .any(|run| matches!(run, DocRun::Text(t) if t.text == wanted)))
                        })
                        .then_some(table)
                })
                .expect("group owning the TAP row");
            assert!(nested.tblp_pr.is_none());
            assert!(
                serde_json::to_value(nested).unwrap()["__tableLayout"]["cellFrame"].is_null(),
                "nondefault TAP remains authoritative even without an active placement: {tap:?}, row {row_index}"
            );
        }
    }
}

#[test]
fn nested_cell_frames_keep_source_facts_and_native_row_identity() {
    for (second, expected_rows) in [(-4, vec![2]), (1441, vec![1, 1])] {
        let bytes = nested_frame_source(nested_owner_frame(-4), nested_owner_frame(second));
        let result = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
            .expect("ordinary nested-cell frames have a shared model consumer");
        let BodyElement::Table(outer) = &result.document.body[0] else {
            panic!("outer")
        };
        assert!(outer.tblp_pr.is_none());
        let nested: Vec<_> = outer.rows[0].cells[0]
            .content
            .iter()
            .filter_map(|block| match block {
                CellElement::Table(table) => Some(table),
                _ => None,
            })
            .collect();
        assert_eq!(
            nested.iter().map(|t| t.rows.len()).collect::<Vec<_>>(),
            expected_rows
        );
        for (index, table) in nested.iter().enumerate() {
            let wire = serde_json::to_value(table).unwrap();
            if index != 0 {
                assert!(wire["__tableLayout"]["cellFrame"].is_null());
                assert_eq!(wire["__tableLayout"]["ordinaryFlow"], true);
                continue;
            }
            assert_eq!(wire["__tableLayout"]["cellFrame"]["hAnchor"], "margin");
            assert_eq!(wire["__tableLayout"]["cellFrame"]["y"], 12.0);
            assert_eq!(wire["__tableLayout"]["ordinaryFlow"], false);
            assert!(
                table.tblp_pr.is_none(),
                "paragraph frames are not authored TAP positioning"
            );
        }
        let paragraphs: Vec<_> = nested
            .iter()
            .flat_map(|table| table.rows.iter())
            .flat_map(|row| row.cells.iter())
            .flat_map(|cell| cell.content.iter())
            .filter_map(|block| match block {
                CellElement::Paragraph(p) => Some(p),
                _ => None,
            })
            .collect();
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(
            paragraphs
                .iter()
                .flat_map(|p| p.runs.iter())
                .filter_map(|run| match run {
                    DocRun::Text(t) => Some(t.text.as_str()),
                    _ => None,
                })
                .collect::<Vec<_>>(),
            ["n", "m"]
        );
        for paragraph in paragraphs {
            let frame = paragraph.frame_pr.as_ref().expect("retained native fact");
            assert_eq!(
                (&*frame.h_anchor, &*frame.v_anchor, &*frame.wrap),
                ("margin", "text", "around")
            );
            assert_eq!(frame.y, Some(12.0));
        }
    }
    for extra in [
        sprm(0x2462, &[1], false),
        sprm(0x8419, &0i16.to_le_bytes(), false),
        sprm(0x841a, &400i16.to_le_bytes(), false),
        sprm(0x442c, &9u16.to_le_bytes(), false),
        sprm(0x443a, &1u16.to_le_bytes(), false),
        sprm(0x261b, &[0x50], false),
        sprm(0x2423, &[4], false),
        sprm(0x442b, &400u16.to_le_bytes(), false),
    ] {
        let bytes = nested_frame_source(
            [nested_owner_frame(-4), extra].concat(),
            nested_owner_frame(-4),
        );
        assert!(
            super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000).is_err()
        );
    }
    let root_frame = [nested_owner_frame(-4), sprm(0x2430, &[0], false)].concat();
    let root = two_row_source(root_frame.clone(), root_frame);
    assert!(super::super::direct_model(&CompoundFile::open(&root).unwrap(), 1_000_000).is_err());
}

/// Two root tables separated by a body paragraph. Each root's outer cell opens
/// with a one-row nested table whose cell paragraph carries `first` or
/// `second` frame SPRMs (empty: unframed).
/// CPs: root 1 inner cell 0..2, inner TTP 2..3, outer cell 3..5, outer TTP
/// 5..6; body paragraph 6..8; root 2 inner cell 8..10, inner TTP 10..11,
/// outer cell 11..13, outer TTP 13..14; final paragraph 14..15.
fn two_root_nested_frame_source(first: Vec<u8>, second: Vec<u8>) -> Vec<u8> {
    let text = "n\r\rx\u{7}\u{7}s\rm\r\ry\u{7}\u{7}\r";
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
        &[
            (0, 2, nested_frame_cell(first)),
            (2, 3, nested_row(500)),
            (3, 5, cell()),
            (5, 6, row(1000)),
            (6, 8, Vec::new()),
            (8, 10, nested_frame_cell(second)),
            (10, 11, nested_row(500)),
            (11, 13, cell()),
            (13, 14, row(1000)),
            (14, units, Vec::new()),
        ],
    )
}

#[test]
fn nested_cell_frame_flow_warning_names_each_root_body_element_once() {
    // The private parser/layout wire is the contract: a fixed code, the
    // native WordDocument stream as part and the final root body index only.
    let warning = |index: usize| {
        serde_json::json!({
            "code": "NATIVE_DOC_NESTED_CELL_FRAME_FLOW",
            "severity": "warning",
            "part": "WordDocument",
            "path": [index],
        })
    };
    let model = |bytes: &[u8]| {
        let result = super::super::direct_model(&CompoundFile::open(bytes).unwrap(), 1_000_000)
            .expect("admitted native model");
        let json = serde_json::to_value(&result.document).unwrap();
        (result.document, json)
    };
    let diagnostics =
        |json: &serde_json::Value| json.get("diagnostics").cloned().unwrap_or_default();

    // A homogeneous first grid acquires placement. A later grid with a
    // different row identity remains a residual frame warning for the root.
    for second in [-4, 1441] {
        let (_, json) = model(&nested_frame_source(
            nested_owner_frame(-4),
            nested_owner_frame(second),
        ));
        assert_eq!(
            diagnostics(&json),
            if second == -4 {
                serde_json::Value::Null
            } else {
                serde_json::json!([warning(0)])
            }
        );
    }

    // Each root table is named by its own final body index.
    let (document, json) = model(&two_root_nested_frame_source(
        nested_owner_frame(-4),
        nested_owner_frame(-4),
    ));
    assert!(matches!(
        document.body[..],
        [
            BodyElement::Table(_),
            BodyElement::Paragraph(_),
            BodyElement::Table(_),
            ..
        ]
    ));
    assert!(
        json.get("diagnostics").is_none(),
        "both homogeneous grids have placement owners"
    );
    let (_, json) = model(&two_root_nested_frame_source(
        Vec::new(),
        nested_owner_frame(-4),
    ));
    assert!(json.get("diagnostics").is_none());

    // Unframed nested tables emit nothing and keep the serialized model free
    // of the private diagnostics member.
    let (_, json) = model(&nested_frame_source(Vec::new(), Vec::new()));
    assert!(json.get("diagnostics").is_none());

    // A root cell frame mirroring its own table position is carried by the
    // positioned table (story.rs), not retained as a cell fact: no warning.
    // CPs: cell 0..2, TTP 2..3, final paragraph 3..4.
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
    let mirrored = with_papx(
        &source,
        &[
            (
                0,
                2,
                [
                    cell(),
                    sprm(0x261b, &[0x60], false),
                    sprm(0x8419, &159i16.to_le_bytes(), false),
                    sprm(0x2423, &[2], false),
                ]
                .concat(),
            ),
            (
                2,
                3,
                [
                    row(1000),
                    sprm(0x360d, &[0x60], false),
                    sprm(0x940f, &159i16.to_le_bytes(), false),
                    cell(),
                ]
                .concat(),
            ),
            (3, units, Vec::new()),
        ],
    );
    let (document, json) = model(&mirrored);
    let BodyElement::Table(table) = &document.body[0] else {
        panic!("positioned root table")
    };
    assert!(table.tblp_pr.is_some());
    let CellElement::Paragraph(paragraph) = &table.rows[0].cells[0].content[0] else {
        panic!("root cell paragraph")
    };
    assert!(paragraph.frame_pr.is_none());
    assert!(json.get("diagnostics").is_none());
}

#[test]
fn nested_cell_frame_flow_warning_does_not_silently_skip_exhausted_allocations() {
    let bytes = nested_frame_source(nested_owner_frame(-4), nested_owner_frame(1441));
    let result = super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
        .expect("public framed-container fixture");
    // Exercise the new allocation boundary with an already retained public
    // fixture. Zero quota refuses the traversal stack; one borrowed-table
    // slot gets past it but cannot retain diagnostic metadata. Neither may
    // return a successful, unreported cell-flow projection.
    for quota in [0, std::mem::size_of::<(&docx_model::DocTable, bool)>()] {
        let mut document = result.document.clone();
        document.diagnostics = Vec::new();
        let mut budget = super::ModelBudget::new(quota);
        assert_eq!(
            super::report_nested_cell_frame_flow(&mut document, &mut budget).unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
        assert!(document.diagnostics.is_empty());
    }
}

#[test]
fn cell_border_superseded_nested_owner_resolves_before_table_projection() {
    // The inner row uses CR + sprmPFInnerTtp at depth two, rather than the
    // depth-one cell/row U+0007 marks. Its active TTP must remain in the context
    // index and reach border resolution before nested-table output planning.
    let text = "n\r\r\u{7}\u{7}\r";
    let source = source_with_typography(
        text,
        &[(text.encode_utf16().count(), 2, 12240, 15840, 1, 720)],
        None,
        None,
        None,
        None,
    );
    let project = |inner_borders: Vec<u8>| {
        let bytes = with_papx(
            &source,
            &[
                (0, 2, nested_cell()),
                (2, 3, [nested_row(500), inner_borders].concat()),
                (3, 4, cell()),
                (4, 5, row(1000)),
                (5, 6, Vec::new()),
            ],
        );
        super::super::direct_model(&CompoundFile::open(&bytes).unwrap(), 1_000_000)
            .map(|result| result.document)
    };
    let blue = cell_border_assignment(0x0f, [0, 0, 0xff, 0, 24, 1, 0, 0]);
    let unresolved = cell_border_assignment(0x0f, [0, 0, 0, 0, 8, 0xff, 0, 0]);
    let control = project(blue.clone()).unwrap();
    let BodyElement::Table(outer) = &control.body[0] else {
        panic!("outer table")
    };
    let inner = outer.rows[0].cells[0]
        .content
        .iter()
        .find_map(|block| match block {
            CellElement::Table(table) => Some(table),
            _ => None,
        })
        .expect("nested table");
    let top = inner.rows[0].cells[0]
        .borders
        .top
        .as_ref()
        .expect("inner top border");
    assert_eq!(top.style, "single");
    assert_eq!(top.color.as_deref(), Some("0000ff"));
    let actual = project([unresolved.clone(), blue.clone()].concat()).unwrap();
    assert_eq!(
        serde_json::to_value(actual).unwrap(),
        serde_json::to_value(control).unwrap()
    );
    for borders in [unresolved.clone(), [blue, unresolved].concat()] {
        let error = project(borders).expect_err("winning nested FF must refuse");
        assert!(
            error.contains("unsupported Word border type 0xFF ignore semantics"),
            "{error}"
        );
    }
}
