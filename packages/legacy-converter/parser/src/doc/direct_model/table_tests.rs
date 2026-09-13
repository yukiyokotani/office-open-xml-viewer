//! Full-container acceptance fixtures for direct binary table projection.

use super::tests::{source_with_typography, with_picture_data};
use crate::cfb::{test_support::build_cfb, CompoundFile};
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
        let cb = (papx.len() + 1) / 2;
        payload -= 1 + papx.len();
        payload &= !1;
        page[payload] = cb as u8;
        page[payload + 1..payload + 1 + papx.len()].copy_from_slice(&papx);
        page[bx + index * 13] = (payload / 2) as u8;
    }
    page[511] = n as u8;
    build_cfb(&[("WordDocument", word), ("0Table", table)])
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
    let paragraph = || {
        Blocks(vec![Block::Paragraph(Box::new(
            docx_model::DocParagraph::default(),
        ))])
    };
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
    let paragraph = || Block::Paragraph(Box::new(docx_model::DocParagraph::default()));
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
