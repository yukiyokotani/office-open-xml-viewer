//! Shared paragraph hard-break normalization policy. Authoritative page and
//! column breaks become body-level pieces; rendered-page cache hints are
//! removed. Whitespace and line-break-only chunks are not visible paragraphs.
//! Leading empty chunks are omitted, while trailing hard breaks are retained
//! and marked as belonging to the source paragraph. splitPgBreakAndParaMark
//! additionally retains the final paragraph mark after a terminal page break.

use crate::{BreakType, ComplexFieldBoundaryWire, DocParagraph, DocRun};

#[cfg(test)]
thread_local! {
    static WORK: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

#[inline]
fn work(units: usize) {
    #[cfg(test)]
    WORK.with(|value| value.set(value.get() + units));
    #[cfg(not(test))]
    let _ = units;
}

#[allow(clippy::large_enum_variant)]
pub enum ParaPiece {
    Para(DocParagraph),
    PageBreak { same_paragraph_as_previous: bool },
    ColumnBreak,
}

/// Normalize hard page and column breaks into body-level pieces. Rendered-page
/// hints are removed. Run payloads and field boundaries move once; only the
/// paragraph metadata, with its owning vectors emptied, is cloned per chunk.
pub fn split_para_on_page_breaks(
    para: DocParagraph,
    split_pg_break_and_para_mark: bool,
) -> Vec<ParaPiece> {
    let mut pieces = Vec::new();
    let result = visit_para_on_page_breaks(para, split_pg_break_and_para_mark, |piece| {
        pieces.push(piece);
        Ok::<(), std::convert::Infallible>(())
    });
    match result {
        Ok(()) => pieces,
        Err(never) => match never {},
    }
}

pub fn visit_para_on_page_breaks<E>(
    mut para: DocParagraph,
    split_pg_break_and_para_mark: bool,
    mut emit: impl FnMut(ParaPiece) -> Result<(), E>,
) -> Result<(), E> {
    let is_hard_break = |run: &DocRun| {
        matches!(
            run,
            DocRun::Break {
                break_type: BreakType::Page | BreakType::Column
            }
        )
    };
    if !para.runs.iter().any(is_hard_break) {
        if !para.complex_field_boundaries.is_empty() {
            let mut retained_prefix = Vec::with_capacity(para.runs.len() + 1);
            retained_prefix.push(0);
            for run in &para.runs {
                work(1);
                let retained = !matches!(
                    run,
                    DocRun::Break {
                        break_type: BreakType::RenderedPage
                    }
                );
                retained_prefix
                    .push(retained_prefix.last().copied().unwrap_or(0) + usize::from(retained));
            }
            for boundary in &mut para.complex_field_boundaries {
                work(1);
                boundary.run_index = retained_prefix[boundary.run_index.min(para.runs.len())];
            }
        }
        para.runs.retain(|run| {
            !matches!(
                run,
                DocRun::Break {
                    break_type: BreakType::RenderedPage
                }
            )
        });
        emit(ParaPiece::Para(para))?;
        return Ok(());
    }

    // ECMA-376 Part 4 §14.8.3.38: only a terminal page break moves the
    // paragraph mark to a new line on the next page. Keep that final empty
    // paragraph with the source paragraph's mark formatting and properties. Word
    // 16 mode 11 also lets this line displace a following near-page-height
    // table; ordinary empty-paragraph layout supplies that behavior unchanged.
    // Rendered-page hints are ignored throughout normalization, not content.
    let retain_mark = split_pg_break_and_para_mark
        && matches!(
            para.runs.iter().rev().find(|run| !matches!(
                run,
                DocRun::Break {
                    break_type: BreakType::RenderedPage
                }
            )),
            Some(DocRun::Break {
                break_type: BreakType::Page
            })
        );
    let runs = std::mem::take(&mut para.runs);
    let boundaries = std::mem::take(&mut para.complex_field_boundaries);
    let run_count = runs.len();
    let mut boundaries_at =
        (!boundaries.is_empty()).then(|| (0..=run_count).map(|_| Vec::new()).collect::<Vec<_>>());
    for boundary in boundaries {
        work(1);
        if let Some(slot) = boundaries_at
            .as_mut()
            .and_then(|values| values.get_mut(boundary.run_index))
        {
            slot.push(boundary);
        }
    }

    let mut chunks: Vec<Vec<DocRun>> = vec![Vec::new()];
    let mut boundary_chunks: Vec<Vec<ComplexFieldBoundaryWire>> = vec![Vec::new()];
    let mut seps = Vec::new();
    for (index, run) in runs.into_iter().enumerate() {
        work(1);
        let local_index = chunks.last().map(Vec::len).unwrap_or(0);
        if let Some(boundaries_at) = &mut boundaries_at {
            for mut boundary in std::mem::take(&mut boundaries_at[index]) {
                boundary.run_index = local_index;
                boundary_chunks
                    .last_mut()
                    .expect("seeded with chunks")
                    .push(boundary);
            }
        }
        match run {
            DocRun::Break {
                break_type: BreakType::Page,
            } => {
                chunks.push(Vec::new());
                boundary_chunks.push(Vec::new());
                seps.push(ParaPiece::PageBreak {
                    same_paragraph_as_previous: false,
                });
            }
            DocRun::Break {
                break_type: BreakType::Column,
            } => {
                chunks.push(Vec::new());
                boundary_chunks.push(Vec::new());
                seps.push(ParaPiece::ColumnBreak);
            }
            DocRun::Break {
                break_type: BreakType::RenderedPage,
            } => {}
            run => chunks.last_mut().expect("seeded").push(run),
        }
    }
    let local_index = chunks.last().map(Vec::len).unwrap_or(0);
    if let Some(boundaries_at) = &mut boundaries_at {
        for mut boundary in std::mem::take(&mut boundaries_at[run_count]) {
            boundary.run_index = local_index;
            boundary_chunks.last_mut().expect("seeded").push(boundary);
        }
    }

    let has_visible = |runs: &[DocRun]| {
        runs.iter().any(|run| {
            matches!(run, DocRun::Text(text) if !text.text.trim().is_empty())
                || matches!(
                    run,
                    DocRun::Field(_) | DocRun::Image(_) | DocRun::Chart(_) | DocRun::Shape(_)
                )
        })
    };
    let mut retained_chunks = chunks.len();
    while retained_chunks > 1 && !has_visible(&chunks[retained_chunks - 1]) {
        work(1);
        retained_chunks -= 1;
    }
    let trailing = seps.split_off(retained_chunks - 1);
    let removed_boundaries = boundary_chunks.split_off(retained_chunks);
    chunks.truncate(retained_chunks);
    let end = chunks.last().map(Vec::len).unwrap_or(0);
    let previous = boundary_chunks.last_mut().expect("seeded with chunks");
    for mut boundary in removed_boundaries.into_iter().flatten() {
        work(1);
        boundary.run_index = end;
        previous.push(boundary);
    }
    if chunks.first().is_some_and(|runs| !has_visible(runs)) && boundary_chunks.len() > 1 {
        let mut migrated = std::mem::take(&mut boundary_chunks[0]);
        for boundary in &mut migrated {
            boundary.run_index = 0;
        }
        migrated.extend(std::mem::take(&mut boundary_chunks[1]));
        boundary_chunks[1] = migrated;
    }

    let visibility: Vec<bool> = chunks.iter().map(|runs| has_visible(runs)).collect();
    let mut emitted = false;
    let mut emitted_paragraph = false;
    for (index, runs) in chunks.into_iter().enumerate() {
        if index == 0 && !has_visible(&runs) {
            continue;
        }
        if index > 0 {
            emit(match seps.get(index - 1) {
                Some(ParaPiece::ColumnBreak) => ParaPiece::ColumnBreak,
                _ => ParaPiece::PageBreak {
                    same_paragraph_as_previous: visibility.get(index - 1).copied().unwrap_or(false),
                },
            })?;
        }
        let mut chunk = para.clone();
        chunk.runs = runs;
        chunk.complex_field_boundaries = std::mem::take(&mut boundary_chunks[index]);
        // A break follows this piece: an interior separator, or a trailing one
        // after the last retained chunk.
        if index < seps.len() || (index + 1 == retained_chunks && !trailing.is_empty()) {
            chunk.ends_at_hard_break = Some(true);
        }
        emit(ParaPiece::Para(chunk))?;
        emitted = true;
        emitted_paragraph = true;
    }
    for sep in trailing {
        emit(match sep {
            ParaPiece::ColumnBreak => ParaPiece::ColumnBreak,
            _ => ParaPiece::PageBreak {
                same_paragraph_as_previous: true,
            },
        })?;
        emitted = true;
    }
    if retain_mark || !emitted {
        if retain_mark && emitted_paragraph {
            // ECMA-376 Part 4 §14.8.3.38 moves the existing paragraph's
            // mark; it does not introduce another list item (§17.9) or source
            // paragraph. Normalization policy: the content piece owns its
            // marker, decorations, frame, outline entry, and source anchors
            // (§17.13.4/§17.13.6, MS-DOCX §2.6.2.3). The synthetic mark line
            // retains font/line metrics, indents, and spacing, but must not
            // repaint those owners or reapply pageBreakBefore. Clear both
            // border projections so typography acquisition cannot restore ink.
            // With no content piece, this is the only paragraph and keeps its
            // ownership metadata; the page-break pieces do not own it.
            para.numbering = None;
            para.paragraph_id = None;
            para.bookmarks.clear();
            para.comment_marks.clear();
            para.run_revisions.clear();
            para.borders = None;
            para.paragraph_typography_acquisition = None;
            para.shading = None;
            para.frame_pr = None;
            para.outline_level = None;
            para.page_break_before = false;
        }
        emit(ParaPiece::Para(para))?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn measured<T>(operation: impl FnOnce() -> T) -> (T, usize) {
        WORK.with(|value| value.set(0));
        let result = operation();
        let work = WORK.with(|value| value.get());
        (result, work)
    }

    #[test]
    fn split_mark_follows_only_the_last_page_break_and_leaves_columns_unchanged() {
        for (last, paragraph_count) in [(BreakType::Page, 1), (BreakType::Column, 0)] {
            let pieces = split_para_on_page_breaks(
                DocParagraph {
                    runs: vec![
                        DocRun::Break {
                            break_type: BreakType::Page,
                        },
                        DocRun::Break { break_type: last },
                    ],
                    ..Default::default()
                },
                true,
            );
            assert_eq!(pieces.len(), 2 + paragraph_count);
            assert!(matches!(pieces[0], ParaPiece::PageBreak { .. }));
            if paragraph_count == 1 {
                assert!(matches!(pieces[1], ParaPiece::PageBreak { .. }));
                assert!(matches!(&pieces[2], ParaPiece::Para(p) if p.runs.is_empty()));
            } else {
                assert!(matches!(pieces[1], ParaPiece::ColumnBreak));
            }
        }
    }

    #[test]
    fn sink_error_stops_before_later_paragraph_emission() {
        let para = DocParagraph {
            runs: vec![
                DocRun::Text(Box::new(crate::TextRun {
                    text: "a".into(),
                    ..Default::default()
                })),
                DocRun::Break {
                    break_type: BreakType::Page,
                },
                DocRun::Text(Box::new(crate::TextRun {
                    text: "b".into(),
                    ..Default::default()
                })),
            ],
            ..Default::default()
        };
        let mut calls = 0;
        let result = visit_para_on_page_breaks(para, false, |_| {
            calls += 1;
            Err("stop")
        });
        assert_eq!(result, Err("stop"));
        assert_eq!(calls, 1);
    }

    #[test]
    fn fragmented_runs_are_processed_without_quadratic_boundary_scans() {
        let mut para = DocParagraph::default();
        for index in 0..20_000 {
            para.runs.push(if index % 2 == 0 {
                DocRun::Text(Box::new(crate::TextRun {
                    text: "x".into(),
                    ..Default::default()
                }))
            } else {
                DocRun::Break {
                    break_type: BreakType::Page,
                }
            });
            para.complex_field_boundaries
                .push(ComplexFieldBoundaryWire {
                    occurrence_id: index as u32,
                    boundary: "start".into(),
                    run_index: index,
                    field_type: "other".into(),
                    instruction: String::new(),
                    hyperlink_anchor: None,
                });
        }
        let mut count = 0;
        let (_, work) = measured(|| {
            visit_para_on_page_breaks(para, false, |_| {
                count += 1;
                Ok::<(), ()>(())
            })
            .unwrap()
        });
        assert_eq!(count, 20_000);
        assert!(work <= 60_000, "work={work}");
    }

    #[test]
    fn no_hard_break_prefix_remap_is_linear_in_runs_and_boundaries() {
        let mut para = DocParagraph::default();
        for index in 0..20_000 {
            para.runs.push(if index % 2 == 0 {
                DocRun::Break {
                    break_type: BreakType::RenderedPage,
                }
            } else {
                DocRun::Text(Box::new(crate::TextRun {
                    text: "x".into(),
                    ..Default::default()
                }))
            });
            para.complex_field_boundaries
                .push(ComplexFieldBoundaryWire {
                    occurrence_id: index as u32,
                    boundary: "start".into(),
                    run_index: index,
                    field_type: "other".into(),
                    instruction: String::new(),
                    hyperlink_anchor: None,
                });
        }
        let (_, work) = measured(|| split_para_on_page_breaks(para, false));
        assert_eq!(work, 40_000);
    }

    #[test]
    fn trailing_boundary_relocation_moves_each_boundary_once() {
        let mut para = DocParagraph::default();
        for index in 0..20_000 {
            para.runs.push(DocRun::Break {
                break_type: BreakType::Page,
            });
            para.complex_field_boundaries
                .push(ComplexFieldBoundaryWire {
                    occurrence_id: index as u32,
                    boundary: "start".into(),
                    run_index: index,
                    field_type: "other".into(),
                    instruction: String::new(),
                    hyperlink_anchor: None,
                });
        }
        let (_, work) = measured(|| split_para_on_page_breaks(para, false));
        assert!(work <= 80_000, "work={work}");
    }
}
