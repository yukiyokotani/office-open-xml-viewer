//! SmartArt fallback: render a diagram from its data model when PowerPoint did
//! **not** persist a prebaked drawing part (`drawingN.xml` / `dsp:spTree`).
//!
//! The canonical path (`shape::parse_smartart_drawing`) replays the cached
//! `dsp:spTree` that Office bakes next to the data. Some producers — and any
//! file whose drawing part was stripped — carry only the semantic data model
//! (`ppt/diagrams/dataN.xml`). Without the layout engine we cannot reproduce
//! the true geometry (hierarchy / cycle / pyramid positions), but we can still
//! surface the *content*: every node's text, indented by its depth in the
//! parent/child tree, as a bulleted list inside the graphicFrame box.
//!
//! This is a deliberate, spec-grounded degradation, not a heuristic layout:
//!
//! * The data part is reached from the slide's `<dgm:relIds r:dm>` through the
//!   containing part's relationships (ECMA-376 §21.4.2.22 relIds → §21.4.2.10
//!   dataModel).
//! * `<dgm:ptLst>` (§21.4.3.6) holds the points; a point's displayable content
//!   lives in its `<dgm:t>` text body (`CT_Pt`/`CT_TextBody`, §21.4.3.5).
//! * Only `node`/`asst` points carry user content; `doc`, `pres`, `parTrans`
//!   and `sibTrans` are structural (ST_PtType, §21.4.7.51) and are skipped.
//! * `<dgm:cxnLst>` (§21.4.3.3) `parOf` connections (ST_CxnType, §21.4.7.23)
//!   define the parent→child tree; `srcOrd` orders siblings under a common
//!   parent (§21.4.3.2 — see [`ParOfEdge`]). A pre-order walk of that tree
//!   yields the reading order and each node's indent depth.
//!
//! Two fallback stages:
//!
//! * **M (content list):** ≥1 displayable node → one synthetic bulleted-list
//!   shape filling the frame, paragraphs = nodes in tree order, `lvl` = depth.
//! * **S (placeholder):** the data part is readable but yields no displayable
//!   text (no `ptLst`, or only structural/empty points) → a bordered box with
//!   a "SmartArt" label, so a readable-but-empty diagram is still visible.
//!
//! A missing `r:dm` relationship, a missing data part, or unparsable XML emits
//! **nothing** (same as before this fallback existed): a broken reference is
//! never turned into spurious output.
//!
//! Text colour: the data model rarely carries run colours (the real diagram
//! derives them from `colors*.xml` / `quickStyle*.xml`, which this fallback
//! intentionally discards), so runs are emitted with `color: null`. The
//! renderer gives THIS synthetic shape a contrast-aware default against the
//! slide background — a documented UX choice, not ECMA-376 semantics (the
//! spec defines no text colour for a diagram without its drawing part). See
//! `packages/pptx/src/smartart-fallback-contrast.ts`.

use crate::parse_preflighted_pptx_xml;
use crate::text::{empty_level_bullets, parse_text_body, ShapeKind};
use crate::types::*;
use crate::{attr, child, read_zip_str, resolve_path, PptxZip};
use std::collections::HashMap;

/// Classify a `<dgm:pt>` `type` (ST_PtType, ECMA-376 §21.4.7.51) as one that
/// carries displayable user content. `node` is the default when the attribute
/// is absent (`CT_Pt/@type` default = "node", §21.4.3.5). `asst` (assistant,
/// used in hierarchy diagrams) also holds real content. All other types —
/// `doc` (document root), `pres` (presentation/layout point), `parTrans`
/// (parent transition) and `sibTrans` (sibling transition) — are structural and
/// never rendered as list items.
pub(crate) fn pt_type_is_content(pt_type: Option<&str>) -> bool {
    matches!(pt_type.unwrap_or("node"), "node" | "asst")
}

/// A `parOf` connection edge (ECMA-376 §21.4.3.2 cxn / ST_CxnType §21.4.7.23):
/// `src` is the parent point's modelId, `dst` the child's. Sibling order under
/// a common parent travels on `src_ord` (`@srcOrd`) — §21.4.3.2's worked
/// example is decisive: six children of one parent (`srcId="0"`) carry
/// `srcOrd="0..5"` while every edge has `destOrd="0"`, i.e. `srcOrd` is the
/// ordinal of this connection among the source's outgoing connections.
/// `dest_ord` (`@destOrd`, the ordinal among the destination's incoming
/// connections, almost always 0) is kept as a secondary key only. Only `parOf`
/// edges build the display tree; `presOf`, `presParOf` and
/// `unknownRelationship` are ignored.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ParOfEdge {
    pub(crate) src: String,
    pub(crate) dst: String,
    pub(crate) src_ord: u32,
    pub(crate) dest_ord: u32,
}

/// Pre-order traversal of the `parOf` tree starting from `root_id`, returning
/// each reachable point's modelId paired with its depth (root's direct children
/// are depth 0). Children are visited in ascending `srcOrd` (§21.4.3.2 — the
/// sibling position under the source), with `destOrd` as a secondary key and
/// declaration order as the stable tie-break. A visited-set guards against
/// cyclic or duplicated connections so a malformed data model cannot loop.
///
/// Pure over IDs — no XML — so the tree logic is unit-tested independently of
/// the point parsing.
pub(crate) fn order_nodes_by_tree(root_id: &str, edges: &[ParOfEdge]) -> Vec<(String, u32)> {
    let mut out: Vec<(String, u32)> = Vec::new();
    let mut visited: std::collections::HashSet<String> = std::collections::HashSet::new();
    visited.insert(root_id.to_owned());
    walk(root_id, 0, edges, &mut visited, &mut out);
    out
}

fn walk(
    parent: &str,
    depth: u32,
    edges: &[ParOfEdge],
    visited: &mut std::collections::HashSet<String>,
    out: &mut Vec<(String, u32)>,
) {
    // Collect this parent's outgoing edges, preserving declaration order as the
    // stable tie-break, then sort by (srcOrd, destOrd) — srcOrd is the sibling
    // position under this parent (§21.4.3.2); a stable sort keeps the tie-break.
    let mut children: Vec<&ParOfEdge> = edges.iter().filter(|e| e.src == parent).collect();
    children.sort_by_key(|e| (e.src_ord, e.dest_ord));
    for edge in children {
        if !visited.insert(edge.dst.clone()) {
            continue; // already placed — avoid cycles / duplicate parents
        }
        out.push((edge.dst.clone(), depth));
        walk(&edge.dst, depth + 1, edges, visited, out);
    }
}

/// Entry point from the graphicFrame walker. `dm_rid` is the `<dgm:relIds r:dm>`
/// value; `rels` are the *referencing part's* relationships and `part_dir` that
/// part's directory (e.g. `ppt/slides`), so `rels[dm_rid]` resolved against
/// `part_dir` is the data part (ECMA-376 §21.4.2.22). Reads the data part,
/// emits the M-stage content list, or the S-stage placeholder when the data
/// model is readable but has no displayable text. Emits nothing (returns
/// `false`) when the relationship is missing or the data part cannot be
/// read/parsed — the caller then leaves the frame empty exactly as before, so
/// a non-diagram graphicData or a broken rel is never turned into output.
///
/// Returns `true` when it produced at least one element.
pub(crate) fn emit_smartart_fallback(
    dm_rid: &str,
    gf_xfrm: &Transform,
    part_dir: &str,
    rels: &HashMap<String, String>,
    theme: &HashMap<String, String>,
    zip: &mut PptxZip,
    out: &mut Vec<SlideElement>,
) -> bool {
    let Some(data_target) = rels.get(dm_rid) else {
        return false;
    };
    let data_path = resolve_path(part_dir, data_target);
    let Ok(data_xml) = read_zip_str(zip, &data_path) else {
        return false;
    };
    let Ok(doc) = parse_preflighted_pptx_xml(&data_xml) else {
        return false;
    };
    let root = doc.root_element();

    // §21.4.3.6 ptLst — index every point by modelId, remembering the doc-root
    // point and each point's `<dgm:t>` text body node.
    let Some(pt_lst) = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "ptLst")
    else {
        return emit_placeholder(gf_xfrm, out);
    };
    let mut points: HashMap<String, roxmltree::Node> = HashMap::new();
    let mut doc_root_id: Option<String> = None;
    for pt in pt_lst
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "pt")
    {
        let Some(model_id) = attr(&pt, "modelId") else {
            continue;
        };
        let pt_type = attr(&pt, "type");
        if pt_type.as_deref() == Some("doc") {
            doc_root_id = Some(model_id.clone());
        }
        points.insert(model_id, pt);
    }

    // §21.4.3.3 cxnLst — collect `parOf` edges (ST_CxnType §21.4.7.23).
    let mut edges: Vec<ParOfEdge> = Vec::new();
    if let Some(cxn_lst) = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "cxnLst")
    {
        for cxn in cxn_lst
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "cxn")
        {
            // CT_Cxn/@type default = "parOf" (§21.4.3.2).
            let cxn_type = attr(&cxn, "type").unwrap_or_else(|| "parOf".into());
            if cxn_type != "parOf" {
                continue;
            }
            let (Some(src), Some(dst)) = (attr(&cxn, "srcId"), attr(&cxn, "destId")) else {
                continue;
            };
            // §21.4.3.2: srcOrd carries the sibling position under the source
            // (parent); destOrd is the ordinal among the destination's incoming
            // connections. Both are xsd:unsignedInt and required by CT_Cxn, but
            // degrade to 0 if absent/unparsable rather than dropping the edge.
            let parse_ord = |name: &str| -> u32 {
                attr(&cxn, name)
                    .and_then(|v| v.parse::<u32>().ok())
                    .unwrap_or(0)
            };
            edges.push(ParOfEdge {
                src,
                dst,
                src_ord: parse_ord("srcOrd"),
                dest_ord: parse_ord("destOrd"),
            });
        }
    }

    // Order the displayable points. Prefer the tree rooted at the `doc` point;
    // if there is none (or it yields nothing), fall back to document order of
    // the content points so text is never dropped.
    let ordered: Vec<(String, u32)> = match &doc_root_id {
        Some(root_id) => order_nodes_by_tree(root_id, &edges),
        None => Vec::new(),
    };

    let mut paragraphs: Vec<Paragraph> = Vec::new();
    let mut default_font_size: Option<f64> = None;
    if !ordered.is_empty() {
        for (model_id, depth) in &ordered {
            let Some(pt) = points.get(model_id) else {
                continue;
            };
            if !pt_type_is_content(attr(pt, "type").as_deref()) {
                continue;
            }
            append_point_paragraphs(
                *pt,
                *depth,
                theme,
                zip,
                &mut paragraphs,
                &mut default_font_size,
            );
        }
    } else {
        // No usable tree: walk content points in document order at depth 0.
        for pt in pt_lst
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "pt")
        {
            if !pt_type_is_content(attr(&pt, "type").as_deref()) {
                continue;
            }
            append_point_paragraphs(pt, 0, theme, zip, &mut paragraphs, &mut default_font_size);
        }
    }

    if paragraphs.is_empty() {
        return emit_placeholder(gf_xfrm, out);
    }

    // M stage: one synthetic bulleted-list shape filling the graphicFrame box.
    // No fill / no stroke — the frame is a transparent container for the text.
    out.push(SlideElement::Shape(text_list_shape(
        gf_xfrm,
        TextBody {
            vertical_anchor: "t".into(),
            paragraphs,
            default_font_size,
            default_bold: None,
            default_italic: None,
            l_ins: 91440,
            r_ins: 91440,
            t_ins: 45720,
            b_ins: 45720,
            wrap: "square".into(),
            vert: "horz".into(),
            auto_fit: "none".into(),
            font_scale: None,
            ln_spc_reduction: None,
            num_col: 1,
            spc_col: 0,
            rtl_col: false,
            text_warp: None,
        },
    )));
    true
}

/// Parse a point's `<dgm:t>` (CT_TextBody) via the shared text-body parser and
/// push its paragraphs, re-based to the node's tree `depth` so nested nodes
/// indent. `depth` maps onto the list level (`lvl`, clamped to the 0..=8 range
/// the renderer supports). The first non-`None` default font size seen becomes
/// the list's default.
fn append_point_paragraphs(
    pt: roxmltree::Node<'_, '_>,
    depth: u32,
    theme: &HashMap<String, String>,
    zip: &mut PptxZip,
    paragraphs: &mut Vec<Paragraph>,
    default_font_size: &mut Option<f64>,
) {
    let Some(t_node) = child(pt, "t") else {
        return;
    };
    // The data part carries no picture-bullet parts we can resolve, so an empty
    // rels stub is correct; the text-body parser only needs it to verify a
    // `buBlip` part exists (it never does here → Bullet::Inherit).
    let empty_rels: HashMap<String, String> = HashMap::new();
    let body = parse_text_body(
        t_node,
        theme,
        &empty_rels,
        "",
        None,
        None,
        Default::default(),
        Default::default(),
        &empty_level_bullets(),
        None,
        None,
        None,
        None,
        None,
        None,
        None,
        None,
        None,
        None,
        None,
        None,
        ShapeKind::Sp,
        zip,
    );
    if default_font_size.is_none() {
        *default_font_size = body.default_font_size;
    }
    let lvl = depth.min(8);
    for mut para in body.paragraphs {
        // Drop wholly empty paragraphs (a node with no runs would otherwise emit
        // a blank bulleted line). Keep paragraphs that carry any run text.
        let has_text = para.runs.iter().any(|r| match r {
            TextRun::Text(t) => !t.text.is_empty(),
            _ => true,
        });
        if !has_text {
            continue;
        }
        // Re-base the list level to the tree depth. A node's own text body may
        // itself contain multiple paragraphs (rare); each takes the node depth.
        para.lvl = lvl;
        paragraphs.push(para);
    }
}

/// S stage: a bordered placeholder box with a "SmartArt" label, used when the
/// data model is readable but has no displayable text (no `ptLst`, or only
/// structural/empty points). Not reached for a missing/unreadable data part —
/// that case emits nothing (see [`emit_smartart_fallback`]).
fn emit_placeholder(gf_xfrm: &Transform, out: &mut Vec<SlideElement>) -> bool {
    let mut para = default_paragraph();
    para.alignment = "ctr".into();
    para.runs.push(TextRun::Text(TextRunData {
        text: "SmartArt".into(),
        ..default_run()
    }));
    let body = TextBody {
        vertical_anchor: "ctr".into(),
        paragraphs: vec![para],
        default_font_size: Some(18.0),
        default_bold: None,
        default_italic: None,
        l_ins: 91440,
        r_ins: 91440,
        t_ins: 45720,
        b_ins: 45720,
        wrap: "square".into(),
        vert: "horz".into(),
        auto_fit: "none".into(),
        font_scale: None,
        ln_spc_reduction: None,
        num_col: 1,
        spc_col: 0,
        rtl_col: false,
        text_warp: None,
    };
    let mut shape = text_list_shape(gf_xfrm, body);
    // A visible 1px grey border marks the diagram's extent.
    shape.stroke = Some(Stroke {
        color: "808080".into(),
        width: 9525,
        fill: None,
        dash_style: None,
        custom_dash: Vec::new(),
        line_cap: None,
        line_join: None,
        miter_limit: None,
        alignment: None,
        head_end: None,
        tail_end: None,
        cmpd: None,
    });
    out.push(SlideElement::Shape(shape));
    true
}

/// Build a synthetic rectangular `ShapeElement` filling the graphicFrame box and
/// carrying `body`. Every non-essential field is zeroed so the shape behaves
/// like a plain transparent text container.
fn text_list_shape(gf_xfrm: &Transform, body: TextBody) -> ShapeElement {
    ShapeElement {
        x: gf_xfrm.x,
        y: gf_xfrm.y,
        width: gf_xfrm.cx,
        height: gf_xfrm.cy,
        rotation: 0.0,
        flip_h: false,
        flip_v: false,
        geometry: "rect".into(),
        fill: None,
        stroke: None,
        text_body: Some(body),
        default_text_color: None,
        cust_geom: None,
        adj: None,
        adj2: None,
        adj3: None,
        adj4: None,
        adj5: None,
        adj6: None,
        adj7: None,
        adj8: None,
        shadow: None,
        inner_shadow: None,
        glow: None,
        soft_edge: None,
        reflection: None,
        id: None,
        name: Some("SmartArt".into()),
        hyperlink: None,
        hyperlink_action: None,
        placeholder_type: None,
        placeholder_idx: None,
        text_rect: None,
        scene3d: None,
        sp3d: None,
    }
}

fn default_paragraph() -> Paragraph {
    Paragraph {
        alignment: "l".into(),
        mar_l: 0,
        mar_r: 0,
        indent: 0,
        space_before: None,
        space_after: None,
        space_line: None,
        lvl: 0,
        bullet: Bullet::None,
        def_font_size: None,
        def_color: None,
        def_bold: None,
        def_italic: None,
        def_font_family: None,
        tab_stops: Vec::new(),
        def_tab_sz: None,
        rtl: false,
        ea_ln_brk: true,
        runs: Vec::new(),
    }
}

fn default_run() -> TextRunData {
    TextRunData {
        text: String::new(),
        bold: None,
        italic: None,
        underline: false,
        underline_style: None,
        underline_color: None,
        strikethrough: false,
        strike_double: false,
        font_size: None,
        color: None,
        font_family: None,
        font_family_ea: None,
        font_family_sym: None,
        baseline: None,
        caps: None,
        letter_spacing: None,
        field_type: None,
        hyperlink: None,
        hyperlink_action: None,
        shadow: None,
        reflection: None,
        outline: None,
        highlight: None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn content_types_only_node_and_asst() {
        // §21.4.7.51 ST_PtType: node (default) and asst carry content.
        assert!(pt_type_is_content(None)); // CT_Pt/@type default = "node"
        assert!(pt_type_is_content(Some("node")));
        assert!(pt_type_is_content(Some("asst")));
        // Structural point types are never displayable list items.
        assert!(!pt_type_is_content(Some("doc")));
        assert!(!pt_type_is_content(Some("pres")));
        assert!(!pt_type_is_content(Some("parTrans")));
        assert!(!pt_type_is_content(Some("sibTrans")));
    }

    fn edge(src: &str, dst: &str, src_ord: u32, dest_ord: u32) -> ParOfEdge {
        ParOfEdge {
            src: src.into(),
            dst: dst.into(),
            src_ord,
            dest_ord,
        }
    }

    #[test]
    fn tree_order_is_preorder_with_depth() {
        // doc(root) -> A(srcOrd0) -> A1; root -> B(srcOrd1) -> B1, B2
        let edges = vec![
            edge("root", "A", 0, 0),
            edge("A", "A1", 0, 0),
            edge("root", "B", 1, 0),
            edge("B", "B1", 0, 0),
            edge("B", "B2", 1, 0),
        ];
        let ordered = order_nodes_by_tree("root", &edges);
        assert_eq!(
            ordered,
            vec![
                ("A".into(), 0),
                ("A1".into(), 1),
                ("B".into(), 0),
                ("B1".into(), 1),
                ("B2".into(), 1),
            ]
        );
    }

    /// §21.4.3.2 — `srcOrd` is the sibling position under the source (parent);
    /// the spec's worked example writes `srcOrd="0..5"` / `destOrd="0"` for six
    /// children of one parent. Real files (sample-9) match: parents order their
    /// children by srcOrd while destOrd stays 0. Here destOrd is set to the
    /// exact opposite order to prove it does NOT govern.
    #[test]
    fn siblings_sorted_by_src_ord_even_when_dest_ord_disagrees() {
        let edges = vec![
            edge("root", "third", 2, 0),
            edge("root", "first", 0, 2),
            edge("root", "second", 1, 1),
        ];
        let ordered = order_nodes_by_tree("root", &edges);
        let ids: Vec<&str> = ordered.iter().map(|(id, _)| id.as_str()).collect();
        assert_eq!(ids, vec!["first", "second", "third"]);
    }

    #[test]
    fn cyclic_connections_do_not_loop() {
        // A -> B -> A would loop without the visited guard.
        let edges = vec![
            edge("root", "A", 0, 0),
            edge("A", "B", 0, 0),
            edge("B", "A", 0, 0),
        ];
        let ordered = order_nodes_by_tree("root", &edges);
        assert_eq!(ordered, vec![("A".into(), 0), ("B".into(), 1)]);
    }

    #[test]
    fn tree_order_ignores_unrelated_root() {
        let edges = vec![edge("other", "X", 0, 0)];
        assert!(order_nodes_by_tree("root", &edges).is_empty());
    }

    // ---- end-to-end: data model → emitted content list / placeholder ----

    use std::io::{Cursor, Write};

    fn zip_with(parts: &[(&str, &[u8])]) -> PptxZip {
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            for (path, bytes) in parts {
                w.start_file(*path, o).unwrap();
                w.write_all(bytes).unwrap();
            }
            w.finish().unwrap();
        }
        PptxZip::new(Cursor::new(buf)).unwrap()
    }

    const DGM_NS: &str = concat!(
        r#"xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" "#,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#
    );

    fn frame() -> Transform {
        Transform {
            x: 100,
            y: 200,
            cx: 3_000_000,
            cy: 2_000_000,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
        }
    }

    /// M stage: a data model with a two-level parOf tree yields one text-list
    /// shape whose paragraphs are the node texts in pre-order, each paragraph's
    /// `lvl` matching its tree depth (§21.4.3.3 parOf / §21.4.3.6 ptLst).
    #[test]
    fn data_model_emits_content_list_with_depths() {
        // doc(root) -> Parent(ord0) -> Child(ord0); root -> Sibling(ord1)
        let data = format!(
            r#"<dgm:dataModel {DGM_NS}>
              <dgm:ptLst>
                <dgm:pt modelId="root" type="doc"/>
                <dgm:pt modelId="P"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Parent</a:t></a:r></a:p></dgm:t></dgm:pt>
                <dgm:pt modelId="C"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Child</a:t></a:r></a:p></dgm:t></dgm:pt>
                <dgm:pt modelId="S"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Sibling</a:t></a:r></a:p></dgm:t></dgm:pt>
                <dgm:pt modelId="pres1" type="pres"/>
              </dgm:ptLst>
              <dgm:cxnLst>
                <dgm:cxn modelId="c1" type="parOf" srcId="root" destId="P" srcOrd="0" destOrd="0"/>
                <dgm:cxn modelId="c2" type="parOf" srcId="P"    destId="C" srcOrd="0" destOrd="0"/>
                <dgm:cxn modelId="c3" type="parOf" srcId="root" destId="S" srcOrd="1" destOrd="1"/>
              </dgm:cxnLst>
            </dgm:dataModel>"#
        );
        let mut zip = zip_with(&[("ppt/diagrams/data1.xml", data.as_bytes())]);
        let mut rels = HashMap::new();
        rels.insert("rId3".to_string(), "../diagrams/data1.xml".to_string());
        let theme = HashMap::new();
        let mut out = Vec::new();
        let produced = emit_smartart_fallback(
            "rId3",
            &frame(),
            "ppt/slides",
            &rels,
            &theme,
            &mut zip,
            &mut out,
        );
        assert!(produced);
        assert_eq!(out.len(), 1, "one synthetic text-list shape");
        let SlideElement::Shape(shape) = &out[0] else {
            panic!("expected a shape");
        };
        // Shape fills the graphicFrame box.
        assert_eq!((shape.x, shape.y), (100, 200));
        assert_eq!((shape.width, shape.height), (3_000_000, 2_000_000));
        let tb = shape.text_body.as_ref().unwrap();
        let lines: Vec<(&str, u32)> = tb
            .paragraphs
            .iter()
            .map(|p| {
                let txt = match p.runs.first() {
                    Some(TextRun::Text(t)) => t.text.as_str(),
                    _ => "",
                };
                (txt, p.lvl)
            })
            .collect();
        assert_eq!(
            lines,
            vec![("Parent", 0), ("Child", 1), ("Sibling", 0)],
            "pre-order traversal with depth-based list level"
        );
    }

    /// Sibling order travels on `srcOrd`, NOT `destOrd`. ECMA-376 §21.4.3.2's
    /// worked example is decisive: six children of one parent (`srcId="0"`)
    /// carry `srcOrd="0..5"` while every edge has `destOrd="0"` — the ordinal
    /// among the source's outgoing connections is the sibling position.
    /// This probe sets the two attributes to opposite orders so a destOrd
    /// (or declaration-order) implementation emits the exact reverse.
    #[test]
    fn sibling_order_follows_src_ord_in_data_model() {
        let data = format!(
            r#"<dgm:dataModel {DGM_NS}>
              <dgm:ptLst>
                <dgm:pt modelId="root" type="doc"/>
                <dgm:pt modelId="a"><dgm:t><a:bodyPr/><a:p><a:r><a:t>First</a:t></a:r></a:p></dgm:t></dgm:pt>
                <dgm:pt modelId="b"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Second</a:t></a:r></a:p></dgm:t></dgm:pt>
                <dgm:pt modelId="c"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Third</a:t></a:r></a:p></dgm:t></dgm:pt>
              </dgm:ptLst>
              <dgm:cxnLst>
                <dgm:cxn modelId="c1" type="parOf" srcId="root" destId="c" srcOrd="2" destOrd="0"/>
                <dgm:cxn modelId="c2" type="parOf" srcId="root" destId="a" srcOrd="0" destOrd="2"/>
                <dgm:cxn modelId="c3" type="parOf" srcId="root" destId="b" srcOrd="1" destOrd="1"/>
              </dgm:cxnLst>
            </dgm:dataModel>"#
        );
        let mut zip = zip_with(&[("ppt/diagrams/data1.xml", data.as_bytes())]);
        let mut rels = HashMap::new();
        rels.insert("rId3".to_string(), "../diagrams/data1.xml".to_string());
        let theme = HashMap::new();
        let mut out = Vec::new();
        emit_smartart_fallback(
            "rId3",
            &frame(),
            "ppt/slides",
            &rels,
            &theme,
            &mut zip,
            &mut out,
        );
        let SlideElement::Shape(shape) = &out[0] else {
            panic!("expected a shape");
        };
        let texts: Vec<&str> = shape
            .text_body
            .as_ref()
            .unwrap()
            .paragraphs
            .iter()
            .filter_map(|p| match p.runs.first() {
                Some(TextRun::Text(t)) => Some(t.text.as_str()),
                _ => None,
            })
            .collect();
        assert_eq!(
            texts,
            vec!["First", "Second", "Third"],
            "srcOrd (0,1,2) must govern sibling order even when destOrd (2,1,0) disagrees"
        );
    }

    /// The `pres`/`parTrans`/`sibTrans`/`doc` structural points must never
    /// surface as list items even though they may sit in the tree.
    #[test]
    fn structural_points_are_excluded_from_the_list() {
        let data = format!(
            r#"<dgm:dataModel {DGM_NS}>
              <dgm:ptLst>
                <dgm:pt modelId="root" type="doc"/>
                <dgm:pt modelId="N"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Node</a:t></a:r></a:p></dgm:t></dgm:pt>
                <dgm:pt modelId="tr" type="sibTrans"><dgm:t><a:bodyPr/><a:p><a:r><a:t>SHOULD_NOT_SHOW</a:t></a:r></a:p></dgm:t></dgm:pt>
              </dgm:ptLst>
              <dgm:cxnLst>
                <dgm:cxn modelId="c1" type="parOf" srcId="root" destId="N"  srcOrd="0" destOrd="0"/>
                <dgm:cxn modelId="c2" type="parOf" srcId="root" destId="tr" srcOrd="1" destOrd="1"/>
              </dgm:cxnLst>
            </dgm:dataModel>"#
        );
        let mut zip = zip_with(&[("ppt/diagrams/data1.xml", data.as_bytes())]);
        let mut rels = HashMap::new();
        rels.insert("rId3".to_string(), "../diagrams/data1.xml".to_string());
        let theme = HashMap::new();
        let mut out = Vec::new();
        emit_smartart_fallback(
            "rId3",
            &frame(),
            "ppt/slides",
            &rels,
            &theme,
            &mut zip,
            &mut out,
        );
        let SlideElement::Shape(shape) = &out[0] else {
            panic!("expected a shape");
        };
        let tb = shape.text_body.as_ref().unwrap();
        assert_eq!(tb.paragraphs.len(), 1);
        let txt = match tb.paragraphs[0].runs.first() {
            Some(TextRun::Text(t)) => t.text.as_str(),
            _ => "",
        };
        assert_eq!(txt, "Node");
    }

    /// With no `doc` root and no cxns, content points are still emitted in
    /// document order at depth 0 — text is never silently dropped.
    #[test]
    fn falls_back_to_document_order_without_a_tree() {
        let data = format!(
            r#"<dgm:dataModel {DGM_NS}>
              <dgm:ptLst>
                <dgm:pt modelId="a"><dgm:t><a:bodyPr/><a:p><a:r><a:t>First</a:t></a:r></a:p></dgm:t></dgm:pt>
                <dgm:pt modelId="b"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Second</a:t></a:r></a:p></dgm:t></dgm:pt>
              </dgm:ptLst>
            </dgm:dataModel>"#
        );
        let mut zip = zip_with(&[("ppt/diagrams/data1.xml", data.as_bytes())]);
        let mut rels = HashMap::new();
        rels.insert("rId3".to_string(), "../diagrams/data1.xml".to_string());
        let theme = HashMap::new();
        let mut out = Vec::new();
        emit_smartart_fallback(
            "rId3",
            &frame(),
            "ppt/slides",
            &rels,
            &theme,
            &mut zip,
            &mut out,
        );
        let SlideElement::Shape(shape) = &out[0] else {
            panic!("expected a shape");
        };
        let tb = shape.text_body.as_ref().unwrap();
        let texts: Vec<&str> = tb
            .paragraphs
            .iter()
            .filter_map(|p| match p.runs.first() {
                Some(TextRun::Text(t)) => Some(t.text.as_str()),
                _ => None,
            })
            .collect();
        assert_eq!(texts, vec!["First", "Second"]);
    }

    /// S stage: a data model with no displayable text emits a bordered
    /// "SmartArt" placeholder rather than nothing.
    #[test]
    fn empty_data_model_emits_placeholder() {
        let data = format!(
            r#"<dgm:dataModel {DGM_NS}>
              <dgm:ptLst>
                <dgm:pt modelId="root" type="doc"/>
                <dgm:pt modelId="pres1" type="pres"/>
              </dgm:ptLst>
            </dgm:dataModel>"#
        );
        let mut zip = zip_with(&[("ppt/diagrams/data1.xml", data.as_bytes())]);
        let mut rels = HashMap::new();
        rels.insert("rId3".to_string(), "../diagrams/data1.xml".to_string());
        let theme = HashMap::new();
        let mut out = Vec::new();
        let produced = emit_smartart_fallback(
            "rId3",
            &frame(),
            "ppt/slides",
            &rels,
            &theme,
            &mut zip,
            &mut out,
        );
        assert!(produced);
        let SlideElement::Shape(shape) = &out[0] else {
            panic!("expected a shape");
        };
        assert!(shape.stroke.is_some(), "placeholder draws a border");
        let tb = shape.text_body.as_ref().unwrap();
        let txt = match tb.paragraphs[0].runs.first() {
            Some(TextRun::Text(t)) => t.text.as_str(),
            _ => "",
        };
        assert_eq!(txt, "SmartArt");
    }

    /// An unreadable data part (missing `rels` entry) emits nothing — a
    /// non-diagram or broken relationship is never turned into spurious output.
    #[test]
    fn missing_data_part_emits_nothing() {
        let mut zip = zip_with(&[]);
        let rels = HashMap::new();
        let theme = HashMap::new();
        let mut out = Vec::new();
        let produced = emit_smartart_fallback(
            "rId3",
            &frame(),
            "ppt/slides",
            &rels,
            &theme,
            &mut zip,
            &mut out,
        );
        assert!(!produced);
        assert!(out.is_empty());
    }
}
