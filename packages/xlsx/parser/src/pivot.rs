use crate::read_zip_string;
use ooxml_common::{
    depth::parse_guarded,
    ns::{is_x_ns, relationships},
};

const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const PIVOT_TABLE_REL_TRANSITIONAL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
const PIVOT_TABLE_REL_STRICT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/pivotTable";
const PIVOT_CACHE_REL_TRANSITIONAL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
const PIVOT_CACHE_REL_STRICT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/pivotCacheDefinition";

use crate::{
    resolve_zip_path, CellRange, Dxf, PivotAxisItem, PivotCacheSource, PivotDataField,
    PivotDiagnostic, PivotDiagnosticReason, PivotLocation, PivotMetadataStatus, PivotPageField,
    PivotPartialReason, PivotTableMetadata, PivotTableStyle, PivotTableStyleElement, XlsxZip,
};

/// Parse pivot metadata without evaluating it. Saved worksheet cells and styles
/// remain authoritative and this module never reads pivotCacheRecords.
pub(crate) fn load_sheet_pivots(
    archive: &mut XlsxZip,
    sheet_path: &str,
    theme_colors: &[String],
) -> (Vec<PivotTableMetadata>, Vec<PivotDiagnostic>) {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return (Vec::new(), Vec::new());
    };
    let rels_path = format!("xl/{sheet_dir}/_rels/{sheet_file}.rels");
    let rels_exists = archive.index_for_name(&rels_path).is_some();
    let Ok(rels_xml) = read_zip_string(archive, &rels_path) else {
        return if rels_exists {
            (
                Vec::new(),
                vec![PivotDiagnostic {
                    part: rels_path,
                    reason: PivotDiagnosticReason::UnreadableWorksheetRelationships,
                }],
            )
        } else {
            (Vec::new(), Vec::new())
        };
    };
    let Ok(rels) = parse_guarded(&rels_xml) else {
        return (
            Vec::new(),
            vec![PivotDiagnostic {
                part: rels_path,
                reason: PivotDiagnosticReason::MalformedWorksheetRelationships,
            }],
        );
    };
    let root = rels.root_element();
    if root.tag_name().name() != "Relationships"
        || root.tag_name().namespace() != Some(PACKAGE_REL_NS)
    {
        return (
            Vec::new(),
            vec![PivotDiagnostic {
                part: rels_path,
                reason: PivotDiagnosticReason::MalformedWorksheetRelationships,
            }],
        );
    }
    let base = format!("xl/{sheet_dir}");
    let mut tables = Vec::new();
    let mut diagnostics = Vec::new();
    let mut targets = Vec::new();
    for relationship in root.children().filter(|node| node.is_element()) {
        if relationship.tag_name().name() != "Relationship"
            || relationship.tag_name().namespace() != Some(PACKAGE_REL_NS)
        {
            diagnostics.push(PivotDiagnostic {
                part: rels_path.clone(),
                reason: PivotDiagnosticReason::MalformedPivotRelationship,
            });
            continue;
        }
        let rel_type = relationship.attribute("Type");
        if relationship.attribute("Id").is_none_or(|id| id.is_empty()) {
            diagnostics.push(PivotDiagnostic {
                part: rels_path.clone(),
                reason: if matches!(
                    rel_type,
                    Some(PIVOT_TABLE_REL_TRANSITIONAL | PIVOT_TABLE_REL_STRICT)
                ) {
                    PivotDiagnosticReason::MalformedPivotRelationship
                } else {
                    PivotDiagnosticReason::MalformedWorksheetRelationships
                },
            });
            continue;
        }
        let Some(rel_type) = rel_type else {
            diagnostics.push(PivotDiagnostic {
                part: rels_path.clone(),
                reason: PivotDiagnosticReason::MalformedWorksheetRelationships,
            });
            continue;
        };
        if rel_type != PIVOT_TABLE_REL_TRANSITIONAL && rel_type != PIVOT_TABLE_REL_STRICT {
            if relationship
                .attribute("Target")
                .is_none_or(|target| target.is_empty())
            {
                diagnostics.push(PivotDiagnostic {
                    part: rels_path.clone(),
                    reason: PivotDiagnosticReason::MalformedWorksheetRelationships,
                });
            }
            continue;
        }
        match relationship.attribute("TargetMode") {
            Some("External") => {
                diagnostics.push(PivotDiagnostic {
                    part: rels_path.clone(),
                    reason: PivotDiagnosticReason::ExternalPivotRelationship,
                });
                continue;
            }
            None | Some("Internal") => {}
            Some(_) => {
                diagnostics.push(PivotDiagnostic {
                    part: rels_path.clone(),
                    reason: PivotDiagnosticReason::MalformedPivotRelationship,
                });
                continue;
            }
        }
        let Some(target) = relationship
            .attribute("Target")
            .filter(|target| !target.is_empty())
        else {
            diagnostics.push(PivotDiagnostic {
                part: rels_path.clone(),
                reason: PivotDiagnosticReason::MalformedPivotRelationship,
            });
            continue;
        };
        targets.push(resolve_zip_path(&base, target));
    }
    let mut workbook_styles: Option<WorkbookTableStyles> = None;
    for part in targets {
        let xml = match read_zip_string(archive, &part) {
            Ok(xml) => xml,
            Err(_) => {
                diagnostics.push(PivotDiagnostic {
                    part,
                    reason: PivotDiagnosticReason::UnreadablePart,
                });
                continue;
            }
        };
        match parse_pivot_table(archive, &part, &xml, theme_colors, &mut workbook_styles) {
            Ok(table) => tables.push(table),
            Err(reason) => diagnostics.push(PivotDiagnostic { part, reason }),
        }
    }
    (tables, diagnostics)
}

fn parse_pivot_table(
    archive: &mut XlsxZip,
    part: &str,
    xml: &str,
    theme_colors: &[String],
    workbook_styles: &mut Option<WorkbookTableStyles>,
) -> Result<PivotTableMetadata, PivotDiagnosticReason> {
    let doc = parse_guarded(xml).map_err(|_| PivotDiagnosticReason::MalformedXml)?;
    let root = doc.root_element();
    if root.tag_name().name() != "pivotTableDefinition" || !is_x_ns(root.tag_name().namespace()) {
        return Err(PivotDiagnosticReason::MalformedXml);
    }
    let name = root
        .attribute("name")
        .filter(|v| !v.trim().is_empty())
        .ok_or(PivotDiagnosticReason::MissingIdentity)?
        .to_string();
    let cache_id = root
        .attribute("cacheId")
        .and_then(|v| v.parse().ok())
        .ok_or(PivotDiagnosticReason::MissingIdentity)?;
    root.attribute("dataCaption")
        .filter(|value| !value.trim().is_empty())
        .ok_or(PivotDiagnosticReason::MissingIdentity)?;
    let location = child(root, "location")
        .and_then(parse_location)
        .ok_or(PivotDiagnosticReason::InvalidLocation)?;

    let (row_fields, malformed_rows) = parse_signed_fields(root, "rowFields");
    let (column_fields, malformed_columns) = parse_signed_fields(root, "colFields");
    let mut malformed_pages = false;
    let page_fields = child(root, "pageFields")
        .map(|parent| {
            parent
                .children()
                .filter(|n| {
                    n.is_element()
                        && n.tag_name().name() == "pageField"
                        && is_x_ns(n.tag_name().namespace())
                })
                .filter_map(|n| match n.attribute("fld").and_then(|v| v.parse().ok()) {
                    Some(field) => Some(PivotPageField {
                        field,
                        item: match n.attribute("item") {
                            Some(value) => match value.parse().ok() {
                                Some(item) => Some(item),
                                None => {
                                    malformed_pages = true;
                                    None
                                }
                            },
                            None => None,
                        },
                        name: n.attribute("name").map(str::to_string),
                    }),
                    None => {
                        malformed_pages = true;
                        None
                    }
                })
                .collect()
        })
        .unwrap_or_default();
    let mut malformed_data_field = false;
    let mut malformed_data_subtotal = false;
    let mut data_semantic_features = Vec::new();
    let data_fields = child(root, "dataFields")
        .map(|parent| {
            parent
                .children()
                .filter(|n| {
                    n.is_element()
                        && n.tag_name().name() == "dataField"
                        && is_x_ns(n.tag_name().namespace())
                })
                .filter_map(|n| match n.attribute("fld").and_then(|v| v.parse().ok()) {
                    Some(field) => {
                        let raw_subtotal = n.attribute("subtotal");
                        let subtotal = raw_subtotal.unwrap_or("sum");
                        let subtotal_valid = is_data_subtotal(subtotal);
                        if !subtotal_valid {
                            malformed_data_subtotal = true;
                        }
                        if n.attribute("showDataAs")
                            .is_some_and(|value| value != "normal")
                        {
                            data_semantic_features.push("dataField.showDataAs".to_string());
                        }
                        if n.attribute("baseField").is_some_and(|value| value != "-1")
                            || n.attribute("baseItem")
                                .is_some_and(|value| value != "1048832")
                        {
                            data_semantic_features.push("dataField.base".to_string());
                        }
                        Some(PivotDataField {
                            field,
                            subtotal: subtotal_valid.then(|| subtotal.to_string()),
                            raw_subtotal: (!subtotal_valid).then(|| subtotal.to_string()),
                            name: n.attribute("name").map(str::to_string),
                        })
                    }
                    None => {
                        malformed_data_field = true;
                        None
                    }
                })
                .collect()
        })
        .unwrap_or_default();

    let mut reasons = unsupported_features(root);
    match parse_bool(root.attribute("dataOnRows")) {
        Ok(true) => reasons.push(PivotPartialReason::UnsupportedSemanticFeature {
            feature: "pivotTable.dataPlacement".into(),
        }),
        Ok(false) => {}
        Err(()) => reasons.push(PivotPartialReason::MalformedField {
            field: "pivotTable.dataOnRows".into(),
        }),
    }
    if let Some(data_position) = root.attribute("dataPosition") {
        if data_position.parse::<u32>().is_ok() {
            reasons.push(PivotPartialReason::UnsupportedSemanticFeature {
                feature: "pivotTable.dataPlacement".into(),
            });
        } else {
            reasons.push(PivotPartialReason::MalformedField {
                field: "pivotTable.dataPosition".into(),
            });
        }
    }
    if malformed_rows {
        reasons.push(PivotPartialReason::MalformedField {
            field: "rowFields".into(),
        });
    }
    if malformed_columns {
        reasons.push(PivotPartialReason::MalformedField {
            field: "columnFields".into(),
        });
    }
    if malformed_pages {
        reasons.push(PivotPartialReason::MalformedField {
            field: "pageFields".into(),
        });
    }
    if malformed_data_field {
        reasons.push(PivotPartialReason::MalformedField {
            field: "dataFields.field".into(),
        });
    }
    if malformed_data_subtotal {
        reasons.push(PivotPartialReason::MalformedField {
            field: "dataFields.subtotal".into(),
        });
    }
    reasons.extend(
        data_semantic_features
            .into_iter()
            .map(|feature| PivotPartialReason::UnsupportedSemanticFeature { feature }),
    );
    let mut refresh_on_load = None;
    let mut cache_invalid = None;
    let mut cache_definition_part = None;
    let mut cache_source = None;
    let mut recorded_extension_uris = extension_uris(root);
    match pivot_cache_target(archive, part) {
        CacheLink::Missing => reasons.push(PivotPartialReason::MissingCacheRelationship),
        CacheLink::Malformed => reasons.push(PivotPartialReason::MalformedCacheRelationships),
        CacheLink::Unreadable => reasons.push(PivotPartialReason::UnreadableCacheRelationships),
        CacheLink::External => reasons.push(PivotPartialReason::ExternalCacheRelationship),
        CacheLink::Ambiguous => reasons.push(PivotPartialReason::AmbiguousCacheRelationship),
        CacheLink::Target(target) => {
            cache_definition_part = Some(target.clone());
            match read_zip_string(archive, &target) {
                Err(_) => reasons.push(PivotPartialReason::UnreadableCacheDefinition),
                Ok(cache_xml) => match parse_guarded(&cache_xml) {
                    Err(_) => reasons.push(PivotPartialReason::MalformedCacheDefinition),
                    Ok(cache_doc) => {
                        let cache_root = cache_doc.root_element();
                        if cache_root.tag_name().name() != "pivotCacheDefinition"
                            || !is_x_ns(cache_root.tag_name().namespace())
                        {
                            reasons.push(PivotPartialReason::MalformedCacheDefinition);
                        } else {
                            match parse_bool(cache_root.attribute("refreshOnLoad")) {
                                Ok(value) => refresh_on_load = Some(value),
                                Err(()) => reasons.push(PivotPartialReason::MalformedField {
                                    field: "refreshOnLoad".into(),
                                }),
                            }
                            match parse_bool(cache_root.attribute("invalid")) {
                                Ok(value) => cache_invalid = Some(value),
                                Err(()) => reasons.push(PivotPartialReason::MalformedField {
                                    field: "cache.invalid".into(),
                                }),
                            }
                            recorded_extension_uris.extend(extension_uris(cache_root));
                            if let Some(source) = child(cache_root, "cacheSource") {
                                let kind = source.attribute("type").unwrap_or("");
                                if kind.is_empty() {
                                    reasons.push(PivotPartialReason::MalformedCacheDefinition);
                                }
                                cache_source = match kind {
                                    "worksheet" => {
                                        let ws = child(source, "worksheetSource");
                                        let relationship_id = ws.and_then(|node| {
                                            node.attribute((relationships::TRANSITIONAL, "id"))
                                                .or_else(|| {
                                                    node.attribute((relationships::STRICT, "id"))
                                                })
                                        });
                                        let sheet = ws
                                            .and_then(|node| node.attribute("sheet"))
                                            .map(str::to_string);
                                        let raw_reference =
                                            ws.and_then(|node| node.attribute("ref"));
                                        let reference = raw_reference
                                            .filter(|value| parse_a1_range(value).is_some())
                                            .map(str::to_string);
                                        if raw_reference.is_some() && reference.is_none() {
                                            reasons.push(PivotPartialReason::MalformedField {
                                                field: "cacheSource.worksheetSource.ref".into(),
                                            });
                                        }
                                        let name = ws
                                            .and_then(|node| node.attribute("name"))
                                            .map(str::to_string);
                                        if sheet.is_none()
                                            && reference.is_none()
                                            && name.is_none()
                                            && relationship_id.is_none()
                                        {
                                            reasons
                                                .push(PivotPartialReason::MalformedCacheDefinition);
                                        }
                                        if relationship_id.is_some() {
                                            reasons.push(
                                                PivotPartialReason::UnresolvedWorksheetSourceRelationship,
                                            );
                                        }
                                        Some(PivotCacheSource::Worksheet {
                                            sheet,
                                            reference,
                                            name,
                                            relationship_id: relationship_id.map(str::to_string),
                                        })
                                    }
                                    "external" => Some(PivotCacheSource::External),
                                    "consolidation" => Some(PivotCacheSource::Consolidation),
                                    "scenario" => Some(PivotCacheSource::Scenario),
                                    _ => None,
                                };
                                if !kind.is_empty() && kind != "worksheet" {
                                    reasons.push(PivotPartialReason::UnsupportedCacheSource {
                                        source_type: kind.to_string(),
                                    });
                                }
                            } else {
                                reasons.push(PivotPartialReason::MalformedCacheDefinition);
                            }
                            if child(cache_root, "cacheFields").is_none() {
                                reasons.push(PivotPartialReason::MalformedCacheDefinition);
                            }
                            for feature in [
                                "tupleCache",
                                "calculatedItems",
                                "calculatedMembers",
                                "dimensions",
                                "measureGroups",
                                "maps",
                            ] {
                                if child(cache_root, feature).is_some() {
                                    reasons.push(PivotPartialReason::UnsupportedSemanticFeature {
                                        feature: feature.to_string(),
                                    });
                                }
                            }
                            if cache_root.descendants().any(|node| {
                                node.is_element()
                                    && node.tag_name().name() == "fieldGroup"
                                    && is_x_ns(node.tag_name().namespace())
                            }) {
                                reasons.push(PivotPartialReason::UnsupportedSemanticFeature {
                                    feature: "fieldGroup".into(),
                                });
                            }
                        }
                    }
                },
            }
        }
    }

    let style = child(root, "pivotTableStyleInfo").and_then(|info| {
        let name = info.attribute("name").filter(|name| !name.is_empty())?;
        let styles =
            workbook_styles.get_or_insert_with(|| WorkbookTableStyles::load(archive, theme_colors));
        // A style that cannot be applied as authored is left off the whole
        // PivotTable and reported, never applied with elements missing.
        let elements = match styles.elements(name, theme_colors) {
            Ok(elements) => elements?,
            Err(field) => {
                reasons.push(PivotPartialReason::MalformedField { field });
                return None;
            }
        };
        let flag = |key: &str| matches!(info.attribute(key), Some("1" | "true"));
        Some(PivotTableStyle {
            name: name.to_string(),
            show_row_headers: flag("showRowHeaders"),
            show_column_headers: flag("showColHeaders"),
            show_row_stripes: flag("showRowStripes"),
            show_column_stripes: flag("showColStripes"),
            show_last_column: flag("showLastColumn"),
            elements,
        })
    });
    let mut items = |parent: &str| {
        axis_items(root, parent).unwrap_or_else(|()| {
            reasons.push(PivotPartialReason::MalformedField {
                field: format!("{parent}.i"),
            });
            Vec::new()
        })
    };
    let row_items = items("rowItems");
    let column_items = items("colItems");

    let status = if reasons.is_empty() {
        PivotMetadataStatus::Complete
    } else {
        PivotMetadataStatus::Partial { reasons }
    };
    Ok(PivotTableMetadata {
        name,
        cache_id,
        location,
        row_fields,
        column_fields,
        page_fields,
        data_fields,
        refresh_on_load,
        cache_invalid,
        cache_definition_part,
        cache_source,
        status,
        extension_uris: recorded_extension_uris,
        style,
        row_items,
        column_items,
    })
}

enum CacheLink {
    Missing,
    Malformed,
    Unreadable,
    External,
    Ambiguous,
    Target(String),
}

fn pivot_cache_target(archive: &mut XlsxZip, pivot_part: &str) -> CacheLink {
    let Some((dir, file)) = pivot_part.rsplit_once('/') else {
        return CacheLink::Missing;
    };
    let rels_path = format!("{dir}/_rels/{file}.rels");
    let rels_exists = archive.index_for_name(&rels_path).is_some();
    let Ok(xml) = read_zip_string(archive, &rels_path) else {
        return if rels_exists {
            CacheLink::Unreadable
        } else {
            CacheLink::Missing
        };
    };
    let Ok(doc) = parse_guarded(&xml) else {
        return CacheLink::Malformed;
    };
    let root = doc.root_element();
    if root.tag_name().name() != "Relationships"
        || root.tag_name().namespace() != Some(PACKAGE_REL_NS)
    {
        return CacheLink::Malformed;
    }
    if root
        .children()
        .filter(|node| node.is_element())
        .any(|node| {
            matches!(
                node.attribute("Type"),
                Some(PIVOT_CACHE_REL_TRANSITIONAL | PIVOT_CACHE_REL_STRICT)
            ) && (node.tag_name().name() != "Relationship"
                || node.tag_name().namespace() != Some(PACKAGE_REL_NS))
        })
    {
        return CacheLink::Malformed;
    }
    let relationships: Vec<_> = doc
        .root_element()
        .children()
        .filter(|n| n.is_element())
        .filter(|n| n.tag_name().name() == "Relationship")
        .filter(|n| n.tag_name().namespace() == Some(PACKAGE_REL_NS))
        .filter(|n| {
            matches!(
                n.attribute("Type"),
                Some(PIVOT_CACHE_REL_TRANSITIONAL | PIVOT_CACHE_REL_STRICT)
            )
        })
        .collect();
    match relationships.as_slice() {
        [relationship] => {
            if relationship.attribute("Id").is_none_or(|id| id.is_empty()) {
                return CacheLink::Malformed;
            }
            if relationship.attribute("TargetMode") == Some("External") {
                return CacheLink::External;
            }
            if relationship
                .attribute("TargetMode")
                .is_some_and(|mode| mode != "Internal")
            {
                return CacheLink::Malformed;
            }
            match relationship.attribute("Target") {
                Some(target) if !target.is_empty() => {
                    CacheLink::Target(resolve_zip_path(dir, target))
                }
                None => CacheLink::Malformed,
                Some(_) => CacheLink::Malformed,
            }
        }
        [] => CacheLink::Missing,
        _ => CacheLink::Ambiguous,
    }
}

fn parse_location(node: roxmltree::Node<'_, '_>) -> Option<PivotLocation> {
    Some(PivotLocation {
        range: parse_a1_range(node.attribute("ref")?)?,
        first_header_row: node.attribute("firstHeaderRow")?.parse().ok()?,
        first_data_row: node.attribute("firstDataRow")?.parse().ok()?,
        first_data_col: node.attribute("firstDataCol")?.parse().ok()?,
    })
}

fn parse_a1_range(value: &str) -> Option<CellRange> {
    let (first, last) = value.split_once(':').unwrap_or((value, value));
    let (left, top) = parse_a1_cell(first)?;
    let (right, bottom) = parse_a1_cell(last)?;
    (top <= bottom && left <= right).then_some(CellRange {
        top,
        left,
        bottom,
        right,
    })
}

fn parse_a1_cell(value: &str) -> Option<(u32, u32)> {
    let value = value.strip_prefix('$').unwrap_or(value);
    let split = value.find(|c: char| !c.is_ascii_alphabetic())?;
    let (letters, row_part) = value.split_at(split);
    let digits = row_part.strip_prefix('$').unwrap_or(row_part);
    if letters.is_empty()
        || digits.is_empty()
        || !letters.chars().all(|c| c.is_ascii_alphabetic())
        || !digits.chars().all(|c| c.is_ascii_digit())
    {
        return None;
    }
    let col = letters.chars().try_fold(0u32, |acc, c| {
        acc.checked_mul(26)?
            .checked_add((c.to_ascii_uppercase() as u8 - b'A' + 1) as u32)
    })?;
    let row = digits.parse().ok()?;
    // SpreadsheetML's A1 grid is bounded by XFD / row 1,048,576.
    (col > 0 && col <= 16_384 && row > 0 && row <= 1_048_576).then_some((col, row))
}

fn parse_signed_fields(root: roxmltree::Node<'_, '_>, parent_name: &str) -> (Vec<i32>, bool) {
    let mut malformed = false;
    let fields = child(root, parent_name)
        .map(|parent| {
            parent
                .children()
                .filter(|n| {
                    n.is_element()
                        && n.tag_name().name() == "field"
                        && is_x_ns(n.tag_name().namespace())
                })
                .filter_map(|n| match n.attribute("x").and_then(|v| v.parse().ok()) {
                    Some(value) => Some(value),
                    None => {
                        malformed = true;
                        None
                    }
                })
                .collect()
        })
        .unwrap_or_default();
    (fields, malformed)
}

fn unsupported_features(root: roxmltree::Node<'_, '_>) -> Vec<PivotPartialReason> {
    [
        "pivotFields",
        "rowItems",
        "colItems",
        "pivotHierarchies",
        "filters",
        "rowHierarchiesUsage",
        "colHierarchiesUsage",
    ]
    .into_iter()
    .filter(|feature| child(root, feature).is_some())
    .map(|feature| PivotPartialReason::UnsupportedSemanticFeature {
        feature: feature.to_string(),
    })
    .collect()
}

fn extension_uris(root: roxmltree::Node<'_, '_>) -> Vec<String> {
    root.descendants()
        .filter(|n| {
            n.is_element() && n.tag_name().name() == "ext" && is_x_ns(n.tag_name().namespace())
        })
        .filter_map(|n| n.attribute("uri"))
        .map(str::to_string)
        .collect()
}

fn is_data_subtotal(value: &str) -> bool {
    matches!(
        value,
        "average"
            | "count"
            | "countNums"
            | "max"
            | "min"
            | "product"
            | "stdDev"
            | "stdDevp"
            | "sum"
            | "var"
            | "varp"
    )
}

fn child<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    node.children().find(|child| {
        child.is_element()
            && child.tag_name().name() == name
            && is_x_ns(child.tag_name().namespace())
    })
}

fn parse_bool(value: Option<&str>) -> Result<bool, ()> {
    match value {
        None | Some("0" | "false") => Ok(false),
        Some("1" | "true") => Ok(true),
        Some(_) => Err(()),
    }
}

/// ECMA-376 §18.10.1.44 `i` items of `rowItems`/`colItems`: the item type
/// (`t`, default `data`) and its field level, `r` (the count of leading
/// fields repeated from the previous item) plus its `x` count less one.
/// `r` is an xsd:unsignedInt (default 0); a value that does not parse, or a
/// level that overflows u32, rejects the whole list (`Err`) rather than
/// yielding a wrapped or clamped level.
fn axis_items(root: roxmltree::Node<'_, '_>, parent: &str) -> Result<Vec<PivotAxisItem>, ()> {
    let Some(items) = child(root, parent) else {
        return Ok(Vec::new());
    };
    items
        .children()
        .filter(|n| {
            n.is_element() && n.tag_name().name() == "i" && is_x_ns(n.tag_name().namespace())
        })
        .map(|item| {
            let repeated: u32 = match item.attribute("r") {
                None => 0,
                Some(value) => value.parse().map_err(|_| ())?,
            };
            let indices = u32::try_from(
                item.children()
                    .filter(|n| n.is_element() && n.tag_name().name() == "x")
                    .count(),
            )
            .map_err(|_| ())?;
            let fields = repeated.checked_add(indices).ok_or(())?;
            Ok(PivotAxisItem {
                kind: item.attribute("t").unwrap_or("data").to_string(),
                depth: fields.saturating_sub(1),
            })
        })
        .collect()
}

/// One `tableStyleElement` (§18.8.41) as authored: `size` defaults to 1 and
/// `dxfId` is optional (an element without one formats nothing, but its
/// `size` still sets a stripe band).
type AuthoredStyleElement = (String, u32, Option<u32>);

/// The workbook's `<tableStyles>` (§18.8.42) with their `<dxfs>`, read once
/// per sheet that has a styled PivotTable. Each `<dxf>` is parsed in place in
/// the styles part, so the namespace declarations in scope there (a default
/// namespace, a prefix bound on an ancestor, Transitional or Strict) apply
/// exactly as they do to the rest of the part.
pub(crate) struct WorkbookTableStyles {
    dxfs: Vec<Dxf>,
    styles: std::collections::HashMap<String, Result<Vec<AuthoredStyleElement>, String>>,
    /// Set when `xl/styles.xml` exists but is not well-formed XML.
    malformed_part: bool,
}

impl WorkbookTableStyles {
    fn load(archive: &mut XlsxZip, theme_colors: &[String]) -> Self {
        match read_zip_string(archive, "xl/styles.xml") {
            Ok(xml) => Self::from_styles_xml(&xml, theme_colors),
            // A workbook without a styles part defines no table styles.
            Err(_) => Self {
                dxfs: Vec::new(),
                styles: std::collections::HashMap::new(),
                malformed_part: false,
            },
        }
    }

    fn from_styles_xml(xml: &str, theme_colors: &[String]) -> Self {
        let mut result = Self {
            dxfs: Vec::new(),
            styles: std::collections::HashMap::new(),
            malformed_part: false,
        };
        let Ok(doc) = parse_guarded(xml) else {
            result.malformed_part = true;
            return result;
        };
        let mut dxfs_seen = false;
        for node in doc
            .descendants()
            .filter(|n| n.is_element() && is_x_ns(n.tag_name().namespace()))
        {
            match node.tag_name().name() {
                // The first `dxfs`, as the cell-style parser reads it.
                "dxfs" if !dxfs_seen => {
                    dxfs_seen = true;
                    result.dxfs = node
                        .children()
                        .filter(|n| n.is_element() && n.tag_name().name() == "dxf")
                        .map(|n| {
                            let mut dxf = crate::styles::parse_dxf(n, theme_colors);
                            explicit_none_edges(n, &mut dxf);
                            dxf
                        })
                        .collect();
                }
                "tableStyle" => {
                    let Some(name) = node.attribute("name") else {
                        continue;
                    };
                    let elements = node
                        .children()
                        .filter(|n| n.is_element() && n.tag_name().name() == "tableStyleElement")
                        .map(|n| {
                            let kind = n
                                .attribute("type")
                                .ok_or_else(|| "tableStyleElement.type".to_string())?;
                            let size = match n.attribute("size") {
                                None => 1,
                                Some(value) => value
                                    .parse()
                                    .map_err(|_| "tableStyleElement.size".to_string())?,
                            };
                            let dxf_id = match n.attribute("dxfId") {
                                None => None,
                                Some(value) => Some(
                                    value
                                        .parse()
                                        .map_err(|_| "tableStyleElement.dxfId".to_string())?,
                                ),
                            };
                            Ok((kind.to_string(), size, dxf_id))
                        })
                        .collect();
                    result.styles.insert(name.to_string(), elements);
                }
                _ => {}
            }
        }
        result
    }

    /// The elements of `name`: a workbook style (§18.8.40), each with its
    /// parsed `<dxf>`, else a built-in Annex G PivotTable style from the
    /// shared preset table, else `Ok(None)` (the name matches no style). A
    /// workbook style that cannot be applied as authored (a malformed
    /// element, a `dxfId` outside `<dxfs>`, an unreadable styles part) is an
    /// `Err` naming the offending field; no element is dropped from a style
    /// that is still applied.
    fn elements(
        &self,
        name: &str,
        theme_colors: &[String],
    ) -> Result<Option<Vec<PivotTableStyleElement>>, String> {
        if let Some(elements) = self.styles.get(name) {
            let elements = elements.as_ref().map_err(Clone::clone)?;
            return elements
                .iter()
                .map(|(kind, size, dxf_id)| {
                    let dxf = match dxf_id {
                        None => Dxf::default(),
                        Some(index) => self
                            .dxfs
                            .get(*index as usize)
                            .cloned()
                            .ok_or_else(|| "tableStyleElement.dxfId".to_string())?,
                    };
                    Ok(PivotTableStyleElement {
                        kind: kind.clone(),
                        size: *size,
                        dxf,
                    })
                })
                .collect::<Result<Vec<_>, String>>()
                .map(Some);
        }
        if let Some(elements) = xlsx_model::style_presets::pivot_style_elements(name, theme_colors)
        {
            return Ok(Some(elements));
        }
        if self.malformed_part {
            return Err("styles.xml".to_string());
        }
        Ok(None)
    }
}

/// A style element's border edge written without a style (`<left/>`, or
/// `style="none"`) is an explicitly cleared edge: CT_BorderPr's style
/// defaults to `none` (§18.8.6), and in a differential format a present
/// edge element overrides the edge of an earlier style element (§18.8.41
/// layering). The shared dxf parser drops such edges, so they are restored
/// here as `none` edges for the PivotTable renderer.
fn explicit_none_edges(dxf_node: roxmltree::Node<'_, '_>, dxf: &mut Dxf) {
    let Some(border) = dxf_node.children().find(|n| {
        n.is_element() && n.tag_name().name() == "border" && is_x_ns(n.tag_name().namespace())
    }) else {
        return;
    };
    let target = dxf.border.get_or_insert_with(Default::default);
    for edge in border.children().filter(|n| n.is_element()) {
        if edge.attribute("style").is_some_and(|style| style != "none") {
            continue;
        }
        let slot = match edge.tag_name().name() {
            "left" => &mut target.left,
            "right" => &mut target.right,
            "top" => &mut target.top,
            "bottom" => &mut target.bottom,
            "horizontal" => &mut target.horizontal,
            "vertical" => &mut target.vertical,
            _ => continue,
        };
        if slot.is_none() {
            *slot = Some(crate::BorderEdge {
                style: "none".into(),
                color: None,
            });
        }
    }
}

#[cfg(test)]
mod style_tests {
    use super::*;

    const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const SML_STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

    fn axis(xml_items: &str) -> Result<Vec<(String, u32)>, ()> {
        let xml = format!(
            r#"<pivotTableDefinition xmlns="{SML}"><rowItems>{xml_items}</rowItems></pivotTableDefinition>"#
        );
        let doc = parse_guarded(&xml).unwrap();
        axis_items(doc.root_element(), "rowItems").map(|items| {
            items
                .into_iter()
                .map(|item| (item.kind, item.depth))
                .collect()
        })
    }

    #[test]
    fn axis_items_carry_type_and_field_level() {
        assert_eq!(
            axis(r#"<i><x/></i><i r="1"><x v="2"/></i><i t="grand"><x/></i>"#).unwrap(),
            vec![("data".into(), 0), ("data".into(), 1), ("grand".into(), 0)]
        );
    }

    /// r + x-count is computed with checked arithmetic, so the boundary holds
    /// identically in debug (overflow panics) and release (it would wrap).
    #[test]
    fn axis_item_level_rejects_u32_overflow_and_unparsable_r() {
        assert_eq!(
            axis(r#"<i r="4294967294"><x/></i>"#).unwrap(),
            vec![("data".into(), 4294967294)]
        );
        assert_eq!(axis(r#"<i r="4294967295"><x/></i>"#), Err(()));
        assert_eq!(
            axis(r#"<i r="4294967295"/>"#).unwrap(),
            vec![("data".into(), 4294967294)]
        );
        assert_eq!(axis(r#"<i r="-1"><x/></i>"#), Err(()));
        assert_eq!(axis(r#"<i r="x"><x/></i>"#), Err(()));
    }

    fn styles(xml: &str) -> WorkbookTableStyles {
        WorkbookTableStyles::from_styles_xml(xml, &[])
    }

    fn one_style(dxfs: &str, elements: &str) -> String {
        format!(
            r#"<styleSheet xmlns="{SML}"><dxfs>{dxfs}</dxfs><tableStyles><tableStyle name="S" pivot="1" table="0">{elements}</tableStyle></tableStyles></styleSheet>"#
        )
    }

    #[test]
    fn style_element_edges_without_a_style_are_cleared_edges() {
        let styles = styles(&one_style(
            r#"<dxf><border><left/><right style="none"/><top style="thin"/><bottom/></border></dxf>"#,
            r#"<tableStyleElement type="firstRowSubheading" dxfId="0"/>"#,
        ));
        let elements = styles.elements("S", &[]).unwrap().unwrap();
        let border = elements[0].dxf.border.as_ref().unwrap();
        assert_eq!(border.left.as_ref().unwrap().style, "none");
        assert_eq!(border.right.as_ref().unwrap().style, "none");
        assert_eq!(border.top.as_ref().unwrap().style, "thin");
        assert_eq!(border.bottom.as_ref().unwrap().style, "none");
        assert!(border.vertical.is_none());
    }

    /// The same style written with a default namespace, with a prefix bound
    /// only on an ancestor of `<dxfs>`, and in the Strict namespace parses to
    /// the same formats: each `<dxf>` is read in place, in its own context.
    #[test]
    fn style_dxfs_parse_identically_under_default_prefixed_and_strict_namespaces() {
        let body = |p: &str| {
            format!(
                r#"<{p}dxfs><{p}dxf><{p}font><{p}b/><{p}i val="0"/><{p}color rgb="FFFF0000"/></{p}font><{p}fill><{p}patternFill><{p}bgColor rgb="FF00FF00"/></{p}patternFill></{p}fill><{p}border><{p}left/><{p}top style="thin"/></{p}border></{p}dxf></{p}dxfs><{p}tableStyles><{p}tableStyle name="S"><{p}tableStyleElement type="wholeTable" dxfId="0"/><{p}tableStyleElement type="firstRowStripe" size="2"/></{p}tableStyle></{p}tableStyles>"#
            )
        };
        let variants = [
            format!(r#"<styleSheet xmlns="{SML}">{}</styleSheet>"#, body("")),
            format!(
                r#"<x:styleSheet xmlns:x="{SML}">{}</x:styleSheet>"#,
                body("x:")
            ),
            format!(
                r#"<styleSheet xmlns="{SML_STRICT}">{}</styleSheet>"#,
                body("")
            ),
            format!(
                r#"<s:styleSheet xmlns:s="{SML_STRICT}">{}</s:styleSheet>"#,
                body("s:")
            ),
        ];
        let rendered: Vec<String> = variants
            .iter()
            .map(|xml| {
                let elements = styles(xml).elements("S", &[]).unwrap().unwrap();
                serde_json::to_string(&elements).unwrap()
            })
            .collect();
        let default = styles(&variants[0]).elements("S", &[]).unwrap().unwrap();
        assert_eq!(default.len(), 2);
        assert_eq!(default[0].dxf.font_toggles.unwrap().bold, Some(true));
        assert_eq!(default[0].dxf.font_toggles.unwrap().italic, Some(false));
        // An element the dxf font omits is no change, not an explicit off.
        assert_eq!(default[0].dxf.font_toggles.unwrap().underline, None);
        assert_eq!(
            default[0]
                .dxf
                .border
                .as_ref()
                .unwrap()
                .left
                .as_ref()
                .unwrap()
                .style,
            "none"
        );
        assert!(default[0].dxf.fill.as_ref().unwrap().fg_color.is_some());
        // An element without dxfId formats nothing but keeps its band size.
        assert_eq!((default[1].size, default[1].dxf.font.is_none()), (2, true));
        for other in &rendered[1..] {
            assert_eq!(other, &rendered[0]);
        }
    }

    #[test]
    fn a_style_that_cannot_be_applied_as_authored_is_an_error_not_a_shorter_style() {
        let out_of_range = styles(&one_style(
            "<dxf/>",
            r#"<tableStyleElement type="wholeTable" dxfId="0"/><tableStyleElement type="headerRow" dxfId="1"/>"#,
        ));
        assert_eq!(
            out_of_range.elements("S", &[]).map(|_| ()),
            Err("tableStyleElement.dxfId".to_string())
        );
        let bad_size = styles(&one_style(
            "<dxf/>",
            r#"<tableStyleElement type="firstRowStripe" size="two" dxfId="0"/>"#,
        ));
        assert_eq!(
            bad_size.elements("S", &[]).map(|_| ()),
            Err("tableStyleElement.size".to_string())
        );
        let malformed = styles("<styleSheet");
        assert_eq!(
            malformed.elements("Custom", &[]).map(|_| ()),
            Err("styles.xml".to_string())
        );
        // A built-in name still resolves, and an unknown name is no style.
        assert!(malformed
            .elements("PivotStyleLight16", &[])
            .unwrap()
            .is_some());
        assert!(styles(&one_style("", ""))
            .elements("NoSuchStyle", &[])
            .unwrap()
            .is_none());
    }

    #[test]
    fn built_in_pivot_styles_resolve_from_annex_g() {
        let styles = styles(&format!(r#"<styleSheet xmlns="{SML}"/>"#));
        let theme: Vec<String> = [
            "#FFFFFF", "#000000", "#E7E6E6", "#44546A", "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
            "#5B9BD5", "#70AD47", "#0563C1", "#954F72",
        ]
        .iter()
        .map(|c| c.to_string())
        .collect();
        let elements = styles
            .elements("PivotStyleLight16", &theme)
            .unwrap()
            .unwrap();
        let kinds: Vec<_> = elements.iter().map(|e| e.kind.as_str()).collect();
        assert!(kinds.contains(&"headerRow") && kinds.contains(&"totalRow"));
        let total = elements.iter().find(|e| e.kind == "totalRow").unwrap();
        assert!(total.dxf.font.as_ref().unwrap().bold);
        assert_eq!(total.dxf.font_toggles.unwrap().bold, Some(true));
        assert!(styles.elements("NoSuchStyle", &theme).unwrap().is_none());
    }
}
