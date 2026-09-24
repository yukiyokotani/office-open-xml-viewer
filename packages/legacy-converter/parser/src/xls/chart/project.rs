//! RawChart -> shared ChartModel. Colors come from verified ShapePropsStream
//! XML when present (MS-XLS 2.4.258), otherwise from the BIFF records: a
//! non-automatic AreaFormat/LineFormat uses its palette index, because Excel
//! ignores their rgb fields on load (MS-XLS 2.4.3 footnotes <23>/<24>).
//! Automatic formatting is left unset so the shared renderer applies its
//! ordinary automatic series colors.
use super::reader::{Cached, Format, GroupKind, RawChart};
use ooxml_common::chart::{ChartModel, ChartSeries};
use ooxml_common::color::{parse_color_node, ThemeResolver, TintMode};

pub(crate) struct Palette<'a> {
    /// Resolved palette color for an Icv (MS-XLS 2.5.161), "#RRGGBB" or "RRGGBB".
    pub icv: &'a dyn Fn(u16) -> Option<String>,
    /// Theme colors in clrScheme order: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink.
    pub theme: [Option<String>; 12],
}

impl ThemeResolver for Palette<'_> {
    fn resolve_scheme_color(&self, name: &str) -> Option<String> {
        let slot = match ooxml_common::color::default_scheme_slot(name) {
            "dk1" => 0,
            "lt1" => 1,
            "dk2" => 2,
            "lt2" => 3,
            "accent1" => 4,
            "accent2" => 5,
            "accent3" => 6,
            "accent4" => 7,
            "accent5" => 8,
            "accent6" => 9,
            "hlink" => 10,
            "folHlink" => 11,
            _ => return None,
        };
        self.theme[slot]
            .as_deref()
            .map(|hex| hex.trim_start_matches('#').to_owned())
    }
}

impl Palette<'_> {
    /// Theme accentN for an automatic series or varied point (N = i mod 6 + 1).
    fn accent(&self, index: usize) -> Option<String> {
        self.theme[4 + index % 6]
            .as_deref()
            .map(|hex| hex.trim_start_matches('#').to_uppercase())
    }
}

struct Paint {
    fill: Option<String>,
    fill_hidden: bool,
    line: Option<String>,
    line_hidden: bool,
    line_width_emu: Option<u32>,
}

/// The shared chart model carries RRGGBB without a leading '#'.
fn hex(value: String) -> String {
    value.trim_start_matches('#').to_uppercase()
}

fn child<'a, 'input>(
    parent: roxmltree::Node<'a, 'input>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'input>> {
    parent
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == name)
}

/// Shape XML (DrawingML spPr) fill/line, resolved with the workbook theme.
/// A verified but empty stream is an empty spPr: automatic formatting, which
/// supersedes the BIFF records (ShapePropsStream is their superset, 2.4.258).
/// Excel then draws the chart-style automatic colors, as for XLSX.
fn xml_paint(xml: &str, palette: &Palette<'_>) -> Option<Paint> {
    if xml.trim().is_empty() {
        return Some(Paint {
            fill: None,
            fill_hidden: false,
            line: None,
            line_hidden: false,
            line_width_emu: None,
        });
    }
    let document = roxmltree::Document::parse(xml).ok()?;
    let root = document.root_element();
    if root.tag_name().name() != "spPr" {
        return None;
    }
    let fill = child(root, "solidFill")
        .and_then(|node| parse_color_node(node, palette, TintMode::PowerPointLinear))
        .map(hex);
    let fill_hidden = child(root, "noFill").is_some();
    let line_node = child(root, "ln");
    let line = line_node
        .and_then(|ln| child(ln, "solidFill"))
        .and_then(|node| parse_color_node(node, palette, TintMode::PowerPointLinear))
        .map(hex);
    let line_hidden = line_node.and_then(|ln| child(ln, "noFill")).is_some();
    let line_width_emu = line_node
        .and_then(|ln| ln.attribute("w"))
        .and_then(|w| w.parse().ok());
    Some(Paint {
        fill,
        fill_hidden,
        line,
        line_hidden,
        line_width_emu,
    })
}

fn biff_paint(format: &Format, palette: &Palette<'_>) -> Paint {
    let u16_at = |bytes: &[u8], at: usize| u16::from_le_bytes([bytes[at], bytes[at + 1]]);
    let (mut fill, mut fill_hidden) = (None, false);
    if let Some(area) = format.area {
        let automatic = u16_at(&area, 10) & 1 != 0;
        let pattern = u16_at(&area, 8);
        if !automatic {
            if pattern == 0 {
                fill_hidden = true;
            } else {
                fill = (palette.icv)(u16_at(&area, 12)).map(hex);
            }
        }
    }
    let (mut line, mut line_hidden) = (None, false);
    let line_width_emu = None;
    if let Some(format) = format.line {
        let automatic = u16_at(&format, 8) & 1 != 0;
        let pattern = u16_at(&format, 4);
        if !automatic {
            if pattern == 5 {
                line_hidden = true;
            } else {
                line = (palette.icv)(u16_at(&format, 10)).map(hex);
                // LineFormat.we (2.4.156) names hairline/narrow/medium/wide
                // weights without a normative length; leave the width to the
                // renderer rather than inventing point values.
            }
        }
    }
    Paint {
        fill,
        fill_hidden,
        line,
        line_hidden,
        line_width_emu,
    }
}

fn paint(format: &Format, palette: &Palette<'_>) -> Paint {
    match format
        .shape_xml
        .get(&0)
        .and_then(|xml| xml_paint(xml, palette))
    {
        Some(xml) => xml,
        None => biff_paint(format, palette),
    }
}

fn number_text(value: f64) -> String {
    if value.fract() == 0.0 && value.abs() < 1e15 {
        format!("{}", value as i64)
    } else {
        format!("{value}")
    }
}

fn series_type(kind: GroupKind) -> &'static str {
    match kind {
        GroupKind::Bar { .. } => "bar",
        GroupKind::Line { .. } => "line",
        GroupKind::Area { .. } => "area",
        GroupKind::Pie { hole, .. } if hole > 0 => "doughnut",
        GroupKind::Pie { .. } | GroupKind::OfPie => "pie",
        GroupKind::Scatter { .. } => "scatter",
        GroupKind::Radar { .. } => "radar",
        GroupKind::Surface => "surface",
    }
}

fn chart_type(kind: GroupKind) -> String {
    match kind {
        GroupKind::Bar {
            horizontal,
            stacked,
            percent,
            ..
        } => ooxml_common::chart::canonical_chart_type(
            "bar",
            if horizontal { "bar" } else { "col" },
            grouping(stacked, percent),
        ),
        GroupKind::Line { stacked, percent } => {
            ooxml_common::chart::canonical_chart_type("line", "", grouping(stacked, percent))
        }
        GroupKind::Area { stacked, percent } => {
            ooxml_common::chart::canonical_chart_type("area", "", grouping(stacked, percent))
        }
        GroupKind::Pie { hole, .. } if hole > 0 => "doughnut".into(),
        GroupKind::Pie { .. } => "pie".into(),
        GroupKind::OfPie => "ofPie".into(),
        GroupKind::Scatter { bubbles: true } => "bubble".into(),
        GroupKind::Scatter { bubbles: false } => "scatter".into(),
        GroupKind::Radar { .. } => "radar".into(),
        GroupKind::Surface => "surface".into(),
    }
}

fn grouping(stacked: bool, percent: bool) -> &'static str {
    match (stacked, percent) {
        (true, true) => "percentStacked",
        (true, false) => "stacked",
        _ => "clustered",
    }
}

/// MarkerFormat.imk (MS-XLS 2.4.160) -> ST_MarkerStyle. The short and long
/// bar markers are the "dot" and "dash" styles of the same Office marker set.
fn marker_symbol(imk: u16) -> Option<&'static str> {
    Some(match imk {
        0 => "none",
        1 => "square",
        2 => "diamond",
        3 => "triangle",
        4 => "x",
        5 => "star",
        6 => "dot",
        7 => "dash",
        8 => "circle",
        9 => "plus",
        _ => return None,
    })
}

struct Marker {
    symbol: Option<&'static str>,
    size_pt: Option<f64>,
    fill: Option<String>,
    line: Option<String>,
}

fn marker(format: &Format, palette: &Palette<'_>) -> Option<Marker> {
    let data = format.marker?;
    let u16_at = |at: usize| u16::from_le_bytes([data[at], data[at + 1]]);
    let flags = u16_at(10);
    if flags & 1 != 0 {
        // fAuto: automatic marker formatting.
        return None;
    }
    let xml = format
        .shape_xml
        .get(&1)
        .and_then(|xml| xml_paint(xml, palette));
    let (fill, line) = match xml {
        Some(paint) => (paint.fill, paint.line),
        None => (
            (flags & 0x10 == 0)
                .then(|| (palette.icv)(u16_at(14)).map(hex))
                .flatten(),
            (flags & 0x20 == 0)
                .then(|| (palette.icv)(u16_at(12)).map(hex))
                .flatten(),
        ),
    };
    let size = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);
    Some(Marker {
        symbol: marker_symbol(u16_at(8)),
        // miSize is in twips (1/20 point).
        size_pt: (size > 0).then(|| f64::from(size) / 20.0),
        fill,
        line,
    })
}

/// AttachedLabel (MS-XLS 2.4.5) flags -> series data labels.
fn data_labels(flags: u16) -> Option<ooxml_common::chart::ChartSeriesDataLabels> {
    let labels = ooxml_common::chart::ChartSeriesDataLabels {
        show_val: flags & 0x01 != 0,
        show_percent: flags & 0x02 != 0 || flags & 0x04 != 0,
        show_cat_name: flags & 0x10 != 0 || flags & 0x04 != 0,
        show_bubble_size: flags & 0x20 != 0,
        show_ser_name: flags & 0x40 != 0,
        ..Default::default()
    };
    (labels.show_val
        || labels.show_percent
        || labels.show_cat_name
        || labels.show_bubble_size
        || labels.show_ser_name)
        .then_some(labels)
}

/// Resolves a BRAI worksheet reference (rgce) to its cells in order.
pub(crate) type References<'a> = dyn Fn(&[u8]) -> Option<Vec<Option<Cached>>> + 'a;

/// Project a raw chart. Returns `None` when there is no drawable series.
/// The chart data cache is authoritative when present (MS-XLS 2.2.3.2);
/// otherwise series parts are read from the referenced worksheet cells.
pub(crate) fn project(
    raw: &RawChart,
    palette: &Palette<'_>,
    references: &References<'_>,
) -> Option<ChartModel> {
    let primary = raw.groups.first()?;
    let values = raw.cache.get(&1);
    let categories = raw.cache.get(&2);
    let bubbles = raw.cache.get(&3);
    let point_count = |cache: Option<&std::collections::BTreeMap<(u16, u16), Cached>>,
                       series: u16| {
        cache.map_or(0, |cache| {
            cache
                .range((series, 0)..=(series, u16::MAX))
                .map(|((_, point), _)| usize::from(*point) + 1)
                .max()
                .unwrap_or(0)
        })
    };
    let mut model = ChartModel {
        chart_type: chart_type(primary.kind),
        title: raw.title.clone(),
        title_present: raw.title.is_some(),
        cat_axis_cross_between: "between".into(),
        val_axis_major_tick_mark: "out".into(),
        cat_axis_major_tick_mark: "out".into(),
        plot_visible_only: Some(raw.plot_visible_only),
        ..ChartModel::default()
    };
    let mut series_models = Vec::new();
    for (index, series) in raw.series.iter().enumerate() {
        if series.trend_or_error {
            continue;
        }
        let key = index as u16;
        let group = raw.groups.get(usize::from(series.group)).unwrap_or(primary);
        let part =
            |numindex: u16, cache: Option<&std::collections::BTreeMap<(u16, u16), Cached>>| {
                let count = point_count(cache, key);
                if count > 0 {
                    return (0..count)
                        .map(|point| cache.and_then(|c| c.get(&(key, point as u16))).cloned())
                        .collect::<Vec<_>>();
                }
                series.references[usize::from(numindex)]
                    .as_deref()
                    .and_then(|rgce| references(rgce))
                    .unwrap_or_default()
            };
        let series_values = part(1, values)
            .into_iter()
            .map(|value| match value {
                Some(Cached::Number(number)) => Some(number),
                _ => None,
            })
            .collect::<Vec<_>>();
        let count = series_values.len();
        let series_categories = part(2, categories)
            .into_iter()
            .map(|value| match value {
                Some(Cached::Text(text)) => text,
                Some(Cached::Number(number)) => number_text(number),
                None => String::new(),
            })
            .collect::<Vec<_>>();
        let series_categories = if series_categories.is_empty() {
            (1..=count).map(|n| n.to_string()).collect()
        } else {
            series_categories
        };
        let series_paint = series
            .series_format
            .as_ref()
            .or(group.default_format.as_ref())
            .map(|format| paint(format, palette));
        let mut model_series = ChartSeries {
            name: series.name.clone().unwrap_or_default(),
            values: series_values,
            series_type: Some(series_type(group.kind).into()),
            use_secondary_axis: (group.axis_group == 1).then_some(true),
            categories: Some(series_categories),
            ..ChartSeries::default()
        };
        // Line-drawn groups take the series color from its line (the fill
        // only paints markers or areas); filled groups use the fill.
        let line_drawn = matches!(
            group.kind,
            GroupKind::Line { .. }
                | GroupKind::Scatter { bubbles: false }
                | GroupKind::Radar { filled: false }
        );
        if let Some(paint) = series_paint {
            model_series.color = if line_drawn {
                paint.line.clone().or(paint.fill)
            } else {
                paint.fill
            };
            model_series.line_color = paint.line;
            model_series.line_width_emu = paint.line_width_emu;
            if paint.line_hidden {
                model_series.line_hidden = Some(true);
            }
        }
        // Automatic series formatting uses the theme accents in series order,
        // matching the XLSX parser's automatic series color.
        if model_series.color.is_none() && !group.varied_colors {
            model_series.color = palette.accent(index);
        }
        if let Some(marker) = series
            .series_format
            .as_ref()
            .and_then(|f| marker(f, palette))
        {
            model_series.show_marker = Some(marker.symbol != Some("none"));
            model_series.marker_symbol = marker.symbol.map(str::to_owned);
            model_series.marker_size = marker.size_pt;
            model_series.marker_fill = marker.fill;
            model_series.marker_line = marker.line;
        }
        if let Some(labels) = series
            .series_format
            .as_ref()
            .and_then(|f| f.data_labels)
            .and_then(data_labels)
        {
            model_series.series_data_labels = Some(labels);
        }
        if let Some(format) = series.series_format.as_ref() {
            model_series.explosion = format.explosion.map(u32::from);
            if format.smooth {
                model_series.smooth = Some(true);
            }
        }
        if !series.point_formats.is_empty() || group.varied_colors {
            let points = count.max(
                series
                    .point_formats
                    .keys()
                    .map(|p| usize::from(*p) + 1)
                    .max()
                    .unwrap_or(0),
            );
            let colors = (0..points)
                .map(|point| {
                    series
                        .point_formats
                        .get(&(point as u16))
                        .and_then(|format| paint(format, palette).fill)
                        .or_else(|| group.varied_colors.then(|| palette.accent(point)).flatten())
                })
                .collect::<Vec<_>>();
            if colors.iter().any(Option::is_some) {
                model_series.data_point_colors = Some(colors);
            }
        }
        let sizes = part(3, bubbles)
            .into_iter()
            .map(|value| match value {
                Some(Cached::Number(number)) => Some(number),
                _ => None,
            })
            .collect::<Vec<_>>();
        if sizes.iter().any(Option::is_some) {
            model_series.bubble_sizes = Some(sizes);
        }
        series_models.push(model_series);
    }
    if series_models.is_empty() {
        return None;
    }
    model.categories = series_models[0].categories.clone().unwrap_or_default();
    model.series = series_models;
    if let Some(legend) = raw.groups.iter().find_map(|g| g.legend) {
        model.show_legend = true;
        model.legend_pos = Some(
            match legend.position {
                0 => "b",
                1 => "tr",
                2 => "t",
                4 => "l",
                _ => "r",
            }
            .into(),
        );
    }
    if let Some(axis) = raw.axes.iter().find(|a| a.kind == 1 && a.axis_group == 0) {
        model.val_min = axis.min;
        model.val_max = axis.max;
    }
    model.val_axis_title = raw.axis_titles.get(&2).cloned();
    model.cat_axis_title = raw.axis_titles.get(&3).cloned();
    if let Some(format) = raw.chart_format.as_ref() {
        let area = paint(format, palette);
        model.chart_bg = area.fill;
        if area.fill_hidden {
            model.chart_fill_hidden = Some(true);
        }
        model.chart_border_color = area.line;
        model.chart_border_width_emu = area.line_width_emu;
        if area.line_hidden {
            model.chart_border_hidden = Some(true);
        }
    }
    if let Some(format) = raw.plot_format.as_ref() {
        let area = paint(format, palette);
        model.plot_area_bg = area.fill;
        if area.fill_hidden {
            model.plot_area_fill_hidden = Some(true);
        }
        model.plot_area_line_color = area.line;
        model.plot_area_line_width_emu = area.line_width_emu.and_then(|w| u32::try_from(w).ok());
        if area.line_hidden {
            model.plot_area_line_hidden = Some(true);
        }
    }
    Some(model)
}
