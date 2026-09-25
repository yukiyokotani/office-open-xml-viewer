//! MS-ODRAW 2.4.24 shape types -> ECMA-376 preset geometry with adjust values.
//!
//! MS-ODRAW names the shape types but not their DrawingML equivalents or how
//! the legacy adjust values (adjustValue..adjust10Value, 2.3.6.10-19, in the
//! 21600-unit shape space) become `a:avLst` guides. The table and formulas
//! below are PowerPoint's own reading: every private-corpus .ppt had its
//! metroBlob (alternative DrawingML, 2.3.4.41) disabled and was saved as .pptx
//! by PowerPoint 16, so the saved `prstGeom`/`avLst` shows how PowerPoint
//! interprets each binary shape type and adjust value. Only shape types seen
//! in that conversion are mapped, and a formula is used only where the
//! conversions determine it:
//! - absent adjust values leave the preset's DrawingML defaults (PowerPoint
//!   writes the defaults for every such shape, including callouts whose
//!   legacy defaults differ);
//! - "W"/"H" distances are fractions of the shape width/height rescaled to
//!   the DrawingML short side, measured on the DrawingML extent (after the
//!   rotated-bounds swap; a 77-degree rotated brace only matches that way).
//! Any adjusted shape type without such evidence is rejected, never guessed.

/// DrawingML preset for a legacy shape type, as PowerPoint converts it.
pub(crate) fn name(kind: u16) -> Option<&'static str> {
    Some(match kind {
        // Not-primitive without vertices, picture frames and text boxes are
        // converted to rectangles.
        0 | 1 | 75 | 202 => "rect",
        2 => "roundRect",
        3 => "ellipse",
        4 => "diamond",
        5 => "triangle",
        6 => "rtTriangle",
        7 => "parallelogram",
        9 => "hexagon",
        10 => "octagon",
        11 => "plus",
        13 => "rightArrow",
        15 => "homePlate",
        16 => "cube",
        20 => "line",
        21 => "plaque",
        22 => "can",
        // MS-ODRAW 2.4.24: distinct from msosptLine. Connectors convert to
        // p:cxnSp presets.
        32 => "straightConnector1",
        34 => "bentConnector3",
        38 => "curvedConnector3",
        41 => "callout1",
        42 => "callout2",
        43 => "callout3",
        44 => "accentCallout1",
        45 => "accentCallout2",
        46 => "accentCallout3",
        47 => "borderCallout1",
        48 => "borderCallout2",
        49 => "borderCallout3",
        50 => "accentBorderCallout1",
        51 => "accentBorderCallout2",
        52 => "accentBorderCallout3",
        53 => "ribbon",
        54 => "ribbon2",
        55 => "chevron",
        56 => "pentagon",
        58 => "star8",
        59 => "star16",
        60 => "star32",
        61 => "wedgeRectCallout",
        62 => "wedgeRoundRectCallout",
        63 => "wedgeEllipseCallout",
        64 => "wave",
        65 => "foldedCorner",
        66 => "leftArrow",
        67 => "downArrow",
        68 => "upArrow",
        69 => "leftRightArrow",
        70 => "upDownArrow",
        71 => "irregularSeal1",
        72 => "irregularSeal2",
        73 => "lightningBolt",
        77 => "leftArrowCallout",
        78 => "rightArrowCallout",
        79 => "upArrowCallout",
        80 => "downArrowCallout",
        81 => "leftRightArrowCallout",
        84 => "bevel",
        85 => "leftBracket",
        86 => "rightBracket",
        87 => "leftBrace",
        88 => "rightBrace",
        92 => "star24",
        94 => "notchedRightArrow",
        96 => "smileyFace",
        97 => "verticalScroll",
        98 => "horizontalScroll",
        102 => "curvedRightArrow",
        103 => "curvedLeftArrow",
        104 => "curvedUpArrow",
        105 => "curvedDownArrow",
        106 => "cloudCallout",
        107 => "ellipseRibbon",
        108 => "ellipseRibbon2",
        109 => "flowChartProcess",
        110 => "flowChartDecision",
        111 => "flowChartInputOutput",
        112 => "flowChartPredefinedProcess",
        113 => "flowChartInternalStorage",
        114 => "flowChartDocument",
        115 => "flowChartMultidocument",
        116 => "flowChartTerminator",
        117 => "flowChartPreparation",
        118 => "flowChartManualInput",
        119 => "flowChartManualOperation",
        120 => "flowChartConnector",
        121 => "flowChartPunchedCard",
        122 => "flowChartPunchedTape",
        123 => "flowChartSummingJunction",
        125 => "flowChartCollate",
        126 => "flowChartSort",
        127 => "flowChartExtract",
        128 => "flowChartMerge",
        130 => "flowChartOnlineStorage",
        131 => "flowChartMagneticTape",
        132 => "flowChartMagneticDisk",
        133 => "flowChartMagneticDrum",
        134 => "flowChartDisplay",
        135 => "flowChartDelay",
        176 => "flowChartAlternateProcess",
        177 => "flowChartOffpageConnector",
        183 => "sun",
        184 => "moon",
        185 => "bracketPair",
        186 => "bracePair",
        187 => "star4",
        188 => "doubleWave",
        _ => return None,
    })
}

const SPACE: f64 = 21_600.0;

/// One DrawingML guide computed from one legacy adjust value (index 0 is
/// adjustValue, 0x147).
#[derive(Clone, Copy)]
enum Rule {
    /// a / 21600
    Scale(usize),
    /// (a - 10800) / 21600: a position measured from the shape centre.
    Centred(usize),
    /// (21600 - a) / 21600
    Complement(usize),
    /// a / 21600 of the width, as a fraction of the short side.
    Width(usize),
    /// (21600 - a) / 21600 of the width, as a fraction of the short side.
    WidthComplement(usize),
    /// a / 21600 of the height, as a fraction of the short side.
    Height(usize),
    /// (21600 - a) / 21600 of the height, as a fraction of the short side.
    HeightComplement(usize),
    /// (21600 - 2a) / 21600: a symmetric band between a and 21600 - a.
    Band(usize),
    /// (10800 - a) / 10800 * 0.5: a star's inner radius.
    Star(usize),
    /// Only the midpoint 10800 is evidenced (it reads as 50%); any other
    /// value is rejected.
    Midpoint(usize),
}

/// Guides per DrawingML slot (adj/adj1 .. adj8) for the evidenced shapes.
/// `None` in a slot keeps the preset default; the legacy values a shape may
/// carry are exactly those named by its rules.
fn rules(kind: u16) -> Option<&'static [Option<Rule>]> {
    use Rule::*;
    Some(match kind {
        2 | 5 => &[Some(Scale(0))],
        // Only square or wide shapes are evidenced (short side = height).
        7 | 9 => &[Some(Width(0))],
        16 => &[Some(Scale(0))],
        38 => &[Some(Midpoint(0))],
        13 | 94 => &[None, Some(WidthComplement(0))],
        15 | 55 => &[Some(WidthComplement(0))],
        58 => &[Some(Star(0))],
        61 | 62 | 63 | 106 => &[Some(Centred(0)), Some(Centred(1))],
        64 | 188 => &[Some(Scale(0))],
        65 => &[Some(Complement(0))],
        66 => &[None, Some(Width(0))],
        67 => &[None, Some(HeightComplement(0))],
        68 => &[None, Some(Height(0))],
        69 => &[Some(Band(1)), Some(Width(0))],
        70 => &[None, Some(Height(1))],
        85 | 86 => &[Some(Height(0))],
        87 | 88 => &[Some(Height(0)), Some(Scale(1))],
        // Callout1 family (callout1, accentCallout1, borderCallout1,
        // accentBorderCallout1): MS-ODRAW and ECMA-376 give the four members
        // the same adjust layout. callout1 is evidenced for all four values;
        // an accentCallout1 metroBlob pair (PowerPoint's own DrawingML beside
        // the binary) confirms the leader point for the family.
        41 | 44 | 47 | 50 => &[
            Some(Scale(3)),
            Some(Scale(2)),
            Some(Scale(1)),
            Some(Scale(0)),
        ],
        // Callout3 family: points in reverse order as for callout1. The x
        // coordinates round trip PowerPoint's defaults; an
        // accentBorderCallout3 metroBlob pair gives the y coordinates of all
        // three leader points (adjust2/4/6Value -> adj7/adj5/adj3). The
        // first point's y (adjust8Value -> adj1) is not evidenced and is
        // rejected.
        43 | 46 | 49 | 52 => &[
            None,
            Some(Scale(6)),
            Some(Scale(5)),
            Some(Scale(4)),
            Some(Scale(3)),
            Some(Scale(2)),
            Some(Scale(1)),
            Some(Scale(0)),
        ],
        53 => &[Some(Scale(1))],
        54 => &[Some(Complement(1))],
        // Evidenced only for tall shapes (short side = width).
        22 => &[Some(Height(0))],
        _ => return None,
    })
}

fn source(rule: Rule) -> usize {
    match rule {
        Rule::Scale(i)
        | Rule::Centred(i)
        | Rule::Complement(i)
        | Rule::Width(i)
        | Rule::WidthComplement(i)
        | Rule::Height(i)
        | Rule::HeightComplement(i)
        | Rule::Band(i)
        | Rule::Star(i)
        | Rule::Midpoint(i) => i,
    }
}

/// DrawingML guide values (100000 = 1) for `kind` from the legacy adjust
/// values and the DrawingML extent. `Ok(None)` means no adjust value is
/// authored, so the preset defaults apply.
pub(crate) fn adjustments(
    kind: u16,
    legacy: &[Option<i32>; 10],
    width: i64,
    height: i64,
) -> Result<Option<[Option<f64>; 8]>, String> {
    if legacy.iter().all(Option::is_none) {
        return Ok(None);
    }
    let unsupported = || {
        format!("UNSUPPORTED:PowerPoint shape type {kind} adjust values have no evidenced DrawingML mapping")
    };
    let rules = rules(kind).ok_or_else(unsupported)?;
    // Every authored value must be consumed by a rule.
    for (index, value) in legacy.iter().enumerate() {
        if value.is_some() && !rules.iter().flatten().any(|rule| source(*rule) == index) {
            return Err(unsupported());
        }
    }
    let (w, h) = (width as f64, height as f64);
    let short = w.min(h);
    if short <= 0.0 {
        return Err(unsupported());
    }
    // Parallelogram, hexagon and cube conversions were observed only with
    // the height as the short side.
    if matches!(kind, 7 | 9 | 16) && w < h || kind == 22 && h < w {
        return Err(unsupported());
    }
    let mut output = [None; 8];
    for (slot, rule) in rules.iter().enumerate() {
        let Some(rule) = rule else { continue };
        let Some(a) = legacy[source(*rule)] else {
            continue;
        };
        let a = f64::from(a);
        let fraction = match *rule {
            Rule::Scale(_) => a / SPACE,
            Rule::Centred(_) => (a - SPACE / 2.0) / SPACE,
            Rule::Complement(_) => (SPACE - a) / SPACE,
            Rule::Width(_) => a / SPACE * w / short,
            Rule::WidthComplement(_) => (SPACE - a) / SPACE * w / short,
            Rule::Height(_) => a / SPACE * h / short,
            Rule::HeightComplement(_) => (SPACE - a) / SPACE * h / short,
            Rule::Band(_) => (SPACE - 2.0 * a) / SPACE,
            Rule::Star(_) => (SPACE / 2.0 - a) / (SPACE / 2.0) * 0.5,
            Rule::Midpoint(_) if a == SPACE / 2.0 => 0.5,
            Rule::Midpoint(_) => return Err(unsupported()),
        };
        output[slot] = Some(fraction * 100_000.0);
    }
    Ok(Some(output))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn adj(kind: u16, values: &[(usize, i32)], w: i64, h: i64) -> [Option<f64>; 8] {
        let mut legacy = [None; 10];
        for &(i, v) in values {
            legacy[i] = Some(v);
        }
        adjustments(kind, &legacy, w, h).unwrap().unwrap()
    }

    fn close(value: Option<f64>, expected: f64) {
        let value = value.unwrap();
        assert!((value - expected).abs() < 1.5, "{value} vs {expected}");
    }

    #[test]
    fn reproduces_powerpoints_own_conversions() {
        // (kind, legacy values, anchor size) -> PowerPoint's saved avLst.
        close(adj(2, &[(0, 2867)], 2448, 437)[0], 13272.0);
        close(adj(7, &[(0, 926)], 3810, 654)[0], 24975.0);
        close(adj(9, &[(0, 5341)], 916, 798)[0], 28383.0);
        close(adj(13, &[(0, 14387)], 1024, 684)[1], 49993.0);
        close(adj(15, &[(0, 16551)], 818, 481)[0], 39752.0);
        close(adj(55, &[(0, 18093)], 1028, 334)[0], 49972.0);
        close(adj(58, &[(0, 2700)], 534, 476)[0], 37500.0);
        close(adj(16, &[(0, 5333)], 1276, 1221)[0], 24690.0);
        close(adj(38, &[(0, 10800)], 427, 242)[0], 50000.0);
        // PowerPoint's defaults round trip (bisect2/defaults-all).
        let callout3 = adj(
            43,
            &[(0, -1800), (2, -3600), (4, -3600), (6, -1800)],
            691,
            692,
        );
        assert_eq!(callout3[0], None);
        close(callout3[1], -8333.0);
        close(callout3[3], -16667.0);
        close(callout3[5], -16667.0);
        close(callout3[7], -8333.0);
        close(adj(53, &[(1, 3600)], 691, 346)[0], 16667.0);
        close(adj(54, &[(1, 18000)], 345, 691)[0], 16667.0);
        close(adj(22, &[(0, 2700)], 547688, 1096963)[0], 25036.0);
        let wedge = adj(62, &[(0, -8503), (1, 13709)], 1440, 672);
        close(wedge[0], -89366.0);
        close(wedge[1], 13468.0);
        close(adj(65, &[(0, 18000)], 492, 484)[0], 16667.0);
        close(adj(68, &[(0, 5184)], 200, 418)[1], 50160.0);
        let both = adj(69, &[(0, 0), (1, 0)], 2686, 368);
        close(both[0], 100000.0);
        close(both[1], 0.0);
        close(adj(69, &[(0, 4957)], 509, 234)[1], 49919.0);
        close(adj(70, &[(1, 6919)], 342, 534)[1], 50016.0);
        close(adj(85, &[(0, 784)], 169, 387)[0], 8312.0);
        let brace = adj(87, &[(0, 574), (1, 6523)], 441, 1384);
        close(brace[0], 8340.0);
        close(brace[1], 30199.0);
        // A 77-degree brace: the DrawingML extent is the swapped anchor.
        close(adj(88, &[(0, 1108)], 255, 414)[0], 8328.0);
        let callout = adj(
            41,
            &[(0, -4816), (1, 26416), (2, 32), (3, 11234)],
            1667,
            325,
        );
        close(callout[0], 52009.0);
        close(callout[1], 148.0);
        close(callout[2], 122296.0);
        close(callout[3], -22296.0);
        // metroBlob pairs: PowerPoint's DrawingML beside the binary values.
        let accent = adj(44, &[(0, -11740), (1, 9416)], 3301229, 3085679);
        assert_eq!((accent[0], accent[1]), (None, None));
        close(accent[2], 43593.0);
        close(accent[3], -54352.0);
        let border3 = adj(
            52,
            &[
                (0, -13516),
                (1, 22284),
                (2, -18037),
                (3, 15089),
                (4, -4655),
                (5, 388),
                (6, -1800),
            ],
            4285543,
            2923781,
        );
        assert_eq!(border3[0], None);
        for (slot, expected) in [
            (1, -8333.0),
            (2, 1794.0),
            (3, -21551.0),
            (4, 69855.0),
            (5, -83506.0),
            (6, 103166.0),
            (7, -62575.0),
        ] {
            let value = border3[slot].unwrap();
            assert!((value - expected).abs() < 2.5, "{slot}: {value} vs {expected}");
        }
    }

    #[test]
    fn unadjusted_shapes_keep_defaults_and_unevidenced_values_fail_closed() {
        assert!(adjustments(43, &[None; 10], 10, 10).unwrap().is_none());
        let mut legacy = [None; 10];
        legacy[0] = Some(-1800);
        assert!(adjustments(42, &legacy, 10, 10).is_err());
        // rightArrow's shaft (adjust2Value) has no evidenced conversion.
        legacy[1] = Some(5400);
        assert!(adjustments(13, &legacy, 10, 10).is_err());
        let mut tall = [None; 10];
        tall[0] = Some(5400);
        assert!(adjustments(9, &tall, 100, 200).is_err());
        assert!(adjustments(999, &tall, 100, 100).is_err());
        assert!(adjustments(16, &tall, 100, 200).is_err());
        assert!(adjustments(22, &tall, 200, 100).is_err());
        // The first leader point's y (adjust8Value) stays unevidenced.
        let mut y = [None; 10];
        y[7] = Some(20000);
        assert!(adjustments(43, &y, 100, 100).is_err());
        let mut curve = [None; 10];
        curve[0] = Some(5400);
        assert!(adjustments(38, &curve, 100, 100).is_err());
    }
}
