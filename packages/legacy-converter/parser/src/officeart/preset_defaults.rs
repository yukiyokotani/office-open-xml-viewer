//! Default adjust values of the DrawingML preset shapes, generated from
//! ECMA-376 Part 1 `presetShapeDefinitions.xml` (each `a:avLst/a:gd` in
//! order, 1/100000 units). An explicit adjust equal to its default is the
//! same geometry as an omitted one (ECMA-376 20.1.9.5). The source defines
//! `upDownArrow` twice with identical values; it is listed once.

#[cfg(any(test, feature = "direct-ppt"))]
const DEFAULTS: &[(&str, &[i32])] = &[
    ("accentBorderCallout1", &[18750, -8333, 112500, -38333]),
    (
        "accentBorderCallout2",
        &[18750, -8333, 18750, -16667, 112500, -46667],
    ),
    (
        "accentBorderCallout3",
        &[18750, -8333, 18750, -16667, 100000, -16667, 112963, -8333],
    ),
    ("accentCallout1", &[18750, -8333, 112500, -38333]),
    (
        "accentCallout2",
        &[18750, -8333, 18750, -16667, 112500, -46667],
    ),
    (
        "accentCallout3",
        &[18750, -8333, 18750, -16667, 100000, -16667, 112963, -8333],
    ),
    ("arc", &[16200000, 0]),
    ("bentArrow", &[25000, 25000, 25000, 43750]),
    ("bentConnector3", &[50000]),
    ("bentConnector4", &[50000, 50000]),
    ("bentConnector5", &[50000, 50000, 50000]),
    ("bentUpArrow", &[25000, 25000, 25000]),
    ("bevel", &[12500]),
    ("blockArc", &[10800000, 0, 25000]),
    ("borderCallout1", &[18750, -8333, 112500, -38333]),
    (
        "borderCallout2",
        &[18750, -8333, 18750, -16667, 112500, -46667],
    ),
    (
        "borderCallout3",
        &[18750, -8333, 18750, -16667, 100000, -16667, 112963, -8333],
    ),
    ("bracePair", &[8333]),
    ("bracketPair", &[16667]),
    ("callout1", &[18750, -8333, 112500, -38333]),
    ("callout2", &[18750, -8333, 18750, -16667, 112500, -46667]),
    (
        "callout3",
        &[18750, -8333, 18750, -16667, 100000, -16667, 112963, -8333],
    ),
    ("can", &[25000]),
    ("chevron", &[50000]),
    ("chord", &[2700000, 16200000]),
    (
        "circularArrow",
        &[12500, 1142319, 20457681, 10800000, 12500],
    ),
    ("cloudCallout", &[-20833, 62500]),
    ("corner", &[50000, 50000]),
    ("cube", &[25000]),
    ("curvedConnector3", &[50000]),
    ("curvedConnector4", &[50000, 50000]),
    ("curvedConnector5", &[50000, 50000, 50000]),
    ("curvedDownArrow", &[25000, 50000, 25000]),
    ("curvedLeftArrow", &[25000, 50000, 25000]),
    ("curvedRightArrow", &[25000, 50000, 25000]),
    ("curvedUpArrow", &[25000, 50000, 25000]),
    ("decagon", &[105146]),
    ("diagStripe", &[50000]),
    ("donut", &[25000]),
    ("doubleWave", &[6250, 0]),
    ("downArrow", &[50000, 50000]),
    ("downArrowCallout", &[25000, 25000, 25000, 64977]),
    ("ellipseRibbon", &[25000, 50000, 12500]),
    ("ellipseRibbon2", &[25000, 50000, 12500]),
    ("foldedCorner", &[16667]),
    ("frame", &[12500]),
    ("gear6", &[15000, 3526]),
    ("gear9", &[10000, 1763]),
    ("halfFrame", &[33333, 33333]),
    ("heptagon", &[102572, 105210]),
    ("hexagon", &[25000, 115470]),
    ("homePlate", &[50000]),
    ("horizontalScroll", &[12500]),
    ("leftArrow", &[50000, 50000]),
    ("leftArrowCallout", &[25000, 25000, 25000, 64977]),
    ("leftBrace", &[8333, 50000]),
    ("leftBracket", &[8333]),
    (
        "leftCircularArrow",
        &[12500, -1142319, 1142319, 10800000, 12500],
    ),
    ("leftRightArrow", &[50000, 50000]),
    ("leftRightArrowCallout", &[25000, 25000, 25000, 48123]),
    (
        "leftRightCircularArrow",
        &[12500, 1142319, 20457681, 11942319, 12500],
    ),
    ("leftRightRibbon", &[50000, 50000, 16667]),
    ("leftRightUpArrow", &[25000, 25000, 25000]),
    ("leftUpArrow", &[25000, 25000, 25000]),
    ("mathDivide", &[23520, 5880, 11760]),
    ("mathEqual", &[23520, 11760]),
    ("mathMinus", &[23520]),
    ("mathMultiply", &[23520]),
    ("mathNotEqual", &[23520, 6600000, 11760]),
    ("mathPlus", &[23520]),
    ("moon", &[50000]),
    ("noSmoking", &[18750]),
    ("nonIsoscelesTrapezoid", &[25000, 25000]),
    ("notchedRightArrow", &[50000, 50000]),
    ("octagon", &[29289]),
    ("parallelogram", &[25000]),
    ("pentagon", &[105146, 110557]),
    ("pie", &[0, 16200000]),
    ("plaque", &[16667]),
    ("plus", &[25000]),
    ("quadArrow", &[22500, 22500, 22500]),
    ("quadArrowCallout", &[18515, 18515, 18515, 48123]),
    ("ribbon", &[16667, 50000]),
    ("ribbon2", &[16667, 50000]),
    ("rightArrow", &[50000, 50000]),
    ("rightArrowCallout", &[25000, 25000, 25000, 64977]),
    ("rightBrace", &[8333, 50000]),
    ("rightBracket", &[8333]),
    ("round1Rect", &[16667]),
    ("round2DiagRect", &[16667, 0]),
    ("round2SameRect", &[16667, 0]),
    ("roundRect", &[16667]),
    ("smileyFace", &[4653]),
    ("snip1Rect", &[16667]),
    ("snip2DiagRect", &[0, 16667]),
    ("snip2SameRect", &[16667, 0]),
    ("snipRoundRect", &[16667, 16667]),
    ("star10", &[42533, 105146]),
    ("star12", &[37500]),
    ("star16", &[37500]),
    ("star24", &[37500]),
    ("star32", &[37500]),
    ("star4", &[12500]),
    ("star5", &[19098, 105146, 110557]),
    ("star6", &[28868, 115470]),
    ("star7", &[34601, 102572, 105210]),
    ("star8", &[37500]),
    ("stripedRightArrow", &[50000, 50000]),
    ("sun", &[25000]),
    ("swooshArrow", &[25000, 16667]),
    ("teardrop", &[100000]),
    ("trapezoid", &[25000]),
    ("triangle", &[50000]),
    ("upArrowCallout", &[25000, 25000, 25000, 64977]),
    ("upDownArrow", &[50000, 50000]),
    ("upDownArrowCallout", &[25000, 25000, 25000, 48123]),
    ("uturnArrow", &[25000, 25000, 25000, 43750, 75000]),
    ("verticalScroll", &[12500]),
    ("wave", &[12500, 0]),
    ("wedgeEllipseCallout", &[-20833, 62500]),
    ("wedgeRectCallout", &[-20833, 62500]),
    ("wedgeRoundRectCallout", &[-20833, 62500, 16667]),
];

/// The preset's default adjust values in `avLst` order, if it has any.
#[cfg(any(test, feature = "direct-ppt"))]
pub(crate) fn defaults(name: &str) -> Option<&'static [i32]> {
    DEFAULTS
        .binary_search_by(|(preset, _)| preset.cmp(&name))
        .ok()
        .map(|index| DEFAULTS[index].1)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn table_is_sorted_and_resolves_known_presets() {
        assert!(DEFAULTS.windows(2).all(|w| w[0].0 < w[1].0));
        assert_eq!(
            defaults("wedgeRoundRectCallout"),
            Some(&[-20833, 62500, 16667][..])
        );
        assert_eq!(defaults("hexagon"), Some(&[25000, 115470][..]));
        assert_eq!(defaults("rect"), None);
    }
}
