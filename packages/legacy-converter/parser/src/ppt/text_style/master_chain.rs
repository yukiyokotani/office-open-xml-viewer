//! Direct-model master text style resolution across indent levels and text
//! types. The byte-conversion path keeps [`Master::levels`]; this resolution
//! is used only by the direct presentation model.
//!
//! MS-PPT 2.9.35 states only that a main master `TextMasterStyleAtom`
//! inherits from the document one. Three further rules are PowerPoint
//! behavior observed in PowerPoint 16 PDF exports of Office-saved decks:
//!
//! - A level (`lstLvl2`-`lstLvl5`) inherits what it leaves unspecified from
//!   the level below it in the same atom. Real atoms store only margins,
//!   indents and a few differences above level 0, yet deeper paragraphs
//!   render with the level-0 typeface, color and bullets; one deck's level-2
//!   body paragraphs render the level-1 bullet although level 0 disables
//!   bullets and level 2 is silent.
//! - Center body, half body and quarter body text (types 5, 7, 8) inherit
//!   from the body style (type 1), and center title (type 6) from the title
//!   style (type 0), before the document style: a center-body paragraph
//!   renders at the body size (28 pt) rather than the document size (18 pt).
//! - Tx_TYPE_OTHER freeform text inherits the document style's levels
//!   (typeface, size, color, bullet glyph), not the main master's type-4
//!   style (a deck whose master type-4 size is 14 pt renders such text at
//!   the document 18 pt).
//!
//! The order in which the three sources combine for one field is not
//! observable in the corpus. Every property is resolved under three natural
//! orders (all own levels first; level by level across own, base and
//! document; level by level across own and base with the document last). A
//! field on which they disagree is recorded as ambiguous and a paragraph or
//! run that would inherit it is rejected.
use super::*;

pub(super) const RTL: u32 = 1;
pub(super) const ALIGN: u32 = 1 << 1;
pub(super) const LINE: u32 = 1 << 2;
pub(super) const BEFORE: u32 = 1 << 3;
pub(super) const AFTER: u32 = 1 << 4;
pub(super) const BULLET: u32 = 1 << 5;
pub(super) const BULLET_HAS_FONT: u32 = 1 << 6;
pub(super) const BULLET_HAS_COLOR: u32 = 1 << 7;
pub(super) const BULLET_HAS_SIZE: u32 = 1 << 8;
pub(super) const BULLET_CHAR: u32 = 1 << 9;
pub(super) const BULLET_FONT: u32 = 1 << 10;
pub(super) const BULLET_SIZE: u32 = 1 << 11;
pub(super) const BULLET_COLOR: u32 = 1 << 12;
pub(super) const MARGIN: u32 = 1 << 13;
pub(super) const INDENT: u32 = 1 << 14;
pub(super) const DEFAULT_TAB: u32 = 1 << 15;
pub(super) const PARAGRAPH: u32 = (1 << 16) - 1;
const BOLD: u32 = 1 << 16;
const ITALIC: u32 = 1 << 17;
const UNDERLINE: u32 = 1 << 18;
const OTHER_STYLE: u32 = 1 << 19;
const SIZE: u32 = 1 << 20;
const FONT: u32 = 1 << 21;
const EA_FONT: u32 = 1 << 22;
const SYMBOL_FONT: u32 = 1 << 23;
const COLOR: u32 = 1 << 24;
pub(super) const CHARACTER: u32 = !PARAGRAPH;

/// Master levels with the fields whose inheritance order is ambiguous.
#[derive(Clone)]
#[cfg(any(test, feature = "direct-ppt"))]
pub(in crate::ppt) struct DirectLevels {
    pub levels: Vec<Level>,
    pub ambiguous: [u32; 5],
}

impl Paragraph {
    /// Fields this exception specifies itself.
    pub(super) fn present(&self) -> u32 {
        let b = &self.bullet;
        [
            (self.rtl.is_some(), RTL),
            (self.align.is_some(), ALIGN),
            (self.spacing[0].is_some(), LINE),
            (self.spacing[1].is_some(), BEFORE),
            (self.spacing[2].is_some(), AFTER),
            (b.enabled.is_some(), BULLET),
            (b.has_font.is_some(), BULLET_HAS_FONT),
            (b.has_color.is_some(), BULLET_HAS_COLOR),
            (b.has_size.is_some(), BULLET_HAS_SIZE),
            (b.character.is_some(), BULLET_CHAR),
            (b.font.is_some(), BULLET_FONT),
            (b.size.is_some(), BULLET_SIZE),
            (b.color.is_some(), BULLET_COLOR),
            (self.margin.is_some(), MARGIN),
            (self.indent.is_some(), INDENT),
            (self.default_tab.is_some(), DEFAULT_TAB),
        ]
        .into_iter()
        .filter_map(|(present, bit)| present.then_some(bit))
        .fold(0, |mask, bit| mask | bit)
    }
}

impl Character {
    /// Fields this exception specifies itself.
    pub(super) fn present(&self) -> u32 {
        [
            (self.mask & 1 != 0, BOLD),
            (self.mask & 2 != 0, ITALIC),
            (self.mask & 4 != 0, UNDERLINE),
            (self.mask & 0x3eb0 != 0, OTHER_STYLE),
            (self.mask & 0x20000 != 0, SIZE),
            (self.font.is_some(), FONT),
            (self.ea.is_some(), EA_FONT),
            (self.symbol.is_some(), SYMBOL_FONT),
            (self.color.is_some(), COLOR),
        ]
        .into_iter()
        .filter_map(|(present, bit)| present.then_some(bit))
        .fold(0, |mask, bit| mask | bit)
    }
}

/// Fields whose effective values differ between two resolved levels.
#[cfg(any(test, feature = "direct-ppt"))]
fn differing(a: &Level, b: &Level) -> u32 {
    let (p, q) = (&a.paragraph, &b.paragraph);
    let (x, y) = (&a.character, &b.character);
    let style = |c: &Character, bits: u16| (c.mask as u16 & bits, c.style & c.mask as u16 & bits);
    [
        (p.rtl != q.rtl, RTL),
        (p.align != q.align, ALIGN),
        (p.spacing[0] != q.spacing[0], LINE),
        (p.spacing[1] != q.spacing[1], BEFORE),
        (p.spacing[2] != q.spacing[2], AFTER),
        (p.bullet.enabled != q.bullet.enabled, BULLET),
        (p.bullet.has_font != q.bullet.has_font, BULLET_HAS_FONT),
        (p.bullet.has_color != q.bullet.has_color, BULLET_HAS_COLOR),
        (p.bullet.has_size != q.bullet.has_size, BULLET_HAS_SIZE),
        (p.bullet.character != q.bullet.character, BULLET_CHAR),
        (p.bullet.font != q.bullet.font, BULLET_FONT),
        (p.bullet.size != q.bullet.size, BULLET_SIZE),
        (p.bullet.color != q.bullet.color, BULLET_COLOR),
        (p.margin != q.margin, MARGIN),
        (p.indent != q.indent, INDENT),
        (p.default_tab != q.default_tab, DEFAULT_TAB),
        (style(x, 1) != style(y, 1), BOLD),
        (style(x, 2) != style(y, 2), ITALIC),
        (style(x, 4) != style(y, 4), UNDERLINE),
        (style(x, 0x3eb0) != style(y, 0x3eb0), OTHER_STYLE),
        (
            (x.mask & 0x20000 != 0).then_some(x.size) != (y.mask & 0x20000 != 0).then_some(y.size),
            SIZE,
        ),
        (x.font != y.font, FONT),
        (x.ea != y.ea, EA_FONT),
        (x.symbol != y.symbol, SYMBOL_FONT),
        (x.color != y.color, COLOR),
    ]
    .into_iter()
    .filter_map(|(differs, bit)| differs.then_some(bit))
    .fold(0, |mask, bit| mask | bit)
}

#[cfg(any(test, feature = "direct-ppt"))]
fn fold<'a>(level: usize, sources: impl Iterator<Item = &'a Level>) -> Level {
    sources.fold(Level::empty(level as u16), |acc, source| {
        acc.inherit(Some(source))
    })
}

/// Resolve five levels from an atom's own levels, an optional base atom
/// (body or title for the derived placeholder types) and the document atom.
#[cfg(any(test, feature = "direct-ppt"))]
pub(super) fn resolve(own: &[Level], base: &[Level], document: &[Level]) -> DirectLevels {
    let mut levels = Vec::with_capacity(5);
    let mut ambiguous = [0; 5];
    for (k, slot) in ambiguous.iter_mut().enumerate() {
        let below = || (0..=k).rev();
        let own_first = fold(
            k,
            below()
                .filter_map(|j| own.get(j))
                .chain(below().filter_map(|j| base.get(j)))
                .chain(below().filter_map(|j| document.get(j))),
        );
        let by_level = fold(
            k,
            below().flat_map(|j| {
                [own.get(j), base.get(j), document.get(j)]
                    .into_iter()
                    .flatten()
            }),
        );
        let document_last = fold(
            k,
            below()
                .flat_map(|j| [own.get(j), base.get(j)].into_iter().flatten())
                .chain(below().filter_map(|j| document.get(j))),
        );
        *slot = differing(&own_first, &by_level) | differing(&own_first, &document_last);
        levels.push(own_first);
    }
    DirectLevels { levels, ambiguous }
}

/// The atom whose levels a derived placeholder type inherits before the
/// document atom: body for center/half/quarter body, title for center title.
#[cfg(any(test, feature = "direct-ppt"))]
pub(super) fn base_type(kind: u16) -> Option<u16> {
    match kind {
        5 | 7 | 8 => Some(1),
        6 => Some(0),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn level(k: u16, edit: impl FnOnce(&mut Level)) -> Level {
        let mut level = Level::empty(k);
        edit(&mut level);
        level
    }

    #[test]
    fn levels_inherit_the_level_below_then_the_base_then_the_document() {
        let own = [
            level(0, |l| {
                l.paragraph.bullet.enabled = Some(false);
                l.character.font = Some(2);
            }),
            level(1, |l| l.paragraph.bullet.enabled = Some(true)),
            level(2, |l| l.paragraph.margin = Some(576)),
        ];
        let document = [level(0, |l| {
            l.character.font = Some(0);
            l.character.color = Some(0x0100_0000);
        })];
        let resolved = resolve(&own, &[], &document);
        let two = &resolved.levels[2];
        assert_eq!(two.paragraph.level, 2);
        assert_eq!(two.paragraph.bullet.enabled, Some(true));
        assert_eq!(two.paragraph.margin, Some(576));
        assert_eq!(two.character.font, Some(2));
        assert_eq!(two.character.color, Some(0x0100_0000));
        assert_eq!(resolved.ambiguous, [0; 5]);
        // Levels absent from every atom still resolve from the chain.
        assert_eq!(resolved.levels[4].paragraph.bullet.enabled, Some(true));
    }

    #[test]
    fn derived_types_use_the_base_before_the_document() {
        let own = [level(0, |l| l.paragraph.align = Some(1))];
        let body = [level(0, |l| {
            l.character.mask |= 0x20000;
            l.character.size = 28;
        })];
        let document = [level(0, |l| {
            l.character.mask |= 0x20000;
            l.character.size = 18;
        })];
        let resolved = resolve(&own, &body, &document);
        assert_eq!(resolved.levels[0].character.size, 28);
        assert_eq!(resolved.levels[0].paragraph.align, Some(1));
        assert_eq!(resolved.ambiguous[0], 0);
        assert_eq!(base_type(5), Some(1));
        assert_eq!(base_type(7), Some(1));
        assert_eq!(base_type(8), Some(1));
        assert_eq!(base_type(6), Some(0));
        assert_eq!(base_type(1), None);
        assert_eq!(base_type(4), None);
    }

    #[test]
    fn unobserved_source_orders_that_disagree_are_ambiguous() {
        // Own level 0 and document level 1 both specify the color; whether a
        // level-1 paragraph sees the own lower level or the document's own
        // level first is not established.
        let own = [level(0, |l| l.character.color = Some(1))];
        let document = [level(0, |_| {}), level(1, |l| l.character.color = Some(2))];
        let resolved = resolve(&own, &[], &document);
        assert_eq!(resolved.ambiguous[0], 0);
        assert_eq!(resolved.ambiguous[1] & COLOR, COLOR);
        assert_eq!(resolved.ambiguous[1] & !COLOR, 0);
    }

    #[test]
    fn presence_masks_cover_each_field_once() {
        let mut paragraph = Level::empty(0).paragraph;
        assert_eq!(paragraph.present(), 0);
        paragraph.bullet.character = Some(0x2022);
        paragraph.margin = Some(0);
        assert_eq!(paragraph.present(), BULLET_CHAR | MARGIN);
        let mut character = Level::empty(0).character;
        assert_eq!(character.present(), 0);
        character.mask = 0x20001;
        assert_eq!(character.present(), BOLD | SIZE);
        assert_eq!(PARAGRAPH & CHARACTER, 0);
    }
}
