//! A bounded, explicit character-property subset. Unknown SPRMs are skipped
//! using their encoded operand size, never interpreted as text or executed.
//! [MS-DOC] 2.2.5, 2.6.1, 2.9.327; ECMA-376 17.3.2 (run properties).

use super::{border::ICO_COLORS, u16_at, u32_at, unsupported};
use std::collections::BTreeMap;

// MS-DOC 2.6.1 sprmCHighlight / 2.9.119 Ico. This is deliberately separate
// from ICO_COLORS: highlight is a symbolic OOXML palette, and controlled
// Word DOC roundtrips distinguish 0x0C darkMagenta from 0x0D darkRed despite
// the duplicated RGB entries in the published Ico table. The control covered
// direct CHPX operands in two fonts; it does not establish style inheritance.
const HIGHLIGHT_COLORS: [&str; 17] = [
    "none",
    "black",
    "blue",
    "cyan",
    "green",
    "magenta",
    "red",
    "yellow",
    "white",
    "darkBlue",
    "darkCyan",
    "darkGreen",
    "darkMagenta",
    "darkRed",
    "darkYellow",
    "darkGray",
    "lightGray",
];

mod direct;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Properties {
    values: BTreeMap<&'static str, String>,
    pub fonts: [Option<usize>; 4],
    font_hint: Option<FontHint>,
    // Sparse style patches must distinguish an absent hint from 0xFF, which
    // explicitly removes inherited guidance but has no ST_Hint equivalent.
    font_hint_present: bool,
    // Raw MS-DOC LID. Resolution is deferred so an assigned locale and an
    // unknown/custom identifier never collapse into the same absent value.
    lang_bidi_lid: Option<u16>,
    // Raw language axes from MS-DOC 2.6.1. Modern properties supply the
    // projected language; compatibility properties remain acquisition metadata,
    // never a guessed fallback or an implicit noProof toggle.
    lang_default_80_lid: Option<u16>,
    lang_east_asia_80_lid: Option<u16>,
    lang_default_lid: Option<u16>,
    lang_east_asia_lid: Option<u16>,
    pub picture: Picture,
    /// Properties projected only by the direct model (see `DirectOnly`).
    direct_only: DirectOnly,
    /// MS-DOC 2.6.1 sprmCSymbol / 2.9.47 CSymbolOperand (ftc, xchar).
    /// Both sprmCPlain and sprmCIstd preserve it.
    symbol: Option<(u16, u16)>,
    /// MS-DOC 2.6.1 sprmCFFldVanish (field text hidden). Both sprmCPlain and
    /// sprmCIstd preserve it.
    field_vanish: Option<bool>,
    /// MS-DOC 2.6.1 insertion revision mark: sprmCFRMarkIns, sprmCIbstRMark
    /// and sprmCDttmRMark. All three survive sprmCPlain and sprmCIstd.
    insertion: InsertionMark,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
struct InsertionMark {
    inserted: Option<bool>,
    author: Option<u16>,
    date: Option<u32>,
}

/// Character properties projected only through dedicated direct DOCX model
/// fields (shading, border, fit text, East Asian layout, grid snapping and
/// line-break clearing). `None` is "not applied", so sparse style patches overlay only what they
/// set, and CPlain/CIstd reset them (neither preserved-property list in
/// MS-DOC 2.6.1 names them).
#[derive(Clone, Debug, Default, PartialEq, Eq)]
struct DirectOnly {
    /// sprmCShd/sprmCShd80 projected fill.
    shading: Option<super::paragraph::ShadingFill>,
    /// sprmCBrc (modern, 8 bytes) / sprmCBrc80 (4 bytes), validated raw.
    border: Option<(bool, [u8; 8])>,
    /// sprmCFitText (dxaFitText twips, FitTextID).
    fit_text: Option<(i32, i32)>,
    /// sprmCFELayout UFEL fTNY / fTNYCompress (horizontal in vertical).
    east_asian: Option<(bool, bool)>,
    /// sprmCFUsePgsuSettings: ECMA-376 17.3.2.34 run snapToGrid.
    snap_to_grid: Option<bool>,
    /// sprmCLbcCRJ raw LBCOperand; only meaningful on U+000B line breaks.
    line_break: Option<u8>,
}

impl DirectOnly {
    fn overlay(&mut self, patch: &Self) {
        if patch.shading.is_some() {
            self.shading = patch.shading.clone();
        }
        if patch.border.is_some() {
            self.border = patch.border;
        }
        if patch.fit_text.is_some() {
            self.fit_text = patch.fit_text;
        }
        if patch.east_asian.is_some() {
            self.east_asian = patch.east_asian;
        }
        if patch.snap_to_grid.is_some() {
            self.snap_to_grid = patch.snap_to_grid;
        }
        if patch.line_break.is_some() {
            self.line_break = patch.line_break;
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum FontHint {
    Default,
    EastAsia,
    ComplexScript,
}

#[derive(Clone, Copy)]
enum LanguageAxis {
    Default,
    EastAsia,
    ComplexScript,
}

enum LanguageResolution {
    Absent,
    Assigned(&'static str),
    Unsupported(u16),
}

#[derive(Clone, Copy)]
pub(super) struct ResolvedLanguages {
    pub(super) default: Option<&'static str>,
    pub(super) east_asia: Option<&'static str>,
    pub(super) bidi: Option<&'static str>,
}

impl FontHint {
    fn xml_value(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::EastAsia => "eastAsia",
            Self::ComplexScript => "cs",
        }
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Picture {
    pub special: bool,
    pub data: bool,
    pub ole: bool,
    pub object: bool,
    pub location: Option<i32>,
}

impl Picture {
    pub fn passive_special(self) -> bool {
        self.special && !self.data && !self.ole && !self.object
    }
    /// MS-DOC 2.6.1: only a special U+0001 without binary-data/OLE flags
    /// denotes a PICFAndOfficeArtData. The caller checks the character code.
    pub fn inline_location(self) -> Result<Option<usize>, String> {
        if !self.passive_special() {
            return Ok(None);
        }
        let location = self
            .location
            .ok_or_else(|| unsupported("Word picture lacks a data location"))?;
        usize::try_from(location)
            .map(Some)
            .map_err(|_| unsupported("negative Word picture data location"))
    }
}

impl Default for Properties {
    fn default() -> Self {
        Self {
            values: BTreeMap::from([("sz", "20".into())]),
            fonts: [None; 4],
            font_hint: None,
            font_hint_present: false,
            lang_bidi_lid: None,
            lang_default_80_lid: None,
            lang_east_asia_80_lid: None,
            lang_default_lid: None,
            lang_east_asia_lid: None,
            picture: Picture::default(),
            direct_only: DirectOnly::default(),
            symbol: None,
            field_vanish: None,
            insertion: InsertionMark::default(),
        }
    }
}

impl Properties {
    /// A resolved style patch, without injecting the document's default size.
    /// Used when list-linked styles are layered onto a paragraph mark.
    pub fn sparse() -> Self {
        Self {
            values: BTreeMap::new(),
            fonts: [None; 4],
            font_hint: None,
            font_hint_present: false,
            lang_bidi_lid: None,
            lang_default_80_lid: None,
            lang_east_asia_80_lid: None,
            lang_default_lid: None,
            lang_east_asia_lid: None,
            picture: Picture::default(),
            direct_only: DirectOnly::default(),
            symbol: None,
            field_vanish: None,
            insertion: InsertionMark::default(),
        }
    }

    pub fn overlay_visible(&mut self, patch: &Self) {
        self.values
            .extend(patch.values.iter().map(|(k, v)| (*k, v.clone())));
        for (current, added) in self.fonts.iter_mut().zip(patch.fonts) {
            if added.is_some() {
                *current = added;
            }
        }
        if patch.font_hint_present {
            self.font_hint = patch.font_hint;
            self.font_hint_present = true;
        }
        if patch.lang_bidi_lid.is_some() {
            self.lang_bidi_lid = patch.lang_bidi_lid;
        }
        for (current, added) in [
            &mut self.lang_default_80_lid,
            &mut self.lang_east_asia_80_lid,
            &mut self.lang_default_lid,
            &mut self.lang_east_asia_lid,
        ]
        .into_iter()
        .zip([
            patch.lang_default_80_lid,
            patch.lang_east_asia_80_lid,
            patch.lang_default_lid,
            patch.lang_east_asia_lid,
        ]) {
            if added.is_some() {
                *current = added;
            }
        }
        self.direct_only.overlay(&patch.direct_only);
        if patch.insertion.inserted.is_some() {
            self.insertion.inserted = patch.insertion.inserted;
        }
        if patch.insertion.author.is_some() {
            self.insertion.author = patch.insertion.author;
        }
        if patch.insertion.date.is_some() {
            self.insertion.date = patch.insertion.date;
        }
        if patch.field_vanish.is_some() {
            self.field_vanish = patch.field_vanish;
        }
        if patch.symbol.is_some() {
            self.symbol = patch.symbol;
        }
        // Object/special flags are not visual run formatting and must not
        // turn numbering text into a picture or an executable object.
    }

    pub fn reset_to(&mut self, paragraph: &Self, preserve_object: bool) {
        // Of the reset exceptions in sprmCPlain/CIstd, these are the supported
        // properties. Revision metadata remains omitted. CIstd preserves
        // CFObj; CPlain does not include it in its exception list.
        let mut picture = self.picture;
        let font_hint = self.font_hint;
        let font_hint_present = self.font_hint_present;
        let symbol = self.symbol;
        let field_vanish = self.field_vanish;
        let insertion = self.insertion;
        if !preserve_object {
            picture.object = paragraph.picture.object;
        }
        let preserved: Vec<_> = ["rtl", "cs", "highlight", "webHidden"]
            .iter()
            .map(|key| (*key, self.values.get(key).cloned()))
            .collect();
        *self = paragraph.clone();
        self.picture = picture;
        // MS-DOC 2.6.1: sprmCPlain and sprmCIstd both leave any previous
        // sprmCIdctHint operand unaffected, including 0xFF (no guidance).
        self.font_hint = font_hint;
        self.font_hint_present = font_hint_present;
        self.symbol = symbol;
        self.field_vanish = field_vanish;
        self.insertion = insertion;
        for (key, value) in preserved {
            if let Some(value) = value {
                self.values.insert(key, value);
            } else {
                self.values.remove(key);
            }
        }
    }

    pub fn apply(&mut self, code: u16, operand: &[u8], style: &Self) -> Result<bool, String> {
        match code {
            0x6a03 => {
                self.picture.location = Some(u32_at(operand, 0)? as i32);
                return Ok(true);
            }
            0x0855 => {
                self.picture.special = match operand[0] {
                    0 => false,
                    1 => true,
                    0x80 => style.picture.special,
                    0x81 => !style.picture.special,
                    _ => return Err(unsupported("invalid Word special-character toggle")),
                };
                return Ok(true);
            }
            0x0806 | 0x080a | 0x0856 => {
                if operand[0] > 1 {
                    return Err(unsupported("invalid Word picture/object flag"));
                }
                let value = operand[0] != 0;
                match code {
                    0x0806 => self.picture.data = value,
                    0x080a => self.picture.ole = value,
                    _ => self.picture.object = value,
                }
                return Ok(true);
            }
            _ => {}
        }
        let flag = match code {
            0x0835 => Some("b"),
            0x0836 => Some("i"),
            0x0837 => Some("strike"),
            0x0838 => Some("outline"),
            0x0839 => Some("shadow"),
            0x083a => Some("smallCaps"),
            0x083b => Some("caps"),
            0x083c => Some("vanish"),
            0x2a53 => Some("dstrike"),
            0x085a => Some("rtl"),
            0x085c => Some("bCs"),
            0x085d => Some("iCs"),
            // MS-DOC 2.6.1 sprmCFNoProof / 2.9.327 ToggleOperand maps to
            // ECMA-376 17.3.2.21. The display-only direct model validates it
            // without introducing proofing UI state.
            0x0875 => Some("noProof"),
            0x0882 => Some("cs"),
            _ => None,
        };
        if let Some(key) = flag {
            if operand.len() != 1 {
                return Err(unsupported("invalid Word character toggle"));
            }
            let base = style.values.get(key).is_some_and(|v| v == "1");
            let value = match operand[0] {
                0 => false,
                1 => true,
                0x80 => base,
                0x81 => !base,
                _ => return Err(unsupported("invalid Word character toggle")),
            };
            self.values
                .insert(key, if value { "1" } else { "0" }.into());
            return Ok(true);
        }
        let (key, value) = match code {
            0x485f => {
                // MS-DOC 2.6.1 sprmCLidBi / 2.9.134 LID: this axis is used for
                // RTL or complex-script presentation. The language itself is
                // not required to belong to either family.
                if operand.len() != 2 {
                    return Err(unsupported("invalid Word complex-script language ID"));
                }
                let lid = u16_at(operand, 0)?;
                // Evidence: styles without a DOCX w:bidi attribute (which
                // inherit the document default) carry LID 0x0400
                // (LOCALE_USER_DEFAULT) in the DOC Word saved from that DOCX.
                // It therefore sets no complex-script language of its own.
                if lid != 0x0400 {
                    self.lang_bidi_lid = Some(lid);
                }
                return Ok(true);
            }
            0x486d | 0x486e | 0x4873 | 0x4874 => {
                // MS-DOC 2.6.1 sprmCRgLid0_80, sprmCRgLid1_80,
                // sprmCRgLid0, and sprmCRgLid1; their operand is the exact
                // two-byte LID from 2.9.134. Keep all four raw axes distinct.
                // Controlled Word roundtrips over the tested flag/order
                // combinations confirm that modern LIDs govern the emitted
                // language. Compatibility LIDs remain typed acquisition
                // metadata only; they never synthesize language or noProof.
                if operand.len() != 2 {
                    return Err(unsupported("invalid Word character language ID"));
                }
                let lid = Some(u16_at(operand, 0)?);
                match code {
                    0x486d => self.lang_default_80_lid = lid,
                    0x486e => self.lang_east_asia_80_lid = lid,
                    0x4873 => self.lang_default_lid = lid,
                    _ => self.lang_east_asia_lid = lid,
                }
                return Ok(true);
            }
            0x6815..=0x6817 => {
                // MS-DOC 2.6.1 identifies these as revision-session IDs for
                // character formatting, inserted text, and deleted text. They
                // do not themselves establish a revision mark. Viewer policy:
                // validate and omit this nonvisual provenance; the actual
                // CFRMarkIns/CFRMarkDel revision properties remain unsupported.
                if operand.len() != 4 {
                    return Err(unsupported("invalid Word character revision session ID"));
                }
                let _ = u32_at(operand, 0)?;
                return Ok(true);
            }
            0x0801 => {
                // MS-DOC 2.6.1 sprmCFRMarkIns (ToggleOperand): text inserted
                // while revision marking was on. Deletions (sprmCFRMarkDel)
                // remain unsupported.
                if operand.len() != 1 {
                    return Err(unsupported("invalid Word character toggle"));
                }
                let base = style.insertion.inserted.unwrap_or(false);
                self.insertion.inserted = Some(match operand[0] {
                    0 => false,
                    1 => true,
                    0x80 => base,
                    0x81 => !base,
                    _ => return Err(unsupported("invalid Word character toggle")),
                });
                return Ok(true);
            }
            0x4804 => {
                // sprmCIbstRMark: a non-negative index into SttbfRMark.
                if operand.len() != 2 || (u16_at(operand, 0)? as i16) < 0 {
                    return Err(unsupported("invalid Word revision author index"));
                }
                self.insertion.author = Some(u16_at(operand, 0)?);
                return Ok(true);
            }
            0x6805 => {
                // sprmCDttmRMark: DTTM of the insertion (MS-DOC 2.9.65).
                if operand.len() != 4 {
                    return Err(unsupported("invalid Word revision date"));
                }
                self.insertion.date = Some(u32_at(operand, 0)?);
                return Ok(true);
            }
            0x0868 | 0x0802 => {
                // ToggleOperand (MS-DOC 2.9.327), relative to the style value.
                if operand.len() != 1 {
                    return Err(unsupported("invalid Word character toggle"));
                }
                let (current, base) = if code == 0x0868 {
                    // sprmCFUsePgsuSettings: "by default, text uses the
                    // document grid"; a corpus DOC/DOCX pair maps it to
                    // ECMA-376 17.3.2.34 w:snapToGrid.
                    (
                        &mut self.direct_only.snap_to_grid,
                        style.direct_only.snap_to_grid.unwrap_or(true),
                    )
                } else {
                    // sprmCFFldVanish: "field text is hidden"; default false.
                    (&mut self.field_vanish, style.field_vanish.unwrap_or(false))
                };
                *current = Some(match operand[0] {
                    0 => false,
                    1 => true,
                    0x80 => base,
                    0x81 => !base,
                    _ => return Err(unsupported("invalid Word character toggle")),
                });
                return Ok(true);
            }
            0x486b => {
                // Not listed in MS-DOC 2.6.1. Its two-byte operand is a
                // Windows code page (the corpus value is 1252). MS-DOC 2.4.1
                // already fixes compressed text to code page 1252 and stores
                // other text as UTF-16, and a Word-saved DOC/DOCX corpus pair
                // has no counterpart for it in the DOCX style. Only that
                // value is accepted as having no display effect.
                if operand.len() != 2 {
                    return Err(unsupported("invalid Word character code page"));
                }
                return Ok(u16_at(operand, 0)? == 1252);
            }
            0x2879 => {
                // MS-DOC 2.6.1 sprmCLbcCRJ / 2.9.129 LBCOperand: where text
                // resumes after a U+000B line break (lbrNone/Left/Right/Both).
                // It MUST NOT be applied to other characters, which ignore it.
                // A Word-saved DOC/DOCX corpus pair carries the undocumented
                // value 0x7C on exactly the four line breaks Word writes as
                // `w:br w:type="textWrapping" w:clear="none"`, and on ordinary
                // text that its DOCX leaves unformatted: only the two low bits
                // select the break type. The value is checked at line breaks.
                self.direct_only.line_break = Some(
                    *operand
                        .first()
                        .ok_or_else(|| unsupported("truncated Word line break type"))?,
                );
                return Ok(true);
            }
            0x6a09 => {
                // MS-DOC 2.6.1 sprmCSymbol / 2.9.47: font index and the
                // Unicode code of the symbol in that font. The font index is
                // validated against the font table at projection.
                if operand.len() != 4 {
                    return Err(unsupported("invalid Word symbol operand"));
                }
                self.symbol = Some((u16_at(operand, 0)?, u16_at(operand, 2)?));
                return Ok(true);
            }
            0x6887 => {
                // MS-DOC 2.6.1 sprmCPbiIBullet: a non-negative CP in the Bullet
                // Pictures document. It only locates the picture used when
                // sprmCPbiGrf enables a picture bullet (handled below).
                if operand.len() != 4 || (u32_at(operand, 0)? as i32) < 0 {
                    return Err(unsupported("invalid Word picture bullet position"));
                }
                return Ok(true);
            }
            0x4888 => {
                // MS-DOC 2.9.176 PbiGrfOperand: fPicBullet (bit 0) states
                // whether the bullet is a picture. A clear bit leaves the text
                // bullet in place, so only an enabled picture bullet, whose
                // image acquisition is not implemented, stays unsupported.
                if operand.len() != 2 {
                    return Err(unsupported("invalid Word picture bullet flags"));
                }
                return Ok(operand[0] & 1 == 0);
            }
            0x0811 => {
                // MS-DOC 2.6.1 sprmCFWebHidden (ToggleOperand): text hidden
                // only in Web Layout view. The direct model is the print/page
                // layout, where the text stays visible, so the validated
                // toggle has no effect. (CPlain/CIstd preserve it, which is
                // moot for a property that is not retained.)
                if operand.len() != 1 || !matches!(operand[0], 0 | 1 | 0x80 | 0x81) {
                    return Err(unsupported("invalid Word web-hidden toggle"));
                }
                return Ok(true);
            }
            0xc81a => {
                // MS-DOC 2.6.1 sprmCFMathPr / 2.9.153 MathPrOperand: cb = 2,
                // jcMath is a DOPMTH mthbpjc value (1..=4). It justifies Office
                // Math equations only; MS-DOC stores no Office Math zones in
                // the text stream (equations are ordinary embedded objects),
                // and note <145> states Word 2007 and later ignore it in
                // compatibility mode while Word 97-2003 never process it.
                if operand.len() != 3 || operand[0] != 2 || !(1..=4).contains(&(operand[1] & 7)) {
                    return Err(unsupported("invalid Word math justification"));
                }
                return Ok(true);
            }
            0xca71 | 0x4866 => {
                // MS-DOC 2.6.1 sprmCShd (SHDOperand) / sprmCShd80 (Shd80).
                return Ok(
                    match super::paragraph::shading_fill(operand, code == 0xca71)? {
                        Some(fill) => {
                            self.direct_only.shading = Some(fill);
                            true
                        }
                        // Valid but a two-color pattern: keep it unsupported.
                        None => false,
                    },
                );
            }
            0xca72 | 0x6865 => {
                // MS-DOC 2.6.1 sprmCBrc (BrcOperand, cb = 8) and sprmCBrc80
                // (Brc80): one border on all four sides of the text.
                let (old, bytes) = if code == 0xca72 {
                    if operand.len() != 9 || operand[0] != 8 {
                        return Err(unsupported("invalid Word character border operand"));
                    }
                    (false, &operand[1..])
                } else {
                    (true, operand)
                };
                if old && operand.len() == 4 && matches!(bytes[1], 0x1a | 0x1b) {
                    return Err(unsupported("invalid Word character Brc80 type"));
                }
                super::border::Border::read(bytes, old)?;
                let size = bytes.len();
                if u32_at(bytes, size - 4)? != u32::MAX {
                    // Brc80/Brc: brcType is byte 1 / 5; byte 3 / 6 holds
                    // dptSpace (5 bits), fShadow (0x20) and fFrame (0x40).
                    let kind = bytes[if old { 1 } else { 5 }];
                    let flags = bytes[if old { 3 } else { 6 }];
                    // The DOCX run-border model has no shadow. fFrame only
                    // reverses a border's appearance across its width
                    // (MS-DOC 2.9.16/17), which is invisible for symmetric
                    // single, double, triple and dash/dot strokes.
                    if flags & 0x20 != 0
                        || (flags & 0x40 != 0 && !matches!(kind, 0 | 1 | 3 | 5..=10 | 22))
                    {
                        return Ok(false);
                    }
                }
                let mut raw = [0; 8];
                raw[..bytes.len()].copy_from_slice(bytes);
                self.direct_only.border = Some((old, raw));
                return Ok(true);
            }
            0xca78 => {
                // MS-DOC 2.9.68 FarEastLayoutOperand: cb = 6, UFEL, ID.
                if operand.len() != 7 || operand[0] != 6 {
                    return Err(unsupported("invalid Word East Asian layout operand"));
                }
                let ufel = u16_at(operand, 1)?;
                let _layout_id = u32_at(operand, 3)?;
                // MS-DOC 2.9.332 UFEL: fTNY (bit 0) is ECMA-376 17.3.2.10
                // eastAsianLayout@vert and fTNYCompress (bit 12) is
                // @vertCompress. Bits that MUST be 0 are ignored as required.
                // fWarichu (two lines in one) has no renderer projection.
                if ufel & 2 != 0 {
                    return Ok(false);
                }
                let vertical = ufel & 1 != 0;
                self.direct_only.east_asian = Some((vertical, vertical && ufel & 0x1000 != 0));
                return Ok(true);
            }
            0xca76 => {
                // MS-DOC 2.9.31 CFitTextOperand: cb = 8, dxaFitText, FitTextID.
                if operand.len() != 9 || operand[0] != 8 {
                    return Err(unsupported("invalid Word fit-text operand"));
                }
                let width = u32_at(operand, 1)? as i32;
                let id = u32_at(operand, 5)? as i32;
                match width {
                    // "A value of zero specifies that the Sprm is ignored."
                    0 => {}
                    // A negative width requests Word's minimum-width fit,
                    // which ECMA-376 17.3.2.14 fitText cannot express.
                    ..=-1 => return Ok(false),
                    _ => self.direct_only.fit_text = Some((width, id)),
                }
                return Ok(true);
            }
            0x2a0c => {
                let index = usize::from(
                    *operand
                        .first()
                        .ok_or_else(|| unsupported("short Word highlight index"))?,
                );
                (
                    "highlight",
                    HIGHLIGHT_COLORS
                        .get(index)
                        .ok_or_else(|| unsupported("invalid Word highlight index"))?
                        .to_string(),
                )
            }
            0x2a42 => {
                // MS-DOC 2.6.1 sprmCIco / 2.9.119 Ico: fixed palette index,
                // with zero representing automatic text color.
                let index = usize::from(
                    *operand
                        .first()
                        .ok_or_else(|| unsupported("short Word text color index"))?,
                );
                (
                    "color",
                    ICO_COLORS
                        .get(index)
                        .ok_or_else(|| unsupported("invalid Word text color index"))?
                        .to_string(),
                )
            }
            0x286f => {
                // MS-DOC 2.6.1 sprmCIdctHint. 0xFF is an explicit absence of
                // guidance and therefore cancels an inherited ST_Hint value.
                self.font_hint = match operand[0] {
                    0 => Some(FontHint::Default),
                    1 => Some(FontHint::EastAsia),
                    2 => Some(FontHint::ComplexScript),
                    0xff => None,
                    _ => return Err(unsupported("invalid Word character font hint")),
                };
                self.font_hint_present = true;
                return Ok(true);
            }
            0x4a4f | 0x4a50 | 0x4a51 | 0x4a5e => {
                let slot = match code {
                    0x4a4f => 0,
                    0x4a50 => 1,
                    0x4a51 => 2,
                    _ => 3,
                };
                let index = u16_at(operand, 0)?;
                if index > i16::MAX as u16 {
                    return Err(unsupported("negative Word font index"));
                }
                self.fonts[slot] = Some(index as usize);
                return Ok(true);
            }
            0x4a43 | 0x4a61 => {
                let size = u16_at(operand, 0)?;
                if size > 3276 || (code == 0x4a43 && size < 2) {
                    return Err(unsupported("invalid Word character size"));
                }
                (if code == 0x4a43 { "sz" } else { "szCs" }, size.to_string())
            }
            0x8840 => ("spacing", (u16_at(operand, 0)? as i16).to_string()),
            0x4845 => {
                let position = u16_at(operand, 0)? as i16;
                if !(-3168..=3168).contains(&position) {
                    return Err(unsupported("invalid Word character position"));
                }
                ("position", position.to_string())
            }
            0x484b => {
                let kern = u16_at(operand, 0)?;
                if kern > 3276 {
                    return Err(unsupported("invalid Word kerning size"));
                }
                ("kern", kern.to_string())
            }
            0x4852 => {
                let value = u16_at(operand, 0)?;
                if !(1..=600).contains(&value) {
                    return Err(unsupported("invalid Word character scaling"));
                }
                ("w", value.to_string())
            }
            0x2a48 => (
                "vertAlign",
                match operand[0] {
                    0 => "baseline",
                    1 => "superscript",
                    2 => "subscript",
                    _ => return Err(unsupported("invalid Word character vertical alignment")),
                }
                .into(),
            ),
            0x2a3e => (
                "u",
                match operand[0] {
                    0 => "none",
                    1 => "single",
                    2 => "words",
                    3 => "double",
                    4 => "dotted",
                    6 => "thick",
                    7 => "dash",
                    9 => "dotDash",
                    10 => "dotDotDash",
                    11 => "wave",
                    20 => "dottedHeavy",
                    23 => "dashedHeavy",
                    25 => "dashDotHeavy",
                    26 => "dashDotDotHeavy",
                    27 => "wavyHeavy",
                    39 => "dashLong",
                    43 => "wavyDouble",
                    55 => "dashLongHeavy",
                    _ => return Err(unsupported("invalid Word underline kind")),
                }
                .into(),
            ),
            0x6870 => ("color", colorref(operand)?),
            0x6877 => ("uColor", colorref(operand)?),
            _ => return Ok(false),
        };
        self.values.insert(key, value);
        Ok(true)
    }

    pub(super) fn resolved_languages(&self) -> Result<ResolvedLanguages, String> {
        Ok(ResolvedLanguages {
            default: self.resolve_language_axis(
                self.lang_default_lid,
                LanguageAxis::Default,
                "default",
            )?,
            east_asia: self.resolve_language_axis(
                self.lang_east_asia_lid,
                LanguageAxis::EastAsia,
                "East Asian",
            )?,
            bidi: self.resolve_language_axis(
                self.lang_bidi_lid,
                LanguageAxis::ComplexScript,
                "complex-script",
            )?,
        })
    }

    fn resolve_language_axis(
        &self,
        lid: Option<u16>,
        axis: LanguageAxis,
        name: &str,
    ) -> Result<Option<&'static str>, String> {
        match self.language(lid, axis) {
            LanguageResolution::Absent => Ok(None),
            LanguageResolution::Assigned(language) => Ok(Some(language)),
            LanguageResolution::Unsupported(lid) => Err(unsupported(format!(
                "unsupported Word {name} language ID 0x{lid:04X}"
            ))),
        }
    }

    fn language(&self, lid: Option<u16>, axis: LanguageAxis) -> LanguageResolution {
        let Some(lid) = lid else {
            return LanguageResolution::Absent;
        };
        match (lid, axis) {
            // Evidence: private DOC files that Word saved from DOCX files. Styles
            // whose DOCX w:lang has w:val="x-none" and w:eastAsia="x-none" carry
            // LID 0x0000 on those axes in the DOC; project the same tag.
            (0x0000, LanguageAxis::Default | LanguageAxis::EastAsia) => {
                return LanguageResolution::Assigned("x-none");
            }
            // MS-LCID LOCALE_CUSTOM_UNSPECIFIED: a language without an LCID.
            // Word wrote it for DOCX tags such as w:val="en-AE" and
            // w:bidi="ae-AR"; the DOC does not retain the tag, so no language
            // is projected (the shared layout treats an unknown tag as absent).
            (0x1000, LanguageAxis::Default | LanguageAxis::ComplexScript) => {
                return LanguageResolution::Absent;
            }
            _ => {}
        }
        match crate::lcid::resolve(u32::from(lid)) {
            crate::lcid::Resolution::Assigned(language) => LanguageResolution::Assigned(language),
            _ => LanguageResolution::Unsupported(lid),
        }
    }
}

fn colorref(operand: &[u8]) -> Result<String, String> {
    match operand.get(3).copied() {
        Some(0) => Ok(format!(
            "{:02X}{:02X}{:02X}",
            operand[0], operand[1], operand[2]
        )),
        Some(0xff) => Ok("auto".into()),
        Some(_) => Err(unsupported("invalid Word COLORREF")),
        None => Err(unsupported("short Word COLORREF")),
    }
}

/// Character subset of [MS-DOC] Prm0.isprm. Non-character properties cannot
/// accidentally enter the character interpreter; unsupported entries warn.
pub fn prm0(prm: u16) -> Option<[u8; 3]> {
    let code: u16 = match (prm >> 1) & 127 {
        0x53 => 0x2a33,
        0x55 => 0x0835,
        0x56 => 0x0836,
        0x57 => 0x0837,
        0x58 => 0x0838,
        0x59 => 0x0839,
        0x5a => 0x083a,
        0x5b => 0x083b,
        0x5c => 0x083c,
        0x5e => 0x2a3e,
        0x68 => 0x2a48,
        0x73 => 0x2a53,
        _ => return None,
    };
    let [a, b] = code.to_le_bytes();
    Some([a, b, (prm >> 8) as u8])
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::sprm::{Budget, Sprms};

    /// The visible direct text run these properties project.
    fn run(properties: &Properties, fonts: &[String]) -> docx_model::TextRun {
        properties
            .direct_text_run("x".into(), fonts)
            .unwrap()
            .expect("visible run")
    }

    fn run_json(properties: &Properties) -> serde_json::Value {
        serde_json::to_value(run(properties, &[])).unwrap()
    }

    fn bidi(properties: &Properties) -> Option<&'static str> {
        properties.resolved_languages().unwrap().bidi
    }

    #[test]
    fn indexed_text_color_uses_shared_palette_and_normal_cascade_order() {
        let base = Properties::default();
        for (index, expected) in ICO_COLORS.iter().enumerate() {
            let mut value = base.clone();
            assert!(value.apply(0x2a42, &[index as u8], &base).unwrap());
            let projected = run(&value, &[]);
            if *expected == "auto" {
                assert!(projected.color_auto && projected.color.is_none());
            } else {
                assert_eq!(projected.color, Some(expected.to_ascii_lowercase()));
            }
        }
        for index in 17..=255u8 {
            assert!(base.clone().apply(0x2a42, &[index], &base).is_err());
        }
        assert!(base.clone().apply(0x2a42, &[], &base).is_err());

        let mut style = base.clone();
        style.apply(0x2a42, &[2], &base).unwrap();
        let mut value = style.clone();
        value.apply(0x6870, &[0x12, 0x34, 0x56, 0], &style).unwrap();
        assert_eq!(run(&value, &[]).color.as_deref(), Some("123456"));
        value.apply(0x2a42, &[6], &style).unwrap();
        assert_eq!(run(&value, &[]).color.as_deref(), Some("ff0000"));
        value.reset_to(&style, false);
        assert_eq!(run(&value, &[]).color.as_deref(), Some("0000ff"));
    }

    #[test]
    fn highlight_uses_symbolic_palette_and_explicit_none() {
        let base = Properties::default();
        for (index, expected) in HIGHLIGHT_COLORS.iter().enumerate() {
            let mut value = base.clone();
            assert!(value.apply(0x2a0c, &[index as u8], &base).unwrap());
            assert_eq!(
                value.values.get("highlight").map(String::as_str),
                Some(*expected)
            );
            let highlight = run(&value, &[]).highlight;
            assert_eq!(
                highlight.as_deref(),
                (*expected != "none").then_some(*expected)
            );
        }
        for index in 17..=255u8 {
            assert!(base.clone().apply(0x2a0c, &[index], &base).is_err());
        }
        assert!(base.clone().apply(0x2a0c, &[], &base).is_err());

        let mut style = base.clone();
        style.apply(0x2a0c, &[12], &base).unwrap();
        let mut value = style.clone();
        value.apply(0x2a0c, &[13], &style).unwrap();
        assert_eq!(run(&value, &[]).highlight.as_deref(), Some("darkRed"));
        // An explicit none overrides the style's highlight.
        value.apply(0x2a0c, &[0], &style).unwrap();
        assert_eq!(run(&value, &[]).highlight, None);

        // MS-DOC 2.6.1 explicitly excludes highlight from both CPlain and
        // CIstd resets. This tests the shared reset primitive used by both.
        value.reset_to(&base, false);
        assert_eq!(run(&value, &[]).highlight, None);
        value.apply(0x2a0c, &[12], &base).unwrap();
        value.reset_to(&base, true);
        assert_eq!(run(&value, &[]).highlight.as_deref(), Some("darkMagenta"));
    }

    #[test]
    fn disabled_picture_bullets_have_no_effect_and_enabled_ones_stay_unsupported() {
        let base = Properties::default();
        let mut value = base.clone();
        assert!(value.apply(0x6887, &[0, 0, 0, 0], &base).unwrap());
        assert!(value.apply(0x4888, &[0, 0], &base).unwrap());
        assert!(value.apply(0x4888, &[0xfe, 0xff], &base).unwrap());
        assert_eq!(value, base);
        assert!(!value.apply(0x4888, &[1, 0], &base).unwrap());
        assert!(base.clone().apply(0x6887, &[0, 0, 0, 0x80], &base).is_err());
        assert!(base.clone().apply(0x6887, &[0, 0, 0], &base).is_err());
        assert!(base.clone().apply(0x4888, &[0], &base).is_err());
    }

    #[test]
    fn insertion_marks_are_style_relative_and_survive_resets() {
        let base = Properties::default();
        let mut value = base.clone();
        assert!(value.apply(0x0801, &[0x81], &base).unwrap());
        assert!(value.apply(0x4804, &1u16.to_le_bytes(), &base).unwrap());
        assert!(value
            .apply(0x6805, &0x86a1_2d2eu32.to_le_bytes(), &base)
            .unwrap());
        assert_eq!(value.direct_insertion(), Some((Some(1), Some(0x86a1_2d2e))));
        // Both CPlain and CIstd preserve the revision-mark properties.
        for preserve_object in [false, true] {
            let mut reset = value.clone();
            reset.reset_to(&base, preserve_object);
            assert_eq!(reset.direct_insertion(), Some((Some(1), Some(0x86a1_2d2e))));
        }
        value.apply(0x0801, &[0], &base).unwrap();
        assert_eq!(value.direct_insertion(), None);
        // Deletions stay unsupported; malformed operands fail.
        assert!(!base.clone().apply(0x0800, &[1], &base).unwrap());
        assert!(base
            .clone()
            .apply(0x4804, &0x8000u16.to_le_bytes(), &base)
            .is_err());
        assert!(base.clone().apply(0x0801, &[2], &base).is_err());
        assert!(base.clone().apply(0x6805, &[0; 3], &base).is_err());
    }

    #[test]
    fn windows_1252_character_code_page_has_no_effect() {
        let base = Properties::default();
        let mut value = base.clone();
        assert!(value.apply(0x486b, &1252u16.to_le_bytes(), &base).unwrap());
        assert_eq!(value, base);
        assert!(!value.apply(0x486b, &932u16.to_le_bytes(), &base).unwrap());
        assert!(base.clone().apply(0x486b, &[0xe4], &base).is_err());
    }

    #[test]
    fn web_hidden_and_math_justification_are_validated_without_print_effect() {
        let base = Properties::default();
        for operand in [0u8, 1, 0x80, 0x81] {
            let mut value = base.clone();
            assert!(value.apply(0x0811, &[operand], &base).unwrap());
            assert_eq!(value, base);
        }
        for operand in [vec![2], vec![0x82], vec![], vec![1, 0]] {
            assert!(base.clone().apply(0x0811, &operand, &base).is_err());
        }
        for jc in 1u8..=4 {
            let mut value = base.clone();
            assert!(value.apply(0xc81a, &[2, jc | 0xf8, 0xff], &base).unwrap());
            assert_eq!(value, base);
        }
        for operand in [vec![2, 0, 0], vec![2, 5, 0], vec![1, 2, 0], vec![2, 2]] {
            assert!(base.clone().apply(0xc81a, &operand, &base).is_err());
        }
    }

    #[test]
    fn revision_session_ids_are_validated_nonvisual_viewer_provenance() {
        let base = Properties::default();
        for code in 0x6815..=0x6817 {
            for value in [0, 0x7856_3412, u32::MAX] {
                let mut properties = base.clone();
                assert!(properties.apply(code, &value.to_le_bytes(), &base).unwrap());
                assert_eq!(run_json(&properties), run_json(&base));
            }
            assert!(base.clone().apply(code, &[0; 3], &base).is_err());
            assert!(base.clone().apply(code, &[0; 5], &base).is_err());
        }

        let mut adjacent = base.clone();
        adjacent
            .apply(0x6816, &0x7856_3412u32.to_le_bytes(), &base)
            .unwrap();
        adjacent.apply(0x0835, &[1], &base).unwrap();
        assert!(run(&adjacent, &[]).bold);
        adjacent.reset_to(&base, false);
        assert_eq!(run_json(&adjacent), run_json(&base));
        // A session ID associated with deletion cannot manufacture a deletion.
        assert!(!adjacent.apply(0x0800, &[1], &base).unwrap());

        let truncated = [0x16, 0x68, 0x12, 0x34, 0x56];
        assert!(Sprms::new(&truncated).next(&mut Budget::default()).is_err());
    }

    #[test]
    fn complex_script_language_is_typed_cascading_and_strictly_resolved() {
        let base = Properties::default();
        let mut style = Properties::sparse();
        assert!(style
            .apply(0x485f, &0x0401u16.to_le_bytes(), &base)
            .unwrap());
        assert_eq!(bidi(&style), Some("ar-SA"));

        let mut patch = Properties::sparse();
        patch
            .apply(0x485f, &0x0411u16.to_le_bytes(), &base)
            .unwrap();
        style.overlay_visible(&patch);
        assert_eq!(bidi(&style), Some("ja-JP"));
        style.overlay_visible(&Properties::sparse());
        assert_eq!(bidi(&style), Some("ja-JP"));
        style
            .apply(0x485f, &0x0409u16.to_le_bytes(), &base)
            .unwrap();
        assert_eq!(bidi(&style), Some("en-US"));

        // Both CPlain and CIstd reset language: neither preserved-property
        // list in MS-DOC 2.6.1 includes language.
        let mut paragraph = base.clone();
        paragraph
            .apply(0x485f, &0x0411u16.to_le_bytes(), &base)
            .unwrap();
        style.reset_to(&paragraph, false);
        assert_eq!(bidi(&style), Some("ja-JP"));
        style
            .apply(0x485f, &0x0401u16.to_le_bytes(), &paragraph)
            .unwrap();
        style.reset_to(&paragraph, true);
        assert_eq!(bidi(&style), Some("ja-JP"));
        for preserve_object in [false, true] {
            let mut reset = style.clone();
            reset.reset_to(&base, preserve_object);
            assert_eq!(bidi(&reset), None);
        }

        assert!(base.clone().apply(0x485f, &[1], &base).is_err());
        assert!(base.clone().apply(0x485f, &[1, 2, 3], &base).is_err());
        for lid in [u16::MAX, 0x0000, 0x007f, 0x0467, 0x040a] {
            let mut unresolved = base.clone();
            unresolved.apply(0x485f, &lid.to_le_bytes(), &base).unwrap();
            assert!(unresolved.resolved_languages().is_err(), "LID {lid:04x}");
        }
        // Word writes 0x0400 for a style without its own complex-script
        // language: it inherits. 0x1000 (a tag without an LCID) projects no
        // language.
        let mut inherited = paragraph.clone();
        inherited
            .apply(0x485f, &0x0400u16.to_le_bytes(), &paragraph)
            .unwrap();
        assert_eq!(bidi(&inherited), Some("ja-JP"));
        let mut custom = paragraph.clone();
        custom
            .apply(0x485f, &0x1000u16.to_le_bytes(), &paragraph)
            .unwrap();
        assert_eq!(bidi(&custom), None);

        let mut truncated = Sprms::new(&[0x5f, 0x48, 0x01]);
        assert!(truncated.next(&mut Budget::default()).is_err());
        let framed = [0x5f, 0x48, 0x01, 0x04, 0x35, 0x08, 0x01];
        let mut framed = Sprms::new(&framed);
        let mut budget = Budget::default();
        assert_eq!(
            framed.next(&mut budget).unwrap(),
            Some((0x485f, &[0x01, 0x04][..]))
        );
        assert_eq!(framed.next(&mut budget).unwrap(), Some((0x0835, &[1][..])));
    }

    #[test]
    fn modern_languages_project_while_compatibility_languages_remain_metadata() {
        let base = Properties::default();
        let cases: [(u16, u16, usize); 4] = [
            (0x486d, 0x0400, 0),
            (0x486e, 0x007f, 1),
            (0x4873, 0x040c, 2),
            (0x4874, 0x0411, 3),
        ];
        let slots = |p: &Properties| {
            [
                p.lang_default_80_lid,
                p.lang_east_asia_80_lid,
                p.lang_default_lid,
                p.lang_east_asia_lid,
            ]
        };

        let mut properties = Properties::sparse();
        for (code, lid, slot) in cases {
            assert!(properties.apply(code, &lid.to_le_bytes(), &base).unwrap());
            assert_eq!(slots(&properties)[slot], Some(lid));
        }
        assert_eq!(
            slots(&properties),
            [Some(0x0400), Some(0x007f), Some(0x040c), Some(0x0411)]
        );
        // Only the modern axes resolve; the Word 97 axes remain metadata.
        let languages = properties.resolved_languages().unwrap();
        assert_eq!(
            (languages.default, languages.east_asia, languages.bidi),
            (Some("fr-FR"), Some("ja-JP"), None)
        );

        // Repeated properties retain their ordered last value without merging
        // the four source axes.
        assert!(properties
            .apply(0x4873, &0x1234u16.to_le_bytes(), &base)
            .unwrap());
        assert_eq!(slots(&properties)[2], Some(0x1234));
        assert_eq!(slots(&properties)[3], Some(0x0411));

        let mut inherited = Properties::sparse();
        inherited
            .apply(0x486d, &0x1111u16.to_le_bytes(), &base)
            .unwrap();
        inherited
            .apply(0x4874, &0x0404u16.to_le_bytes(), &base)
            .unwrap();
        let mut patch = Properties::sparse();
        patch
            .apply(0x486e, &0x3333u16.to_le_bytes(), &base)
            .unwrap();
        inherited.overlay_visible(&patch);
        assert_eq!(
            slots(&inherited),
            [Some(0x1111), Some(0x3333), None, Some(0x0404)]
        );
        inherited.overlay_visible(&Properties::sparse());
        assert_eq!(
            slots(&inherited),
            [Some(0x1111), Some(0x3333), None, Some(0x0404)]
        );

        // CPlain and CIstd use this reset primitive. Language is not among
        // their preserved properties, so all four axes come from the reset
        // target while bidi/noProof remain independent properties.
        let mut paragraph = Properties::sparse();
        paragraph
            .apply(0x4873, &0x0409u16.to_le_bytes(), &base)
            .unwrap();
        inherited
            .apply(0x485f, &0x0401u16.to_le_bytes(), &base)
            .unwrap();
        inherited.apply(0x0875, &[1], &base).unwrap();
        let languages = inherited.resolved_languages().unwrap();
        assert_eq!(
            (languages.east_asia, languages.bidi),
            (Some("zh-TW"), Some("ar-SA"))
        );
        assert_eq!(
            inherited.values.get("noProof").map(String::as_str),
            Some("1")
        );
        for preserve_object in [false, true] {
            let mut reset = inherited.clone();
            reset.reset_to(&paragraph, preserve_object);
            assert_eq!(slots(&reset), [None, None, Some(0x0409), None]);
            assert_eq!(reset.lang_bidi_lid, paragraph.lang_bidi_lid);
            assert!(!reset.values.contains_key("noProof"));
        }
    }

    #[test]
    fn western_and_east_asian_language_ids_require_exact_operands() {
        let base = Properties::default();
        for code in [0x486d, 0x486e, 0x4873, 0x4874] {
            assert!(base.clone().apply(code, &[], &base).is_err());
            assert!(base.clone().apply(code, &[1], &base).is_err());
            assert!(base.clone().apply(code, &[1, 2, 3], &base).is_err());
        }

        let mut truncated = Sprms::new(&[0x6d, 0x48, 0x01]);
        assert!(truncated.next(&mut Budget::default()).is_err());
        let framed = [0x6d, 0x48, 0x01, 0x04, 0x35, 0x08, 0x01];
        let mut framed = Sprms::new(&framed);
        let mut budget = Budget::default();
        assert_eq!(
            framed.next(&mut budget).unwrap(),
            Some((0x486d, &[0x01, 0x04][..]))
        );
        assert_eq!(framed.next(&mut budget).unwrap(), Some((0x0835, &[1][..])));
    }

    #[test]
    fn underline_color_is_an_attribute_and_follows_style_resets() {
        let base = Properties::default();
        let mut style = base.clone();
        style.apply(0x2a3e, &[11], &base).unwrap();
        style.apply(0x6877, &[0x12, 0x34, 0x56, 0], &base).unwrap();
        let underline = |p: &Properties| {
            let run = run(p, &[]);
            (run.underline, run.underline_style, run.underline_color)
        };
        assert_eq!(
            underline(&style),
            (true, Some("wave".into()), Some("123456".into()))
        );
        let mut direct = style.clone();
        direct.apply(0x6877, &[0, 0, 0, 0xff], &style).unwrap();
        assert_eq!(underline(&direct).2.as_deref(), Some("auto"));
        direct.reset_to(&base, false);
        assert_eq!(underline(&direct), (false, None, None));
        assert!(base.clone().apply(0x6877, &[], &base).is_err());
        assert!(base.clone().apply(0x6877, &[1, 2, 3, 1], &base).is_err());

        let mut color_only = base.clone();
        color_only
            .apply(0x6877, &[0x12, 0x34, 0x56, 0], &base)
            .unwrap();
        // MS-OI29500 2.1.100(c): color alone keeps its authorship without
        // activating an underline.
        assert_eq!(underline(&color_only), (false, None, None));
        let wire = run(&color_only, &[]).typography_acquisition.unwrap();
        let wire = wire.underline.unwrap();
        assert_eq!(
            (wire.val.raw, wire.color.raw.as_deref()),
            (None, Some("123456"))
        );
        color_only.apply(0x2a3e, &[4], &base).unwrap();
        assert_eq!(
            underline(&color_only),
            (true, Some("dotted".into()), Some("123456".into()))
        );
    }

    #[test]
    fn picture_metadata_survives_style_reset_without_becoming_run_formatting() {
        let base = Properties::default();
        let mut props = base.clone();
        for (code, bytes) in [
            (0x0855, vec![1]),
            (0x6a03, 123i32.to_le_bytes().to_vec()),
            (0x0856, vec![1]),
        ] {
            assert!(props.apply(code, &bytes, &base).unwrap());
        }
        props.reset_to(&base, true);
        assert!(props.picture.object);
        assert_eq!(props.picture.inline_location().unwrap(), None);
        props.reset_to(&base, false);
        assert_eq!(props.picture.inline_location().unwrap(), Some(123));
        assert_eq!(run_json(&props), run_json(&base));
        for code in [0x0806, 0x080a, 0x0856] {
            let mut active = props.clone();
            active.apply(code, &[1], &base).unwrap();
            active.apply(0x6a03, &(-1i32).to_le_bytes(), &base).unwrap();
            assert_eq!(
                active.picture.inline_location().unwrap(),
                None,
                "active objects must not dereference even an invalid offset"
            );
            assert!(active.apply(code, &[2], &base).is_err());
        }
        props.apply(0x6a03, &(-1i32).to_le_bytes(), &base).unwrap();
        assert!(props.picture.inline_location().is_err());
        props.apply(0x0855, &[0], &base).unwrap();
        assert_eq!(props.picture.inline_location().unwrap(), None);
    }
    #[test]
    fn repeated_toggle_is_relative_to_style_not_previous_direct_value() {
        let mut style = Properties::default();
        style.apply(0x0835, &[1], &Properties::default()).unwrap();
        let mut direct = style.clone();
        direct.apply(0x0835, &[0x81], &style).unwrap();
        direct.apply(0x0835, &[0x81], &style).unwrap();
        assert!(!run(&direct, &[]).bold);
        direct.apply(0x0835, &[0x80], &style).unwrap();
        assert_eq!(direct, style);
        assert!(direct.apply(0x0835, &[2], &style).is_err());
        assert!(direct.apply(0x0835, &[], &style).is_err());
        assert!(direct.apply(0x0835, &[1, 0], &style).is_err());
    }

    #[test]
    fn no_proof_is_a_validated_style_relative_toggle_and_resets_normally() {
        let base = Properties::default();
        let mut style = base.clone();
        assert!(style.apply(0x0875, &[1], &base).unwrap());
        let no_proof = |p: &Properties| p.values.get("noProof").cloned();
        assert_eq!(no_proof(&style).as_deref(), Some("1"));

        for (operand, expected) in [(0, "0"), (1, "1"), (0x80, "1"), (0x81, "0")] {
            let mut direct = style.clone();
            assert!(direct.apply(0x0875, &[operand], &style).unwrap());
            assert_eq!(no_proof(&direct).as_deref(), Some(expected));
        }

        let mut false_style = base.clone();
        false_style.apply(0x0875, &[0], &base).unwrap();
        let mut opposite = false_style.clone();
        opposite.apply(0x0875, &[0x81], &false_style).unwrap();
        assert_eq!(no_proof(&opposite).as_deref(), Some("1"));
        opposite.apply(0x0875, &[0], &false_style).unwrap();
        assert_eq!(no_proof(&opposite).as_deref(), Some("0"));

        let mut inherited = style.clone();
        inherited.overlay_visible(&Properties::sparse());
        assert_eq!(no_proof(&inherited).as_deref(), Some("1"));
        for preserve_object in [false, true] {
            let mut reset = inherited.clone();
            reset.reset_to(&false_style, preserve_object);
            assert_eq!(no_proof(&reset).as_deref(), Some("0"));
        }

        for operand in [vec![], vec![2], vec![0x82], vec![1, 0]] {
            assert!(base.clone().apply(0x0875, &operand, &style).is_err());
        }

        assert!(Sprms::new(&[0x75, 0x08])
            .next(&mut Budget::default())
            .is_err());
        let mut adjacent = Sprms::new(&[0x75, 0x08, 1, 0x35, 0x08, 1]);
        let mut budget = Budget::default();
        assert_eq!(
            adjacent.next(&mut budget).unwrap(),
            Some((0x0875, &[1][..]))
        );
        assert_eq!(
            adjacent.next(&mut budget).unwrap(),
            Some((0x0835, &[1][..]))
        );
    }

    #[test]
    fn preserves_font_slots_without_embedding_fonts() {
        let mut p = Properties::default();
        let base = p.clone();
        p.apply(0x4a4f, &[0, 0], &base).unwrap();
        p.apply(0x4a50, &[1, 0], &base).unwrap();
        p.apply(0x4a43, &[24, 0], &base).unwrap();
        let projected = run(&p, &["A & \"B\"".into(), "CJK".into()]);
        assert_eq!(projected.font_family.as_deref(), Some("A & \"B\""));
        assert_eq!(projected.font_family_east_asia.as_deref(), Some("CJK"));
        assert_eq!(projected.font_size, 12.0);
        assert!(p.direct_text_run("x".into(), &[]).is_err());
    }

    #[test]
    fn preserves_all_documented_font_hints_without_changing_font_slots() {
        let mut p = Properties::default();
        let base = p.clone();
        for (code, bytes) in [
            (0x4a4f, [0, 0]),
            (0x4a50, [1, 0]),
            (0x4a51, [2, 0]),
            (0x4a5e, [3, 0]),
        ] {
            assert!(p.apply(code, &bytes, &base).unwrap());
        }
        let fonts = ["ASCII", "East Asia", "High ANSI", "Complex Script"].map(String::from);
        for (value, expected) in [(0, "default"), (1, "eastAsia"), (2, "cs")] {
            assert!(p.apply(0x286f, &[value], &base).unwrap());
            let projected = run(&p, &fonts);
            assert_eq!(projected.font_hint.as_deref(), Some(expected));
            assert_eq!(
                [
                    projected.font_family.as_deref(),
                    projected.font_family_east_asia.as_deref(),
                    projected.font_family_high_ansi.as_deref(),
                    projected.font_family_cs.as_deref(),
                ],
                [
                    Some("ASCII"),
                    Some("East Asia"),
                    Some("High ANSI"),
                    Some("Complex Script")
                ]
            );
        }
    }

    #[test]
    fn no_guidance_cancels_inherited_hint() {
        let base = Properties::default();
        assert_eq!(run(&base, &[]).font_hint, None);

        let mut inherited = base.clone();
        inherited.apply(0x286f, &[1], &base).unwrap();
        assert_eq!(run(&inherited, &[]).font_hint.as_deref(), Some("eastAsia"));
        inherited.apply(0x286f, &[0xff], &base).unwrap();
        assert_eq!(run(&inherited, &[]).font_hint, None);
        assert!(inherited.apply(0x286f, &[3], &base).is_err());
    }

    #[test]
    fn style_reset_preserves_direction_but_removes_direct_font_size() {
        let base = Properties::default();
        let mut p = base.clone();
        p.apply(0x085a, &[1], &base).unwrap();
        p.apply(0x4a43, &[40, 0], &base).unwrap();
        p.reset_to(&base, false);
        let projected = run(&p, &[]);
        assert_eq!(projected.font_size, 10.0);
        assert_eq!(projected.rtl, Some(true));
    }
}
