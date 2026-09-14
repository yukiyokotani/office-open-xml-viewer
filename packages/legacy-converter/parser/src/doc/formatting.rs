//! Resolve the shared DOC formatting cascade for ordinary OOXML and direct-model
//! projection. [MS-DOC] 2.4.6, STSH/STD, SttbfFfn and FFN apply. Direct-model
//! projection includes bounded table-style conditional color; other conditional
//! formatting and advanced paragraph properties remain gated.

use super::character::{self, Properties};
use super::fkp::{self, Index, Kind};
use super::sprm::{self, Budget, Sprms};
use super::{numbering, paragraph, table, u16_at, unsupported};
use std::collections::{BTreeMap, BTreeSet};
use std::rc::Rc;

mod cnf;
#[cfg(feature = "direct-doc")]
mod direct;
mod table_style;
mod tapx;
pub(in crate::doc) use table_style::TableFormattingKey;

/// Table-aware caches trade recomputation for a fixed retained-entry bound.
/// A document can otherwise present the Cartesian product of paragraph and
/// table style IDs even though each individual style table is bounded.
const MAX_TABLE_AWARE_CACHE_ENTRIES: usize = 256;

struct Style<'a> {
    base: usize,
    kind: u16,
    chpx: &'a [u8],
    papx: &'a [u8],
    table: Option<TableStylePropertySets<'a>>,
    language_compatibility: StyleLanguageCompatibility,
}

/// Raw, ordered MS-DOC 2.9.270 StkTableGRLPUPX members. Acquisition is
/// deliberately lossless: interpretation and inheritance are separate slices.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct TableStylePropertySets<'a> {
    tapx: &'a [u8],
    papx: &'a [u8],
    chpx: &'a [u8],
}

#[derive(Clone)]
struct CachedParagraphProperties {
    properties: paragraph::Properties,
    /// True only while the effective alignment still comes from the table
    /// style. A later paragraph-style/direct/list write clears this marker.
    table_alignment: bool,
}

/// Raw MS-DOC 2.9.112 GRFSTD language-compatibility facts. Interpretation is
/// deferred until the _80/modern language precedence is established.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[allow(dead_code)] // Acquired now; consumed by the subsequent language-resolution slice.
pub(in crate::doc) struct StyleLanguageCompatibility {
    f97_lids_set: bool,
    f_copy_lang: bool,
}

#[allow(dead_code)] // Acquired now; consumed by the subsequent resolution slice.
impl StyleLanguageCompatibility {
    pub(in crate::doc) fn compatibility_lids_applied(self) -> bool {
        self.f97_lids_set
    }

    /// `None` represents MS-DOC's requirement to ignore fCopyLang when
    /// f97LidsSet is clear; the raw bit remains retained in this value.
    pub(in crate::doc) fn copy_language(self) -> Option<bool> {
        self.f97_lids_set.then_some(self.f_copy_lang)
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct DirectParagraphProperties {
    bidi: Option<bool>,
    alignment: Option<(u16, u8)>,
    absolute_indents: [Option<(u16, i16)>; 6],
}

impl DirectParagraphProperties {
    fn overlay(&mut self, later: Self) {
        if later.bidi.is_some() {
            self.bidi = later.bidi;
        }
        if later.alignment.is_some() {
            self.alignment = later.alignment;
        }
        for value in later.absolute_indents.into_iter().flatten() {
            self.push_absolute_indent(value);
        }
    }

    fn push_absolute_indent(&mut self, value: (u16, i16)) {
        if let Some(index) = self
            .absolute_indents
            .iter()
            .position(|entry| entry.is_some_and(|entry| entry.0 == value.0))
        {
            self.absolute_indents[index..].rotate_left(1);
            self.absolute_indents[5] = None;
        }
        if let Some(slot) = self.absolute_indents.iter_mut().find(|slot| slot.is_none()) {
            *slot = Some(value);
        } else {
            // There are exactly six recognized absolute-indent codes. A full
            // array with no matching code therefore cannot accept another one.
            debug_assert!(false, "absolute-indent code set exceeded its fixed bound");
        }
    }
}

pub struct Formatting<'a> {
    pub characters: Index<'a>,
    paragraphs: Index<'a>,
    fonts: Vec<String>,
    defaults: Properties,
    styles: Vec<Option<Style<'a>>>,
    paragraph_cache: BTreeMap<(usize, Option<TableFormattingKey>), Properties>,
    paragraph_marker_styles: BTreeMap<usize, Properties>,
    paragraph_layout_cache:
        BTreeMap<(usize, Option<TableFormattingKey>), CachedParagraphProperties>,
    table_style_cache: BTreeMap<usize, Rc<table_style::Profile>>,
    effective_nfib: u16,
    interpret_table_styles: bool,
    data: &'a [u8],
    budget: Budget,
    numbering: numbering::Tables<'a>,
    pub numbering_output: numbering::output::Store,
    pub unsupported_character_properties: bool,
    pub unsupported_paragraph_properties: bool,
    pub unsupported_piece_properties: bool,
    pub missing_tables: bool,
    pub unsupported_table_properties: bool,
}

pub(in crate::doc) struct ResolvedParagraph {
    pub(in crate::doc) properties: paragraph::Properties,
    pub(in crate::doc) numbering: Option<(numbering::Reference, Properties)>,
    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) paragraph_mark: Option<Properties>,
}

impl<'a> Formatting<'a> {
    #[allow(dead_code)] // Internal producer seam for the next resolution slice.
    pub(in crate::doc) fn style_language_compatibility(
        &self,
        index: usize,
    ) -> Option<StyleLanguageCompatibility> {
        self.styles
            .get(index)
            .and_then(Option::as_ref)
            .map(|style| style.language_compatibility)
    }

    pub fn read(word: &'a [u8], table: &'a [u8], data: &'a [u8]) -> Result<Self, String> {
        let fonts = read_fonts(fkp::table_part(word, table, 0x112)?)?;
        let (defaults, styles) = read_styles(fkp::table_part(word, table, 0xa2)?)?;
        let characters = Index::read(word, table, Kind::Character)?;
        let paragraphs = Index::read(word, table, Kind::Paragraph)?;
        let missing_tables =
            characters.is_empty() || paragraphs.is_empty() || styles.is_empty() || fonts.is_empty();
        Ok(Self {
            characters,
            paragraphs,
            fonts,
            defaults,
            styles,
            paragraph_cache: BTreeMap::new(),
            paragraph_marker_styles: BTreeMap::new(),
            paragraph_layout_cache: BTreeMap::new(),
            table_style_cache: BTreeMap::new(),
            effective_nfib: 0x00c1,
            interpret_table_styles: false,
            data,
            budget: Budget::default(),
            numbering: numbering::Tables::read(word, table)?,
            numbering_output: numbering::output::Store::default(),
            unsupported_character_properties: false,
            unsupported_paragraph_properties: false,
            unsupported_piece_properties: false,
            missing_tables,
            unsupported_table_properties: false,
        })
    }

    pub fn paragraph_style(&self, end_fc: usize) -> Result<usize, String> {
        if !self.paragraphs.is_empty() && self.paragraphs.at(end_fc).is_none() {
            return Err(unsupported("Word paragraph mark outside formatting ranges"));
        }
        fkp::paragraph_style(&self.paragraphs, end_fc)
    }

    pub(in crate::doc) fn configure_table_styles(
        &mut self,
        effective_nfib: u16,
        interpret_table_styles: bool,
    ) {
        debug_assert!(self.table_style_cache.is_empty());
        self.effective_nfib = effective_nfib;
        self.interpret_table_styles = interpret_table_styles;
    }

    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) fn use_raw_table_shading(&self) -> bool {
        self.effective_nfib > 0x00d9 && self.interpret_table_styles
    }

    pub(in crate::doc) fn resolve_table_style_id(&self, selected: Option<usize>) -> Option<usize> {
        let selected = selected?;
        // [MS-DOC] 2.6.3 sprmTIstd: an empty, missing, or wrong-kind style is
        // equivalent to applying istd 0x000B. Absence remains distinct because
        // 2.4.6.6 Part 1 step 6.3 skips table-style formatting when TIstd was
        // never applied. Validation of style 0x000B itself belongs to the later
        // style-property application slice.
        Some(
            self.styles
                .get(selected)
                .and_then(Option::as_ref)
                .filter(|style| style.kind == 3)
                .map_or(0x000b, |_| selected),
        )
    }

    pub fn paragraph_xml(
        &mut self,
        style: usize,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<String, String> {
        let mut resolved = self.resolve_paragraph(style, fc, prm, prcs)?;
        if let Some((reference, marker)) = resolved.numbering {
            let ppr = resolved.properties.xml();
            let rpr = self.byte_run_xml(&marker)?;
            let id = self.numbering_output.activate(
                &self.numbering,
                reference,
                ppr,
                rpr,
                super::MAX_DOCUMENT_XML_BYTES,
            )?;
            resolved.properties.numbering = id.map(|id| (id, reference.level));
        }
        Ok(resolved.properties.xml())
    }

    pub(in crate::doc) fn resolve_paragraph(
        &mut self,
        style: usize,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<ResolvedParagraph, String> {
        self.resolve_paragraph_with_table(style, None, fc, prm, prcs)
    }

    fn resolve_paragraph_with_table(
        &mut self,
        style: usize,
        table_style: Option<TableFormattingKey>,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<ResolvedParagraph, String> {
        let key = (style, table_style);
        let cached = if let Some(cached) = self.paragraph_layout_cache.get(&key) {
            cached.clone()
        } else {
            let mut props = paragraph::Properties::default();
            let mut table_alignment = self.apply_table_paragraph_style(&mut props, table_style)?;
            for id in self.chain(style, 1)? {
                let bytes = self.styles[id].as_ref().expect("validated style").papx;
                if self.apply_paragraph(&mut props, bytes)?.alignment.is_some() {
                    table_alignment = false;
                }
            }
            let cached = CachedParagraphProperties {
                properties: props,
                table_alignment,
            };
            if self.paragraph_layout_cache.len() >= MAX_TABLE_AWARE_CACHE_ENTRIES {
                self.paragraph_layout_cache.clear();
            }
            self.paragraph_layout_cache.insert(key, cached.clone());
            cached
        };
        let mut props = cached.properties;
        let mut table_alignment = cached.table_alignment;
        let direct = self
            .paragraphs
            .at(fc)
            .map_or(&[][..], |(_, run)| run.properties);
        let mut direct_properties = if !direct.is_empty() {
            self.apply_paragraph(&mut props, &direct[2..])?
        } else {
            DirectParagraphProperties::default()
        };
        if prm & 1 != 0 {
            let bytes = prcs
                .get((prm >> 1) as usize)
                .ok_or_else(|| unsupported("Word piece property index outside CLX"))?;
            let filter = if self.interpret_table_styles {
                sprm::TopLevelFilter::Paragraph
            } else {
                sprm::TopLevelFilter::All
            };
            direct_properties.overlay(self.apply_paragraph_with_filter(&mut props, bytes, filter)?);
        } else if let Some(bytes) = paragraph::prm0(prm) {
            direct_properties.overlay(self.apply_paragraph(&mut props, &bytes)?);
        }
        if direct_properties.alignment.is_some() {
            table_alignment = false;
        }
        if let Some(reference) = numbering::Reference::new(props.ilfo, props.ilvl)? {
            let selected = self.numbering.resolve(reference)?;
            let level = *selected.level;
            let linked = selected.list.styles[usize::from(reference.level)];
            let original = props.clone();
            // MS-DOC 2.4.6.3 part 3 / 2.4.6.6 part 2: list paragraph
            // properties follow style/direct/PCD properties. Body character
            // runs continue to resolve their own original style and CHPX.
            if linked != 0xfff {
                for id in self.chain(usize::from(linked), 1)? {
                    let bytes = self.styles[id].as_ref().expect("validated style").papx;
                    if self.apply_paragraph(&mut props, bytes)?.alignment.is_some() {
                        table_alignment = false;
                    }
                }
            }
            if self
                .apply_paragraph(&mut props, level.papx)?
                .alignment
                .is_some()
            {
                table_alignment = false;
            }
            // MS-DOC 2.4.6.3 describes applying list properties after the
            // paragraph's style/direct properties. Word-produced evidence
            // establishes narrower precedence exceptions for explicitly
            // authored direct bidi and alignment: retain them after list
            // application. This does not alter other paragraph properties or
            // the same properties inherited only through a style.
            if let Some(value) = direct_properties.bidi {
                props.set_bidi(value);
            }
            if let Some((code, value)) = direct_properties.alignment {
                // Reuse the paragraph property's enum/range validation.
                props.apply(code, &[value])?;
                table_alignment = false;
            }
            // Replay only explicitly direct absolute twip indents, after the
            // list and after final direct bidi restoration. The fixed-size
            // capture preserves PAPX→PCD chronology without retaining style,
            // default, relative-nest, or character-unit properties.
            // Word-produced positive-iLfo controls establish this for signed
            // absolute values, explicit zero, and both paragraph directions;
            // this is observed Office precedence, not the literal list-last
            // ordering in MS-DOC 2.4.6.3. No coordinate correction is applied.
            for (code, value) in direct_properties.absolute_indents.into_iter().flatten() {
                props.apply(code, &value.to_le_bytes())?;
            }
            if reference.preserve_indent {
                props.preserve_list_indent(&original);
            }

            if table_alignment && props.is_bidi() {
                // The Office controls establish unconditional table PAPX
                // alignment only for LTR paragraphs. Keep RTL behavior on the
                // pre-existing paragraph/list cascade until separately proven,
                // and keep it behind the admission gate when TIstd support is
                // broadened later.
                self.unsupported_paragraph_properties = true;
                props.clear_alignment();
            }

            #[cfg(feature = "direct-doc")]
            let paragraph_mark =
                self.run_properties_with_table(style, table_style, fc, prm, prcs)?;
            #[cfg(feature = "direct-doc")]
            let mut marker = paragraph_mark.clone();
            #[cfg(not(feature = "direct-doc"))]
            let mut marker = self.run_properties(style, fc, prm, prcs)?;
            let mut baseline = self.paragraph_base_with_table(style, table_style)?;
            if linked != 0xfff {
                // Resolve the linked style's toggles against its own base
                // chain once, then overlay its explicit visible properties.
                // Reapplying that chain to an already styled mark would
                // invert relative toggles twice and inject a default size.
                let id = usize::from(linked);
                let patch = if let Some(patch) = self.paragraph_marker_styles.get(&id) {
                    patch.clone()
                } else {
                    let mut patch = Properties::sparse();
                    self.apply_style(&mut patch, id, 1)?;
                    self.paragraph_marker_styles.insert(id, patch.clone());
                    patch
                };
                marker.overlay_visible(&patch);
                baseline.overlay_visible(&patch);
            }
            let mut character_style = baseline.clone();
            self.apply_direct(&mut marker, &mut character_style, &baseline, level.chpx)?;
            return Ok(ResolvedParagraph {
                properties: props,
                numbering: Some((reference, marker)),
                #[cfg(feature = "direct-doc")]
                paragraph_mark: Some(paragraph_mark),
            });
        }
        if table_alignment && props.is_bidi() {
            // See the bounded LTR compatibility note in the numbered branch.
            self.unsupported_paragraph_properties = true;
            props.clear_alignment();
        }
        Ok(ResolvedParagraph {
            properties: props,
            numbering: None,
            #[cfg(feature = "direct-doc")]
            paragraph_mark: None,
        })
    }

    /// Apply the validated table-style PAPX alignment established by the
    /// Office controls. This remains narrower than a general interpretation
    /// of the basedOn array wording in MS-DOC 2.4.6.5.
    fn apply_table_paragraph_style(
        &mut self,
        props: &mut paragraph::Properties,
        table_style: Option<TableFormattingKey>,
    ) -> Result<bool, String> {
        let Some(table_style) = table_style else {
            return Ok(false);
        };
        let profile = self.table_style_profile(table_style.selected_style)?;
        if self
            .styles
            .get(table_style.selected_style)
            .and_then(Option::as_ref)
            .filter(|style| style.kind == 3)
            .is_none()
        {
            // sprmTIstd falls back to style 0x000B for an invalid selection.
            // If that slot is itself unavailable, retain the existing table
            // formatting gate instead of failing before its generic result.
            self.unsupported_table_properties = true;
            return Ok(false);
        }
        let mut alignment = profile.paragraph_alignment;
        for condition in table_style.matches.into_iter().flatten() {
            if let Some(patch) = profile.conditional_paragraph_alignment.get(&condition) {
                alignment = Some(*patch);
            }
        }
        if let Some(alignment) = alignment {
            alignment.apply(props);
            Ok(true)
        } else {
            Ok(false)
        }
    }

    fn apply_paragraph<'b>(
        &mut self,
        props: &mut paragraph::Properties,
        bytes: &'b [u8],
    ) -> Result<DirectParagraphProperties, String>
    where
        'a: 'b,
    {
        self.apply_paragraph_with_filter(props, bytes, sprm::TopLevelFilter::All)
    }

    fn apply_paragraph_with_filter<'b>(
        &mut self,
        props: &mut paragraph::Properties,
        bytes: &'b [u8],
        filter: sprm::TopLevelFilter,
    ) -> Result<DirectParagraphProperties, String>
    where
        'a: 'b,
    {
        let mut direct = DirectParagraphProperties::default();
        sprm::paragraph_properties(
            bytes,
            self.data,
            &mut self.budget,
            filter,
            |code, operand, _| {
                if (code >> 10) & 7 == 1
                    && !matches!(code, 0x2416 | 0x2417 | 0x6649 | 0x664a | 0x244b | 0x244c)
                    && !props.apply(code, operand)?
                {
                    self.unsupported_paragraph_properties = true;
                }
                if code == 0x2441 {
                    // Properties::apply validated this Bool8 operand above.
                    direct.bidi = Some(operand[0] == 1);
                } else if matches!(code, 0x2403 | 0x2461) {
                    // Properties::apply validated the alignment enum above.
                    direct.alignment = Some((code, operand[0]));
                } else if matches!(code, 0x840e | 0x840f | 0x845d | 0x845e | 0x8411 | 0x8460) {
                    // Properties::apply validated the signed XAS operand above.
                    direct.push_absolute_indent((code, u16_at(operand, 0)? as i16));
                }
                Ok(())
            },
        )?;
        Ok(direct)
    }

    pub fn table_properties(
        &mut self,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<table::Properties, String> {
        self.table_properties_with_policy(fc, prm, prcs, false)
    }

    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) fn table_properties_native(
        &mut self,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<table::Properties, String> {
        if !self.interpret_table_styles {
            return Err(unsupported(
                "native Word table-style acquisition is not enabled",
            ));
        }
        self.table_properties_with_policy(fc, prm, prcs, true)
    }

    fn table_properties_with_policy(
        &mut self,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
        interpret_table_styles: bool,
    ) -> Result<table::Properties, String> {
        let mut properties = table::Properties::default();
        // MS-DOC 2.4.3: structural flags are direct, never inherited from STSH.
        let direct = self
            .paragraphs
            .at(fc)
            .map_or(&[][..], |(_, r)| r.properties);
        let direct = if direct.is_empty() {
            direct
        } else {
            &direct[2..]
        };
        let piece = if prm & 1 != 0 {
            *prcs
                .get((prm >> 1) as usize)
                .ok_or_else(|| unsupported("Word piece property index outside CLX"))?
        } else {
            &[]
        };
        let shading_policy = table::TableShadingPolicy {
            effective_nfib: self.effective_nfib,
            interpret_table_styles,
        };
        #[cfg(feature = "direct-doc")]
        let mut native_geometry = table::NativeGeometry::default();
        for (bytes, complex_piece) in [(direct, false), (piece, true)] {
            #[cfg(feature = "direct-doc")]
            if interpret_table_styles {
                native_geometry.begin_source();
            }
            sprm::paragraph_properties(
                bytes,
                self.data,
                &mut self.budget,
                if interpret_table_styles && complex_piece {
                    sprm::TopLevelFilter::Paragraph
                } else {
                    sprm::TopLevelFilter::All
                },
                |code, operand, _| {
                    // MS-DOC 2.4.6.1 step 5 filters the top-level complex PCD
                    // before first-position indirection is evaluated. Nested
                    // PTableProps/PHugePapx data remains unfiltered.
                    if interpret_table_styles && complex_piece && (code >> 10) & 7 == 5 {
                        // Retain table facts reached through paragraph data,
                        // but fail native admission until their acquisition
                        // order relative to direct table properties is known.
                        self.unsupported_table_properties = true;
                    }
                    #[cfg(feature = "direct-doc")]
                    if interpret_table_styles {
                        properties.row.reset_row_properties_at_tistd(code);
                    }
                    #[cfg(feature = "direct-doc")]
                    if interpret_table_styles
                        && properties.row.apply_native_cant_split(code, operand)?
                    {
                        return Ok(());
                    }
                    #[cfg(feature = "direct-doc")]
                    if interpret_table_styles
                        && properties.row.apply_style_aware_margins(code, operand)?
                    {
                        return Ok(());
                    }
                    #[cfg(feature = "direct-doc")]
                    if interpret_table_styles {
                        match properties.row.apply_style_aware_borders(code, operand)? {
                            table::StyleAwareBorderApply::Handled => return Ok(()),
                            table::StyleAwareBorderApply::HandledUnsupported => {
                                self.unsupported_table_properties = true;
                                return Ok(());
                            }
                            table::StyleAwareBorderApply::Unhandled => {}
                        }
                    }
                    match properties
                        .row
                        .apply_style_aware_shading(code, operand, shading_policy)?
                    {
                        table::StyleAwareShadingApply::Handled => return Ok(()),
                        table::StyleAwareShadingApply::HandledUnsupported => {
                            self.unsupported_table_properties = true;
                            return Ok(());
                        }
                        table::StyleAwareShadingApply::Unhandled => {}
                    }
                    #[cfg(feature = "direct-doc")]
                    if interpret_table_styles {
                        match native_geometry.apply(&mut properties.row, code, operand)? {
                            table::NativeGeometryApply::Handled => return Ok(()),
                            table::NativeGeometryApply::HandledUnsupported => {
                                self.unsupported_table_properties = true;
                                return Ok(());
                            }
                            table::NativeGeometryApply::Unhandled => {}
                        }
                    }
                    if !properties.apply(code, operand)? && (code >> 10) & 7 == 5 {
                        self.unsupported_table_properties = true;
                    }
                    Ok(())
                },
            )?;
        }
        if prm & 1 == 0 {
            if let Some([a, b, value]) = table::prm0(prm) {
                properties.apply(u16::from_le_bytes([a, b]), &[value])?;
            }
        }
        Ok(properties)
    }

    fn chain(&mut self, mut id: usize, kind: u16) -> Result<Vec<usize>, String> {
        let mut result = Vec::new();
        let mut visited = BTreeSet::new();
        while id != 0xfff {
            self.budget.take()?;
            // Explicit resource policy. No recursive stack growth, cycles, or
            // exponentially expanded arrays of inherited property records.
            if result.len() >= 256 || !visited.insert(id) {
                return Err(unsupported("cyclic or excessive Word style inheritance"));
            }
            let Some(style) = self.styles.get(id).and_then(Option::as_ref) else {
                // Default Paragraph Font (10) is commonly a latent empty style.
                if (id == 10 && kind == 2) || (id == 0 && self.styles.is_empty()) {
                    break;
                }
                return Err(unsupported("Word style index references a missing style"));
            };
            if style.kind != kind {
                return Err(unsupported("Word style inheritance changes style kind"));
            }
            result.push(id);
            id = style.base;
        }
        result.reverse();
        Ok(result)
    }

    fn apply_style(&mut self, props: &mut Properties, id: usize, kind: u16) -> Result<(), String> {
        for id in self.chain(id, kind)? {
            let baseline = props.clone();
            let mut sprms = Sprms::new(self.styles[id].as_ref().expect("validated style").chpx);
            while let Some((code, operand)) = sprms.next(&mut self.budget)? {
                if !props.apply(code, operand, &baseline)? {
                    self.unsupported_character_properties = true;
                }
            }
        }
        Ok(())
    }

    fn paragraph_base(&mut self, id: usize) -> Result<Properties, String> {
        self.paragraph_base_with_table(id, None)
    }

    fn paragraph_base_with_table(
        &mut self,
        id: usize,
        table_style: Option<TableFormattingKey>,
    ) -> Result<Properties, String> {
        let key = (id, table_style);
        if let Some(value) = self.paragraph_cache.get(&key) {
            return Ok(value.clone());
        }
        let mut props = self.defaults.clone();
        self.apply_table_character_style(&mut props, table_style)?;
        self.apply_style(&mut props, id, 1)?;
        if self.paragraph_cache.len() >= MAX_TABLE_AWARE_CACHE_ENTRIES {
            self.paragraph_cache.clear();
        }
        self.paragraph_cache.insert(key, props.clone());
        Ok(props)
    }

    /// A caller caches this result for a consecutive (paragraph style, CHPX,
    /// PCD) range. Properties are not decoded or allocated once per character.
    pub fn run_xml(
        &mut self,
        paragraph_style: usize,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<String, String> {
        let props = self.run_properties(paragraph_style, fc, prm, prcs)?;
        self.byte_run_xml(&props)
    }

    fn byte_run_xml(&mut self, properties: &Properties) -> Result<String, String> {
        let (xml, omitted_language) = properties.byte_xml(&self.fonts)?;
        if omitted_language {
            self.unsupported_character_properties = true;
        }
        Ok(xml)
    }

    pub fn inline_picture_location(
        &mut self,
        style: usize,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<Option<usize>, String> {
        self.run_properties(style, fc, prm, prcs)?
            .picture
            .inline_location()
    }

    pub fn passive_special_character(
        &mut self,
        style: usize,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<bool, String> {
        Ok(self
            .run_properties(style, fc, prm, prcs)?
            .picture
            .passive_special())
    }

    fn run_properties(
        &mut self,
        paragraph_style: usize,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<Properties, String> {
        self.run_properties_with_table(paragraph_style, None, fc, prm, prcs)
    }

    fn run_properties_with_table(
        &mut self,
        paragraph_style: usize,
        table_style: Option<TableFormattingKey>,
        fc: usize,
        prm: u16,
        prcs: &[&[u8]],
    ) -> Result<Properties, String> {
        let paragraph = self.paragraph_base_with_table(paragraph_style, table_style)?;
        let mut props = paragraph.clone();
        let mut style = paragraph.clone();
        let direct = match self.characters.at(fc) {
            Some((_, run)) => run.properties,
            None if self.characters.is_empty() => &[],
            None => return Err(unsupported("Word character outside formatting ranges")),
        };
        self.apply_direct(&mut props, &mut style, &paragraph, direct)?;
        if prm & 1 != 0 {
            let bytes = prcs
                .get((prm >> 1) as usize)
                .ok_or_else(|| unsupported("Word piece property index outside CLX"))?;
            self.apply_direct(&mut props, &mut style, &paragraph, bytes)?;
        } else if let Some(bytes) = character::prm0(prm) {
            self.apply_direct(&mut props, &mut style, &paragraph, &bytes)?;
        } else if prm != 0 && paragraph::prm0(prm).is_none() && table::prm0(prm).is_none() {
            self.unsupported_piece_properties = true;
        }
        Ok(props)
    }

    fn apply_direct(
        &mut self,
        props: &mut Properties,
        style: &mut Properties,
        paragraph: &Properties,
        bytes: &[u8],
    ) -> Result<(), String> {
        let mut sprms = Sprms::new(bytes);
        while let Some((code, operand)) = sprms.next(&mut self.budget)? {
            if (code >> 10) & 7 != 2 {
                continue;
            }
            match code {
                0x4a30 => {
                    let id = u16_at(operand, 0)? as usize;
                    let mut character_style = paragraph.clone();
                    self.apply_style(&mut character_style, id, 2)?;
                    // Reset exceptions survive both the reset and style application.
                    props.reset_to(&character_style, true);
                    *style = character_style;
                }
                0x2a33 => {
                    props.reset_to(paragraph, false);
                    *style = paragraph.clone();
                }
                _ => {
                    if !props.apply(code, operand, style)? {
                        self.unsupported_character_properties = true;
                    }
                }
            }
        }
        Ok(())
    }
}

fn read_fonts(bytes: &[u8]) -> Result<Vec<String>, String> {
    if bytes.is_empty() {
        return Ok(Vec::new());
    }
    let count = u16_at(bytes, 0)? as usize;
    if count > 0x7ff0 || u16_at(bytes, 2)? != 0 {
        return Err(unsupported("invalid Word font table header"));
    }
    let mut offset = 4;
    let mut fonts = Vec::with_capacity(count);
    for _ in 0..count {
        let size = *bytes
            .get(offset)
            .ok_or_else(|| unsupported("truncated Word font record"))? as usize;
        offset += 1;
        let font = bytes
            .get(offset..offset + size)
            .ok_or_else(|| unsupported("truncated Word font data"))?;
        let name = font
            .get(39..)
            .ok_or_else(|| unsupported("short Word font data"))?;
        let units: Vec<_> = name
            .chunks_exact(2)
            .map(|b| u16::from_le_bytes([b[0], b[1]]))
            .take_while(|u| *u != 0)
            .collect();
        if units.is_empty() || units.len() * 2 + 2 > name.len() {
            return Err(unsupported("unterminated or empty Word font name"));
        }
        fonts.push(
            String::from_utf16(&units)
                .map_err(|_| unsupported("invalid Unicode Word font name"))?,
        );
        offset += size;
    }
    Ok(fonts)
}

fn read_styles(bytes: &[u8]) -> Result<(Properties, Vec<Option<Style<'_>>>), String> {
    let mut defaults = Properties::default();
    if bytes.is_empty() {
        return Ok((defaults, Vec::new()));
    }
    let header_size = u16_at(bytes, 0)? as usize;
    let header = bytes
        .get(2..2 + header_size)
        .filter(|v| v.len() >= 18)
        .ok_or_else(|| unsupported("truncated Word stylesheet header"))?;
    let count = u16_at(header, 0)? as usize;
    let base_size = u16_at(header, 2)? as usize;
    if !(15..4094).contains(&count) || ![10, 18].contains(&base_size) {
        return Err(unsupported("invalid Word stylesheet header"));
    }
    for (slot, field) in [12, 14, 16, 18].iter().enumerate() {
        if *field + 2 <= header.len() {
            let font = u16_at(header, *field)?;
            if font > i16::MAX as u16 {
                return Err(unsupported("negative default Word font index"));
            }
            defaults.fonts[slot] = Some(font as usize);
        }
    }
    let mut offset = 2 + header_size;
    let mut styles = Vec::with_capacity(count);
    for _ in 0..count {
        let size = u16_at(bytes, offset)? as usize;
        if size > i16::MAX as usize {
            return Err(unsupported("negative Word style size"));
        }
        offset += 2;
        let std = bytes
            .get(offset..offset + size)
            .ok_or_else(|| unsupported("truncated Word style definition"))?;
        offset += size + size % 2;
        if size == 0 {
            styles.push(None);
            continue;
        }
        let kind_and_base = u16_at(std, 2)?;
        let kind = kind_and_base & 15;
        let count = (u16_at(std, 4)? & 15) as usize;
        // MS-DOC 2.9.260 StdfBase requires exactly the three ordered
        // StkTableGRLPUPX members for a non-revision-marked table style.
        if kind == 3 && count != 3 {
            return Err(unsupported("invalid Word table style property set count"));
        }
        // MS-DOC 2.9.260 StdfBase: GRFSTD follows the two-byte bchUpe
        // at offset 6, in both accepted STD header sizes.
        let grfstd = u16_at(std, 8)?;
        let language_compatibility = StyleLanguageCompatibility {
            f97_lids_set: grfstd & (1 << 2) != 0,
            f_copy_lang: grfstd & (1 << 3) != 0,
        };
        let name_len = u16_at(std, base_size)? as usize;
        let mut p = base_size + 2 + name_len * 2;
        if u16_at(std, p)? != 0 {
            return Err(unsupported("unterminated Word style name"));
        }
        p += 2;
        let mut chpx = &[][..];
        let mut papx = &[][..];
        let mut table_sets = [None; 3];
        for i in 0..count {
            let n = u16_at(std, p)? as usize;
            p += 2;
            let upx = std
                .get(p..p + n)
                .ok_or_else(|| unsupported("truncated Word style properties"))?;
            // Other sets include paragraph/table properties and old revision
            // formatting. Never apply those as current character properties.
            if (kind == 1 && i == 1) || (kind == 2 && i == 0) {
                chpx = upx;
            }
            if kind == 1 && i == 0 {
                papx = upx
                    .get(2..)
                    .ok_or_else(|| unsupported("missing Word style paragraph index"))?;
            }
            if kind == 3 {
                table_sets[i] = Some(upx);
            }
            p += n + n % 2;
        }
        let table = (kind == 3).then(|| TableStylePropertySets {
            tapx: table_sets[0].expect("validated table style property set count"),
            papx: table_sets[1].expect("validated table style property set count"),
            chpx: table_sets[2].expect("validated table style property set count"),
        });
        styles.push(Some(Style {
            base: (kind_and_base >> 4) as usize,
            kind,
            chpx,
            papx,
            table,
            language_compatibility,
        }));
    }
    Ok((defaults, styles))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::table_style_condition;

    #[cfg(feature = "direct-doc")]
    fn parse_direct_fixture(ppr: &str, rpr: &str, mark_rpr: &str) -> serde_json::Value {
        use std::io::{Cursor, Write};
        use zip::write::SimpleFileOptions;

        let mark_inner = mark_rpr
            .strip_prefix("<w:rPr>")
            .and_then(|value| value.strip_suffix("</w:rPr>"))
            .expect("resolved run properties");
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr>{ppr}<w:rPr>{mark_inner}</w:rPr></w:pPr><w:r>{rpr}<w:t>x</w:t></w:r></w:p></w:body></w:document>"#
        );
        let mut bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
            zip.start_file("word/document.xml", SimpleFileOptions::default())
                .unwrap();
            zip.write_all(xml.as_bytes()).unwrap();
            zip.finish().unwrap();
        }
        let parsed: serde_json::Value =
            serde_json::from_str(&docx_parser::parse_docx_native(&bytes).unwrap()).unwrap();
        parsed["body"][0].clone()
    }

    #[test]
    fn reads_font_names_after_ffn_metadata_not_as_latin1() {
        let mut bytes = vec![1, 0, 0, 0];
        let mut font = vec![0; 39];
        for unit in "日本語 Font\0".encode_utf16() {
            font.extend(unit.to_le_bytes());
        }
        bytes.push(font.len() as u8);
        bytes.extend(font);
        assert_eq!(read_fonts(&bytes).unwrap(), ["日本語 Font"]);
        bytes.pop();
        assert!(read_fonts(&bytes).is_err());
    }

    fn empty() -> Formatting<'static> {
        Formatting {
            characters: Index::default(),
            paragraphs: Index::default(),
            fonts: vec![],
            defaults: Properties::default(),
            styles: vec![],
            paragraph_cache: BTreeMap::new(),
            paragraph_marker_styles: BTreeMap::new(),
            paragraph_layout_cache: BTreeMap::new(),
            table_style_cache: BTreeMap::new(),
            effective_nfib: 0x00c1,
            interpret_table_styles: false,
            data: &[],
            budget: Budget::default(),
            numbering: numbering::Tables::default(),
            numbering_output: numbering::output::Store::default(),
            unsupported_character_properties: false,
            unsupported_paragraph_properties: false,
            unsupported_piece_properties: false,
            missing_tables: true,
            unsupported_table_properties: false,
        }
    }

    fn stylesheet_with_style_flags(base_size: u16, flags: &[(u16, u16, u16)]) -> Vec<u8> {
        let mut header = vec![0; 18];
        header[0..2].copy_from_slice(&15u16.to_le_bytes());
        header[2..4].copy_from_slice(&base_size.to_le_bytes());
        let mut bytes = Vec::new();
        bytes.extend((header.len() as u16).to_le_bytes());
        bytes.extend(header);
        for &(base, kind, grfstd) in flags {
            let cupx = if kind == 1 { 2 } else { 1 };
            let property_bytes = if kind == 1 { 6 } else { 2 };
            let mut std = vec![0; usize::from(base_size) + 4 + property_bytes];
            std[2..4].copy_from_slice(&((base << 4) | kind).to_le_bytes());
            std[4..6].copy_from_slice(&(cupx as u16).to_le_bytes());
            let std_size = std.len() as u16;
            std[6..8].copy_from_slice(&std_size.to_le_bytes());
            std[8..10].copy_from_slice(&grfstd.to_le_bytes());
            let properties = usize::from(base_size) + 4;
            if kind == 1 {
                // StkParaGRLPUPX: a minimal two-byte istd PAPX, then empty CHPX.
                std[properties..properties + 2].copy_from_slice(&2u16.to_le_bytes());
                std[properties + 4..properties + 6].copy_from_slice(&0u16.to_le_bytes());
            } else {
                // StkCharGRLPUPX: one empty CHPX.
                std[properties..properties + 2].copy_from_slice(&0u16.to_le_bytes());
            }
            bytes.extend((std.len() as u16).to_le_bytes());
            bytes.extend(std);
        }
        for _ in flags.len()..15 {
            bytes.extend(0u16.to_le_bytes());
        }
        bytes
    }

    fn stylesheet_with_property_sets(base_size: u16, kind: u16, sets: &[&[u8]]) -> Vec<u8> {
        let mut header = vec![0; 18];
        header[0..2].copy_from_slice(&15u16.to_le_bytes());
        header[2..4].copy_from_slice(&base_size.to_le_bytes());
        let mut std = vec![0; usize::from(base_size) + 4];
        std[2..4].copy_from_slice(&((0x0fff << 4) | kind).to_le_bytes());
        std[4..6].copy_from_slice(&(sets.len() as u16).to_le_bytes());
        for set in sets {
            std.extend((set.len() as u16).to_le_bytes());
            std.extend(*set);
            if set.len() % 2 != 0 {
                std.push(0);
            }
        }
        let std_size = std.len() as u16;
        std[6..8].copy_from_slice(&std_size.to_le_bytes());

        let mut bytes = Vec::new();
        bytes.extend((header.len() as u16).to_le_bytes());
        bytes.extend(header);
        bytes.extend(std_size.to_le_bytes());
        bytes.extend(std);
        for _ in 1..15 {
            bytes.extend(0u16.to_le_bytes());
        }
        bytes
    }

    #[test]
    fn retains_table_style_property_sets_in_specified_order() {
        for base_size in [10, 18] {
            let bytes = stylesheet_with_property_sets(
                base_size,
                3,
                &[
                    &[0x04, 0x34, 1],
                    &[0x0b, 0, 0x41, 0x24, 1],
                    &[0x35, 0x08, 1],
                ],
            );
            let (_, styles) = read_styles(&bytes).unwrap();
            let style = styles[0].as_ref().unwrap();
            assert_eq!(style.kind, 3);
            assert_eq!(style.base, 0x0fff);
            let table = style.table.as_ref().unwrap();
            assert_eq!(table.tapx, [0x04, 0x34, 1]);
            assert_eq!(table.papx, [0x0b, 0, 0x41, 0x24, 1]);
            assert_eq!(table.chpx, [0x35, 0x08, 1]);
            assert!(style.papx.is_empty());
            assert!(style.chpx.is_empty());
        }
    }

    #[test]
    fn retains_empty_table_style_property_sets_without_consuming_padding() {
        let bytes = stylesheet_with_property_sets(10, 3, &[&[], &[0, 0], &[]]);
        let (_, styles) = read_styles(&bytes).unwrap();
        let style = styles[0].as_ref().unwrap();
        let table = style.table.as_ref().unwrap();
        assert!(table.tapx.is_empty());
        assert_eq!(table.papx, [0, 0]);
        assert!(table.chpx.is_empty());
    }

    #[test]
    fn tistd_selection_uses_default_table_style_for_invalid_slots() {
        fn style(kind: u16) -> Option<Style<'static>> {
            Some(Style {
                base: 0xfff,
                kind,
                chpx: &[],
                papx: &[],
                table: (kind == 3).then_some(TableStylePropertySets {
                    tapx: &[],
                    papx: &[],
                    chpx: &[],
                }),
                language_compatibility: StyleLanguageCompatibility::default(),
            })
        }

        let mut formatting = empty();
        formatting.styles.resize_with(15, || None);
        formatting.styles[3] = style(3);
        formatting.styles[4] = style(1);
        formatting.styles[11] = style(3);

        assert_eq!(formatting.resolve_table_style_id(None), None);
        assert_eq!(formatting.resolve_table_style_id(Some(3)), Some(3));
        assert_eq!(formatting.resolve_table_style_id(Some(4)), Some(11));
        assert_eq!(formatting.resolve_table_style_id(Some(5)), Some(11));
        assert_eq!(formatting.resolve_table_style_id(Some(99)), Some(11));
        formatting.styles[11] = None;
        assert_eq!(formatting.resolve_table_style_id(Some(5)), Some(11));
    }

    #[test]
    fn table_chpx_color_inheritance_composes_base_to_derived() {
        // Controlled DOC files rendered by Word isolate table-style CHPX color:
        // base red, child blue, empty child red, grandchild green, and a red
        // child overriding a green base. This records the observed composition
        // direction despite the "prepend" wording in MS-DOC 2.4.6.5. It does
        // not establish TAPX, PAPX, size, toggle, or conditional-style behavior.
        const RED: &[u8] = &[0x70, 0x68, 0xff, 0x00, 0x00, 0x00];
        const BLUE: &[u8] = &[0x70, 0x68, 0x00, 0x00, 0xff, 0x00];
        const GREEN: &[u8] = &[0x70, 0x68, 0x00, 0x80, 0x00, 0x00];
        const ORDINARY_POISON: &[u8] = &[0xff];

        fn table_style(base: usize, chpx: &'static [u8]) -> Option<Style<'static>> {
            Some(Style {
                base,
                kind: 3,
                // Deliberately distinguish ordinary CHPX from StkTableGRLPUPX
                // CHPX so this test cannot pass through the legacy alias.
                chpx: ORDINARY_POISON,
                papx: &[],
                table: Some(TableStylePropertySets {
                    tapx: &[],
                    papx: &[],
                    chpx,
                }),
                language_compatibility: StyleLanguageCompatibility::default(),
            })
        }

        let mut formatting = empty();
        formatting.styles = vec![
            table_style(0xfff, RED),   // 0: base
            table_style(0, BLUE),      // 1: derived conflict
            table_style(0, &[]),       // 2: derived missing CHPX
            table_style(1, GREEN),     // 3: grandchild conflict
            table_style(0xfff, GREEN), // 4: reverse-value base
            table_style(4, RED),       // 5: reverse-value child
        ];

        assert_eq!(formatting.chain(0, 3).unwrap(), [0]);
        assert_eq!(formatting.chain(1, 3).unwrap(), [0, 1]);
        assert_eq!(formatting.chain(2, 3).unwrap(), [0, 2]);
        assert_eq!(formatting.chain(3, 3).unwrap(), [0, 1, 3]);
        assert_eq!(formatting.chain(5, 3).unwrap(), [4, 5]);

        for (id, expected) in [
            (0, "FF0000"),
            (1, "0000FF"),
            (2, "FF0000"),
            (3, "008000"),
            (5, "FF0000"),
        ] {
            let mut properties = Properties::default();
            for style_id in formatting.chain(id, 3).unwrap() {
                let baseline = properties.clone();
                let raw = formatting.styles[style_id]
                    .as_ref()
                    .unwrap()
                    .table
                    .as_ref()
                    .unwrap()
                    .chpx;
                let mut sprms = Sprms::new(raw);
                while let Some((code, operand)) = sprms.next(&mut formatting.budget).unwrap() {
                    assert!(properties.apply(code, operand, &baseline).unwrap());
                }
            }
            assert!(
                properties
                    .xml(&[])
                    .unwrap()
                    .contains(&format!("<w:color w:val=\"{expected}\"/>")),
                "style {id} did not resolve to {expected}"
            );
        }
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn table_chpx_absolute_size_inherits_and_direct_size_wins_for_run_and_mark() {
        let mut formatting = observed_table_style_formatting();
        for (id, half_points) in [(0, 28), (1, 36), (3, 44), (4, 44), (5, 28)] {
            formatting.styles[id]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .chpx = leaked(vec![0x43, 0x4a, half_points, 0]);
        }
        formatting.styles[2]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .chpx = &[];

        for (id, expected) in [(0, 14.0), (1, 18.0), (2, 14.0), (3, 22.0), (5, 14.0)] {
            let run = formatting
                .direct_text_run(7, table_key(id), 0, 0, &[], "x".into())
                .unwrap()
                .unwrap();
            assert_eq!(run.font_size, expected, "table style {id}");
        }
        let paragraph = formatting
            .direct_paragraph(7, table_key(1), 0, 0, &[])
            .unwrap()
            .paragraph;
        assert_eq!(paragraph.default_font_size, Some(18.0));

        let direct_size = &[0x43, 0x4a, 24, 0][..];
        let run = formatting
            .direct_text_run(7, table_key(3), 0, 1, &[direct_size], "x".into())
            .unwrap()
            .unwrap();
        assert_eq!(run.font_size, 12.0);
        let paragraph = formatting
            .direct_paragraph(7, table_key(3), 0, 1, &[direct_size])
            .unwrap()
            .paragraph;
        assert_eq!(paragraph.default_font_size, Some(12.0));
        assert!(!formatting.unsupported_character_properties);
    }

    #[test]
    fn table_chpx_size_validates_bounds_and_keeps_other_size_axis_gated() {
        for half_points in [1u16, 3277] {
            let mut formatting = observed_table_style_formatting();
            formatting.styles[0]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .chpx = leaked([vec![0x43, 0x4a], half_points.to_le_bytes().to_vec()].concat());
            assert!(formatting.table_style_selector_profile(Some(0)).is_err());
        }

        let mut complex_script = observed_table_style_formatting();
        complex_script.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .chpx = &[0x61, 0x4a, 28, 0];
        complex_script
            .table_style_selector_profile(Some(0))
            .unwrap();
        assert!(complex_script.unsupported_character_properties);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn inherited_conditional_character_fields_compose_by_property() {
        let mut formatting = observed_table_style_formatting();
        let base = ccnf(
            table_style_condition::FIRST_ROW,
            &[0x70, 0x68, 0xff, 0, 0, 0, 0x43, 0x4a, 28, 0],
        );
        formatting.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .chpx = leaked(base);
        formatting.styles[1]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .chpx = leaked(ccnf(
            table_style_condition::FIRST_ROW,
            &[0x70, 0x68, 0, 0, 0xff, 0],
        ));
        let key = TableFormattingKey {
            selected_style: 1,
            matches: [
                None,
                None,
                None,
                Some(table_style_condition::FIRST_ROW),
                None,
            ],
        };
        let run = formatting
            .direct_text_run(7, Some(key), 0, 0, &[], "x".into())
            .unwrap()
            .unwrap();
        assert_eq!(run.color.as_deref(), Some("0000ff"));
        assert_eq!(run.font_size, 14.0);

        let inherited = TableFormattingKey {
            selected_style: 2,
            matches: [
                None,
                None,
                None,
                Some(table_style_condition::FIRST_ROW),
                None,
            ],
        };
        let run = formatting
            .direct_text_run(7, Some(inherited), 0, 0, &[], "x".into())
            .unwrap()
            .unwrap();
        assert_eq!(run.color.as_deref(), Some("ff0000"));
        assert_eq!(run.font_size, 14.0);
        assert!(!formatting.unsupported_character_properties);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn table_chpx_ascii_and_high_ansi_fonts_inherit_and_direct_fonts_win() {
        fn font_axes(index: u16) -> Vec<u8> {
            [
                vec![0x4f, 0x4a],
                index.to_le_bytes().to_vec(),
                vec![0x51, 0x4a],
                index.to_le_bytes().to_vec(),
            ]
            .concat()
        }

        let mut formatting = observed_table_style_formatting();
        formatting.fonts = vec![
            "Times New Roman".into(),
            "Courier New".into(),
            "Arial".into(),
        ];
        for (id, index) in [(0, 0), (1, 1), (3, 2), (4, 2), (5, 0)] {
            formatting.styles[id]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .chpx = leaked(font_axes(index));
        }
        formatting.styles[2]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .chpx = &[];

        for (id, expected) in [
            (0, "Times New Roman"),
            (1, "Courier New"),
            (2, "Times New Roman"),
            (3, "Arial"),
            (5, "Times New Roman"),
        ] {
            let run = formatting
                .direct_text_run(7, table_key(id), 0, 0, &[], "x".into())
                .unwrap()
                .unwrap();
            assert_eq!(run.font_family.as_deref(), Some(expected));
            assert_eq!(run.font_family_high_ansi.as_deref(), Some(expected));
        }

        let paragraph = formatting
            .direct_paragraph(7, table_key(1), 0, 0, &[])
            .unwrap()
            .paragraph;
        let mark = paragraph.paragraph_mark_font_facts.unwrap();
        assert_eq!(mark.font_family.as_deref(), Some("Courier New"));
        assert_eq!(mark.font_family_high_ansi.as_deref(), Some("Courier New"));

        let direct_fonts = leaked(font_axes(1));
        let run = formatting
            .direct_text_run(7, table_key(3), 0, 1, &[direct_fonts], "x".into())
            .unwrap()
            .unwrap();
        assert_eq!(run.font_family.as_deref(), Some("Courier New"));
        assert_eq!(run.font_family_high_ansi.as_deref(), Some("Courier New"));
        let paragraph = formatting
            .direct_paragraph(7, table_key(3), 0, 1, &[direct_fonts])
            .unwrap()
            .paragraph;
        let mark = paragraph.paragraph_mark_font_facts.unwrap();
        assert_eq!(mark.font_family.as_deref(), Some("Courier New"));
        assert_eq!(mark.font_family_high_ansi.as_deref(), Some("Courier New"));
        assert!(!formatting.unsupported_character_properties);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn table_chpx_fonts_validate_indices_and_keep_other_axes_and_conditions_gated() {
        let mut negative = observed_table_style_formatting();
        negative.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .chpx = &[0x4f, 0x4a, 0, 0x80];
        assert!(negative.table_style_selector_profile(Some(0)).is_err());

        let mut outside = observed_table_style_formatting();
        outside.fonts = vec!["Times New Roman".into()];
        outside.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .chpx = &[0x4f, 0x4a, 1, 0];
        assert!(outside
            .direct_text_run(7, table_key(0), 0, 0, &[], "x".into())
            .is_err());

        for code in [0x4a50u16, 0x4a5e] {
            let mut other_axis = observed_table_style_formatting();
            other_axis.styles[0]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .chpx = leaked([code.to_le_bytes().as_slice(), &[0, 0]].concat());
            other_axis.table_style_selector_profile(Some(0)).unwrap();
            assert!(other_axis.unsupported_character_properties);
        }

        let mut conditional = observed_table_style_formatting();
        conditional.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .chpx = leaked(ccnf(table_style_condition::FIRST_ROW, &[0x4f, 0x4a, 0, 0]));
        assert_eq!(
            conditional.table_style_selector_profile(Some(0)).unwrap().2,
            0
        );
        assert!(conditional.unsupported_character_properties);
    }

    fn observed_table_style(
        base: usize,
        papx: &'static [u8],
        chpx: &'static [u8],
    ) -> Option<Style<'static>> {
        Some(Style {
            base,
            kind: 3,
            // These ordinary aliases are deliberately malformed: table-style
            // projection must read only the StkTableGRLPUPX members.
            chpx: &[0xff],
            papx: &[0xff],
            table: Some(TableStylePropertySets {
                tapx: &[],
                papx,
                chpx,
            }),
            language_compatibility: StyleLanguageCompatibility::default(),
        })
    }

    fn observed_table_style_formatting() -> Formatting<'static> {
        const RED: &[u8] = &[0x42, 0x2a, 6];
        const BLUE: &[u8] = &[0x70, 0x68, 0x00, 0x00, 0xff, 0x00];
        const GREEN: &[u8] = &[0x70, 0x68, 0x00, 0x80, 0x00, 0x00];
        const BASE_LEFT: &[u8] = &[0, 0, 0x61, 0x24, 0];
        const CHILD_CENTER: &[u8] = &[1, 0, 0x61, 0x24, 1];
        const EMPTY_CHILD: &[u8] = &[2, 0];
        const GRANDCHILD_RIGHT: &[u8] = &[3, 0, 0x61, 0x24, 2];
        const REVERSE_BASE_RIGHT: &[u8] = &[4, 0, 0x61, 0x24, 2];
        const REVERSE_CHILD_LEFT: &[u8] = &[5, 0, 0x61, 0x24, 0];
        const NONDEFAULT_EMPTY_CHILD: &[u8] = &[6, 0];

        let mut formatting = empty();
        formatting.styles = vec![
            observed_table_style(0xfff, BASE_LEFT, RED),
            observed_table_style(0, CHILD_CENTER, BLUE),
            observed_table_style(0, EMPTY_CHILD, &[]),
            observed_table_style(1, GRANDCHILD_RIGHT, GREEN),
            observed_table_style(0xfff, REVERSE_BASE_RIGHT, GREEN),
            observed_table_style(4, REVERSE_CHILD_LEFT, RED),
            observed_table_style(4, NONDEFAULT_EMPTY_CHILD, &[]),
            Some(Style {
                base: 0xfff,
                kind: 1,
                chpx: &[],
                papx: &[],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            }),
            Some(Style {
                base: 0xfff,
                kind: 2,
                chpx: BLUE,
                papx: &[],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            }),
        ];
        formatting
    }

    fn table_key(id: usize) -> Option<TableFormattingKey> {
        Some(TableFormattingKey::unconditional(id))
    }

    fn leaked(bytes: Vec<u8>) -> &'static [u8] {
        Box::leak(bytes.into_boxed_slice())
    }

    fn cnf(code: u16, condition: u16, properties: &[u8]) -> Vec<u8> {
        let cb = u8::try_from(2 + properties.len()).unwrap();
        let mut bytes = Vec::from(code.to_le_bytes());
        bytes.push(cb);
        bytes.extend(condition.to_le_bytes());
        bytes.extend(properties);
        bytes
    }

    fn ccnf(condition: u16, properties: &[u8]) -> Vec<u8> {
        cnf(0xca85, condition, properties)
    }

    #[cfg(feature = "direct-doc")]
    fn table_style_shading(background: [u8; 3], pattern: u16) -> Vec<u8> {
        let mut bytes = vec![0x87, 0xd6, 10, 0, 0, 0, 255];
        bytes.extend(background);
        bytes.push(0);
        bytes.extend(pattern.to_le_bytes());
        bytes
    }

    #[cfg(feature = "direct-doc")]
    fn table_style_shading_auto() -> Vec<u8> {
        vec![0x87, 0xd6, 10, 0, 0, 0, 255, 0, 0, 0, 255, 0, 0]
    }

    #[cfg(feature = "direct-doc")]
    fn table_style_shading_nil() -> Vec<u8> {
        [vec![0x87, 0xd6, 10], vec![255; 8], vec![0, 0]].concat()
    }

    #[cfg(feature = "direct-doc")]
    fn table_style_margin(code: u16, sides: u8, unit: u8, width: u16) -> Vec<u8> {
        let [lo, hi] = width.to_le_bytes();
        test_prl(code, &[6, 0, 1, sides, unit, lo, hi])
    }

    #[cfg(feature = "direct-doc")]
    fn table_style_borders(color: [u8; 3], width: u8) -> Vec<u8> {
        let mut operand = vec![48];
        for _ in 0..6 {
            operand.extend([color[0], color[1], color[2], 0, width, 1, 0, 0]);
        }
        test_prl(0xd613, &operand)
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn unconditional_table_borders_compose_base_to_child_and_empty_inherits() {
        let mut formatting = observed_table_style_formatting();
        formatting.configure_table_styles(0x0112, true);
        formatting.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .tapx = leaked(table_style_borders([0xff, 0, 0], 8));
        formatting.styles[1]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .tapx = leaked(table_style_borders([0, 0, 0xff], 16));

        for (style, color, width) in [(0, "ff0000", 1.0), (1, "0000ff", 2.0), (2, "ff0000", 1.0)] {
            let borders = formatting.table_borders(Some(style)).unwrap();
            for value in borders {
                let spec = value.unwrap().decode().unwrap().direct_spec();
                assert_eq!(spec.color.as_deref(), Some(color));
                assert_eq!(spec.width, width);
                assert_eq!(spec.style, "single");
            }
        }
        assert!(!formatting.unsupported_table_properties);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn conditional_and_nil_table_style_borders_stay_behind_admission_gate() {
        let mut conditional = observed_table_style_formatting();
        conditional.configure_table_styles(0x0112, true);
        conditional.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .tapx = leaked(cnf(
            0xd66a,
            table_style_condition::FIRST_ROW,
            &table_style_borders([0xff, 0, 0], 8),
        ));
        assert!(conditional
            .table_borders(Some(0))
            .unwrap()
            .iter()
            .all(Option::is_none));
        assert!(conditional.unsupported_table_properties);

        let mut nil = observed_table_style_formatting();
        nil.configure_table_styles(0x0112, true);
        let mut operand = vec![48];
        for _ in 0..6 {
            operand.extend([0xff; 8]);
        }
        nil.styles[0].as_mut().unwrap().table.as_mut().unwrap().tapx =
            leaked(test_prl(0xd613, &operand));
        assert!(nil
            .table_borders(Some(0))
            .unwrap()
            .iter()
            .all(Option::is_none));
        assert!(nil.unsupported_table_properties);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn table_style_margin_overlap_is_gated_but_disjoint_sides_resolve() {
        for (style_side, expected_gate) in [(0x02, true), (0x08, false)] {
            let mut formatting = observed_table_style_formatting();
            formatting.configure_table_styles(0x0112, true);
            formatting.styles[0]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .tapx = leaked(table_style_margin(0xd634, 0x02, 3, 360));
            formatting.styles[1]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .tapx = leaked(table_style_margin(0xd63e, style_side, 3, 720));

            let (defaults, cells) = formatting.table_cell_margins(Some(1)).unwrap();
            let mut row = table::Row::default();
            row.apply(0x7621, &[0, 1, 0xe8, 3]).unwrap();
            row.resolve_style_aware_margins(defaults, cells);
            if expected_gate {
                assert_eq!(row.cells[0].margins[1], Some(720));
                assert_eq!(row.cells[0].margins[3], Some(108));
            } else {
                assert_eq!(row.cells[0].margins[1], Some(360));
                assert_eq!(row.cells[0].margins[3], Some(720));
            }
            assert_eq!(formatting.unsupported_table_properties, expected_gate);
        }
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn conditional_table_style_margin_remains_gated_and_does_not_project() {
        let mut formatting = observed_table_style_formatting();
        formatting.configure_table_styles(0x0112, true);
        formatting.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .tapx = leaked(cnf(
            0xd66a,
            table_style_condition::FIRST_ROW,
            &table_style_margin(0xd63e, 0x02, 3, 360),
        ));

        let (defaults, cells) = formatting.table_cell_margins(Some(0)).unwrap();
        let mut row = table::Row::default();
        row.apply(0x7621, &[0, 1, 0xe8, 3]).unwrap();
        row.resolve_style_aware_margins(defaults, cells);
        assert_eq!(row.cells[0].margins[1], Some(108));
        assert!(formatting.unsupported_table_properties);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn table_style_margin_profile_checks_range_unit_and_width_boundary() {
        for (code, operand) in [
            (0xd634, [6, 0, 2, 0x02, 3, 0, 0]),
            (0xd63e, [6, 0, 1, 0x02, 0, 0, 0]),
            (0xd63e, [6, 0, 1, 0x02, 3, 0xc1, 0x7b]),
        ] {
            let mut formatting = observed_table_style_formatting();
            formatting.configure_table_styles(0x0112, true);
            formatting.styles[0]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .tapx = leaked(test_prl(code, &operand));
            assert!(formatting.table_cell_margins(Some(0)).is_err());
        }

        let mut boundary = observed_table_style_formatting();
        boundary.configure_table_styles(0x0112, true);
        boundary.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .tapx = leaked(table_style_margin(0xd63e, 0x02, 3, 31_680));
        assert!(boundary.table_cell_margins(Some(0)).is_ok());
        assert!(!boundary.unsupported_table_properties);
    }

    fn conditional_color_formatting() -> Formatting<'static> {
        const RED: &[u8] = &[0x42, 0x2a, 6];
        const BLUE: &[u8] = &[0x70, 0x68, 0x00, 0x00, 0xff, 0x00];
        const GREEN: &[u8] = &[0x70, 0x68, 0x00, 0x80, 0x00, 0x00];
        const MAGENTA: &[u8] = &[0x70, 0x68, 0xff, 0x00, 0xff, 0x00];
        const CYAN: &[u8] = &[0x70, 0x68, 0x00, 0xff, 0xff, 0x00];
        const BLACK: &[u8] = &[0x70, 0x68, 0, 0, 0, 0];

        let mut chpx = RED.to_vec();
        for (condition, color) in [
            (table_style_condition::HORIZONTAL_ODD, BLUE),
            (table_style_condition::VERTICAL_ODD, GREEN),
            (table_style_condition::FIRST_COLUMN, MAGENTA),
            (table_style_condition::FIRST_ROW, CYAN),
            (table_style_condition::TOP_LEFT, BLACK),
        ] {
            chpx.extend(ccnf(condition, color));
        }
        let mut formatting = observed_table_style_formatting();
        let sets = formatting.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap();
        sets.tapx = &[0x88, 0x34, 1, 0x89, 0x34, 1];
        sets.chpx = leaked(chpx);
        formatting
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn table_style_shading_layers_unconditional_and_ordered_conditions() {
        let red = table_style_shading([255, 0, 0], 0);
        let blue = table_style_shading([0, 0, 255], 0);
        let green = table_style_shading([0, 128, 0], 0);
        let mut tapx = red;
        tapx.extend(cnf(0xd66a, table_style_condition::HORIZONTAL_ODD, &blue));
        tapx.extend(cnf(0xd66a, table_style_condition::FIRST_ROW, &green));

        let mut formatting = observed_table_style_formatting();
        formatting.configure_table_styles(0x0112, true);
        formatting.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .tapx = leaked(tapx);
        assert_eq!(
            formatting.table_style_selector_profile(Some(0)).unwrap().2,
            table_style_condition::HORIZONTAL_ODD | table_style_condition::FIRST_ROW
        );
        let key = TableFormattingKey {
            selected_style: 0,
            matches: [
                Some(table_style_condition::HORIZONTAL_ODD),
                None,
                None,
                Some(table_style_condition::FIRST_ROW),
                None,
            ],
        };
        assert!(formatting
            .table_cell_shading(Some(key))
            .unwrap()
            .unwrap()
            .xml()
            .contains("w:fill=\"008000\""));
        assert!(!formatting.unsupported_table_properties);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn table_style_shading_distinguishes_empty_nil_and_auto_descendants() {
        let cases = [
            (
                Some(table_style_shading([0, 0, 255], 0)),
                Some("0000FF"),
                true,
                false,
            ),
            (None, Some("FF0000"), true, false),
            (Some(table_style_shading_nil()), None, false, false),
            (Some(table_style_shading_auto()), None, true, false),
        ];
        for (child, expected_fill, expected_value, expected_gate) in cases {
            let mut formatting = observed_table_style_formatting();
            formatting.configure_table_styles(0x0112, true);
            formatting.styles[0]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .tapx = leaked(table_style_shading([255, 0, 0], 0));
            formatting.styles[1]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .tapx = child.map_or(&[], leaked);

            let shading = formatting.table_cell_shading(table_key(1)).unwrap();
            match expected_fill {
                Some(fill) => assert!(shading
                    .unwrap()
                    .xml()
                    .contains(&format!("w:fill=\"{fill}\""))),
                None => {
                    assert_eq!(shading.is_some(), expected_value);
                    if let Some(shading) = shading {
                        assert_eq!(shading.direct_background(), None);
                    }
                }
            }
            assert_eq!(formatting.unsupported_table_properties, expected_gate);
        }
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn table_style_shading_keeps_xml_gated_and_applies_conditional_nil_as_noop() {
        let mut xml = observed_table_style_formatting();
        xml.styles[0].as_mut().unwrap().table.as_mut().unwrap().tapx =
            leaked(table_style_shading([255, 0, 0], 0));
        xml.table_style_selector_profile(Some(0)).unwrap();
        assert!(xml.unsupported_table_properties);

        let mut conditional_nil = observed_table_style_formatting();
        conditional_nil.configure_table_styles(0x0112, true);
        let mut tapx = table_style_shading([255, 0, 0], 0);
        tapx.extend(cnf(
            0xd66a,
            table_style_condition::FIRST_ROW,
            &table_style_shading_nil(),
        ));
        conditional_nil.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .tapx = leaked(tapx);
        assert_eq!(
            conditional_nil
                .table_style_selector_profile(Some(0))
                .unwrap()
                .2,
            table_style_condition::FIRST_ROW
        );
        let key = TableFormattingKey {
            selected_style: 0,
            matches: [
                None,
                None,
                None,
                Some(table_style_condition::FIRST_ROW),
                None,
            ],
        };
        assert!(conditional_nil
            .table_cell_shading(Some(key))
            .unwrap()
            .unwrap()
            .xml()
            .contains("w:fill=\"FF0000\""));
        assert!(!conditional_nil.unsupported_table_properties);

        let mut unmapped = observed_table_style_formatting();
        unmapped.configure_table_styles(0x0112, true);
        unmapped.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .tapx = leaked(table_style_shading([255, 0, 0], 0x23));
        assert!(unmapped.table_cell_shading(table_key(0)).unwrap().is_none());
        assert!(unmapped.unsupported_table_properties);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn empty_table_condition_does_not_compete_with_present_nil_across_property_families() {
        const BLACK: &[u8] = &[0x70, 0x68, 0, 0, 0, 0];
        const GREEN: &[u8] = &[0x70, 0x68, 0, 0x80, 0, 0];
        let first_row = table_style_condition::FIRST_ROW;
        let last_row = table_style_condition::LAST_ROW;

        for (first_row_tapx, expected_presence, expected_row, expected_fill, expected_text) in [
            (None, last_row, last_row, "0000FF", "008000"),
            (Some(Vec::new()), last_row, last_row, "0000FF", "008000"),
            (
                Some(table_style_shading_nil()),
                first_row | last_row,
                first_row,
                "FF0000",
                "000000",
            ),
        ] {
            let mut formatting = observed_table_style_formatting();
            formatting.configure_table_styles(0x0112, true);

            let mut tapx = table_style_shading([255, 0, 0], 0);
            tapx.extend(cnf(0xd66a, last_row, &table_style_shading([0, 0, 255], 0)));
            if let Some(first_row_tapx) = first_row_tapx {
                tapx.extend(cnf(0xd66a, first_row, &first_row_tapx));
            }
            let mut chpx = BLACK.to_vec();
            chpx.extend(ccnf(last_row, GREEN));
            let sets = formatting.styles[0]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap();
            sets.tapx = leaked(tapx);
            sets.chpx = leaked(chpx);

            let (_, _, presence) = formatting.table_style_selector_profile(Some(0)).unwrap();
            assert_eq!(presence, expected_presence);
            assert_eq!(presence & first_row, expected_presence & first_row);

            // A singleton row with both TTlp edge flags selects the first
            // supported row condition. Empty TCnf contributes no presence;
            // authored ShdNil does, even though its shading application is a
            // no-op. Exercise the resulting key across TAPX and CHPX.
            let matches = [None, None, None, Some(expected_row), None];
            let key = formatting
                .table_formatting_key(Some(0), Some(96), matches)
                .unwrap()
                .unwrap();
            assert_eq!(key.matches, matches);
            assert!(formatting
                .table_cell_shading(Some(key))
                .unwrap()
                .unwrap()
                .xml()
                .contains(&format!("w:fill=\"{expected_fill}\"")));
            let run = formatting
                .direct_text_run(7, Some(key), 0, 0, &[], "x".into())
                .unwrap()
                .unwrap();
            assert_eq!(run.color.as_deref(), Some(expected_text));
            assert!(!formatting.unsupported_table_properties);
            assert!(!formatting.unsupported_character_properties);
        }
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn conditional_color_keys_layer_in_doc_order_before_direct_formatting_and_resets() {
        let mut formatting = conditional_color_formatting();
        assert_eq!(
            formatting.table_style_selector_profile(Some(0)).unwrap(),
            (
                Some(1),
                Some(1),
                table_style_condition::HORIZONTAL_ODD
                    | table_style_condition::VERTICAL_ODD
                    | table_style_condition::FIRST_COLUMN
                    | table_style_condition::FIRST_ROW
                    | table_style_condition::TOP_LEFT,
            )
        );
        let key = |matches| TableFormattingKey {
            selected_style: 0,
            matches,
        };
        let none = key([None; 5]);
        let horizontal = key([
            Some(table_style_condition::HORIZONTAL_ODD),
            None,
            None,
            None,
            None,
        ]);
        let vertical = key([
            None,
            Some(table_style_condition::VERTICAL_ODD),
            None,
            None,
            None,
        ]);
        let all = key([
            Some(table_style_condition::HORIZONTAL_ODD),
            Some(table_style_condition::VERTICAL_ODD),
            Some(table_style_condition::FIRST_COLUMN),
            Some(table_style_condition::FIRST_ROW),
            Some(table_style_condition::TOP_LEFT),
        ]);

        for (key, expected) in [
            (none, "ff0000"),
            (horizontal, "0000ff"),
            (vertical, "008000"),
            (all, "000000"),
        ] {
            let run = formatting
                .direct_text_run(7, Some(key), 0, 0, &[], "x".into())
                .unwrap()
                .unwrap();
            assert_eq!(run.color.as_deref(), Some(expected));
        }
        assert_eq!(formatting.table_style_cache.len(), 1);
        assert_eq!(formatting.paragraph_cache.len(), 4);

        // Direct CHPX remains last in MS-DOC 2.4.6.6.
        let direct_black = [0x70, 0x68, 0, 0, 0, 0];
        let run = formatting
            .direct_text_run(7, Some(horizontal), 0, 1, &[&direct_black], "x".into())
            .unwrap()
            .unwrap();
        assert_eq!(run.color.as_deref(), Some("000000"));

        // A character style can override the condition; CPlain resets to the
        // same conditional table-aware paragraph baseline.
        let select_blue_then_plain = [0x30, 0x4a, 8, 0, 0x33, 0x2a, 0];
        let run = formatting
            .direct_text_run(
                7,
                Some(vertical),
                0,
                1,
                &[&select_blue_then_plain],
                "x".into(),
            )
            .unwrap()
            .unwrap();
        assert_eq!(run.color.as_deref(), Some("008000"));
    }

    #[test]
    fn absent_ttlp_and_table_mark_keys_apply_only_unconditional_color() {
        let mut formatting = conditional_color_formatting();
        let absent = formatting
            .table_formatting_key(Some(0), None, [None; 5])
            .unwrap()
            .unwrap();
        let table_mark = formatting
            .table_formatting_key(Some(0), None, [None; 5])
            .unwrap()
            .unwrap();
        assert_eq!(absent, table_mark);
        assert_eq!(absent.matches, [None; 5]);
        let mut properties = Properties::default();
        formatting
            .apply_table_character_style(&mut properties, Some(absent))
            .unwrap();
        assert!(properties.xml(&[]).unwrap().contains("w:val=\"FF0000\""));
    }

    #[test]
    fn empty_edge_ccnf_has_no_supported_presence_or_missing_edge_gate() {
        let mut formatting = observed_table_style_formatting();
        let sets = formatting.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap();
        sets.tapx = &[0x89, 0x34, 1];
        sets.chpx = leaked(ccnf(table_style_condition::FIRST_COLUMN, &[]));

        assert_eq!(
            formatting.table_style_selector_profile(Some(0)).unwrap(),
            (None, Some(1), 0)
        );
        let _ = formatting
            .table_formatting_key(Some(0), Some(1 << 7), [None; 5])
            .unwrap();
        assert!(!formatting.unsupported_character_properties);
    }

    #[test]
    fn invalid_and_unsupported_conditional_properties_fail_closed() {
        let mut invalid = observed_table_style_formatting();
        invalid.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .chpx = &[0x85, 0xca, 2, 0, 0];
        assert!(invalid.table_style_selector_profile(Some(0)).is_err());

        let mut invalid_paragraph = observed_table_style_formatting();
        invalid_paragraph.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .papx = &[0, 0, 0x66, 0xc6, 2, 3, 0];
        assert!(invalid_paragraph
            .resolve_paragraph_with_table(7, table_key(0), 0, 0, &[])
            .is_err());

        let mut unsupported = observed_table_style_formatting();
        let mut chpx = ccnf(table_style_condition::FIRST_COLUMN, &[0x35, 0x08, 1]);
        chpx.extend(ccnf(
            table_style_condition::FIRST_ROW,
            &[0x85, 0xca, 2, 1, 0],
        ));
        unsupported.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .chpx = leaked(chpx);
        let _ = unsupported.table_style_selector_profile(Some(0)).unwrap();
        assert!(unsupported.unsupported_character_properties);

        let mut paragraph = observed_table_style_formatting();
        let mut papx = vec![0, 0];
        papx.extend(cnf(
            0xc666,
            table_style_condition::FIRST_ROW,
            &[0x03, 0x24, 1],
        ));
        paragraph.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .papx = leaked(papx);
        paragraph
            .resolve_paragraph_with_table(7, table_key(0), 0, 0, &[])
            .unwrap();
        assert!(paragraph.unsupported_paragraph_properties);

        let mut table = observed_table_style_formatting();
        let table_shading = [0x87, 0xd6, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
        table.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .tapx = leaked(cnf(
            0xd66a,
            table_style_condition::FIRST_ROW,
            &table_shading,
        ));
        let _ = table.table_style_selector_profile(Some(0)).unwrap();
        assert!(table.unsupported_table_properties);

        let mut invalid_band = observed_table_style_formatting();
        invalid_band.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .tapx = &[0x88, 0x34, 0];
        assert!(invalid_band.table_style_selector_profile(Some(0)).is_err());
    }

    #[test]
    fn invalid_table_tapx_never_publishes_a_partial_profile() {
        // A valid band before an invalid property must not survive a failed
        // build in the bounded profile cache.
        for invalid in [
            vec![0x35, 0x08, 1],
            vec![0x60, 0xd6, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            vec![0x87, 0xd6, 9, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ] {
            let mut formatting = observed_table_style_formatting();
            let mut bytes = vec![0x88, 0x34, 3];
            bytes.extend(invalid);
            formatting.styles[0]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .tapx = leaked(bytes);
            assert!(formatting.table_style_selector_profile(Some(0)).is_err());
            assert!(formatting.table_style_cache.is_empty());
            formatting.styles[0]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .tapx = &[0x88, 0x34, 1];
            assert_eq!(
                formatting.table_style_selector_profile(Some(0)).unwrap(),
                (Some(1), None, 0)
            );
        }
    }

    #[test]
    fn supported_conditional_inheritance_and_conflicting_bands_are_independent() {
        // Office controls establish inherited supported CCnf composition.
        // Conflicting inherited band widths remain a separate unsupported
        // TAPX property.
        let mut conditional = conditional_color_formatting();
        conditional.table_style_selector_profile(Some(1)).unwrap();
        let mut properties = Properties::default();
        conditional
            .apply_table_character_style(
                &mut properties,
                Some(TableFormattingKey {
                    selected_style: 1,
                    matches: [
                        None,
                        None,
                        None,
                        Some(table_style_condition::FIRST_ROW),
                        None,
                    ],
                }),
            )
            .unwrap();
        assert!(properties.xml(&[]).unwrap().contains("w:val=\"00FFFF\""));
        assert!(!conditional.unsupported_character_properties);

        for (child_width, unsupported) in [(1, false), (2, true)] {
            let mut formatting = observed_table_style_formatting();
            formatting.styles[0]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .tapx = &[0x88, 0x34, 1];
            formatting.styles[1]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .tapx = leaked(vec![0x88, 0x34, child_width]);
            formatting.table_style_selector_profile(Some(1)).unwrap();
            assert_eq!(formatting.unsupported_table_properties, unsupported);
        }
    }

    #[test]
    fn table_papx_alignment_uses_embedded_style_index_and_observed_descendant_priority() {
        // Word-produced LTR controls establish these base/child/grandchild and
        // reverse-value results. The test is deliberately limited to PJc and
        // does not generalize the 2.4.6.5 array wording to other PAPX.
        let mut formatting = observed_table_style_formatting();
        for (table_style, expected) in [
            (0, "left"),
            (1, "center"),
            (2, "left"),
            (3, "right"),
            (4, "right"),
            (5, "left"),
            (6, "right"),
        ] {
            let resolved = formatting
                .resolve_paragraph_with_table(7, table_key(table_style), 0, 0, &[])
                .unwrap();
            assert!(
                resolved
                    .properties
                    .xml()
                    .contains(&format!("<w:jc w:val=\"{expected}\"/>")),
                "table style {table_style} did not resolve to {expected}"
            );
        }

        formatting.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .papx = &[1, 0, 0x03, 0x24, 0];
        formatting.paragraph_layout_cache.clear();
        formatting.table_style_cache.clear();
        let error = match formatting.resolve_paragraph_with_table(7, table_key(0), 0, 0, &[]) {
            Err(error) => error,
            Ok(_) => panic!("mismatched embedded table style index must fail"),
        };
        assert!(error.contains("mismatched style index"), "{error}");
    }

    #[test]
    fn conditional_pjc_layers_in_selector_order_and_direct_formatting_wins() {
        let mut formatting = observed_table_style_formatting();
        let sets = formatting.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap();
        let mut papx = vec![0, 0, 0x61, 0x24, 0];
        for (condition, value) in [
            (table_style_condition::HORIZONTAL_ODD, 1),
            (table_style_condition::VERTICAL_ODD, 2),
            (table_style_condition::FIRST_COLUMN, 1),
            (table_style_condition::FIRST_ROW, 2),
            (table_style_condition::TOP_LEFT, 1),
        ] {
            papx.extend(cnf(0xc666, condition, &[0x61, 0x24, value]));
        }
        sets.papx = leaked(papx);

        assert_eq!(
            formatting.table_style_selector_profile(Some(0)).unwrap().2,
            table_style_condition::HORIZONTAL_ODD
                | table_style_condition::VERTICAL_ODD
                | table_style_condition::FIRST_COLUMN
                | table_style_condition::FIRST_ROW
                | table_style_condition::TOP_LEFT
        );
        let key = Some(TableFormattingKey {
            selected_style: 0,
            matches: [
                Some(table_style_condition::HORIZONTAL_ODD),
                Some(table_style_condition::VERTICAL_ODD),
                Some(table_style_condition::FIRST_COLUMN),
                Some(table_style_condition::FIRST_ROW),
                Some(table_style_condition::TOP_LEFT),
            ],
        });
        let resolved = formatting
            .resolve_paragraph_with_table(7, key, 0, 0, &[])
            .unwrap();
        assert!(resolved
            .properties
            .xml()
            .contains("<w:jc w:val=\"center\"/>"));

        let direct_right = &[0x61, 0x24, 2][..];
        let resolved = formatting
            .resolve_paragraph_with_table(7, key, 0, 1, &[direct_right])
            .unwrap();
        assert!(resolved
            .properties
            .xml()
            .contains("<w:jc w:val=\"right\"/>"));
        assert!(!formatting.unsupported_paragraph_properties);

        let direct_bidi = &[0x41, 0x24, 1][..];
        let resolved = formatting
            .resolve_paragraph_with_table(7, key, 0, 1, &[direct_bidi])
            .unwrap();
        assert!(resolved.properties.xml().contains("<w:jc w:val=\"left\"/>"));
        assert!(formatting.unsupported_paragraph_properties);
    }

    #[test]
    fn table_style_pjc80_is_gated_and_does_not_create_condition_presence() {
        let mut unconditional = observed_table_style_formatting();
        unconditional.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .papx = &[0, 0, 0x03, 0x24, 1];
        let resolved = unconditional
            .resolve_paragraph_with_table(7, table_key(0), 0, 0, &[])
            .unwrap();
        assert!(resolved.properties.xml().contains("<w:jc w:val=\"left\"/>"));
        assert!(unconditional.unsupported_paragraph_properties);

        let mut conditional = observed_table_style_formatting();
        conditional.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .papx = leaked(
            [
                &[0, 0][..],
                cnf(0xc666, table_style_condition::FIRST_ROW, &[0x03, 0x24, 1]).as_slice(),
            ]
            .concat(),
        );
        assert_eq!(
            conditional.table_style_selector_profile(Some(0)).unwrap().2,
            0
        );
        assert!(conditional.unsupported_paragraph_properties);
    }

    #[test]
    fn inherited_conditional_pjc_overrides_child_unconditional_alignment() {
        let mut formatting = observed_table_style_formatting();
        let mut papx = vec![0, 0];
        papx.extend(cnf(
            0xc666,
            table_style_condition::FIRST_ROW,
            &[0x61, 0x24, 2],
        ));
        formatting.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap()
            .papx = leaked(papx);
        let key = TableFormattingKey {
            selected_style: 1,
            matches: [
                None,
                None,
                None,
                Some(table_style_condition::FIRST_ROW),
                None,
            ],
        };
        let resolved = formatting
            .resolve_paragraph_with_table(7, Some(key), 0, 0, &[])
            .unwrap();
        assert!(resolved
            .properties
            .xml()
            .contains("<w:jc w:val=\"right\"/>"));
        assert!(!formatting.unsupported_paragraph_properties);
    }

    #[test]
    fn cached_table_alignment_profile_is_reused_across_paragraph_contexts() {
        let mut formatting = observed_table_style_formatting();
        formatting.styles.push(Some(Style {
            base: 0xfff,
            kind: 1,
            chpx: &[],
            papx: &[],
            table: None,
            language_compatibility: StyleLanguageCompatibility::default(),
        }));

        formatting
            .resolve_paragraph_with_table(7, table_key(1), 0, 0, &[])
            .unwrap();
        let before = formatting.budget.remaining();
        let resolved = formatting
            .resolve_paragraph_with_table(9, table_key(1), 0, 0, &[])
            .unwrap();
        let after = formatting.budget.remaining();

        // Only the new ordinary paragraph-style chain consumes one operation;
        // the cached table chain and its PAPX/UPX are not scanned again.
        assert_eq!(before - after, 1);
        assert!(resolved
            .properties
            .xml()
            .contains("<w:jc w:val=\"center\"/>"));
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn direct_projection_layers_observed_table_color_and_ltr_alignment_before_higher_sources() {
        let mut formatting = observed_table_style_formatting();

        // Resolve two table identities through the same paragraph style. This
        // exercises both table-aware caches through the typed projection seam.
        let base = formatting
            .direct_paragraph(7, table_key(0), 0, 0, &[])
            .unwrap()
            .paragraph;
        assert_eq!(base.alignment, "left");
        assert_eq!(base.paragraph_mark_color.as_deref(), Some("ff0000"));
        let child = formatting
            .direct_paragraph(7, table_key(1), 0, 0, &[])
            .unwrap()
            .paragraph;
        assert_eq!(child.alignment, "center");
        assert_eq!(child.paragraph_mark_color.as_deref(), Some("0000ff"));
        let empty_child = formatting
            .direct_text_run(7, table_key(2), 0, 0, &[], "x".into())
            .unwrap()
            .unwrap();
        assert_eq!(empty_child.color.as_deref(), Some("ff0000"));
        let grandchild = formatting
            .direct_text_run(7, table_key(3), 0, 0, &[], "x".into())
            .unwrap()
            .unwrap();
        assert_eq!(grandchild.color.as_deref(), Some("008000"));
        let nondefault_empty_child = formatting
            .direct_text_run(7, table_key(6), 0, 0, &[], "x".into())
            .unwrap()
            .unwrap();
        assert_eq!(nondefault_empty_child.color.as_deref(), Some("008000"));

        // MS-DOC 2.4.6.6 places direct CHPX after the table and paragraph
        // styles. This direct black Ccv is the only direct override established
        // by the Office control; no direct-alignment claim is made from it.
        let black = [0x70, 0x68, 0, 0, 0, 0];
        let direct = formatting
            .direct_text_run(7, table_key(1), 0, 1, &[&black], "x".into())
            .unwrap()
            .unwrap();
        assert_eq!(direct.color.as_deref(), Some("000000"));

        // Character styles sit above the table contribution; CPlain then
        // returns to the table-aware paragraph baseline, preserving red.
        let select_blue_then_plain = [0x30, 0x4a, 8, 0, 0x33, 0x2a, 0];
        let reset = formatting
            .direct_text_run(
                7,
                table_key(0),
                0,
                1,
                &[&select_blue_then_plain],
                "x".into(),
            )
            .unwrap()
            .unwrap();
        assert_eq!(reset.color.as_deref(), Some("ff0000"));

        // Table PAPX alignment is currently bounded to the observed LTR case.
        let rtl = [0x41, 0x24, 1];
        let paragraph = formatting
            .direct_paragraph(7, table_key(4), 0, 1, &[&rtl])
            .unwrap()
            .paragraph;
        assert_eq!(paragraph.alignment, "left");
        assert!(formatting.unsupported_paragraph_properties);

        // The ordinary paragraph style and direct PAPX retain their normative
        // positions above the table contribution.
        formatting.styles[7].as_mut().unwrap().chpx = &[0x70, 0x68, 0x80, 0x00, 0x80, 0x00];
        formatting.styles[7].as_mut().unwrap().papx = &[0x03, 0x24, 1];
        formatting.paragraph_cache.clear();
        formatting.paragraph_layout_cache.clear();
        let styled = formatting
            .direct_paragraph(7, table_key(4), 0, 0, &[])
            .unwrap()
            .paragraph;
        assert_eq!(styled.alignment, "center");
        assert_eq!(styled.paragraph_mark_color.as_deref(), Some("800080"));
        let direct_left = [0x03, 0x24, 0];
        let direct = formatting
            .direct_paragraph(7, table_key(1), 0, 1, &[&direct_left])
            .unwrap()
            .paragraph;
        assert_eq!(direct.alignment, "left");
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn unsupported_table_style_properties_keep_the_admission_gates_closed() {
        let mut formatting = observed_table_style_formatting();
        let sets = formatting.styles[0]
            .as_mut()
            .unwrap()
            .table
            .as_mut()
            .unwrap();
        sets.tapx = &[0x04, 0x34, 1];
        sets.papx = &[0, 0, 0x41, 0x24, 1];
        sets.chpx = &[0x35, 0x08, 1];

        let _ = formatting
            .direct_paragraph(7, table_key(0), 0, 0, &[])
            .unwrap();
        assert!(formatting.unsupported_table_properties);
        assert!(formatting.unsupported_paragraph_properties);
        assert!(formatting.unsupported_character_properties);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn table_aware_formatting_caches_have_a_fixed_entry_bound() {
        let mut formatting = empty();
        for id in 0..=MAX_TABLE_AWARE_CACHE_ENTRIES {
            let embedded_style = Box::leak(
                u16::try_from(id)
                    .unwrap()
                    .to_le_bytes()
                    .to_vec()
                    .into_boxed_slice(),
            );
            formatting
                .styles
                .push(observed_table_style(0xfff, embedded_style, &[]));
            if id == 0x000b {
                formatting.styles[id]
                    .as_mut()
                    .unwrap()
                    .table
                    .as_mut()
                    .unwrap()
                    .tapx = &[0x17, 0xf6, 3, 0, 0];
            }
        }
        let paragraph_style = formatting.styles.len();
        formatting.styles.push(Some(Style {
            base: 0xfff,
            kind: 1,
            chpx: &[],
            papx: &[],
            table: None,
            language_compatibility: StyleLanguageCompatibility::default(),
        }));

        for table_style in 0..=MAX_TABLE_AWARE_CACHE_ENTRIES {
            let paragraph = formatting
                .direct_paragraph(paragraph_style, table_key(table_style), 0, 0, &[])
                .unwrap()
                .paragraph;
            assert_eq!(paragraph.alignment, "left");
        }
        assert!(formatting.paragraph_cache.len() <= MAX_TABLE_AWARE_CACHE_ENTRIES);
        assert!(formatting.paragraph_layout_cache.len() <= MAX_TABLE_AWARE_CACHE_ENTRIES);
        assert!(formatting.table_style_cache.len() <= MAX_TABLE_AWARE_CACHE_ENTRIES);
    }

    #[test]
    fn retains_short_table_style_papx_losslessly_and_rejects_truncated_lpupx() {
        for papx in [&[][..], &[0x01][..]] {
            let bytes = stylesheet_with_property_sets(10, 3, &[&[0x11], papx, &[0x31]]);
            let (_, styles) = read_styles(&bytes).unwrap();
            assert_eq!(
                styles[0].as_ref().unwrap().table.as_ref().unwrap().papx,
                papx
            );
        }

        let mut bytes = stylesheet_with_property_sets(10, 3, &[&[0x11], &[0, 0], &[0x31]]);
        let papx_length = 22 + 10 + 4 + 2 + 1 + 1;
        bytes[papx_length..papx_length + 2].copy_from_slice(&u16::MAX.to_le_bytes());
        let error = read_styles(&bytes)
            .err()
            .expect("truncated LPUpxPapx must fail");
        assert!(error.contains("truncated Word style properties"), "{error}");
    }

    #[test]
    fn enforces_only_the_table_style_property_set_count() {
        for count in [0, 1, 2, 4] {
            let sets = vec![&[][..]; count];
            let bytes = stylesheet_with_property_sets(10, 3, &sets);
            let error = read_styles(&bytes)
                .err()
                .expect("invalid table cupx must fail");
            assert!(error.contains("table style property set count"), "{error}");
        }

        assert!(read_styles(&stylesheet_with_property_sets(10, 1, &[&[0, 0], &[], &[]],)).is_ok());
        assert!(read_styles(&stylesheet_with_property_sets(18, 2, &[&[], &[]])).is_ok());
    }

    #[test]
    fn paragraph_and_character_style_acquisition_remains_unchanged() {
        let paragraph = stylesheet_with_property_sets(10, 1, &[&[0x34, 0x12, 0x21], &[0x31]]);
        let (_, styles) = read_styles(&paragraph).unwrap();
        let style = styles[0].as_ref().unwrap();
        assert!(style.table.is_none());
        assert_eq!(style.papx, [0x21]);
        assert_eq!(style.chpx, [0x31]);

        let character = stylesheet_with_property_sets(18, 2, &[&[0x41]]);
        let (_, styles) = read_styles(&character).unwrap();
        let style = styles[0].as_ref().unwrap();
        assert!(style.table.is_none());
        assert!(style.papx.is_empty());
        assert_eq!(style.chpx, [0x41]);
    }

    #[test]
    fn reads_raw_style_language_compatibility_from_both_std_base_sizes() {
        for base_size in [10, 18] {
            for kind in [1, 2] {
                for bits in 0..4u16 {
                    let grfstd = 0xf000 | (bits << 2);
                    let bytes = stylesheet_with_style_flags(base_size, &[(0x0fff, kind, grfstd)]);
                    let (_, styles) = read_styles(&bytes).unwrap();
                    let style = styles[0].as_ref().unwrap();
                    assert_eq!(style.kind, kind);
                    let facts = style.language_compatibility;
                    assert_eq!(facts.f97_lids_set, bits & 1 != 0);
                    assert_eq!(facts.f_copy_lang, bits & 2 != 0);
                    assert_eq!(facts.compatibility_lids_applied(), bits & 1 != 0);
                    assert_eq!(
                        facts.copy_language(),
                        (bits & 1 != 0).then_some(bits & 2 != 0)
                    );
                }
            }
        }
    }

    #[test]
    fn style_language_compatibility_is_per_style_and_strictly_bounded() {
        for kind in [1, 2] {
            let bytes =
                stylesheet_with_style_flags(10, &[(0x0fff, kind, 0x000c), (0, kind, 0x0008)]);
            let (_, styles) = read_styles(&bytes).unwrap();
            assert_eq!(styles[0].as_ref().unwrap().base, 0x0fff);
            assert_eq!(styles[1].as_ref().unwrap().base, 0);
            assert_eq!(styles[0].as_ref().unwrap().kind, kind);
            assert_eq!(styles[1].as_ref().unwrap().kind, kind);
            let mut formatting = empty();
            formatting.styles = styles;
            assert_eq!(
                formatting
                    .style_language_compatibility(0)
                    .unwrap()
                    .copy_language(),
                Some(true)
            );
            // The derived style's raw bits are its own; inheritance is not
            // applied during acquisition, and fCopyLang is ignored while
            // f97LidsSet=false.
            let derived = formatting.style_language_compatibility(1).unwrap();
            assert!(!derived.compatibility_lids_applied());
            assert_eq!(derived.copy_language(), None);
            assert!(formatting.style_language_compatibility(14).is_none());
            assert!(formatting.style_language_compatibility(15).is_none());
        }

        for base_size in [10, 18] {
            let mut short_header = stylesheet_with_style_flags(base_size, &[]);
            short_header[0..2].copy_from_slice(&17u16.to_le_bytes());
            assert!(read_styles(&short_header).is_err());

            let mut short_std = stylesheet_with_style_flags(base_size, &[(0x0fff, 1, 0)]);
            short_std[20..22].copy_from_slice(&(base_size - 2).to_le_bytes());
            assert!(read_styles(&short_std).is_err());
        }
    }

    fn with_direct_paragraph_runs(
        runs: &[(u32, u32, Vec<u8>)],
        data: Vec<u8>,
    ) -> Formatting<'static> {
        assert!(!runs.is_empty());
        assert!(runs
            .windows(2)
            .all(|pair| pair[0].1 == pair[1].0 && pair[0].0 < pair[0].1));
        assert!(runs.last().unwrap().0 < runs.last().unwrap().1);
        let mut word = vec![0u8; 1024];
        let table: Vec<u8> = [runs[0].0, runs.last().unwrap().1, 1]
            .into_iter()
            .flat_map(u32::to_le_bytes)
            .collect();
        word[0x106..0x10a].copy_from_slice(&12u32.to_le_bytes());
        let page = &mut word[512..1024];
        for (index, (start, _, _)) in runs.iter().enumerate() {
            page[index * 4..index * 4 + 4].copy_from_slice(&start.to_le_bytes());
        }
        let count = runs.len();
        page[count * 4..count * 4 + 4].copy_from_slice(&runs.last().unwrap().1.to_le_bytes());
        let bx = (count + 1) * 4;
        let mut payload = (bx + count * 13 + 1) & !1;
        for (index, (_, _, properties)) in runs.iter().enumerate() {
            if properties.is_empty() {
                continue;
            }
            page[bx + index * 13] = (payload / 2) as u8;
            if properties.len() % 2 == 0 {
                page[payload] = 0;
                page[payload + 1] = (properties.len() / 2) as u8;
                page[payload + 2..payload + 2 + properties.len()].copy_from_slice(properties);
                payload += 2 + properties.len();
            } else {
                page[payload] = properties.len().div_ceil(2) as u8;
                page[payload + 1..payload + 1 + properties.len()].copy_from_slice(properties);
                payload += 1 + properties.len();
            }
            payload = (payload + 1) & !1;
            assert!(payload < 511);
        }
        page[511] = count as u8;
        // The index borrows its inputs, so leak this small test fixture.
        Formatting::read(
            Box::leak(word.into_boxed_slice()),
            Box::leak(table.into_boxed_slice()),
            Box::leak(data.into_boxed_slice()),
        )
        .unwrap()
    }

    fn with_direct_paragraph(properties: &[u8]) -> Formatting<'static> {
        with_direct_paragraph_runs(&[(100, 110, properties.to_vec())], Vec::new())
    }

    fn test_prl(code: u16, operand: &[u8]) -> Vec<u8> {
        [code.to_le_bytes().as_slice(), operand].concat()
    }

    fn test_prc_data(grpprl: Vec<u8>) -> Vec<u8> {
        assert!(grpprl.len() >= 10);
        let mut data = Vec::with_capacity(2 + grpprl.len());
        data.extend(u16::try_from(grpprl.len()).unwrap().to_le_bytes());
        data.extend(grpprl);
        data
    }

    fn level_bidi_formatting(value: u8) -> numbering::Tables<'static> {
        numbering::Tables {
            lists: vec![numbering::List {
                id: 42,
                styles: [0xfff; 9],
                simple: true,
                hybrid: false,
                auto_number: false,
                levels: vec![numbering::Level {
                    start: Some(1),
                    format: 0,
                    justification: 0,
                    legal: false,
                    restart: Some(0),
                    follow: 0,
                    tentative: false,
                    papx: Box::leak(vec![0x41, 0x24, value, 0x03, 0x24, 2].into_boxed_slice()),
                    chpx: &[],
                    text: &[0, 0, b'.', 0],
                    placeholders: [Some((1, 0)), None, None, None, None, None, None, None, None],
                }],
            }],
            overrides: vec![numbering::Override {
                list_index: 0,
                first_cp: None,
                auto_number_field: None,
                levels: vec![],
            }],
        }
    }

    fn level_paragraph_formatting(papx: Vec<u8>) -> numbering::Tables<'static> {
        let mut tables = level_bidi_formatting(0);
        tables.lists[0].levels[0].papx = Box::leak(papx.into_boxed_slice());
        tables
    }

    fn list_piece(ilfo: i16) -> Vec<u8> {
        [vec![0x0b, 0x46], ilfo.to_le_bytes().to_vec()].concat()
    }

    #[test]
    fn typed_paragraph_resolution_does_not_activate_numbering() {
        let piece = list_piece(1);
        let mut resolved_formatting = empty();
        resolved_formatting.numbering = level_bidi_formatting(1);
        let resolved = resolved_formatting
            .resolve_paragraph(0, 0, 1, &[&piece])
            .unwrap();
        assert_eq!(resolved.properties.numbering, None);
        let (reference, marker) = resolved.numbering.unwrap();
        assert_eq!((reference.index, reference.level), (0, 0));
        assert!(marker.xml(&[]).unwrap().starts_with("<w:rPr>"));
        assert_eq!(
            resolved_formatting.numbering_output.xml(10_000).unwrap(),
            None
        );

        let after_resolution = resolved_formatting
            .paragraph_xml(0, 0, 1, &[&piece])
            .unwrap();
        let after_numbering = resolved_formatting.numbering_output.xml(10_000).unwrap();
        let mut adapter_only = empty();
        adapter_only.numbering = level_bidi_formatting(1);
        let expected = adapter_only.paragraph_xml(0, 0, 1, &[&piece]).unwrap();
        assert_eq!(after_resolution, expected);
        assert_eq!(
            after_numbering,
            adapter_only.numbering_output.xml(10_000).unwrap()
        );
    }

    #[test]
    fn typed_resolver_and_xml_adapter_preserve_piece_reference_errors() {
        let mut typed = empty();
        let typed_error = match typed.resolve_paragraph(0, 0, 1, &[]) {
            Ok(_) => panic!("invalid piece reference unexpectedly resolved"),
            Err(error) => error,
        };
        assert!(typed_error.contains("outside CLX"), "{typed_error}");
        assert_eq!(typed.numbering_output.xml(10_000).unwrap(), None);

        let mut adapter = empty();
        let adapter_error = adapter.paragraph_xml(0, 0, 1, &[]).unwrap_err();
        assert_eq!(adapter_error, typed_error);
        assert_eq!(adapter.numbering_output.xml(10_000).unwrap(), None);
        assert_eq!(
            adapter.unsupported_paragraph_properties,
            typed.unsupported_paragraph_properties
        );
        assert_eq!(
            adapter.unsupported_piece_properties,
            typed.unsupported_piece_properties
        );

        let unsupported = [0x00, 0x24, 0];
        let mut typed = empty();
        let resolved = typed.resolve_paragraph(0, 0, 1, &[&unsupported]).unwrap();
        assert!(typed.unsupported_paragraph_properties);
        assert_eq!(typed.numbering_output.xml(10_000).unwrap(), None);

        let mut adapter = empty();
        let xml = adapter.paragraph_xml(0, 0, 1, &[&unsupported]).unwrap();
        assert!(adapter.unsupported_paragraph_properties);
        assert_eq!(xml, resolved.properties.xml());
        assert_eq!(adapter.numbering_output.xml(10_000).unwrap(), None);
    }

    #[test]
    fn typed_plain_and_suppressed_numbering_leave_output_inactive() {
        let mut typed = empty();
        let resolved = typed.resolve_paragraph(0, 0, 0, &[]).unwrap();
        assert!(resolved.numbering.is_none());
        let typed_xml = resolved.properties.xml();
        let mut adapter = empty();
        assert_eq!(adapter.paragraph_xml(0, 0, 0, &[]).unwrap(), typed_xml);

        for piece in [
            list_piece(-2047),
            [list_piece(1), vec![0x0a, 0x26, 12]].concat(),
        ] {
            let mut formatting = empty();
            let resolved = formatting.resolve_paragraph(0, 0, 1, &[&piece]).unwrap();
            assert!(resolved.numbering.is_none());
            assert_eq!(formatting.numbering_output.xml(10_000).unwrap(), None);
            let expected = resolved.properties.xml();
            assert_eq!(
                formatting.paragraph_xml(0, 0, 1, &[&piece]).unwrap(),
                expected
            );
            assert_eq!(formatting.numbering_output.xml(10_000).unwrap(), None);
        }
    }

    #[test]
    fn marker_font_validation_remains_in_xml_adapter_before_activation() {
        let piece = list_piece(1);
        let mut formatting = empty();
        formatting.numbering = level_bidi_formatting(1);
        formatting.numbering.lists[0].levels[0].chpx =
            Box::leak(vec![0x4f, 0x4a, 1, 0].into_boxed_slice());
        let resolved = formatting.resolve_paragraph(0, 0, 1, &[&piece]).unwrap();
        assert!(resolved.numbering.is_some());
        assert_eq!(formatting.numbering_output.xml(10_000).unwrap(), None);
        let error = formatting.paragraph_xml(0, 0, 1, &[&piece]).unwrap_err();
        assert!(
            error.contains("font index outside empty font table"),
            "{error}"
        );
        assert_eq!(formatting.numbering_output.xml(10_000).unwrap(), None);
    }

    #[test]
    fn byte_adapter_omits_only_unresolved_complex_script_language_and_warns() {
        let base = Properties::default();
        let mut unresolved = base.clone();
        unresolved
            .apply(0x485f, &u16::MAX.to_le_bytes(), &base)
            .unwrap();
        assert!(unresolved.xml(&[]).is_err());

        let mut formatting = empty();
        let xml = formatting.byte_run_xml(&unresolved).unwrap();
        assert!(!xml.contains("<w:lang"));
        assert!(formatting.unsupported_character_properties);

        let mut invalid_font = unresolved.clone();
        invalid_font.fonts[0] = Some(1);
        assert!(formatting
            .byte_run_xml(&invalid_font)
            .unwrap_err()
            .contains("font index outside empty font table"));

        let mut overridden = unresolved;
        overridden
            .apply(0x485f, &0x0401u16.to_le_bytes(), &base)
            .unwrap();
        let mut formatting = empty();
        let xml = formatting.byte_run_xml(&overridden).unwrap();
        assert!(xml.contains("<w:lang w:bidi=\"ar-SA\"/>"));
        assert!(!formatting.unsupported_character_properties);
    }

    #[test]
    fn compatibility_languages_are_metadata_and_modern_languages_project() {
        const COMPATIBILITY: [[u8; 4]; 2] = [[0x6d, 0x48, 0x34, 0x12], [0x6e, 0x48, 0x34, 0x12]];
        for direct in &COMPATIBILITY {
            let mut formatting = empty();
            let xml = formatting.run_xml(0, 0, 1, &[direct]).unwrap();
            assert_eq!(xml, "<w:rPr><w:sz w:val=\"20\"/></w:rPr>");
            assert!(!formatting.unsupported_character_properties);

            let mut formatting = empty();
            formatting.styles = vec![Some(Style {
                kind: 1,
                base: 0xfff,
                chpx: direct,
                papx: &[],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            })];
            let xml = formatting.run_xml(0, 0, 0, &[]).unwrap();
            assert_eq!(xml, "<w:rPr><w:sz w:val=\"20\"/></w:rPr>");
            assert!(!formatting.unsupported_character_properties);
        }

        for direct in [[0x73, 0x48, 0x0c, 0x04], [0x74, 0x48, 0x11, 0x04]] {
            let mut formatting = empty();
            let xml = formatting.run_xml(0, 0, 1, &[&direct]).unwrap();
            assert!(xml.contains(if direct[0] == 0x73 {
                "w:val=\"fr-FR\""
            } else {
                "w:eastAsia=\"ja-JP\""
            }));
            assert!(!formatting.unsupported_character_properties);
        }

        for direct in [[0x73, 0x48, 0x00, 0x04], [0x74, 0x48, 0xff, 0xff]] {
            let mut formatting = empty();
            let xml = formatting.run_xml(0, 0, 1, &[&direct]).unwrap();
            assert!(!xml.contains("<w:lang"));
            assert!(formatting.unsupported_character_properties);
        }
    }

    #[test]
    fn byte_language_policy_preserves_resolved_axes_when_one_axis_is_unresolved() {
        let base = Properties::default();
        let mut properties = base.clone();
        properties
            .apply(0x4873, &0x040cu16.to_le_bytes(), &base)
            .unwrap();
        properties
            .apply(0x4874, &u16::MAX.to_le_bytes(), &base)
            .unwrap();
        properties
            .apply(0x485f, &0x0401u16.to_le_bytes(), &base)
            .unwrap();
        properties.apply(0x0875, &[1], &base).unwrap();

        assert!(properties.xml(&[]).is_err());
        let mut formatting = empty();
        let xml = formatting.byte_run_xml(&properties).unwrap();
        assert!(xml.contains("w:val=\"fr-FR\""), "{xml}");
        assert!(!xml.contains("w:eastAsia="), "{xml}");
        assert!(xml.contains("w:bidi=\"ar-SA\""), "{xml}");
        assert!(xml.contains("<w:noProof w:val=\"1\"/>"), "{xml}");
        assert!(formatting.unsupported_character_properties);
    }

    #[test]
    fn direct_absolute_indent_axes_replay_after_list_in_last_write_order() {
        let mut f = with_direct_paragraph(&[
            0, 0, 0x41, 0x24, 1, // direct RTL
            0x0e, 0x84, 0xd0, 0x02, // physical right 720 => RTL logical left
            0x5e, 0x84, 0x68, 0x01, // later logical left 360 wins
            0x0f, 0x84, 0xf0, 0x00, // physical left 240 => RTL logical right
            0x5d, 0x84, 0xe0, 0x01, // later logical right 480 wins
        ]);
        f.numbering =
            level_paragraph_formatting(vec![0x5e, 0x84, 0xd0, 0x02, 0x5d, 0x84, 0xd0, 0x02]);
        let piece = list_piece(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        assert!(
            xml.contains("<w:ind w:left=\"360\" w:right=\"480\""),
            "{xml}"
        );
    }

    #[test]
    fn direct_zero_and_first_line_replay_but_absent_axis_inherits_list() {
        let mut f = with_direct_paragraph(&[
            0, 0, 0x0f, 0x84, 0, 0, // explicit physical-left zero
            0x60, 0x84, 0x68, 0x01, // explicit first line +360
        ]);
        f.numbering = level_paragraph_formatting(vec![
            0x5e, 0x84, 0xd0, 0x02, // list left 720
            0x5d, 0x84, 0x68, 0x01, // list right 360 (must remain: absent direct)
            0x60, 0x84, 0x98, 0xfe, // list hanging 360
        ]);
        let piece = list_piece(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        assert!(
            xml.contains("<w:ind w:left=\"0\" w:right=\"360\" w:firstLine=\"360\""),
            "{xml}"
        );
    }

    #[test]
    fn ltr_physical_logical_and_first_line_use_latest_direct_writes() {
        let mut f = with_direct_paragraph(&[
            0, 0, 0x0f, 0x84, 0xd0, 0x02, // physical left 720
            0x5e, 0x84, 0xa0, 0x05, // later logical left 1440
            0x11, 0x84, 0x68, 0x01, // first80 +360
            0x60, 0x84, 0x98, 0xfe, // later first -360
        ]);
        f.numbering = level_paragraph_formatting(vec![0x5e, 0x84, 0x68, 0x01, 0x60, 0x84, 0, 0]);
        let piece = list_piece(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        assert!(xml.contains("<w:ind w:left=\"1440\""), "{xml}");
        assert!(xml.contains("w:hanging=\"360\""), "{xml}");
    }

    #[test]
    fn pcd_absolute_indent_is_the_later_direct_layer() {
        let mut f = with_direct_paragraph(&[0, 0, 0x5e, 0x84, 0xd0, 0x02]);
        f.numbering = level_paragraph_formatting(vec![0x5e, 0x84, 0x68, 0x01]);
        let piece = [list_piece(1), vec![0x5e, 0x84, 0xa0, 0x05]].concat();
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        assert!(xml.contains("<w:ind w:left=\"1440\""), "{xml}");
    }

    #[test]
    fn repeated_same_axis_writes_never_evict_an_unchanged_opposite_axis() {
        let mut bytes = vec![
            0, 0, 0x41, 0x24, 1, // RTL
            0x0e, 0x84, 0, 0, // physical right zero => RTL logical left
        ];
        for value in 1_i16..=7 {
            bytes.extend([0x0f, 0x84]); // repeatedly replace physical left
            bytes.extend(value.to_le_bytes());
        }
        let mut f = with_direct_paragraph(&bytes);
        f.numbering =
            level_paragraph_formatting(vec![0x5e, 0x84, 0xd0, 0x02, 0x5d, 0x84, 0xd0, 0x02]);
        let piece = list_piece(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        assert!(xml.contains("<w:ind w:left=\"0\" w:right=\"7\""), "{xml}");
    }

    #[test]
    fn all_six_codes_survive_when_pcd_repeats_one_papx_code() {
        let mut f = with_direct_paragraph(&[
            0, 0, 0x0e, 0x84, 10, 0, 0x0f, 0x84, 20, 0, 0x5d, 0x84, 30, 0, 0x5e, 0x84, 40, 0, 0x11,
            0x84, 50, 0, 0x60, 0x84, 60, 0,
        ]);
        f.numbering = level_paragraph_formatting(vec![0x5e, 0x84, 0xd0, 0x02]);
        let piece = [list_piece(1), vec![0x0e, 0x84, 70, 0]].concat();
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        // PCD's repeated physical-right code moves last without dropping the
        // other five distinct codes; logical-left remains later for LTR left.
        assert!(xml.contains("<w:ind w:left=\"40\" w:right=\"70\""), "{xml}");
        assert!(xml.contains("w:firstLine=\"60\""), "{xml}");
    }

    #[test]
    fn negative_list_reference_preservation_stays_authoritative() {
        let mut f = with_direct_paragraph(&[0, 0, 0x5e, 0x84, 0xd0, 0x02, 0x60, 0x84, 0x98, 0xfe]);
        f.numbering =
            level_paragraph_formatting(vec![0x5e, 0x84, 0xa0, 0x05, 0x60, 0x84, 0x68, 0x01]);
        let piece = list_piece(-1);
        let xml = f.paragraph_xml(0, 100, 1, &[&piece]).unwrap();
        assert!(xml.contains("<w:ind w:left=\"720\""), "{xml}");
        assert!(xml.contains("w:hanging=\"360\""), "{xml}");
    }

    #[test]
    fn explicit_direct_bidi_survives_list_level_formatting() {
        let mut f = with_direct_paragraph(&[0, 0, 0x41, 0x24, 0]);
        f.numbering = level_bidi_formatting(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]]).unwrap();
        assert!(xml.contains("<w:bidi w:val=\"0\"/>"));
    }

    #[test]
    fn explicit_direct_bidi_and_physical_alignment_survive_list_level_formatting() {
        let mut f = with_direct_paragraph(&[0, 0, 0x41, 0x24, 0, 0x03, 0x24, 0]);
        f.numbering = level_bidi_formatting(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]]).unwrap();
        assert!(xml.contains("<w:bidi w:val=\"0\"/>"));
        assert!(xml.contains("<w:jc w:val=\"left\"/>"));
    }

    #[test]
    fn direct_bidi_does_not_protect_absent_alignment_from_the_list_level() {
        let mut f = with_direct_paragraph(&[0, 0, 0x41, 0x24, 0]);
        f.numbering = level_bidi_formatting(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]]).unwrap();
        assert!(xml.contains("<w:bidi w:val=\"0\"/>"));
        assert!(xml.contains("<w:jc w:val=\"right\"/>"));
    }

    #[test]
    fn direct_alignment_is_independent_and_uses_its_last_explicit_write() {
        let mut f = with_direct_paragraph(&[0, 0, 0x03, 0x24, 2, 0x61, 0x24, 0]);
        f.numbering = level_bidi_formatting(1);
        let xml = f.paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]]).unwrap();
        assert!(xml.contains("<w:bidi w:val=\"1\"/>"));
        // The later logical PJc left replaces the earlier physical PJc80 right.
        assert!(xml.contains("<w:jc w:val=\"left\"/>"));
    }

    #[test]
    fn direct_bidi_controls_are_explicit_last_write_and_piece_override() {
        let mut explicit_true = with_direct_paragraph(&[0, 0, 0x41, 0x24, 1]);
        explicit_true.numbering = level_bidi_formatting(0);
        assert!(explicit_true
            .paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]])
            .unwrap()
            .contains("<w:bidi w:val=\"1\"/>"));

        let mut absent = with_direct_paragraph(&[]);
        absent.numbering = level_bidi_formatting(1);
        assert!(absent
            .paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]])
            .unwrap()
            .contains("<w:bidi w:val=\"1\"/>"));

        let mut sequential = with_direct_paragraph(&[0, 0, 0x41, 0x24, 1, 0x41, 0x24, 0]);
        sequential.numbering = level_bidi_formatting(1);
        assert!(sequential
            .paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0]])
            .unwrap()
            .contains("<w:bidi w:val=\"0\"/>"));

        let mut piece_override = with_direct_paragraph(&[0, 0, 0x41, 0x24, 0]);
        piece_override.numbering = level_bidi_formatting(0);
        assert!(piece_override
            .paragraph_xml(0, 100, 1, &[&[0x0b, 0x46, 1, 0, 0x41, 0x24, 1]])
            .unwrap()
            .contains("<w:bidi w:val=\"1\"/>"));
    }

    #[test]
    fn paragraph_and_character_styles_resolve_before_piece_overrides() {
        let mut f = empty();
        f.styles = vec![
            Some(Style {
                kind: 1,
                base: 0xfff,
                chpx: &[0x43, 0x4a, 24, 0, 0x35, 8, 1],
                papx: &[],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            }),
            Some(Style {
                kind: 1,
                base: 0,
                chpx: &[0x43, 0x4a, 32, 0],
                papx: &[],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            }),
            Some(Style {
                kind: 2,
                base: 0xfff,
                chpx: &[0x36, 8, 1],
                papx: &[],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            }),
        ];
        let xml = f
            .run_xml(1, 0, 1, &[&[0x30, 0x4a, 2, 0, 0x35, 8, 0x81]])
            .unwrap();
        assert!(xml.contains("w:sz w:val=\"32\""));
        assert!(xml.contains("w:i w:val=\"1\""));
        assert!(xml.contains("w:b w:val=\"0\""));
        let xml = f.run_xml(1, 0, 0x0100 | (0x55 << 1), &[]).unwrap();
        assert!(xml.contains("w:sz w:val=\"32\""));
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn direct_run_and_mark_use_the_existing_style_and_piece_cascade() {
        let mut f = empty();
        f.fonts = ["ASCII", "East Asia", "High ANSI", "Complex Script"]
            .map(String::from)
            .to_vec();
        f.styles = vec![
            Some(Style {
                kind: 1,
                base: 0xfff,
                chpx: &[0x43, 0x4a, 24, 0, 0x4f, 0x4a, 0, 0, 0x35, 8, 1],
                papx: &[],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            }),
            Some(Style {
                kind: 1,
                base: 0,
                chpx: &[0x50, 0x4a, 1, 0, 0x51, 0x4a, 2, 0, 0x5e, 0x4a, 3, 0],
                papx: &[],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            }),
            Some(Style {
                kind: 2,
                base: 0xfff,
                chpx: &[0x36, 8, 1],
                papx: &[],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            }),
        ];
        let piece = [
            0x30, 0x4a, 2, 0, // character style
            0x35, 8, 0x81, // toggle inherited bold off
            0x70, 0x68, 0, 0, 0, 0xff, // auto color
        ];
        let ppr = f
            .resolve_paragraph(1, 0, 1, &[&piece])
            .unwrap()
            .properties
            .xml();
        let rpr = f.run_xml(1, 0, 1, &[&piece]).unwrap();
        let expected = parse_direct_fixture(&ppr, &rpr, &rpr);

        let direct_run = f
            .direct_text_run(1, None, 0, 1, &[&piece], "x".into())
            .unwrap()
            .unwrap();
        let mut expected_run = expected["runs"][0].clone();
        expected_run.as_object_mut().unwrap().remove("type");
        assert_eq!(serde_json::to_value(direct_run).unwrap(), expected_run);

        let direct = f.direct_paragraph(1, None, 0, 1, &[&piece]).unwrap();
        let mut expected_paragraph = expected;
        let object = expected_paragraph.as_object_mut().unwrap();
        object.remove("type");
        object.remove("styleId");
        object.insert("runs".into(), serde_json::json!([]));
        assert_eq!(
            serde_json::to_value(&direct.paragraph).unwrap(),
            expected_paragraph
        );
        assert!(direct.numbering.is_none());
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn direct_font_validation_precedes_hidden_mark_and_run_filtering() {
        let mut f = empty();
        let base = f.defaults.clone();
        assert!(f.defaults.apply(0x083c, &[1], &base).unwrap());
        assert!(f
            .defaults
            .apply(0x4a4f, &1u16.to_le_bytes(), &base)
            .unwrap());
        assert!(f.direct_paragraph(0, None, 0, 0, &[]).is_err());
        assert!(f
            .direct_text_run(0, None, 0, 0, &[], "hidden".into())
            .is_err());
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn direct_projection_uses_physical_chpx_before_piece_overrides() {
        let mut word = vec![0u8; 1024];
        let table: Vec<u8> = [100u32, 110, 1]
            .into_iter()
            .flat_map(u32::to_le_bytes)
            .collect();
        word[0xfe..0x102].copy_from_slice(&12u32.to_le_bytes());
        let page = &mut word[512..];
        page[..4].copy_from_slice(&100u32.to_le_bytes());
        page[4..8].copy_from_slice(&110u32.to_le_bytes());
        page[8] = 32;
        page[64..68].copy_from_slice(&[3, 0x35, 0x08, 1]);
        page[511] = 1;
        let mut f = Formatting::read(&word, &table, &[]).unwrap();
        assert!(
            f.direct_text_run(0, None, 100, 0, &[], "x".into())
                .unwrap()
                .unwrap()
                .bold
        );
        assert!(
            f.direct_paragraph(0, None, 109, 0, &[])
                .unwrap()
                .paragraph
                .paragraph_mark_font_facts
                .unwrap()
                .bold
        );
        let clear = [0x35, 0x08, 0];
        assert!(
            !f.direct_text_run(0, None, 100, 1, &[&clear], "x".into())
                .unwrap()
                .unwrap()
                .bold
        );
        assert!(f
            .direct_text_run(0, None, 110, 0, &[], "outside".into())
            .is_err());
        let hidden = [0x3c, 0x08, 1];
        assert!(f
            .direct_text_run(0, None, 100, 1, &[&hidden], "hidden".into())
            .unwrap()
            .is_none());
        let picture = [0x55, 0x08, 1, 0x03, 0x6a, 123, 0, 0, 0, 0x3c, 0x08, 1];
        let facts = f
            .direct_inline_picture_facts(0, None, 100, 1, &[&picture])
            .unwrap();
        assert_eq!(facts.location, Some(123));
        assert!(facts.vanish);
        let special = [0x55, 0x08, 1];
        let host = f
            .direct_anchor_host_metrics(0, None, 100, 1, &[&special])
            .unwrap()
            .unwrap();
        assert_eq!(host.font_size, 10.0);
        assert!(f
            .direct_anchor_host_metrics(0, None, 100, 1, &[&picture])
            .unwrap()
            .is_none());
        assert!(f.direct_anchor_host_metrics(0, None, 100, 0, &[]).is_err());
        assert!(
            f.direct_paragraph(0, None, 109, 1, &[&hidden])
                .unwrap()
                .paragraph
                .mark_vanish
        );
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn direct_paragraph_retains_numbering_reference_and_marker_without_activation() {
        let mut f = empty();
        f.numbering = level_bidi_formatting(1);
        f.numbering.lists[0].levels[0].chpx = &[0x35, 8, 1];
        let piece = [0x0b, 0x46, 1, 0];
        let direct = f.direct_paragraph(0, None, 0, 1, &[&piece]).unwrap();
        let (reference, marker) = direct.numbering.expect("resolved numbering");
        assert_eq!(reference.level, 0);
        assert!(
            !direct
                .paragraph
                .paragraph_mark_font_facts
                .as_ref()
                .unwrap()
                .bold
        );
        assert!(marker.direct_font_facts(&[]).unwrap().bold);
        assert!(direct.paragraph.numbering.is_none());
        assert_eq!(f.numbering_output.xml(10_000).unwrap(), None);
    }

    #[test]
    fn font_hint_resolves_style_direct_no_guidance_and_reset_semantics() {
        let mut f = empty();
        f.styles = vec![
            Some(Style {
                kind: 1,
                base: 0xfff,
                chpx: &[0x6f, 0x28, 1], // eastAsia
                papx: &[],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            }),
            Some(Style {
                kind: 1,
                base: 0,
                chpx: &[0x6f, 0x28, 0], // inherited eastAsia -> default
                papx: &[],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            }),
            Some(Style {
                kind: 2,
                base: 0xfff,
                chpx: &[0x6f, 0x28, 0], // must not replace a pre-CIstd hint
                papx: &[],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            }),
        ];

        assert!(f
            .run_xml(0, 0, 0, &[])
            .unwrap()
            .contains("<w:rFonts w:hint=\"eastAsia\"/>"));
        assert!(f
            .run_xml(1, 0, 0, &[])
            .unwrap()
            .contains("<w:rFonts w:hint=\"default\"/>"));

        // PCD/direct formatting is later than paragraph-style inheritance.
        let cs = f.run_xml(1, 0, 1, &[&[0x6f, 0x28, 2]]).unwrap();
        assert!(cs.contains("<w:rFonts w:hint=\"cs\"/>"), "{cs}");

        // Both CIstd and CPlain preserve the previous sprmCIdctHint operand.
        let reset = f
            .run_xml(
                0,
                0,
                1,
                &[&[
                    0x6f, 0x28, 2, // cs
                    0x30, 0x4a, 2, 0, // CIstd 2
                    0x33, 0x2a, 0, // CPlain
                ]],
            )
            .unwrap();
        assert!(reset.contains("<w:rFonts w:hint=\"cs\"/>"), "{reset}");

        // 0xFF is a valid explicit cancellation with no ST_Hint equivalent.
        let cancelled = f
            .run_xml(0, 0, 1, &[&[0x6f, 0x28, 0xff, 0x33, 0x2a, 0]])
            .unwrap();
        assert!(!cancelled.contains("w:hint="), "{cancelled}");
    }

    #[test]
    fn font_hint_rejects_invalid_values_and_truncated_operands() {
        let mut f = empty();
        assert!(f
            .run_xml(0, 0, 1, &[&[0x6f, 0x28, 3]])
            .unwrap_err()
            .contains("invalid Word character font hint"));
        assert!(f
            .run_xml(0, 0, 1, &[&[0x6f, 0x28]])
            .unwrap_err()
            .contains("truncated Word formatting operand"));
        assert!(!f.run_xml(0, 0, 0, &[]).unwrap().contains("w:hint="));
    }

    #[test]
    fn list_marker_style_patches_do_not_toggle_twice_or_inject_default_sizes() {
        for (chpx, bold) in [(&[][..], "1"), (&[0x35, 0x08, 0x81][..], "0")] {
            let mut f = empty();
            f.styles = vec![Some(Style {
                kind: 1,
                base: 0xfff,
                chpx: &[0x35, 0x08, 0x81],
                papx: &[],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            })];
            f.numbering = numbering::Tables {
                lists: vec![numbering::List {
                    id: 42,
                    styles: [0; 9],
                    simple: true,
                    hybrid: false,
                    auto_number: false,
                    levels: vec![numbering::Level {
                        start: Some(1),
                        format: 0,
                        justification: 0,
                        legal: false,
                        restart: Some(0),
                        follow: 0,
                        tentative: false,
                        papx: &[],
                        chpx,
                        text: &[0, 0, b'.', 0],
                        placeholders: [
                            Some((1, 0)),
                            None,
                            None,
                            None,
                            None,
                            None,
                            None,
                            None,
                            None,
                        ],
                    }],
                }],
                overrides: vec![numbering::Override {
                    list_index: 0,
                    first_cp: None,
                    auto_number_field: None,
                    levels: vec![],
                }],
            };
            let piece: &[u8] = &[0x0b, 0x46, 1, 0, 0x43, 0x4a, 40, 0];
            for _ in 0..2 {
                f.paragraph_xml(0, 0, 1, &[piece]).unwrap();
            }
            let xml = f.numbering_output.xml(10000).unwrap().unwrap();
            assert!(xml.contains(&format!("<w:b w:val=\"{bold}\"/>")));
            assert!(xml.contains("<w:sz w:val=\"40\"/>"));
            assert!(!xml.contains("<w:sz w:val=\"20\"/>"));
            assert_eq!(xml.matches("<w:num w:numId=").count(), 1);
            assert!(f
                .run_xml(0, 0, 1, &[piece])
                .unwrap()
                .contains("<w:b w:val=\"1\"/>"));
        }
    }

    #[test]
    fn list_marker_hint_patch_can_cancel_and_level_chpx_can_override_it() {
        let build = |level_chpx: &'static [u8]| {
            let mut f = empty();
            f.styles = vec![
                Some(Style {
                    kind: 1,
                    base: 0xfff,
                    chpx: &[0x6f, 0x28, 1], // paragraph inherits eastAsia
                    papx: &[],
                    table: None,
                    language_compatibility: StyleLanguageCompatibility::default(),
                }),
                Some(Style {
                    kind: 1,
                    base: 0,
                    chpx: &[0x6f, 0x28, 0xff], // linked marker style: no guidance
                    papx: &[],
                    table: None,
                    language_compatibility: StyleLanguageCompatibility::default(),
                }),
            ];
            f.numbering = numbering::Tables {
                lists: vec![numbering::List {
                    id: 42,
                    styles: [1; 9],
                    simple: true,
                    hybrid: false,
                    auto_number: false,
                    levels: vec![numbering::Level {
                        start: Some(1),
                        format: 0,
                        justification: 0,
                        legal: false,
                        restart: Some(0),
                        follow: 0,
                        tentative: false,
                        papx: &[],
                        chpx: level_chpx,
                        text: &[0, 0, b'.', 0],
                        placeholders: [
                            Some((1, 0)),
                            None,
                            None,
                            None,
                            None,
                            None,
                            None,
                            None,
                            None,
                        ],
                    }],
                }],
                overrides: vec![numbering::Override {
                    list_index: 0,
                    first_cp: None,
                    auto_number_field: None,
                    levels: vec![],
                }],
            };
            f.paragraph_xml(0, 0, 1, &[&[0x0b, 0x46, 1, 0]]).unwrap();
            f.numbering_output.xml(10000).unwrap().unwrap()
        };

        let cancelled = build(&[]);
        assert!(!cancelled.contains("w:hint="), "{cancelled}");
        let overridden = build(&[0x6f, 0x28, 2]);
        assert!(
            overridden.contains("<w:rFonts w:hint=\"cs\"/>"),
            "{overridden}"
        );
    }

    #[test]
    fn paragraph_border_style_cascade_and_piece_resets_do_not_mutate_the_cache() {
        let mut f = empty();
        f.styles = vec![
            Some(Style {
                kind: 1,
                base: 0xfff,
                chpx: &[],
                papx: &[0x24, 0x64, 8, 1, 2, 0, 0x26, 0x64, 8, 1, 2, 0],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            }),
            Some(Style {
                kind: 1,
                base: 0,
                chpx: &[],
                papx: &[0x50, 0xc6, 8, 0xff, 0, 0, 0, 16, 3, 0, 0],
                table: None,
                language_compatibility: StyleLanguageCompatibility::default(),
            }),
        ];
        let before = f.paragraph_xml(1, 0, 0, &[]).unwrap();
        assert!(before.contains("<w:top w:val=\"single\""));
        assert!(before.contains("<w:bottom w:val=\"double\""));
        let cleared = f
            .paragraph_xml(1, 0, 1, &[&[0x50, 0xc6, 8, 0, 0, 0, 0xff, 0, 0, 0, 0]])
            .unwrap();
        assert!(cleared.contains("<w:top w:val=\"single\""));
        assert!(cleared.contains("<w:bottom w:val=\"none\""));
        assert_eq!(before, f.paragraph_xml(1, 0, 0, &[]).unwrap());
        assert!(!f.unsupported_paragraph_properties);
    }

    #[test]
    fn rejects_style_cycles_and_invalid_complex_piece_references() {
        let mut f = empty();
        f.styles = vec![Some(Style {
            kind: 1,
            base: 0,
            chpx: &[],
            papx: &[],
            table: None,
            language_compatibility: StyleLanguageCompatibility::default(),
        })];
        assert!(f.run_xml(0, 0, 0, &[]).unwrap_err().contains("cyclic"));
        f.styles.clear();
        assert!(f.run_xml(0, 0, 1, &[]).unwrap_err().contains("outside CLX"));
    }

    #[test]
    fn custom_tabs_resolve_style_additions_then_direct_deletions_and_replacements() {
        let mut f = empty();
        f.styles = vec![Some(Style {
            kind: 1,
            base: 0xfff,
            chpx: &[],
            // Two tabs: 720 left/dotted, 1440 right/no leader.
            papx: &[0x0d, 0xc6, 8, 0, 2, 0xd0, 2, 0xa0, 5, 8, 2],
            table: None,
            language_compatibility: StyleLanguageCompatibility::default(),
        })];
        let original = f.paragraph_xml(0, 0, 0, &[]).unwrap();
        assert!(original.contains("<w:tab w:val=\"left\" w:pos=\"720\" w:leader=\"dot\"/>"));
        // Delete at 740 (+20 twips from inherited 720), replace 1440 with center.
        let modified = f
            .paragraph_xml(0, 0, 1, &[&[0x0d, 0xc6, 7, 1, 0xe4, 2, 1, 0xa0, 5, 1]])
            .unwrap();
        assert!(!modified.contains("w:pos=\"720\""));
        assert!(modified.contains("<w:tab w:val=\"center\" w:pos=\"1440\" w:leader=\"none\"/>"));
        assert_eq!(f.paragraph_xml(0, 0, 0, &[]).unwrap(), original);
        assert!(!f.unsupported_paragraph_properties);
    }
    #[test]
    fn inherited_paragraph_layout_is_overridden_by_piece_properties() {
        let mut f = empty();
        f.styles = vec![Some(Style {
            kind: 1,
            base: 0xfff,
            chpx: &[],
            papx: &[0x12, 0x64, 0xd4, 0xfe, 0, 0, 0x13, 0xa4, 240, 0],
            table: None,
            language_compatibility: StyleLanguageCompatibility::default(),
        })];
        let xml = f
            .paragraph_xml(0, 0, 1, &[&[0x13, 0xa4, 0, 0, 0x07, 0x24, 1]])
            .unwrap();
        assert!(xml.contains("w:line=\"300\" w:lineRule=\"exact\""));
        assert!(xml.contains("w:before=\"0\""));
        assert!(xml.contains("<w:pageBreakBefore w:val=\"1\"/>"));
        assert!(f
            .paragraph_xml(0, 0, 0, &[])
            .unwrap()
            .contains("w:before=\"240\""));
    }

    #[test]
    fn paragraph_mark_physical_offset_selects_direct_layout_before_piece_override() {
        let mut word = vec![0u8; 1024];
        // A single FKP can contain both paragraph runs; its BTE covers both.
        let table: Vec<u8> = [100u32, 120, 1]
            .into_iter()
            .flat_map(u32::to_le_bytes)
            .collect();
        word[0x106..0x10a].copy_from_slice(&12u32.to_le_bytes());
        let page = &mut word[512..1024];
        for (i, fc) in [100u32, 110, 120].into_iter().enumerate() {
            page[i * 4..i * 4 + 4].copy_from_slice(&fc.to_le_bytes());
        }
        page[12] = 32;
        page[25] = 48;
        page[64..74].copy_from_slice(&[5, 0, 0, 0x13, 0xa4, 120, 0, 0x61, 0x24, 1]);
        page[96..106].copy_from_slice(&[5, 0, 0, 0x13, 0xa4, 240, 0, 0x61, 0x24, 2]);
        page[511] = 2;
        let mut f = Formatting::read(&word, &table, &[]).unwrap();
        let first = f.paragraph_xml(0, 109, 0, &[]).unwrap();
        assert!(first.contains("w:before=\"120\""));
        assert!(first.contains("<w:jc w:val=\"center\"/>"));
        let second = f.paragraph_xml(0, 110, 1, &[&[0x13, 0xa4, 0, 0]]).unwrap();
        assert!(second.contains("w:before=\"0\""));
        assert!(second.contains("<w:jc w:val=\"right\"/>"));
        assert!(f
            .paragraph_xml(0, 120, 0, &[])
            .unwrap()
            .contains("w:before=\"0\""));
    }

    #[test]
    fn paragraph_data_indirection_replaces_tail_and_rejects_cycles() {
        let data = [10, 0, 0x12, 0x64, 0xd4, 0xfe, 0, 0, 0x13, 0xa4, 120, 0];
        let mut f = empty();
        f.data = &data;
        let mut props = paragraph::Properties::default();
        f.apply_paragraph(&mut props, &[0x46, 0x66, 0, 0, 0, 0, 0x13, 0xa4, 240, 0])
            .unwrap();
        assert!(props.xml().contains("w:before=\"120\""));
        assert!(props.xml().contains("w:line=\"300\""));
        // A non-first PHugePapx must be ignored, including an invalid pointer.
        f.apply_paragraph(&mut props, &[0x07, 0x24, 1, 0x46, 0x66, 255, 255, 255, 255])
            .unwrap();
        assert!(f
            .apply_paragraph(&mut props, &[0x46, 0x66, 255, 255, 255, 255])
            .is_err());
        let cyclic = [10, 0, 0x46, 0x66, 0, 0, 0, 0, 0x13, 0xa4, 0, 0];
        f.data = &cyclic;
        assert!(f
            .apply_paragraph(&mut props, &[0x46, 0x66, 0, 0, 0, 0])
            .unwrap_err()
            .contains("cyclic"));
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn native_complex_paragraph_filter_controls_huge_papx_first_position() {
        let data = test_prc_data(
            [
                test_prl(0x2403, &[2]),
                test_prl(0x2407, &[0]),
                test_prl(0x2407, &[0]),
                test_prl(0x2407, &[0]),
            ]
            .concat(),
        );
        let mut formatting = empty();
        formatting.data = Box::leak(data.into_boxed_slice());

        let filtered_then_huge = [
            test_prl(0xd608, &[1, 0]),
            test_prl(0x6646, &0u32.to_le_bytes()),
            test_prl(0x2403, &[1]),
        ]
        .concat();
        let mut props = paragraph::Properties::default();
        formatting
            .apply_paragraph_with_filter(
                &mut props,
                &filtered_then_huge,
                sprm::TopLevelFilter::Paragraph,
            )
            .unwrap();
        assert!(props.xml().contains("w:val=\"right\""));

        let paragraph_then_huge = [
            test_prl(0x2407, &[0]),
            test_prl(0x6646, &u32::MAX.to_le_bytes()),
            test_prl(0x2403, &[1]),
        ]
        .concat();
        let mut props = paragraph::Properties::default();
        formatting
            .apply_paragraph_with_filter(
                &mut props,
                &paragraph_then_huge,
                sprm::TopLevelFilter::Paragraph,
            )
            .unwrap();
        assert!(props.xml().contains("w:val=\"center\""));
    }

    #[test]
    fn table_properties_use_physical_papx_data_and_then_complex_piece_properties() {
        let data = test_prc_data(
            [
                test_prl(0x2416, &[1]),
                test_prl(0x2417, &[1]),
                test_prl(0x7621, &[0, 1, 0xe8, 3]),
                test_prl(0x563a, &4u16.to_le_bytes()),
                test_prl(0x740a, &[0, 0, 0x20, 0]),
                test_prl(0x3404, &[1]),
            ]
            .concat(),
        );
        let first_papx = [
            vec![0, 0],
            test_prl(0x646b, &0u32.to_le_bytes()),
            // Compatibility TDefTable for readers that ignore PTableProps.
            // Processing PTableProps replaces this tail per MS-DOC 2.6.2.
            test_prl(0xd608, &[6, 0, 1, 0, 0, 0xd0, 7]),
        ]
        .concat();
        let neighboring_papx = [
            vec![0, 0],
            test_prl(0x2416, &[1]),
            test_prl(0x2417, &[1]),
            test_prl(0x7621, &[0, 1, 0x90, 1]),
            test_prl(0x563a, &7u16.to_le_bytes()),
            test_prl(0x740a, &[0, 0, 0x40, 0]),
            test_prl(0x3404, &[1]),
        ]
        .concat();
        let mut formatting = with_direct_paragraph_runs(
            &[(100, 110, first_papx), (110, 120, neighboring_papx)],
            data,
        );

        let base = formatting.table_properties(109, 0, &[]).unwrap();
        assert!(base.in_table && base.row_end && base.row.header);
        assert_eq!(base.row.cells.len(), 1);
        assert_eq!(base.row.cells[0].width, 1000);
        assert_eq!(base.row.table_style, Some(4));
        assert_eq!(base.row.table_style_options, Some(0x20));

        let piece = [
            test_prl(0x563a, &9u16.to_le_bytes()),
            test_prl(0x740a, &[0, 0, 0x40, 3]),
            test_prl(0x7623, &[0, 1, 0xdc, 5]),
            test_prl(0x3404, &[0]),
        ]
        .concat();
        let overridden = formatting.table_properties(109, 1, &[&piece]).unwrap();
        assert_eq!(overridden.row.cells[0].width, 1500);
        assert_eq!(overridden.row.table_style, Some(9));
        assert_eq!(overridden.row.table_style_options, Some(0x0340));
        assert!(!overridden.row.header);

        let neighbor = formatting.table_properties(110, 0, &[]).unwrap();
        assert_eq!(neighbor.row.cells[0].width, 400);
        assert_eq!(neighbor.row.table_style, Some(7));
        assert_eq!(neighbor.row.table_style_options, Some(0x40));
        assert!(neighbor.row.header);
        assert!(!formatting.table_properties(120, 0, &[]).unwrap().in_table);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn native_table_acquisition_enables_raw_shading_only_after_word_2000() {
        let definition = test_prl(0xd608, &[6, 0, 1, 0, 0, 0xd0, 7]);
        let raw = test_prl(0xd670, &[10, 0, 0, 0, 255, 255, 0, 0, 0, 0, 0]);
        let papx = [vec![0, 0], definition, raw].concat();

        let mut native = with_direct_paragraph(&papx);
        native.configure_table_styles(0x00da, true);
        let properties = native.table_properties_native(109, 0, &[]).unwrap();
        assert!(matches!(
            properties.row.cells[0].prepared_shading,
            Some(table::PreparedCellShading::Explicit(_))
        ));
        assert!(!native.unsupported_table_properties);

        let mut old_native = with_direct_paragraph(&papx);
        old_native.configure_table_styles(0x00d9, true);
        let properties = old_native.table_properties_native(109, 0, &[]).unwrap();
        assert!(properties.row.cells[0].prepared_shading.is_none());
        assert!(old_native.unsupported_table_properties);

        let mut xml = with_direct_paragraph(&papx);
        xml.configure_table_styles(0x00da, false);
        let properties = xml.table_properties(109, 0, &[]).unwrap();
        assert!(properties.row.cells[0].prepared_shading.is_none());
        assert!(xml.unsupported_table_properties);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn native_complex_piece_applies_only_paragraph_sprms_until_table_props_data() {
        let papx = [vec![0, 0], test_prl(0x2416, &[1]), test_prl(0x2417, &[1])].concat();
        let definition = test_prl(0xd608, &[6, 0, 1, 0, 0, 0xd0, 7]);
        let width = test_prl(0x7623, &[0, 1, 0xdc, 5]);

        let raw_piece = [definition.clone(), width.clone()].concat();
        let mut raw = with_direct_paragraph(&papx);
        raw.configure_table_styles(0x0112, true);
        let properties = raw.table_properties_native(109, 1, &[&raw_piece]).unwrap();
        assert!(properties.row.cells.is_empty());
        assert!(!raw.unsupported_table_properties);

        let data = test_prc_data(
            [
                definition.clone(),
                width,
                test_prl(0x2407, &[0]),
                test_prl(0x2407, &[0]),
            ]
            .concat(),
        );
        let wrapped_piece = test_prl(0x646b, &0u32.to_le_bytes());
        let mut wrapped = with_direct_paragraph(&papx);
        wrapped.data = Box::leak(data.clone().into_boxed_slice());
        wrapped.configure_table_styles(0x0112, true);
        let properties = wrapped
            .table_properties_native(109, 1, &[&wrapped_piece])
            .unwrap();
        assert_eq!(properties.row.cells[0].width, 1500);
        assert!(wrapped.unsupported_table_properties);

        let nonfirst_huge = [definition, test_prl(0x6646, &0u32.to_le_bytes())].concat();
        let mut first_huge = with_direct_paragraph(&papx);
        first_huge.data = Box::leak(data.into_boxed_slice());
        first_huge.configure_table_styles(0x0112, true);
        let properties = first_huge
            .table_properties_native(109, 1, &[&nonfirst_huge])
            .unwrap();
        assert_eq!(properties.row.cells[0].width, 1500);
        assert!(first_huge.unsupported_table_properties);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn native_tistd_resets_preceding_cell_margins_without_changing_xml_mode() {
        let papx = [
            vec![0, 0],
            test_prl(0x7621, &[0, 1, 0xe8, 3]),
            test_prl(0xd632, &[6, 0, 1, 0x02, 3, 0xd0, 2]),
            test_prl(0x563a, &1u16.to_le_bytes()),
        ]
        .concat();

        let mut native = with_direct_paragraph(&papx);
        native.configure_table_styles(0x00da, true);
        let mut properties = native.table_properties_native(109, 0, &[]).unwrap();
        properties.row.resolve_style_aware_margins(
            table::MarginPatch::default(),
            table::MarginPatch::default(),
        );
        assert_eq!(properties.row.cells[0].margins[1], Some(108));
        assert_eq!(properties.row.cells[0].width, 1000);

        let mut xml = with_direct_paragraph(&papx);
        xml.configure_table_styles(0x00da, false);
        let properties = xml.table_properties(109, 0, &[]).unwrap();
        assert_eq!(properties.row.cells[0].margins[1], Some(720));
        assert_eq!(properties.row.cells[0].width, 1000);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn native_margin_acquisition_distinguishes_style_defaults_and_later_row_overrides() {
        // Native bordered controls isolate D634 by removing the later direct
        // Nil that would mask it. D63E remains above row D634 even when the
        // latter is authored after TIstd ([MS-DOC] 2.6.3).
        let cases = [
            (table_style_margin(0xd634, 0x02, 3, 288), vec![], 288),
            (table_style_margin(0xd634, 0x02, 3, 0), vec![], 0),
            (table_style_margin(0xd634, 0x02, 0, 0), vec![], 0),
            (
                table_style_margin(0xd634, 0x02, 3, 288),
                table_style_margin(0xd634, 0x02, 0, 0),
                0,
            ),
            (
                table_style_margin(0xd63e, 0x02, 3, 288),
                table_style_margin(0xd634, 0x02, 3, 72),
                288,
            ),
            (
                [
                    table_style_margin(0xd634, 0x02, 3, 288),
                    table_style_margin(0xd63e, 0x02, 3, 72),
                ]
                .concat(),
                table_style_margin(0xd634, 0x02, 0, 0),
                72,
            ),
        ];
        for (tapx, direct, expected) in cases {
            let papx = [
                vec![0, 0],
                test_prl(0x7621, &[0, 1, 0xe8, 3]),
                test_prl(0x563a, &0u16.to_le_bytes()),
                direct,
            ]
            .concat();
            let mut native = with_direct_paragraph(&papx);
            native.configure_table_styles(0x0112, true);
            native.styles = observed_table_style_formatting().styles;
            native.styles[0]
                .as_mut()
                .unwrap()
                .table
                .as_mut()
                .unwrap()
                .tapx = leaked(tapx);
            let (defaults, cells) = native.table_cell_margins(Some(0)).unwrap();
            assert!(!native.unsupported_table_properties);
            let mut properties = native.table_properties_native(109, 0, &[]).unwrap();
            properties.row.resolve_style_aware_margins(defaults, cells);
            assert_eq!(properties.row.cells[0].margins[1], Some(expected));
            assert_eq!(properties.row.cells[0].width, 1000);
            // Scalar style selection still carries the independent admission
            // prerequisite; a correct margin projection does not remove it.
            assert!(native.unsupported_table_properties);
        }
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn native_table_acquisition_resets_authored_shading_at_each_tistd() {
        let definition = test_prl(0xd608, &[6, 0, 1, 0, 0, 0xd0, 7]);
        let tistd = |style: u16| test_prl(0x563a, &style.to_le_bytes());
        let compatibility = test_prl(0xd612, &[10, 0, 0, 0, 255, 0x12, 0x34, 0x56, 0, 0, 0]);
        let raw_nil = test_prl(0xd670, &[10, 255, 255, 255, 255, 255, 255, 255, 255, 0, 0]);

        let acquire = |ordered: Vec<u8>| {
            let papx = [vec![0, 0], ordered].concat();
            let mut formatting = with_direct_paragraph(&papx);
            formatting.configure_table_styles(0x0112, true);
            let properties = formatting.table_properties_native(109, 0, &[]).unwrap();
            (properties, formatting.unsupported_table_properties)
        };

        let (before_reset, unsupported) = acquire(
            [
                definition.clone(),
                tistd(1),
                compatibility.clone(),
                raw_nil.clone(),
            ]
            .concat(),
        );
        assert!(unsupported, "the independent TIstd admission gate remains");
        assert!(matches!(
            &before_reset.row.cells[0].prepared_shading,
            Some(table::PreparedCellShading::Explicit(shading))
                if shading.xml().contains("w:fill=\"123456\"")
        ));

        let (reversed, unsupported) = acquire(
            [
                definition.clone(),
                tistd(1),
                raw_nil.clone(),
                compatibility.clone(),
            ]
            .concat(),
        );
        assert!(unsupported, "the independent TIstd admission gate remains");
        assert!(matches!(
            &reversed.row.cells[0].prepared_shading,
            Some(table::PreparedCellShading::Explicit(shading))
                if shading.xml().contains("w:fill=\"123456\"")
        ));

        let (after_reset, unsupported) = acquire(
            [
                definition,
                tistd(1),
                compatibility,
                raw_nil.clone(),
                tistd(2),
                raw_nil,
            ]
            .concat(),
        );
        assert!(unsupported, "the independent TIstd admission gate remains");
        assert_eq!(after_reset.row.table_style, Some(2));
        assert!(matches!(
            after_reset.row.cells[0].prepared_shading,
            Some(table::PreparedCellShading::StyleDeferred)
        ));
        assert_eq!(after_reset.row.cells[0].compatibility_shading, None);
        assert!(after_reset.row.cells[0].raw_nil_authored);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn native_table_acquisition_retains_d635_authored_state_across_tistd() {
        let definition = test_prl(0xd608, &[6, 0, 1, 0, 0, 0xe8, 3]);
        let preferred = test_prl(0xd635, &[5, 0, 1, 3, 0xa0, 5]);
        let nil = test_prl(0xd635, &[5, 0, 1, 0, 0, 0]);
        let tistd = |style: u16| test_prl(0x563a, &style.to_le_bytes());

        let acquire = |ordered: Vec<u8>| {
            let papx = [vec![0, 0], ordered].concat();
            let mut formatting = with_direct_paragraph(&papx);
            formatting.configure_table_styles(0x0112, true);
            formatting.table_properties_native(109, 0, &[]).unwrap().row
        };

        for ordered in [
            [definition.clone(), preferred.clone(), tistd(1)].concat(),
            [definition.clone(), tistd(1), preferred.clone()].concat(),
            [definition.clone(), preferred.clone(), tistd(1), tistd(2)].concat(),
        ] {
            let row = acquire(ordered);
            assert_eq!(row.cells[0].width, 1000);
            assert_eq!(
                row.cells[0].preferred,
                Some(table::PreferredWidth::Dxa(1440))
            );
        }

        let row = acquire([definition, preferred, tistd(1), nil].concat());
        assert_eq!(row.cells[0].width, 1000);
        assert_eq!(row.cells[0].preferred, None);
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn native_table_acquisition_preserves_tdxacol_across_same_count_tdef() {
        fn definition(boundaries: &[i16]) -> Vec<u8> {
            let count = boundaries.len() - 1;
            let mut operand = vec![0, 0, count as u8];
            for boundary in boundaries {
                operand.extend_from_slice(&boundary.to_le_bytes());
            }
            operand.resize(operand.len() + count * 20, 0);
            let cb = (operand.len() - 1) as u16;
            operand[..2].copy_from_slice(&cb.to_le_bytes());
            test_prl(0xd608, &operand)
        }
        let first = definition(&[0, 1500, 6000, 9000]);
        let tistd = test_prl(0x563a, &21u16.to_le_bytes());

        for (width, second, expected) in [
            (
                test_prl(0x7623, &[1, 2, 0xdc, 5]),
                definition(&[0, 1500, 3000, 6000]),
                [1500, 1500, 3000],
            ),
            (
                test_prl(0x7623, &[1, 2, 0xb8, 0xb]),
                definition(&[0, 2500, 4000, 9000]),
                [2500, 3000, 5000],
            ),
        ] {
            for ordered in [
                [
                    first.clone(),
                    width.clone(),
                    second.clone(),
                    test_prl(0x3615, &[0]),
                ]
                .concat(),
                [
                    first.clone(),
                    second.clone(),
                    width.clone(),
                    test_prl(0x3615, &[0]),
                ]
                .concat(),
            ] {
                let papx = [vec![0, 0], ordered.clone()].concat();
                let mut formatting = with_direct_paragraph(&papx);
                formatting.configure_table_styles(0x0112, true);
                let properties = formatting.table_properties_native(109, 0, &[]).unwrap();
                assert_eq!(
                    properties
                        .row
                        .cells
                        .iter()
                        .map(|cell| cell.width)
                        .collect::<Vec<_>>(),
                    expected
                );
                assert!(
                    !formatting.unsupported_table_properties,
                    "the proven geometry profile has its own open gate"
                );

                let papx = [vec![0, 0], [ordered, tistd.clone()].concat()].concat();
                let mut formatting = with_direct_paragraph(&papx);
                formatting.configure_table_styles(0x0112, true);
                let properties = formatting.table_properties_native(109, 0, &[]).unwrap();
                assert_eq!(properties.row.table_style, Some(21));
                assert!(
                    formatting.unsupported_table_properties,
                    "the independent TIstd admission gate remains"
                );
            }
        }

        let papx = [
            vec![0, 0],
            [
                first,
                test_prl(0x7623, &[1, 2, 0xb8, 0xb]),
                definition(&[0, 2500, 4000, 9000]),
            ]
            .concat(),
        ]
        .concat();
        let mut xml = with_direct_paragraph(&papx);
        let properties = xml.table_properties(109, 0, &[]).unwrap();
        assert_eq!(
            properties
                .row
                .cells
                .iter()
                .map(|cell| cell.width)
                .collect::<Vec<_>>(),
            [2500, 1500, 5000],
            "legacy XML acquisition retains raw Prl ordering"
        );
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn native_table_acquisition_preserves_varied_tdxacol_ranges_and_order() {
        fn definition(boundaries: &[i16]) -> Vec<u8> {
            let count = boundaries.len() - 1;
            let mut operand = vec![0, 0, count as u8];
            for boundary in boundaries {
                operand.extend_from_slice(&boundary.to_le_bytes());
            }
            operand.resize(operand.len() + count * 20, 0);
            let cb = (operand.len() - 1) as u16;
            operand[..2].copy_from_slice(&cb.to_le_bytes());
            test_prl(0xd608, &operand)
        }
        let first = definition(&[0, 1500, 6000, 9000]);
        let second = definition(&[0, 2500, 4000, 9000]);
        let width = |first: u8, limit: u8, value: u16| {
            let [lo, hi] = value.to_le_bytes();
            test_prl(0x7623, &[first, limit, lo, hi])
        };

        for (widths, expected) in [
            (vec![width(0, 1, 1200)], [1200, 1500, 5000]),
            (vec![width(2, 3, 4800)], [2500, 1500, 4800]),
            (
                vec![width(0, 1, 1200), width(2, 3, 4800)],
                [1200, 1500, 4800],
            ),
            (
                vec![width(0, 2, 2000), width(1, 3, 4000)],
                [2000, 4000, 4000],
            ),
            (
                vec![width(1, 3, 4000), width(0, 2, 2000)],
                [2000, 2000, 4000],
            ),
        ] {
            for before_second_definition in [true, false] {
                let mut ordered = vec![first.clone()];
                if before_second_definition {
                    ordered.extend(widths.clone());
                    ordered.push(second.clone());
                } else {
                    ordered.push(second.clone());
                    ordered.extend(widths.clone());
                }
                ordered.push(test_prl(0x3615, &[0]));
                let papx = [vec![0, 0], ordered.concat()].concat();
                let mut formatting = with_direct_paragraph(&papx);
                formatting.configure_table_styles(0x0112, true);
                let properties = formatting.table_properties_native(109, 0, &[]).unwrap();
                assert_eq!(
                    properties
                        .row
                        .cells
                        .iter()
                        .map(|cell| cell.width)
                        .collect::<Vec<_>>(),
                    expected
                );
                assert!(!formatting.unsupported_table_properties);
            }
        }
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn native_geometry_gate_tracks_only_competitors_to_persisted_widths() {
        fn definition() -> Vec<u8> {
            let mut operand = vec![0x46, 0, 3];
            for boundary in [0i16, 1000, 2000, 3000] {
                operand.extend(boundary.to_le_bytes());
            }
            operand.extend([0; 60]);
            test_prl(0xd608, &operand)
        }
        let width = test_prl(0x7623, &[0, 1, 0xd0, 7]);
        let merge = test_prl(0x5624, &[0, 2]);
        let autofit = test_prl(0x3615, &[1]);
        let bidi = test_prl(0x560b, &1u16.to_le_bytes());
        let preferred = test_prl(0xd635, &[5, 0, 1, 3, 100, 0]);
        let vertical_merge = test_prl(0xd62b, &[2, 0, 3]);
        let malformed_vertical_merge = test_prl(0xd62b, &[1, 0]);
        let vertical_center = test_prl(0xd62c, &[3, 0, 1, 1]);
        let vertical_top = test_prl(0xd62c, &[3, 0, 1, 0]);
        let empty_vertical_center = test_prl(0xd62c, &[3, 1, 1, 1]);
        let malformed_vertical_alignment = test_prl(0xd62c, &[2, 0, 1]);
        let acquire = |ordered: Vec<u8>| {
            let papx = [vec![0, 0], ordered].concat();
            let mut formatting = with_direct_paragraph(&papx);
            formatting.configure_table_styles(0x0112, true);
            formatting.table_properties_native(109, 0, &[]).unwrap();
            formatting.unsupported_table_properties
        };

        assert!(!acquire(
            [definition(), width.clone(), merge.clone()].concat()
        ));
        assert!(!acquire(
            [definition(), preferred.clone(), definition()].concat()
        ));
        assert!(!acquire([definition(), vertical_merge.clone()].concat()));
        assert!(acquire([definition(), malformed_vertical_merge].concat()));
        assert!(!acquire(
            [
                definition(),
                width.clone(),
                vertical_top.clone(),
                definition()
            ]
            .concat()
        ));
        assert!(!acquire(
            [
                definition(),
                width.clone(),
                empty_vertical_center,
                definition()
            ]
            .concat()
        ));
        assert!(acquire(
            [definition(), malformed_vertical_alignment].concat()
        ));
        assert!(acquire(
            [
                definition(),
                vertical_center.clone(),
                width.clone(),
                definition(),
            ]
            .concat()
        ));
        assert!(acquire(
            [
                definition(),
                width.clone(),
                vertical_center.clone(),
                definition(),
            ]
            .concat()
        ));
        assert!(acquire(
            [definition(), width.clone(), preferred.clone(), definition()].concat()
        ));
        assert!(acquire(
            [bidi.clone(), definition(), width.clone(), definition()].concat()
        ));
        for late in [merge, autofit, bidi, preferred, vertical_center] {
            assert!(acquire(
                [definition(), width.clone(), definition(), late].concat()
            ));
        }
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn native_tistd_resets_only_independently_authored_row_properties_in_data_order() {
        let before_tistd = [
            test_prl(0x2416, &[1]),
            test_prl(0x2417, &[1]),
            test_prl(0x7621, &[0, 1, 0xe8, 3]),
            test_prl(0x548a, &2u16.to_le_bytes()),
            test_prl(0x3404, &[1]),
            test_prl(0x3466, &[1]),
            test_prl(0x3465, &[1]),
            test_prl(0x360d, &[0x20]),
            test_prl(0x940e, &721u16.to_le_bytes()),
            test_prl(0x9410, &120u16.to_le_bytes()),
            test_prl(0x9601, &720u16.to_le_bytes()),
            test_prl(0x9602, &180u16.to_le_bytes()),
            test_prl(0x9407, &720u16.to_le_bytes()),
            test_prl(0xf614, &[3, 0x70, 0x17]),
            test_prl(0x3615, &[1]),
            test_prl(0x740a, &[0, 0, 0x20, 0]),
            test_prl(0x560b, &1u16.to_le_bytes()),
            test_prl(0x5664, &0u16.to_le_bytes()),
        ]
        .concat();
        let reset = test_prl(0x563a, &1u16.to_le_bytes());
        let data = test_prc_data([before_tistd.clone(), reset.clone()].concat());
        let papx = [vec![0, 0], test_prl(0x646b, &0u32.to_le_bytes())].concat();

        let mut native = with_direct_paragraph_runs(&[(100, 110, papx.clone())], data.clone());
        native.configure_table_styles(0x0112, true);
        let properties = native.table_properties_native(109, 0, &[]).unwrap();
        assert_eq!(properties.row.alignment, (0, false));
        assert!(!properties.row.header);
        assert!(!properties.row.cant_split);
        let (position, overlap) = properties.row.position.direct();
        let position = position.unwrap();
        assert_eq!(position.tblp_x, 36.0);
        assert_eq!(position.left_from_text, 120.0 / 20.0);
        assert_eq!(overlap, None);
        assert_eq!(properties.row.left, 720);
        assert_eq!(properties.row.gap, 180);
        assert_eq!(properties.row.height, 720);
        assert_eq!(
            properties.row.preferred_width,
            Some(table::PreferredWidth::Dxa(6000))
        );
        assert!(properties.row.autofit);
        assert_eq!(properties.row.table_style_options, Some(0x20));
        assert_eq!(properties.row.cells[0].width, 1000);
        assert!(properties.row.bidi);
        assert!(!properties.row.identity.contains_key(&0x3465));
        assert!(native.unsupported_table_properties);

        let mut xml = with_direct_paragraph_runs(&[(100, 110, papx.clone())], data);
        xml.configure_table_styles(0x0112, false);
        let properties = xml.table_properties(109, 0, &[]).unwrap();
        assert_eq!(properties.row.alignment, (2, false));
        assert!(properties.row.header);
        assert!(properties.row.cant_split);
        assert!(properties.row.position.xml().contains("tblOverlap"));
        assert!(properties.row.identity.contains_key(&0x3465));

        let after_tistd = [
            reset.clone(),
            test_prl(0x548a, &1u16.to_le_bytes()),
            test_prl(0x3404, &[1]),
            test_prl(0x3466, &[1]),
            test_prl(0x3465, &[1]),
        ]
        .concat();
        let data = test_prc_data(after_tistd);
        let mut native = with_direct_paragraph_runs(&[(100, 110, papx.clone())], data);
        native.configure_table_styles(0x0112, true);
        let properties = native.table_properties_native(109, 0, &[]).unwrap();
        assert_eq!(properties.row.alignment, (1, false));
        assert!(properties.row.header);
        assert!(properties.row.cant_split);
        assert!(properties.row.position.xml().contains("tblOverlap"));

        let repeated = test_prc_data(
            [
                before_tistd,
                reset.clone(),
                test_prl(0x548a, &1u16.to_le_bytes()),
                test_prl(0x3404, &[1]),
                test_prl(0x3466, &[1]),
                test_prl(0x3465, &[1]),
                reset,
            ]
            .concat(),
        );
        let mut native = with_direct_paragraph_runs(&[(100, 110, papx)], repeated);
        native.configure_table_styles(0x0112, true);
        let properties = native.table_properties_native(109, 0, &[]).unwrap();
        assert_eq!(properties.row.alignment, (0, false));
        assert!(!properties.row.header);
        assert!(!properties.row.cant_split);
        assert!(!properties.row.position.xml().contains("tblOverlap"));
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn native_current_word_cant_split_policy_is_ordered_and_not_selected_by_nfib() {
        fn native_value(papx: &[u8], effective_nfib: u16) -> bool {
            let mut formatting = with_direct_paragraph(papx);
            formatting.configure_table_styles(effective_nfib, true);
            formatting
                .table_properties_native(109, 0, &[])
                .unwrap()
                .row
                .cant_split
        }

        let modern_then_legacy =
            [vec![0, 0], test_prl(0x3466, &[1]), test_prl(0x3403, &[0])].concat();
        assert!(native_value(&modern_then_legacy, 0x00d9));
        assert!(native_value(&modern_then_legacy, 0x0112));

        let mut xml = with_direct_paragraph(&modern_then_legacy);
        xml.configure_table_styles(0x0112, false);
        assert!(!xml.table_properties(109, 0, &[]).unwrap().row.cant_split);

        let repeated_reset = [
            vec![0, 0],
            test_prl(0x3466, &[1]),
            test_prl(0x563a, &1u16.to_le_bytes()),
            test_prl(0x3403, &[1]),
            test_prl(0x3466, &[1]),
            test_prl(0x563a, &2u16.to_le_bytes()),
            test_prl(0x3403, &[1]),
        ]
        .concat();
        assert!(!native_value(&repeated_reset, 0x0112));

        let modern_after_reset = [repeated_reset, test_prl(0x3466, &[1])].concat();
        assert!(native_value(&modern_after_reset, 0x0112));
    }

    #[cfg(feature = "direct-doc")]
    #[test]
    fn table_acquisition_normalizes_explicit_overlap_false_to_omission() {
        fn acquired(overlap: Option<u8>, native: bool) -> (bool, bool) {
            let mut papx = vec![0, 0];
            if let Some(value) = overlap {
                papx.extend(test_prl(0x3465, &[value]));
            }
            let mut formatting = with_direct_paragraph(&papx);
            formatting.configure_table_styles(0x0112, native);
            let properties = if native {
                formatting.table_properties_native(109, 0, &[])
            } else {
                formatting.table_properties(109, 0, &[])
            }
            .unwrap();
            (
                properties.row.identity.contains_key(&0x3465),
                properties.row.position.xml().contains("tblOverlap"),
            )
        }

        for native in [false, true] {
            assert_eq!(acquired(None, native), (false, false));
            assert_eq!(acquired(Some(0), native), (false, false));
            assert_eq!(acquired(Some(1), native), (true, true));
        }
    }

    #[test]
    fn table_properties_follow_mixed_data_indirection_and_reject_bad_records() {
        // A: apply in-table, ignore the non-first PHugePapx, then follow
        // PTableProps to B. B: follow its first PHugePapx to C and ignore its
        // tail. Every complete PrcData has cbGrpprl >= 10.
        let offset_b = 17u32;
        let offset_c = 29u32;
        let record_a = test_prc_data(
            [
                test_prl(0x2416, &[1]),
                test_prl(0x6646, &u32::MAX.to_le_bytes()),
                test_prl(0x646b, &offset_b.to_le_bytes()),
            ]
            .concat(),
        );
        let record_b = test_prc_data(
            [
                test_prl(0x6646, &offset_c.to_le_bytes()),
                test_prl(0x563a, &99u16.to_le_bytes()),
            ]
            .concat(),
        );
        let record_c = test_prc_data(
            [
                test_prl(0x2417, &[1]),
                test_prl(0x7621, &[0, 1, 0xe8, 3]),
                test_prl(0x3404, &[1]),
            ]
            .concat(),
        );
        assert_eq!(record_a.len(), offset_b as usize);
        assert_eq!(record_a.len() + record_b.len(), offset_c as usize);
        let data = [record_a, record_b, record_c].concat();
        let papx = [vec![0, 0], test_prl(0x646b, &0u32.to_le_bytes())].concat();
        let mut formatting = with_direct_paragraph_runs(&[(100, 110, papx.clone())], data);
        let properties = formatting.table_properties(109, 0, &[]).unwrap();
        assert!(properties.in_table && properties.row_end && properties.row.header);
        assert_eq!(properties.row.cells[0].width, 1000);
        assert_eq!(properties.row.table_style, None);

        let cycle = test_prc_data(
            [
                test_prl(0x646b, &0u32.to_le_bytes()),
                test_prl(0x563a, &1u16.to_le_bytes()),
            ]
            .concat(),
        );
        let mut formatting = with_direct_paragraph_runs(&[(100, 110, papx.clone())], cycle);
        assert!(formatting
            .table_properties(109, 0, &[])
            .err()
            .unwrap()
            .contains("cyclic"));

        let mut truncated = 10u16.to_le_bytes().to_vec();
        truncated.extend([0u8; 9]);
        let mut formatting = with_direct_paragraph_runs(&[(100, 110, papx)], truncated);
        assert!(formatting
            .table_properties(109, 0, &[])
            .err()
            .unwrap()
            .contains("outside Data stream"));
    }

    #[test]
    fn supported_paragraph_prm_does_not_report_character_loss() {
        let mut f = empty();
        for code in [0x09, 0x18, 0x19] {
            f.run_xml(0, 0, 0x0100 | (code << 1), &[]).unwrap();
        }
        assert!(!f.unsupported_piece_properties && !f.unsupported_character_properties);
    }
}
