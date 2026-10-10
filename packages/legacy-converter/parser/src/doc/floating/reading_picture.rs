//! Explicit reading-policy picture properties; never a general OfficeArt whitelist.
//! MS-ODRAW 2.2.12/2.3 establish shape properties over document defaults.
//! Same-scope contradictory assignments have no inferred chronological winner.
use super::{records, unsupported};
use crate::officeart::{properties, Record};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Scope {
    Document,
    Shape,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct Raw<'a> {
    pub scope: Scope,
    pub opid: u16,
    pub value: u32,
    pub complex: Option<&'a [u8]>,
}

#[derive(Default)]
pub(super) struct Defaults<'a> {
    pub primary: Option<Record<'a>>,
    pub tertiary: Option<Record<'a>>,
    pub duplicate: bool,
    pub unacquired_secondary: bool,
}
impl<'a> Defaults<'a> {
    pub fn register(&mut self, record: Record<'a>) {
        let slot = match record.kind {
            0xf00b => &mut self.primary,
            0xf122 => &mut self.tertiary,
            0xf121 => {
                self.unacquired_secondary = true;
                return;
            }
            _ => return,
        };
        self.duplicate |= slot.replace(record).is_some();
    }
}

#[derive(Clone, Copy, Default)]
struct Booleans {
    used: u16,
    values: u16,
}
impl Booleans {
    fn add(&mut self, raw: Raw<'_>, mask: u16) -> Result<(), String> {
        ordinary(raw)?;
        let used = (raw.value >> 16) as u16 & mask;
        let values = raw.value as u16 & used;
        if (self.values ^ values) & self.used & used != 0 {
            return Err(unsupported("conflicting reading picture Boolean owners"));
        }
        self.used |= used;
        self.values = (self.values & !used) | values;
        Ok(())
    }
    fn over(self, defaults: Self) -> Self {
        Self {
            used: self.used | defaults.used,
            values: self.values | (defaults.values & !self.used),
        }
    }
    fn effective(self, builtin: u16) -> u16 {
        self.values | (builtin & !self.used)
    }
}

#[derive(Default)]
struct Owner<'a> {
    names: [Option<Raw<'a>>; 3], // pibName, wzName, wzDescription
    pib_flags: Option<u32>,
    blip: Booleans,
    fill: Booleans,
    line: Booleans,
    carriers: [Option<Raw<'a>>; 2], // fillBlip, lineFillBlip: not name strings
    line_present: bool,
}

fn ordinary(raw: Raw<'_>) -> Result<(), String> {
    if raw.opid & 0xc000 != 0 || raw.complex.is_some() {
        return Err(unsupported(
            "invalid reading picture ordinary property flags",
        ));
    }
    Ok(())
}
fn singleton<'a>(slot: &mut Option<Raw<'a>>, raw: Raw<'a>) -> Result<(), String> {
    if slot.is_some_and(|old| old.value != raw.value || old.complex != raw.complex) {
        return Err(unsupported("conflicting reading picture metadata owners"));
    }
    // Agreeing values share semantic ownership. Keep the last encoded representative
    // for metadata; this is not an archive of every declaration or an opid rewrite.
    *slot = Some(raw);
    Ok(())
}
fn text(raw: Raw<'_>) -> Result<(), String> {
    // fBid is undefined and MUST be ignored for these names (2.3.4/2.3.23).
    let Some(bytes) = raw.complex else {
        return if raw.value == 0 {
            Ok(())
        } else {
            Err(unsupported("invalid reading picture name scalar"))
        };
    };
    if bytes.len() < 2 || bytes.len() % 2 != 0 || bytes[bytes.len() - 2..] != [0, 0] {
        return Err(unsupported("invalid reading picture Unicode name framing"));
    }
    let words = bytes[..bytes.len() - 2]
        .as_chunks::<2>()
        .0
        .iter()
        .map(|word| u16::from_le_bytes(*word));
    for decoded in char::decode_utf16(words) {
        match decoded {
            Ok(c) if c != '\0' => {}
            _ => return Err(unsupported("invalid reading picture Unicode name")),
        }
    }
    Ok(())
}

impl<'a> Owner<'a> {
    fn add(&mut self, property: properties::Property<'a>, scope: Scope) -> Result<(), String> {
        let raw = Raw {
            scope,
            opid: property.opid,
            value: property.value,
            complex: property.complex,
        };
        let id = raw.opid & 0x3fff;
        match id {
            0x7f => ordinary(raw), // 2.3.20.1: edit-only protection, unused bits ignored.
            0x105 | 0x380 | 0x381 => {
                text(raw)?;
                let index = match id {
                    0x105 => 0,
                    0x380 => 1,
                    _ => 2,
                };
                singleton(&mut self.names[index], raw)
            }
            0x106 => {
                ordinary(raw)?;
                // 2.4.8: File/URL mutually exclusive; DoNotSave requires link,
                // and a link requires File/URL. Preserve the semantic refusal.
                let flags = raw.value;
                if flags & !15 != 0
                    || flags & 3 == 3
                    || (flags & 4 != 0 && flags & 8 == 0)
                    || (flags & 8 != 0 && flags & 3 == 0)
                {
                    return Err(unsupported("invalid reading picture name flags"));
                }
                if self.pib_flags.is_some_and(|old| old != flags) {
                    return Err(unsupported("conflicting reading picture name flags"));
                }
                self.pib_flags = Some(flags);
                Ok(())
            }
            0x13f => self.blip.add(raw, 0x007f), // 2.3.23.35: seven default-false members.
            0x1bf => {
                // Preserve the existing conservative fill subset. Acquiring
                // fFilled ownership is not permission to admit other authored
                // fill/hit-test/anchor members, even when they appear unused.
                if raw.value as u16 & 0x006f != 0 {
                    return Err(unsupported("reading picture fill members are not acquired"));
                }
                self.fill.add(raw, 0x007f)
            }
            0x1ff => {
                // 2.3.8.38 reserved value bits MUST be zero even if unused.
                if raw.value & 0x0180 != 0 {
                    return Err(unsupported("invalid reading picture reserved line bits"));
                }
                self.line_present = true;
                self.line.add(raw, 0x027f)
            }
            0x186 | 0x1c5 => {
                // MS-ODRAW 2.2.8: when fComplex=1, fBid MUST be ignored and
                // op is the bounded payload length. Preserve the authored opid.
                // Property-specific producer flags prescribe fBid=0 for complex data
                // (2.3.7.7/2.3.8.6); ignoring this bit here is reading-consumer handling,
                // not a claim that the producer encoding conforms. Scalar indices
                // still require fBid=1. Framing, inactive-owner and budget gates remain.
                if raw.complex.is_none() && raw.opid & 0x4000 == 0 {
                    return Err(unsupported("invalid reading picture latent BLIP flags"));
                }
                singleton(&mut self.carriers[usize::from(id == 0x1c5)], raw)
            }
            0x180 | 0x1c4 => {
                ordinary(raw)?;
                // Closed first class: solid paint excludes pattern/texture BLIPs.
                if raw.value != 0 {
                    return Err(unsupported(
                        "reading picture non-solid latent paint is not acquired",
                    ));
                }
                self.line_present |= id == 0x1c4;
                Ok(())
            }
            0x81..=0x84 if scope == Scope::Document => {
                ordinary(raw)?;
                // Picture gate independently requires no owned textbox story.
                if raw.value > 0x132f540 {
                    return Err(unsupported("invalid Word textbox margin"));
                }
                Ok(())
            }
            // Keep the existing picture placement/crop subset, shape scope only.
            0x100..=0x104 | 0x384..=0x387 | 0x38f..=0x392 | 0x3aa | 0x3bf
                if scope == Scope::Shape && raw.complex.is_none() =>
            {
                Ok(())
            }
            4 if scope == Scope::Shape && raw.complex.is_none() && raw.value == 0 => Ok(()),
            0x383 => Err(unsupported(
                "reading picture retains a stored/empty contour property",
            )),
            _ => Err(unsupported(
                "reading picture property/default is not acquired",
            )),
        }
    }
    fn table(
        &mut self,
        record: Record<'a>,
        scope: Scope,
        budget: &mut usize,
    ) -> Result<(), String> {
        let callback = |property| self.add(property, scope);
        match record.kind {
            0xf00b => properties::visit(record, budget, callback),
            0xf122 => properties::visit_tertiary(record, budget, callback),
            _ => Err(unsupported("invalid reading picture property table kind")),
        }
    }
}

/// Borrowed result. Raw latent bytes are not a validated/selected OfficeArtBlip.
/// Owned projection must charge the model budget before copying any payload.
pub(super) struct Acquired<'a> {
    pub names: [Option<Raw<'a>>; 3],
    pub carriers: [Option<Raw<'a>>; 2],
}
pub(super) fn acquire<'a>(
    shape: Record<'a>,
    defaults: &Defaults<'a>,
    budget: &mut usize,
) -> Result<Acquired<'a>, String> {
    if defaults.unacquired_secondary {
        return Err(unsupported(
            "reading picture secondary properties are not acquired",
        ));
    }
    if defaults.duplicate {
        return Err(unsupported(
            "duplicate reading picture document property tables",
        ));
    }
    let mut document = Owner::default();
    for table in [defaults.primary, defaults.tertiary].into_iter().flatten() {
        document.table(table, Scope::Document, budget)?;
    }
    let mut local = Owner::default();
    for table in records(shape.payload, budget)? {
        // MS-ODRAW 2.2.10 permits movie in SecondaryFOPT. No passive
        // picture proof is inferred from leaving this table unconsumed.
        if table.kind == 0xf121 {
            return Err(unsupported(
                "reading picture secondary properties are not acquired",
            ));
        }
        if matches!(table.kind, 0xf00b | 0xf122) {
            local.table(table, Scope::Shape, budget)?;
        }
    }
    let blip = local.blip.over(document.blip);
    if blip.effective(0) != 0 {
        return Err(unsupported(
            "reading picture active BLIP effects are not acquired",
        ));
    }
    let flags = local.pib_flags.or(document.pib_flags).unwrap_or(0);
    if flags != 0 {
        return Err(unsupported(
            "reading picture external name semantics are not acquired",
        ));
    }
    let names = std::array::from_fn(|i| local.names[i].or(document.names[i]));
    let carriers: [Option<Raw<'a>>; 2] =
        std::array::from_fn(|i| local.carriers[i].or(document.carriers[i]));
    let fill = local.fill.over(document.fill);
    let line = local.line.over(document.line);
    if fill.effective(0x1c) & 0x62 != 0 {
        return Err(unsupported("reading picture fill effects are not acquired"));
    }
    let needs_inactive =
        |raw: Option<Raw<'_>>| raw.is_some_and(|raw| raw.complex.is_some() || raw.value != 0);
    if needs_inactive(carriers[0]) && (fill.used & 0x10 == 0 || fill.effective(0x1c) & 0x50 != 0) {
        return Err(unsupported(
            "reading picture latent fill has no proved inactive owner",
        ));
    }
    // fLine defaults true; an unused local zero is not a disabled line. Bit0
    // fNoLineDrawDash can draw a line even when fLine=false (2.3.8.38).
    if (needs_inactive(carriers[1]) || local.line_present || document.line_present)
        && (line.used & 8 == 0 || line.effective(0x2e) & 0x249 != 0)
    {
        return Err(unsupported(
            "reading picture latent line has no proved inactive owner",
        ));
    }
    Ok(Acquired { names, carriers })
}

impl Acquired<'_> {
    pub fn into_wire(
        self,
        remaining_bytes: usize,
    ) -> Result<Option<Box<docx_model::NativePictureMetadataWire>>, String> {
        if self
            .names
            .iter()
            .chain(self.carriers.iter())
            .all(Option::is_none)
        {
            return Ok(None);
        }
        // Preflight every owned name/raw allocation before copying. The final
        // ImageRun payload counts actual capacities before model publication.
        let mut required = std::mem::size_of::<docx_model::NativePictureMetadataWire>();
        for raw in self.names.iter().chain(self.carriers.iter()).flatten() {
            let bytes = raw.complex.map_or(0, |b| b.len());
            let decoded_bound = if matches!(raw.opid & 0x3fff, 0x105 | 0x380 | 0x381) {
                bytes.checked_mul(2).ok_or("OUTPUT_TOO_LARGE")?
            } else {
                0
            };
            required = required
                .checked_add(bytes)
                .and_then(|v| v.checked_add(decoded_bound))
                .ok_or("OUTPUT_TOO_LARGE")?;
        }
        if required > remaining_bytes {
            return Err("OUTPUT_TOO_LARGE".into());
        }
        let convert = |raw: Raw<'_>,
                       is_name: bool|
         -> Result<docx_model::NativePicturePropertyWire, String> {
            let text = if is_name {
                let bytes = raw.complex.unwrap_or(&[]);
                let content = if bytes.is_empty() {
                    &[][..]
                } else {
                    &bytes[..bytes.len() - 2]
                };
                let capacity = content.len().checked_mul(3).ok_or("OUTPUT_TOO_LARGE")? / 2;
                let mut decoded = String::with_capacity(capacity);
                for c in char::decode_utf16(
                    content
                        .as_chunks::<2>()
                        .0
                        .iter()
                        .map(|word| u16::from_le_bytes(*word)),
                ) {
                    decoded
                        .push(c.map_err(|_| unsupported("invalid reading picture Unicode name"))?);
                }
                Some(decoded)
            } else {
                None
            };
            Ok(docx_model::NativePicturePropertyWire {
                scope: match raw.scope {
                    Scope::Document => docx_model::NativePicturePropertyScopeWire::DocumentDefault,
                    Scope::Shape => docx_model::NativePicturePropertyScopeWire::Shape,
                },
                opid: raw.opid,
                value: raw.value,
                text,
                raw_bytes: raw.complex.unwrap_or(&[]).to_vec(),
                retention: if is_name {
                    docx_model::NativePicturePropertyRetentionWire::PassiveName
                } else if raw.complex.is_some() {
                    docx_model::NativePicturePropertyRetentionWire::InactiveOpaqueNotDecoded
                } else if raw.value == 0 {
                    docx_model::NativePicturePropertyRetentionWire::IgnoredZeroIndex
                } else {
                    docx_model::NativePicturePropertyRetentionWire::InactiveIndexNotResolved
                },
            })
        };
        let [pib, name, description] = self.names;
        let [fill, line] = self.carriers;
        Ok(Some(Box::new(docx_model::NativePictureMetadataWire {
            blip_name: pib.map(|raw| convert(raw, true)).transpose()?,
            shape_name: name.map(|raw| convert(raw, true)).transpose()?,
            description: description.map(|raw| convert(raw, true)).transpose()?,
            inactive_fill_carrier: fill.map(|raw| convert(raw, false)).transpose()?,
            inactive_line_carrier: line.map(|raw| convert(raw, false)).transpose()?,
        })))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reading_picture_metadata_budget_charges_names_and_opaque_bytes_before_owned_projection() {
        // Actual property framing + ownership + projection; no mock validator.
        let name = "界"
            .repeat(100)
            .encode_utf16()
            .chain([0])
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>();
        let body = [
            0x8381u16.to_le_bytes().as_slice(),
            &(name.len() as u32).to_le_bytes(),
            name.as_slice(),
        ]
        .concat();
        let table = [
            (0x13u16).to_le_bytes().as_slice(),
            &0xf00bu16.to_le_bytes(),
            &(body.len() as u32).to_le_bytes(),
            body.as_slice(),
        ]
        .concat();
        let shape = Record {
            version: 15,
            instance: 0,
            kind: 0xf004,
            payload: &table,
        };
        let acquired = acquire(shape, &Defaults::default(), &mut 100).unwrap();
        assert_eq!(acquired.into_wire(512).unwrap_err(), "OUTPUT_TOO_LARGE");
        let acquired = acquire(shape, &Defaults::default(), &mut 100).unwrap();
        let wire = acquired.into_wire(4096).unwrap().unwrap();
        assert_eq!(
            wire.description.as_ref().unwrap().text.as_deref(),
            Some("界".repeat(100).as_str())
        );
        assert_eq!(wire.description.as_ref().unwrap().raw_bytes, name);
        let empty = Record {
            version: 15,
            instance: 0,
            kind: 0xf004,
            payload: &[],
        };
        assert!(acquire(empty, &Defaults::default(), &mut 100)
            .unwrap()
            .into_wire(0)
            .unwrap()
            .is_none());
        assert!(acquire(shape, &Defaults::default(), &mut 1).is_err());
    }

    #[test]
    fn reading_picture_metadata_refuses_unacquired_secondary_property_owners() {
        let secondary = [
            0x0003u16.to_le_bytes().as_slice(),
            &0xf121u16.to_le_bytes(),
            &0u32.to_le_bytes(),
        ]
        .concat();
        let movie_owner = Record {
            version: 15,
            instance: 0,
            kind: 0xf004,
            payload: &secondary,
        };
        assert!(acquire(movie_owner, &Defaults::default(), &mut 100).is_err());
        let mut defaults = Defaults::default();
        defaults.register(records(&secondary, &mut 100).unwrap()[0]);
        let empty = Record {
            version: 15,
            instance: 0,
            kind: 0xf004,
            payload: &[],
        };
        assert!(acquire(empty, &defaults, &mut 100).is_err());
    }
}
