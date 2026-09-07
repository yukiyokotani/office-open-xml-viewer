//! Neutral OfficeArt gradient facts (MS-ODRAW 2.2.51 and 2.2.61).

use super::{unsupported, ByteSpan};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct ShadeColor {
    pub color: u32,
    /// Raw 16.16 position; validated to the inclusive [0, 1] range.
    pub position: u32,
}

#[derive(Clone)]
pub(crate) struct Storage<T> {
    colors: Option<T>,
    specified: bool,
    invalid_scalar: bool,
}
impl<T> Default for Storage<T> {
    fn default() -> Self {
        Self {
            colors: None,
            specified: false,
            invalid_scalar: false,
        }
    }
}
pub(crate) type Borrowed<'a> = Storage<&'a [u8]>;
pub(crate) type Spanned = Storage<ByteSpan>;

impl<T: Clone> Storage<T> {
    pub(crate) fn set(&mut self, value: T) {
        self.colors = Some(value);
        self.specified = true;
        self.invalid_scalar = false;
    }
    pub(crate) fn scalar(&mut self, value: u32) {
        self.colors = None;
        self.specified = true;
        self.invalid_scalar = value != 0;
    }
    pub(crate) fn inherit(&self, parent: &Self) -> Self {
        if self.specified {
            self.clone()
        } else {
            parent.clone()
        }
    }
}
impl<'a> Borrowed<'a> {
    pub(crate) fn decode(
        &self,
        work_budget: &mut usize,
        model_budget: &mut usize,
    ) -> Result<Option<Vec<ShadeColor>>, String> {
        if self.invalid_scalar {
            return Err(unsupported("invalid scalar OfficeArt shade color property"));
        }
        self.colors
            .map(|value| decode(value, work_budget, model_budget))
            .transpose()
    }
}
impl Spanned {
    pub(crate) fn view<'a>(&self, backing: &'a [u8]) -> Result<Borrowed<'a>, String> {
        Ok(Storage {
            colors: self
                .colors
                .as_ref()
                .map(|value| value.view(backing))
                .transpose()?,
            specified: self.specified,
            invalid_scalar: self.invalid_scalar,
        })
    }
}

fn decode(
    value: &[u8],
    work_budget: &mut usize,
    model_budget: &mut usize,
) -> Result<Vec<ShadeColor>, String> {
    if value.len() < 6 {
        return Err(unsupported("truncated OfficeArt shade color array"));
    }
    let count = usize::from(u16::from_le_bytes([value[0], value[1]]));
    let allocated = usize::from(u16::from_le_bytes([value[2], value[3]]));
    let encoded = u16::from_le_bytes([value[4], value[5]]);
    if allocated < count || !matches!(encoded, 8 | 0xfff0) {
        return Err(unsupported("invalid OfficeArt shade color array header"));
    }
    let width = if encoded == 8 { 8 } else { 4 };
    if value.len()
        != 6usize
            .checked_add(
                count
                    .checked_mul(width)
                    .ok_or_else(|| unsupported("OfficeArt shade color array is too large"))?,
            )
            .ok_or_else(|| unsupported("OfficeArt shade color array is too large"))?
    {
        return Err(unsupported("invalid OfficeArt shade color array length"));
    }
    *work_budget = work_budget
        .checked_sub(count)
        .ok_or_else(|| unsupported("OfficeArt shade color work budget exceeded"))?;
    let charge = count
        .checked_mul(std::mem::size_of::<ShadeColor>())
        .ok_or_else(|| unsupported("OfficeArt shade color model budget exceeded"))?;
    *model_budget = model_budget
        .checked_sub(charge)
        .ok_or_else(|| unsupported("OfficeArt shade color model budget exceeded"))?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(count)
        .map_err(|_| unsupported("OfficeArt shade color allocation failed"))?;
    let mut previous = 0;
    for i in 0..count {
        let at = 6 + i * width;
        let color = u32::from_le_bytes(value[at..at + 4].try_into().unwrap());
        let position = if width == 8 {
            u32::from_le_bytes(value[at + 4..at + 8].try_into().unwrap())
        } else {
            0
        };
        if position > 65536 || (i != 0 && position < previous) {
            return Err(unsupported("invalid OfficeArt shade color position"));
        }
        previous = position;
        result.push(ShadeColor { color, position });
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    fn decoded(value: &[u8]) -> Result<Vec<ShadeColor>, String> {
        let (mut work, mut model) = (usize::MAX, usize::MAX);
        decode(value, &mut work, &mut model)
    }

    #[test]
    fn decodes_full_truncated_empty_and_duplicate_positions() {
        let full = [
            2, 0, 2, 0, 8, 0, 1, 0, 0, 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, 1, 0,
        ];
        assert_eq!(
            decoded(&full).unwrap(),
            vec![
                ShadeColor {
                    color: 1,
                    position: 0
                },
                ShadeColor {
                    color: 2,
                    position: 65536
                }
            ]
        );
        let short = [1, 0, 1, 0, 0xf0, 0xff, 3, 0, 0, 0];
        assert_eq!(
            decoded(&short).unwrap(),
            vec![ShadeColor {
                color: 3,
                position: 0
            }]
        );
        assert!(decoded(&[0, 0, 0, 0, 8, 0]).unwrap().is_empty());
        let duplicate = [
            2, 0, 2, 0, 8, 0, 1, 0, 0, 0, 9, 0, 0, 0, 2, 0, 0, 0, 9, 0, 0, 0,
        ];
        assert_eq!(decoded(&duplicate).unwrap()[1].position, 9);
    }

    #[test]
    fn rejects_invalid_headers_positions_and_lengths() {
        assert!(decoded(&[2, 0, 1, 0, 8, 0]).is_err());
        assert!(decoded(&[0, 0, 0, 0, 7, 0]).is_err());
        let out_of_range = [1, 0, 1, 0, 8, 0, 1, 0, 0, 0, 1, 0, 1, 0];
        assert!(decoded(&out_of_range).is_err());
        let descending = [
            2, 0, 2, 0, 8, 0, 1, 0, 0, 0, 1, 0, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0,
        ];
        assert!(decoded(&descending).is_err());
        let valid = [1, 0, 1, 0, 8, 0, 1, 0, 0, 0, 0, 0, 0, 0];
        assert!(decoded(&valid[..13]).is_err());
        let mut trailing = valid.to_vec();
        trailing.push(0);
        assert!(decoded(&trailing).is_err());
    }

    #[test]
    fn charges_work_and_retained_bytes_before_allocation() {
        let valid = [1, 0, 1, 0, 8, 0, 1, 0, 0, 0, 0, 0, 0, 0];
        let bytes = std::mem::size_of::<ShadeColor>();
        let (mut work, mut model) = (1, bytes);
        assert_eq!(decode(&valid, &mut work, &mut model).unwrap().len(), 1);
        assert_eq!((work, model), (0, 0));
        assert!(decode(&valid, &mut 0, &mut bytes.clone()).is_err());
        assert!(decode(&valid, &mut 1, &mut (bytes - 1)).is_err());
    }

    #[test]
    fn explicit_reset_vetoes_inheritance_and_spans_validate_backing() {
        let bytes = vec![1, 0, 1, 0, 8, 0, 7, 0, 0, 0, 0, 0, 0, 0];
        let span = ByteSpan::new(0..bytes.len(), bytes.len(), "gradient").unwrap();
        let mut parent = Spanned::default();
        parent.set(span);
        let inherited = Spanned::default().inherit(&parent);
        let moved = bytes;
        assert_eq!(
            inherited
                .view(&moved)
                .unwrap()
                .decode(&mut 1, &mut 8)
                .unwrap()
                .unwrap()[0]
                .color,
            7
        );
        assert!(inherited.view(&moved[..moved.len() - 1]).is_err());
        let mut reset = Spanned::default();
        reset.scalar(0);
        assert!(reset
            .inherit(&parent)
            .view(&moved)
            .unwrap()
            .decode(&mut 1, &mut 8)
            .unwrap()
            .is_none());
        let mut invalid = Spanned::default();
        invalid.scalar(1);
        assert!(invalid
            .inherit(&parent)
            .view(&moved)
            .unwrap()
            .decode(&mut 1, &mut 8)
            .unwrap_err()
            .contains("invalid scalar"));
    }
}
