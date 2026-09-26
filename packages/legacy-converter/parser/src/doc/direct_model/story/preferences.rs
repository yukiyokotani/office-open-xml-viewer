//! Style inheritance for row preferences that the table projection checks
//! against the physical row geometry (see `tables::check_row_preferences`).

use super::*;

/// [MS-DOC] 2.6.3 sprmTIstd applies the selected style's table properties
/// before later direct Prls. Direct sprmTWidthIndent/sprmTWidthBefore are
/// admitted only after the selection (see table::NativeAdmission), so they
/// override the style's inherited values; otherwise the style values are the
/// row's effective preferences.
pub(super) fn resolve(
    prepared: &mut [PreparedParagraph],
    index: &table_context::Index,
    formatting: &mut formatting::Formatting<'_>,
) -> Result<(), String> {
    for table_context in index.tables() {
        for row_context in &table_context.rows {
            let (indent, before) = formatting.table_row_preferences(row_context.table_style)?;
            let row = &mut prepared
                .get_mut(row_context.ttp_id)
                .ok_or_else(|| unsupported("Word table preference TTP outside story"))?
                .table_properties
                .row;
            if row.preferred_indent.is_none() {
                row.preferred_indent = indent;
            }
            if row.preferred_before.is_none() {
                row.preferred_before = before;
            }
        }
    }
    Ok(())
}
