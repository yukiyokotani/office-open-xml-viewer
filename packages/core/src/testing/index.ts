// Test-only utilities shared across the docx / pptx / xlsx suites. Never
// imported by production code; exposed via the `@silurus/ooxml-core/testing`
// subpath.
export { findMissingExportsFromUrl, formatMissing } from './export-completeness';
export { buildCfbFixture, buildCfbWithStreams, type CfbStream } from './cfb-fixture';
export { ENCRYPTED_DOCX_SPIN0_BASE64, encryptedDocxSpin0 } from './encrypted-fixture';
