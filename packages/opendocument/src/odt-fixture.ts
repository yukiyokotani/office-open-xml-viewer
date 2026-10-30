/** Synthetic ODT packages for tests: a conforming leading `mimetype`, the
 * manifest and the given content/styles bodies (ODF 1.3 namespaces). */
import { storedZip } from '@silurus/ooxml-core/testing';

export const ODT_MEDIA_TYPE = 'application/vnd.oasis.opendocument.text';

const NS = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
].join(' ');

export interface OdtFixtureOptions {
  readonly body: string;
  readonly styles?: string;
  readonly automaticStyles?: string;
  readonly masterContent?: string;
  readonly manifestExtra?: string;
  readonly mediaType?: string;
  /** Further package files after the required ones. */
  readonly extraFiles?: ReadonlyArray<readonly [string, string]>;
}

export function buildOdtFixture(options: OdtFixtureOptions): Uint8Array<ArrayBuffer> {
  const mediaType = options.mediaType ?? ODT_MEDIA_TYPE;
  const fonts = '<office:font-face-decls><style:font-face style:name="Liberation Serif" svg:font-family="&apos;Liberation Serif&apos;" style:font-family-generic="roman" style:font-pitch="variable"/></office:font-face-decls>';
  const styles = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS} office:version="1.3">${fonts}`
    + '<office:styles><style:default-style style:family="paragraph"><style:text-properties style:font-name="Liberation Serif" fo:font-size="12pt" style:font-name-asian="Liberation Serif" style:font-size-asian="12pt" style:font-name-complex="Liberation Serif" style:font-size-complex="12pt" style:letter-kerning="true"/></style:default-style>'
    + `<style:style style:name="Standard" style:family="paragraph"/>${options.styles ?? ''}</office:styles>`
    + '<office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/></style:page-layout></office:automatic-styles>'
    + `<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1">${options.masterContent ?? ''}</style:master-page></office:master-styles></office:document-styles>`;
  const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3">${fonts}`
    + `<office:automatic-styles>${options.automaticStyles ?? ''}</office:automatic-styles>`
    + `<office:body><office:text>${options.body}</office:text></office:body></office:document-content>`;
  const manifest = '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">'
    + `<manifest:file-entry manifest:full-path="/" manifest:media-type="${mediaType}"/>${options.manifestExtra ?? ''}</manifest:manifest>`;
  return storedZip([
    ['mimetype', mediaType],
    ['META-INF/manifest.xml', manifest],
    ['styles.xml', styles],
    ['content.xml', content],
    ...(options.extraFiles ?? []),
  ]);
}
