/** Synthetic ODS packages for tests: a conforming leading `mimetype`, the
 * manifest and the given content/styles bodies (ODF 1.3 namespaces). */
import { storedZip } from '@silurus/ooxml-core/testing';

export const ODS_MEDIA_TYPE = 'application/vnd.oasis.opendocument.spreadsheet';

export const ODS_NS = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"',
  'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"',
].join(' ');

export interface OdsFixtureOptions {
  /** The children of `office:spreadsheet` (tables, named expressions). */
  readonly body: string;
  readonly styles?: string;
  readonly automaticStyles?: string;
  readonly mediaType?: string;
  /** Further package files after the required ones. */
  readonly extraFiles?: ReadonlyArray<readonly [string, string]>;
}

export function buildOdsFixture(options: OdsFixtureOptions): Uint8Array<ArrayBuffer> {
  const mediaType = options.mediaType ?? ODS_MEDIA_TYPE;
  const fonts = '<office:font-face-decls><style:font-face style:name="Liberation Sans" svg:font-family="&apos;Liberation Sans&apos;" style:font-family-generic="swiss"/></office:font-face-decls>';
  const styles = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${ODS_NS} office:version="1.3">${fonts}`
    + '<office:styles><style:default-style style:family="table-cell"><style:paragraph-properties style:tab-stop-distance="1.25cm"/><style:text-properties style:font-name="Liberation Sans" fo:font-size="10pt"/></style:default-style>'
    + `<style:style style:name="Default" style:family="table-cell"/>${options.styles ?? ''}</office:styles></office:document-styles>`;
  const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODS_NS} office:version="1.3">${fonts}`
    + `<office:automatic-styles>${options.automaticStyles ?? ''}</office:automatic-styles>`
    + `<office:body><office:spreadsheet>${options.body}</office:spreadsheet></office:body></office:document-content>`;
  const manifest = '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">'
    + `<manifest:file-entry manifest:full-path="/" manifest:media-type="${mediaType}"/></manifest:manifest>`;
  return storedZip([
    ['mimetype', mediaType],
    ['META-INF/manifest.xml', manifest],
    ['styles.xml', styles],
    ['content.xml', content],
    ...(options.extraFiles ?? []),
  ]);
}
