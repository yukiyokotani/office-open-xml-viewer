/** Synthetic ODP packages for tests: a conforming leading `mimetype`, the
 * manifest and the given content/styles bodies (ODF 1.3 namespaces). */
import { storedZip } from '@silurus/ooxml-core/testing';

export const ODP_MEDIA_TYPE = 'application/vnd.oasis.opendocument.presentation';

const NS = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"',
  'xmlns:drawooo="http://openoffice.org/2010/draw"',
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"',
].join(' ');

export interface OdpFixtureOptions {
  /** `draw:page` elements (master "M" exists). */
  readonly pages: string;
  /** Extra common styles (office:styles). */
  readonly styles?: string;
  /** Content automatic styles. */
  readonly automaticStyles?: string;
  /** Objects on the master page "M". */
  readonly masterContent?: string;
  /** Master page drawing-page style properties. */
  readonly masterPageProperties?: string;
  /** office:presentation children before the pages (declarations). */
  readonly decls?: string;
  readonly manifestExtra?: string;
  readonly mediaType?: string;
  readonly extraFiles?: ReadonlyArray<readonly [string, string]>;
}

export function buildOdpFixture(options: OdpFixtureOptions): Uint8Array<ArrayBuffer> {
  const mediaType = options.mediaType ?? ODP_MEDIA_TYPE;
  const styles = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS} office:version="1.3">`
    + '<office:styles><style:default-style style:family="graphic"><style:graphic-properties svg:stroke-color="#3465a4" draw:fill-color="#729fcf"/><style:text-properties fo:font-size="18pt" fo:font-family="Liberation Sans" fo:color="#000000"/></style:default-style>'
    + `${options.styles ?? ''}</office:styles>`
    + '<office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="28cm" fo:page-height="15.75cm"/></style:page-layout>'
    + `<style:style style:name="Mdp1" style:family="drawing-page"><style:drawing-page-properties ${options.masterPageProperties ?? 'draw:fill="solid" draw:fill-color="#ffffff"'}/></style:style></office:automatic-styles>`
    + `<office:master-styles><style:master-page style:name="M" style:page-layout-name="PM1" draw:style-name="Mdp1">${options.masterContent ?? ''}</style:master-page></office:master-styles></office:document-styles>`;
  const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3">`
    + `<office:automatic-styles>${options.automaticStyles ?? ''}</office:automatic-styles>`
    + `<office:body><office:presentation>${options.decls ?? ''}${options.pages}</office:presentation></office:body></office:document-content>`;
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
