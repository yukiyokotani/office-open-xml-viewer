// XML text helpers. Parts are written by hand as strings so that every byte of
// the specimen is visible in its case module; these helpers only escape.

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export function esc(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Attribute list from an object; undefined/null values are omitted, order kept. */
export function attrs(object) {
  return Object.entries(object)
    .filter(([, value]) => value !== undefined && value !== null)
    .map(([name, value]) => ` ${name}="${esc(value)}"`)
    .join('');
}

/** Namespace URIs and the prefixes used in `covers` shorthand. */
export const NS = {
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  pr: 'http://schemas.openxmlformats.org/package/2006/relationships',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  x: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  v: 'urn:schemas-microsoft-com:vml',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
};

export const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
