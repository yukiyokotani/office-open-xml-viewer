// Open Packaging Conventions scaffolding (ECMA-376 Part 2): the content-type
// stream and relationship parts. Nothing else is added to a package.

import { XML_DECL, attrs } from './xml.mjs';

export const RELS_CT = 'application/vnd.openxmlformats-package.relationships+xml';

export function contentTypes(overrides) {
  const body = Object.entries(overrides)
    .map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`)
    .join('');
  return `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + `<Default Extension="rels" ContentType="${RELS_CT}"/>`
    + `<Default Extension="xml" ContentType="application/xml"/>`
    + `${body}</Types>`;
}

/** @param {Array<[id: string, type: string, target: string, mode?: 'External']>} rels type is the
 * last segment of the officeDocument/2006/relationships URI, or a whole URI
 * for a relationship type defined elsewhere. */
export function relationships(rels) {
  const body = rels
    .map(([id, type, target, mode]) => `<Relationship${attrs({
      Id: id,
      Type: type.startsWith('http') ? type : `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}`,
      Target: target,
      TargetMode: mode,
    })}/>`)
    .join('');
  return `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
}

/** Relationship part path for a source part ('' = the package). */
export function relsPath(part) {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
}
