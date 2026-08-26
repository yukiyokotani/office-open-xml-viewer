function hasDataAttribute(candidate: EventTarget, dataAttribute: string): boolean {
  return (candidate as { dataset?: DOMStringMap }).dataset?.[dataAttribute] !== undefined;
}

/** Return true only when the event targets a marked element owned by `root`.
 * The root scope prevents one Viewer from treating another Viewer's comment UI
 * as its own interaction boundary. */
export function eventTargetsDataAttributeWithin(
  event: Event,
  root: Node,
  dataAttribute: string,
): boolean {
  const path = typeof event.composedPath === 'function' ? event.composedPath() : [];
  if (path.length > 0) {
    let marked = false;
    for (const candidate of path) {
      if (candidate === root) return marked;
      if (hasDataAttribute(candidate, dataAttribute)) marked = true;
    }
  }

  const target = event.target as Node | null;
  if (!target || !root.contains(target)) return false;
  let element = target as HTMLElement | null;
  while (element) {
    if (hasDataAttribute(element, dataAttribute)) return true;
    if (element === root) break;
    element = element.parentElement;
  }
  return false;
}
