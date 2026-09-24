/** Desktop macOS font environment; iPadOS may advertise MacIntel. */
export function isMacDesktop(): boolean {
  const nav = typeof navigator !== 'undefined' ? navigator : undefined;
  return Boolean(nav && /Mac/.test(nav.platform || nav.userAgent || '')
    && !(nav.platform === 'MacIntel' && nav.maxTouchPoints > 1));
}
