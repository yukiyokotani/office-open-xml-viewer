export interface CanvasFontRoute {
  /** Complete CSS family list. Consumers must not append another fallback tail. */
  readonly familyList: string;
  readonly scope: 'registered' | 'native' | 'generic';
  /** Syntax identity only; this never asserts portable glyph geometry. */
  readonly fingerprint: string;
  /** Canvas `lang` (HTML §4.12.5.1.4, a BCP 47 tag) the route shapes text
   * in; absent leaves the context's own. The language selects OpenType forms
   * (`locl`) and positioning, so it is part of the route: text measured,
   * compared, cached or painted through a route uses that route's language
   * ({@link canvasFontRouteWithLanguage}). */
  readonly lang?: string;
}

export function createCanvasFontRoute(
  familyList: string,
  scope: CanvasFontRoute['scope'],
): CanvasFontRoute {
  const normalized = familyList.trim();
  if (!normalized) throw new TypeError('Canvas font route requires a family list');
  return Object.freeze({
    familyList: normalized,
    scope,
    fingerprint: `canvas-font-route-v1:${encodeURIComponent(scope)}:${encodeURIComponent(normalized)}`,
  });
}

const languageRoutes = new WeakMap<CanvasFontRoute, Map<string, CanvasFontRoute>>();
/** Languages interned per base route. Past it a derived route is created
 * afresh: its fingerprint, the identity consumers compare and key by, is the
 * same string, so only object reuse is lost. */
const INTERNED_LANGUAGES_PER_ROUTE = 64;
const faceRoutes = new WeakMap<CanvasFontRoute, CanvasFontRoute>();

/**
 * `route` shaping in `lang`. The language joins the fingerprint, so every
 * consumer that compares, merges or caches by route identity keeps text of
 * different languages apart, and whatever applies the route to a context
 * applies `lang` with it (core internal/canvas-language.ts). A route without
 * a language is unchanged, fingerprint included.
 *
 * One object per base route and language: routes are also compared and
 * weakly keyed by identity. The table lives and dies with the base route.
 */
export function canvasFontRouteWithLanguage(
  route: Readonly<CanvasFontRoute>,
  lang: string | undefined,
): CanvasFontRoute {
  if (lang === undefined || lang === route.lang) return route;
  if (route.lang !== undefined) throw new TypeError('Canvas font route already has a language');
  let routes = languageRoutes.get(route);
  if (!routes) languageRoutes.set(route, routes = new Map());
  let derived = routes.get(lang);
  if (!derived) {
    derived = Object.freeze({
      familyList: route.familyList,
      scope: route.scope,
      // Encoded family lists and scopes contain no ':'.
      fingerprint: `${route.fingerprint}:lang:${encodeURIComponent(lang)}`,
      lang: lang,
    });
    if (routes.size < INTERNED_LANGUAGES_PER_ROUTE) routes.set(lang, derived);
    faceRoutes.set(derived, route);
  }
  return derived;
}

/**
 * The faces `route` resolves to, whatever language it shapes in: the route
 * itself, or the route its language was added to. A named family is matched
 * in list order under any language (CSS Fonts 4 §5), so a rule that reads a
 * face's own metrics (line geometry, average width) compares this identity.
 * Whatever measures, compares or paints shaped text compares the route.
 */
export function canvasFontRouteFaces(route: Readonly<CanvasFontRoute>): CanvasFontRoute {
  return route.lang === undefined ? route : faceRoutes.get(route) ?? route;
}

export function canvasFontString(
  route: Readonly<CanvasFontRoute>,
  sizePx: number,
  weight: number,
  style: 'normal' | 'italic',
): string {
  if (!Number.isFinite(sizePx) || sizePx < 0) throw new RangeError('Canvas font size must be finite and non-negative');
  if (!Number.isFinite(weight) || weight < 1 || weight > 1000) {
    throw new RangeError('Canvas font weight must be finite and between 1 and 1000');
  }
  if (!route.familyList.trim()) throw new TypeError('Canvas font route requires a family list');
  return `${style} ${weight} ${sizePx}px ${route.familyList}`;
}
