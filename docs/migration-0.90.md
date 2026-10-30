# Migrating from 0.89 to 0.90: font sources

Version 0.90 replaces the boolean `useGoogleFonts` load option with `fonts`, an
ordered list of font sources. Migration is required only if you pass
`useGoogleFonts`. The default is unchanged: with no sources, the library
requests no fonts from any network host and the canvas uses the fonts installed
on the viewer's machine.

`useGoogleFonts` still works in 0.90, with a deprecation warning. It is removed
in 0.91.0.

## Replace `useGoogleFonts`

| 0.89 | 0.90 |
| --- | --- |
| `useGoogleFonts: true` | `fonts: [googleFonts()]` |
| `useGoogleFonts: false` | omit the option |

```ts
import { DocxViewer, googleFonts } from '@silurus/ooxml/docx';

const viewer = new DocxViewer(canvas, { fonts: [googleFonts()] });
```

The factories are exported from `@silurus/ooxml` and from the
`@silurus/ooxml/docx`, `@silurus/ooxml/xlsx` and `@silurus/ooxml/pptx` entries.
They are the same objects in every entry. `fonts` works the same way for DOCX,
XLSX and PPTX, on every engine `load()` and every Viewer, in `mode: 'main'` and
`mode: 'worker'`.

`fonts: [googleFonts()]` sends the same stylesheet requests as
`useGoogleFonts: true`. In 0.90, DOCX and XLSX also load only the
`unicode-range` subsets that intersect the characters the document paints, as
PPTX already did.

While the alias exists:

- `useGoogleFonts: true` is converted to `fonts: [googleFonts()]` before
  loading starts.
- `useGoogleFonts: false` does nothing, but still warns.
- `useGoogleFonts: true` together with `fonts` throws a `TypeError`. Add
  `googleFonts()` to `fonts` instead.
- The warning is printed once per JavaScript realm (page, popup or worker):

  ```text
  [ooxml] useGoogleFonts is deprecated and will be removed in 0.91.0. Replace it with fonts: [googleFonts()] — import { googleFonts } from '@silurus/ooxml' (also exported from the docx/xlsx/pptx entries). A custom CSS origin is googleFonts({ cssOrigin }); other sources are fontCssApi(), fontCss(), fontFiles() and fontResolver(). See docs/migration-0.90.md.
  ```

The VS Code extension keeps its `ooxmlViewer.useGoogleFonts` setting. It now
passes `fonts: [googleFonts()]`, so its users see no warning.

## Font source kinds

### A Google-compatible mirror or proxy: `googleFonts({ cssOrigin })`

```ts
fonts: [googleFonts({ cssOrigin: 'https://fonts.googleapis.cn' })]
fonts: [googleFonts({ cssOrigin: 'https://fonts.internal.example:8443' })]
```

`cssOrigin` is a scheme, host and optional port (`http:` or `https:`). One
trailing slash is accepted. A path, query or user name throws a `TypeError`
when `googleFonts()` is called, before any load. The requests keep the Google
CSS2 path and query (`/css2?family=…&display=swap`), so the origin must serve
that API. The library sends no stylesheet request to `fonts.googleapis.com`.

`cssOrigin` sets only where the stylesheet requests start:

- Redirects are followed. The stylesheet is read from the final URL, which can
  be on another origin.
- Font file URLs come from the returned CSS and are used as given. Relative
  URLs are resolved against the stylesheet's final URL. Absolute URLs can
  point anywhere, for example at `fonts.gstatic.com` if the mirror passes
  Google's CSS through unchanged.

For network isolation, the mirror must control its redirects and rewrite the
font URLs it returns to hosts you allow. Your CSP (`connect-src` and
`font-src`) is the enforcement point. Two documents that use different
origins share no stylesheet, cache entry or registered face.

A proxy that serves the Google CSS2 API under a path prefix
(`https://proxy.example/google-fonts/css2?…`) is not an origin, so
`cssOrigin` cannot express it. Use [`fontCssApi()`](#another-css-api-fontcssapi)
with the full template, the families the proxy serves with their CSS2 axes,
and the default substitutes:

```ts
import { fontCssApi, googleFontsSubstitutions } from '@silurus/ooxml';

const TEXT = 'ital,wght@0,400;0,700;1,400;1,700';
fonts: [fontCssApi({
  urlTemplate: 'https://proxy.example/google-fonts/css2?family={family}:{query}&display=swap',
  families: { Carlito: TEXT, Caladea: TEXT, 'Noto Sans JP': 'wght@400;700' },
  substitutes: googleFontsSubstitutions,
})]
```

List every family the proxy should serve: CSS2 rejects an axis query a family
does not have, and without `families` every demanded family is requested.

### Substitutes for `googleFonts()`

`googleFonts()` uses `googleFontsSubstitutions` by default: the
[Office substitutions](#substitutions) followed by Google's names for the
native Noto CJK families (`Noto Sans CJK JP` → `Noto Sans JP`, …). Pass
`substitutes` to replace that list. To extend it, spread it:

```ts
import { googleFonts, googleFontsSubstitutions } from '@silurus/ooxml';

fonts: [googleFonts({
  substitutes: [...googleFontsSubstitutions, { for: 'Corporate Sans', use: 'Open Sans' }],
})]
```

`googleFonts()` requests only the families in its catalog (the Office
stand-ins, the Noto script families and a few popular web fonts), so a
substitute's `use` must be one of them. For other Google families, use
`fontCssApi()` with your own `families` list.

### Another CSS API: `fontCssApi()`

```ts
import { fontCssApi, officeFontSubstitutions } from '@silurus/ooxml';

fonts: [fontCssApi({
  urlTemplate: 'https://fonts.example/css?family={family}:{query}',
  query: '400,700',
  substitutes: officeFontSubstitutions,
})]
```

- `urlTemplate` is an absolute `http(s)` URL. It must contain `{family}` and
  may contain `{query}`.
- `{family}` is percent-encoded, with `+` for spaces.
- `{query}` is inserted verbatim from `query` or from the `families`
  allow-list, never from document content. `query` is required when the
  template has `{query}`, unless every `families` entry supplies its own query.
- `families` is an optional allow-list (`{ 'Open Sans': true, Lato: 'wght@400' }`).
  When present, only those families are requested.
- `credentials`: optional, see [Network, CORS and cookies](#network-cors-and-cookies).

The source sends one stylesheet request per served family. Without
`families`, every demanded family is requested, and an error response (such as
HTTP 400 for a family the API does not serve) is a failure that makes the family
final (see [Chaining](#chaining)). List the served families when the source is
not the last in the chain.

### A self-hosted stylesheet: `fontCss()`

```ts
fonts: [fontCss({ url: '/fonts/office.css', substitutes: officeFontSubstitutions })]
```

The stylesheet may describe any number of families. Documents that use the
same configuration share one fetch while any of them is open. Only `@font-face`
rules with `url()` sources are read. `local()` sources, `@import` and
everything outside `@font-face` are ignored, and no script in the page runs.
`credentials` is optional, as for `fontCssApi()`.

#### Hosted font services

A provider's static CSS, such as an Adobe Fonts web project, a Monotype or
Morisawa static kit, Bunny Fonts or a Fontsource CDN stylesheet, works with
`fontCss()` (one stylesheet) or `fontCssApi()` (one request per family). The
library only reads the `@font-face` `url()` rules. It does not run the
provider's embed code (`<script>` or JavaScript kits), and it sends no usage
or page-view tracking. Loading the embed code, any usage reporting the
provider requires, and the font licences (including whether the licence
covers rendering in a document viewer) are your application's responsibility.

Services that subset fonts per page text at run time (for example
TypeSquare, FONTPLUS or Adobe Fonts dynamic CJK kits) are not supported. Such
services generate faces for the text their script sees in the page, which a
canvas viewer does not expose to them. Serve static faces for those families,
for example with `unicode-range` subsets, or answer with `fontResolver()`.

### Bundled files: `fontFiles()`

```ts
fonts: [fontFiles({
  faces: [
    { family: 'Carlito', url: '/fonts/Carlito-Regular.woff2', weight: '400' },
    { family: 'Carlito', url: '/fonts/Carlito-Bold.woff2', weight: '700' },
    { family: 'Noto Sans JP', url: '/fonts/NotoSansJP-kana.woff2', unicodeRange: 'U+3000-30FF' },
    { family: 'Brand', url: 'data:font/woff2;base64,…' },
    { family: 'Brand Display', data: brandDisplayBytes },
  ],
  substitutes: officeFontSubstitutions,
})]
```

Each face has exactly one of these:

- `url`: an `http(s)`, `blob:` or `data:` URL.
- `data`: a non-empty `ArrayBuffer`. Views, `SharedArrayBuffer`s and
  resizable buffers are rejected.

`weight`, `style`, `stretch` and `unicodeRange` take standard CSS
`@font-face` descriptor values. A `data:` URL is accepted and its decoded size
counts against the same limits as a `data` buffer. `credentials` is optional,
as for `fontCssApi()`.

Relative `url` values in `fontFiles()` and `fontCss()` are resolved against
the page's base URL when the factory is called. Build these descriptors on the
page, not in a worker.

### Authenticated or private services: `fontResolver()`

```ts
import { fontResolver, googleFonts, officeFontSubstitutions } from '@silurus/ooxml';

const privateFonts = fontResolver(async (demand, { signal }) => {
  const response = await fetch('/api/fonts/resolve', {
    method: 'POST',
    body: JSON.stringify({ families: demand.families }),
    signal,
  });
  return response.json(); // FontFaceDeclaration[]: { family, url | data, … }
}, { substitutes: officeFontSubstitutions });

fonts: [privateFonts, googleFonts({ cssOrigin: 'https://fonts.internal.example' })]
```

The function runs in the page that created it. In worker mode, the worker
asks the page to run it, so tokens and network clients stay in your
application. The worker never fetches the URLs the function returns: once a
family is claimed, the page fetches its `http(s)` and `blob:` faces and sends
the worker the bytes, so cookies, HTTP authentication and Referer checks see
the page in both modes. The page fetches only URLs the function returned for
that document. In main mode the browser loads the URLs, as before.

The `demand` argument contains:

- `families`: the authored families the host does not have and that no
  earlier call already answered.
- `fallbackFamilies`: the library's script fallback families for the text,
  such as Noto families for CJK.

Return every face you have for each family you answer: all weights, styles
and `unicode-range` subsets. A family is resolved once per document, so a face
you leave out is not asked for again. The renderer picks the weight and style
of each run, and loads only the `unicode-range` subsets of the characters it
paints. A resolver may answer only some families. With `substitutes`, a
family may also be answered with faces of its substitute's `use` family.

The function returns `FontFaceDeclaration` objects with the same fields as
`fontFiles()` faces. URLs must be absolute.

Documents that need the same families from the same function at the same time
share one call. `signal` is aborted only when no document is waiting for that
call any more: every document sharing it was destroyed or reached its
resolution deadline, or, in worker mode, the worker cancelled the request.
Destroying one of several documents that share a call does not abort it.

If the answer is not an array or contains a malformed declaration, every family
of that call keeps its host fallback, and a `console.warn` reports the problem.
If the function throws, the same happens and the warning includes the error
message. More than 2,048 declarations or a face over 30 MiB is a
[limit](#resource-limits) breach instead, which rejects the load.

## Chaining

Sources are consulted per family, in option order:

1. Each source tries the family itself, then the `use` family of each of its
   matching `substitutes`.
2. The first source that has a face for the family claims it. Later sources
   are never consulted for that family.
3. A source that does not serve the family passes it to the next source.
   Examples: a family outside a `fontCssApi` allow-list, a family missing from a
   loaded stylesheet or a file list, or a family a resolver did not answer.
4. A source that tries a family and fails also makes the family final, with no
   retry anywhere else. Failures include an unreachable or empty stylesheet, an
   HTTP error, a resolver error and the resolution deadline (see
   [Timing](#resolver-buffers-and-timing)). A `fontCss` stylesheet that fails
   to load fails every family it is asked for.
5. A face that is claimed but fails to load also keeps its host fallback. It
   does not re-enter the chain.

Each failure keeps the family on the fonts installed on the viewer's machine.
Each font load step prints one `console.warn` listing the families that newly
failed. There is no strict mode
and no callback. The library therefore never retries a failing mirror against
the public Google origin: it requests Google's stylesheets only when
`googleFonts()` without `cssOrigin` is in the list. Where a working mirror's
redirects and font URLs lead is up to the mirror (see
[the mirror section](#a-google-compatible-mirror-or-proxy-googlefonts-cssorigin-)). Failures are not cached, so a document opened later
retries the same source.

Within one document a family is decided once. A family that a source claimed
or failed is final. A family that no source served is asked again, through the
whole chain, only when later text in the document proves a substitute script
that its earlier request excluded. For example, a family first seen in Latin
text and later in Arabic text is requested again with its Arabic-only
substitutes. A `fontResolver` is not called again for it: its first answer is
claimed again under the wider scope.

What paints, highest first:

1. A font embedded in the document.
2. An installed face with the authored name. Installed families are never
   requested from a source.
3. A source face for the authored family.
4. A source face for a declared substitute.
5. The host's fallback chain.

## Substitutions

`officeFontSubstitutions` (`OFFICE_FONT_SUBSTITUTIONS` inside the library) is
the library's list of openly licensed stand-ins for Office faces:

- Calibri → Carlito
- Cambria → Caladea
- Franklin Gothic Book and Franklin Gothic Medium → Libre Franklin
- several Arabic faces → Noto Naskh Arabic or Noto Sans Arabic, for Arabic
  text only

Any source can use it, or a list of your own, through `substitutes`:

```ts
substitutes: [...officeFontSubstitutions, { for: 'Corporate Sans', use: 'Inter' }]
```

An entry is `{ for, use, script? }`:

- `for`: the authored family, matched case-insensitively.
- `use`: the family the source serves.
- `script`: optional. `'arabic'` limits the substitute to Arabic text.

The renderers use a substitution only when its face actually loaded.
`googleFonts()` already includes `officeFontSubstitutions`, plus Google's names
for the Noto CJK families.

## Requirements and limits

### Network, CORS and cookies

- Fonts from another origin need CORS headers (`Access-Control-Allow-Origin`).
  This applies to stylesheets and font files, as with Google Fonts today.
- `fontCssApi()`, `fontCss()` and `fontFiles()` take `credentials`:
  - `'same-origin'` (the default): stylesheets are fetched with it, and font
    files are loaded by the browser, which also sends credentials to the same
    origin only (the CSS font-loading rules).
  - `'omit'`: no cookies or HTTP authentication are sent at all. The library
    fetches the stylesheet and the font files itself, with
    `credentials: 'omit'`. It fetches every face of a claimed family when the
    family is claimed, not only the `unicode-range` subsets the text needs.
    Documents in the same realm share each fetched file while any of them is
    open, and the bytes count like `data` faces (see
    [Resource limits](#resource-limits)).
  - `'include'` and custom headers are not supported.
- `googleFonts()` takes no `credentials`: its requests go to another origin,
  which receives no cookies with the default.
- `fontResolver()` URLs are loaded with the CSS font-loading rules: by the
  browser in main mode, and by the page with `credentials: 'same-origin'` in
  worker mode.
- Cookies are never sent to another origin. For an authenticated service on
  another origin, use `fontResolver` and return `data` bytes or a signed URL.
- A hosted source receives the end user's IP address and User-Agent.

### Content Security Policy

- Stylesheet origins (`fontCssApi`, `fontCss`, `googleFonts`) need
  `connect-src`. The library downloads the stylesheet with `fetch()`; it does
  not inject a `<link>`.
- Font file origins need `font-src`. Add `data:` or `blob:` to `font-src` when
  you use those URLs. Font files the library downloads itself (a source with
  `credentials: 'omit'`, and resolver URLs in worker mode) need `connect-src`
  as well.
- `data` buffers are not fetched, so they need no `font-src` entry.
- For Google Fonts, the policy is `connect-src https://fonts.googleapis.com`
  and `font-src https://fonts.gstatic.com`. A mirror uses its own two origins.

### Resolver buffers and timing

- A `data` buffer belongs to the library once it is passed to `fontFiles()` or
  returned from a resolver. Do not modify it. Documents in the same realm may
  share it. In worker mode the worker receives a copy, and your buffer is not
  detached.
- A font load step runs each time a document needs new fonts or characters:
  for example once per DOCX document, once per PPTX slide that adds them, and
  once per XLSX workbook plus once per sheet that adds them. Layout waits for
  the step to finish.
- Each step has two phases, each with its own 15 s deadline:
  1. **Resolution** walks the sources in order. One deadline covers the whole
     walk, so a slow first source leaves less time for the next. Families
     still unresolved when it expires fail and keep their host fallback.
  2. **Face loading** fetches the faces the library fetches itself (see
     [Network, CORS and cookies](#network-cors-and-cookies)) and waits for
     the faces to load (`FontFace.load()`). Faces not loaded when it expires
     keep their host fallback for that step.

  A step can therefore take up to about 30 s, for example 14 s of resolution
  followed by 14 s of face loading.
- In worker mode the page owns the deadline of a resolver call. The worker
  sends the time left of its resolution deadline with each request.
- A face file the library fetches itself (a resolver face in worker mode, a
  source with `credentials: 'omit'`) is fetched in the face-loading phase,
  within its deadline, at most six at a time. A face that cannot be fetched
  fails like a face that cannot load: it is left out with a `console.warn`,
  and the family is not passed to another source or substitute. For XLSX in
  worker mode, the viewer and the worker share each fetch and its outcome.
- A resolver is called only for families that no earlier call answered and no
  live cache entry covers. Within a document that means once per family. For
  XLSX in worker mode, the
  viewer and the worker share the call. Documents open at the same time share
  answers from the same function. A document opened after every earlier one
  was closed calls the function again.
- In worker mode, the worker waits silently while the resolver runs. A slow
  resolver combined with a short `workerTimeoutMs` can make the load time out.

### Resource limits

Font sources are application configuration, so they have fixed limits. These
limits are not part of `resourceLimits`.

| Limit | Value |
| --- | --- |
| Stylesheet size | 1 MiB per stylesheet |
| Declarations | 2,048 per document |
| Requested families | 256 per document |
| One face (`data`, the decoded body of a `data:` URL, or a font file the library downloads) | 30 MiB |
| Retained `data` bytes | 256 MiB per realm, across open documents |
| URL length (non-`data:` URLs and composed `fontCssApi` requests) | 8 KiB |

Exceeding a limit fails closed. The error is an `OoxmlError` with
`code === 'font-source-limit'`, and `error.violation` holds `{ resource:
'font-source', quota, metric, limit, observed }`. Limits that can be checked
up front make the factory throw immediately. Otherwise the error rejects
`load()`, or the slide or sheet operation that needed the fonts, and that
document's later font loads reject too. A limit breach means a misconfigured
or hostile source, not a network problem.

The library does not count the bytes of `url` faces the browser downloads.
Font files the library downloads itself (see
[Network, CORS and cookies](#network-cors-and-cookies)) count like `data`
faces. A document whose painted text has more than 65,536 distinct
characters loads every subset; that is not an error.

## Licences

The library ships no font binaries. Your application is responsible for the
licences of the fonts it serves, and for any embed code, usage tracking or
other terms a font provider requires (see
[Hosted font services](#hosted-font-services)). Google Fonts' Carlito, Caladea, Libre Franklin
and Noto families are licensed under the SIL Open Font License.

## Removal in 0.91

0.91.0 removes `useGoogleFonts` from `LoadOptions`. Passing it then throws a
`TypeError`. Replace it with `fonts` before upgrading.
