# Third-Party Notices

`@silurus/ooxml-opendocument` is MIT-licensed (see [LICENSE](./LICENSE)).
Its JavaScript helpers are from this repository. Its WASM binary is built from
`opendocument-parser` and the shared model/common crates of this repository.
The viewer, renderers, fonts and optional rendering engines are supplied by the
separate `@silurus/ooxml` peer and are not bundled in this package.

## Rust / WebAssembly

The normal dependency graph of the `odt` reader on `wasm32-unknown-unknown` uses
third-party crates under the licenses below (procedural macros used during
compilation included; no development dependencies ship).

| License | Crates |
|---|---|
| MIT OR Apache-2.0 | bitflags, bumpalo, cfg-if, console_error_panic_hook, crc32fast, equivalent, fdeflate, flate2, gif, hashbrown, indexmap, itoa, js-sys, once_cell, png, proc-macro2, quote, roxmltree, serde, serde_core, serde_derive, serde_json, syn, typed-path, wasm-bindgen, wasm-bindgen-macro, wasm-bindgen-macro-support, wasm-bindgen-shared, weezl |
| MIT | quick-xml, simd-adler32, zip, zmij |
| (Apache-2.0 OR MIT) AND Unicode-3.0 | unicode-ident |
| 0BSD OR Apache-2.0 OR MIT | adler2 |
| MIT OR Zlib OR Apache-2.0 | miniz_oxide |
| MIT OR Apache-2.0 OR Zlib | zune-core, zune-jpeg |
| Unlicense OR MIT | memchr |
