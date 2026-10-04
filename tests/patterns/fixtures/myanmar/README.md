# Myanmar Canvas fixture

`paint.woff2` is a regular, static subset of Noto Sans Myanmar, renamed
Myanmar Paint Fixture. Its original copyright and SIL Open Font License are
in [OFL.txt](OFL.txt). It is used only by browser tests and is not a runtime font.

Upstream: [google/fonts at 8b0a1d0](https://github.com/google/fonts/tree/8b0a1d0f5983c89bc2b93f1b5fb55f9e252744b5/ofl/notosansmyanmar),
`NotoSansMyanmar[wdth,wght].ttf`, SHA-256
`7abbbfbe2514105d7ce94937aee3feb2ba89b73a256c8b77b5866bd9b83e32ec`.

With `fonttools[woff]==4.60.1`, run `python build-paint-font.py <upstream.ttf>`.
The generator fixes weight 400 and width 100, retains the listed Myanmar
base/marks and dotted circle with layout closure, and preserves GSUB, GPOS and
GDEF. Keeping the real shaping tables makes isolated-mark repair observable;
a cmap-only mock cannot detect the regression.
