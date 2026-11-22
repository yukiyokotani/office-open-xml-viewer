# Release

## Rules

Minor releases come from `main`; patch releases come from the `v0.x` release
branch with only the chosen fixes backported by PR. Before tagging, no page
may move farther from Office than in the previous release without an
explanation the user approves.

## Which branch

- **Minor release** (`0.x.0`): from `main`.
- **Patch release** (`0.x.y`): from the release branch `v0.x`, created from
  tag `v0.x.0`. Backport only the fixes chosen for the patch, as cherry-pick
  PRs against that branch; never push to it directly, and never cut a patch
  from `main` while `main` carries features. Verify each fix in the
  environment where it was reported before release. A backport PR goes through
  the same merge gate with the release branch in place of `main`: its VRT
  baseline is the head of the release branch, and `main` is never merged into
  it.

## Release gate: no regression against Office

Before tagging, run a release-level VRT of the candidate against the
**previous release tag** over the full demo and private corpora ([VRT](vrt.md), with
the tag's SHA as the baseline). Judge every changed page, sheet or
slide against its Office PDF and classify it as closer, farther or equal.

A change toward Office is fine: the previous release is not ground truth. A
page that matched Office in the previous release and is now farther from it is
a regression. It blocks the release unless it gets a fix, or an explained
reason the user approves (for example, the old match was coincidental and the
new rule is the correct general one).

## Checklist

Prepare a dedicated PR named `release/<version>` and do all release changes
there.

1. Update README screenshots in `docs/images/{pptx,docx,xlsx}.png` with
   `pnpm screenshots:readme`.
2. Run the gardening pass in [documentation](documentation.md#gardening-pass)
   for a minor release, and review README against the current implementation:
   - ESM-only package format,
   - install/import examples,
   - feature support table,
   - bundle-size notes,
   - deprecated API or stale version references.
3. Sync public API documentation:
   - `site/src/lib/api-reference.ts`,
   - `site/src/components/Capabilities.astro` when capabilities changed.
   - `site/src/pages/bundle-size.astro` with the latest production measurements.
     Measure with `node scripts/measure-bundle-size.mjs` after `pnpm build:wasm && pnpm build`.
4. Add a new top `CHANGELOG.md` section for the version being released:
   `## <version> — YYYY-MM-DD`.
5. Bump all versions consistently:
   - root `package.json`,
   - `packages/{core,pptx,xlsx,docx,markdown,node,vscode-extension}/package.json`,
   - `site/package.json`,
   - `packages/mcp-server/Cargo.toml`.

   After bumping the root version, run `pnpm sync:versions` to synchronize the
   companion packages `packages/legacy-office/package.json`
   (`@silurus/ooxml-legacy-office`) and `packages/opendocument/package.json`
   (`@silurus/ooxml-opendocument`), their exact `@silurus/ooxml` peer pins, and
   the compiled companion runtime version (`src/companion-version.ts`). Run
   `pnpm install --lockfile-only` and `pnpm check:versions`; CI and publishing
   reject any mismatch. All three npm packages publish from the same tag
   (`scripts/publish-packages.mjs`): the viewer first, then legacy-office, then
   opendocument; retries skip exact versions already present on npm.
   `pnpm build:wasm` builds one OpenDocument reader per `./<format>` export of
   the opendocument package through its `wasm:<format>` script, and
   `pnpm test:packed-packages` expects exactly those formats' assets in the
   tarball, so a new format needs only its export, script and build entry.
   `publish.yml` authenticates with the `NPM_TOKEN` secret (plus OIDC for
   provenance). Before the first release of a new companion package, the
   repository owner must make sure that token can create new packages in the
   `@silurus` scope (or publish the package once by hand and grant the token
   access to it); a token limited to existing packages fails with E404/E403
   after the viewer has already published, and the retry dispatch then
   publishes only the missing companions.
6. Run `cargo check -p ooxml-mcp-server` after bumping the MCP server so
   `Cargo.lock` follows.
7. Open PR `chore(release): <version>` against the release's base branch
   (`main` for a minor release, the release branch for a patch).
8. Merge through the merge gate in [Git, PRs and issues](git-and-prs.md): a
   merge commit with `gh pr merge <number> --merge --match-head-commit <sha>`.
9. Pull the release's base branch, create and push the annotated tag for the
   version being released: `git tag -a v<version> -m "v<version>"` and
   `git push origin v<version>`.
10. Create the GitHub Release with notes based on the changelog and a full
    changelog comparison link.

Reference images under `packages/*/tests/visual/references/` are outside the release
checklist unless the user explicitly asks to update them.

## Release notes

Follow the [documentation style](documentation.md#style): lead with the
visible outcome, say whether migration is required, and put any technical
qualification in a final `Technical note` section.

## VS Code extension

The VS Code extension version matches the npm library version. Pushing a `v*`
tag triggers `.github/workflows/publish-vscode-extension.yml` to publish the
extension. For manual packaging checks, use the workflow dispatch dry run.
