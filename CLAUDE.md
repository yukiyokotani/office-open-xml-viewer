@AGENTS.md

## Claude-specific

### Claude Autonomous Work

Between 01:00 and 09:00 local time, Claude sessions may proceed without asking
for confirmation for non-destructive work:

- code changes,
- tests,
- WASM builds,
- feature-branch commits and pushes,
- package scripts.

Still ask before destructive operations, reference image updates, direct changes
to `main`, or anything that conflicts with `AGENTS.md`.

### Commit Attribution

When Claude authors a commit, include a co-author trailer with the actual model
used for that task, for example:

```text
Co-Authored-By: Claude <model-name> <noreply@anthropic.com>
```

If Claude starts or directs Codex for the work, include both the Claude/Fable
and Codex co-author trailers with the models actually used for that task, as
described in `AGENTS.md`.
