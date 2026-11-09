@AGENTS.md

## Claude Autonomous Work

Between 01:00 and 09:00 local time, Claude sessions may proceed without asking
for confirmation for non-destructive work:

- code changes,
- tests,
- WASM builds,
- feature-branch commits and pushes,
- package scripts.

Still ask before destructive operations, reference image updates, direct changes
to `main`, or anything that conflicts with `AGENTS.md`.
