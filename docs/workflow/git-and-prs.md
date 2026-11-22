# Git, PRs and issues

How work moves from a branch to `main`, and what may be committed.

## Standard Development Flow

1. Start from `main` and run `git pull --ff-only`.
2. If the checkout is dirty, preserve unrelated work before switching or
   pulling. Prefer `git stash push -u -m <clear-message>` for work that
   appears to belong to another session.
3. Create a feature branch named `<agent>/<topic>`, for example
   `codex/<topic>` or `claude/<topic>`.
4. When creating or entering a Git worktree, make the root checkout's ignored
   `spec/` directory available as a local symlink. For a worktree under
   `.claude/worktrees/<name>`, run `ln -s ../../../spec spec`; elsewhere, link
   `spec` to the main checkout's `spec/`. Do not copy or commit it.
5. Work with TDD for bug fixes: explore the failing sample or behavior, add a
   focused failing test, make it pass, and refactor only where it reduces real
   complexity.
6. Verify with the narrowest meaningful tests first, then broader checks when
   the touched surface warrants it.
7. Commit only the intended source/docs/test files. Do not commit private
   samples or generated local artifacts.
8. Push the feature branch and open the PR as a draft. Take it through the
   merge gate below, then merge with a merge commit. Do not squash.

## Worktrees and Parallel Sessions

- Before editing, make sure `spec` exists in the worktree and read the
  `AGENTS.md` of each package you will edit.
- A task scoped to one format edits that format's package. Reading other
  packages is fine. Edit `packages/core`, `packages/ooxml-common`, or another
  format only for a coherent fix described in Cross-Package Integration, or
  when the user asks for a cross-package change.
- Treat unexpected dirty files as work that belongs to the user or to another
  session. Do not revert or overwrite them; preserve them as in step 2 of the development flow.

## Git, PRs and Issues

- Never push directly to `main` or to a release branch. Use PRs.
- Open every PR with `gh pr create --draft`. Keep it draft, including through
  rebases, until the adversarial review has passed and the VRT is done.
- Merge gate, one PR at a time in dependency order: adversarial review rounds
  until PASS → one scoped VRT against the head of the PR's base branch, which
  is the latest `main` except for a backport ([VRT](vrt.md)) →
  `gh pr ready` → CI green on the current head → merge with
  `gh pr merge <n> --merge --match-head-commit <full sha>`. Copy SHAs from
  `gh pr view <n> --json headRefOid` output; never retype or expand one.
- Do not open PRs faster than they can land. Retarget and rebase the next
  stacked PR after each merge.
- A fix that makes some documents better and others worse is not merged by
  default; bring per-file evidence to the user. Layout time and bundle size
  are hard constraints, not prices for small fidelity gains.
- Fix regressions from merged work forward: keep the behavior the original
  change was for and fix the cause. Do not revert.
- Set-aside issues that are in scope get the `deferred` label and a one-line
  comment giving the reason; propose deferral to the user first. Skip
  `deferred` issues unless the user brings one back. Do not create release
  milestones.
- Close an issue once its fix is merged, with a comment naming the merge
  commit. Leave partly fixed issues open.
- When asking the user to decide, give each option's pros and cons and a
  recommendation.
- Before `git push`, set `git config http.postBuffer 524288000` if large
  pushes are expected.
- Commit messages follow the repository style: subject `fix(scope): ...`,
  `test(scope): ...`, `refactor(scope): ...`; the body explains the root
  cause, the fix, and verification, citing specification sections or observed
  Office behavior when relevant.
- PR titles, PR descriptions, commit messages, and public docs must be safe
  for an OSS repository:
  - no local absolute paths, usernames, home directories, machine names, or
    other personal environment details;
  - do not quote or summarize the concrete contents of private or copyrighted
    samples;
  - describe private-sample-driven work by the implementation problem, the
    OOXML/Office behavior mismatch, the class of affected documents, and the
    verification performed;
  - keep local-only sample filenames and visual details out of public PR text
    unless the user explicitly approves publishing them.
- Agent-authored commits include a co-author trailer naming the model
  actually used for the task, e.g.
  `Co-Authored-By: Codex <model-name> <noreply@openai.com>` or
  `Co-Authored-By: Claude <model-name> <noreply@anthropic.com>`. When one
  agent directs, delegates to, or reviews for another, include a trailer for
  each agent.

## Private Samples and Storybook

Private Office samples and local-only sample stories are gitignored and must
not be committed:

- `packages/{docx,xlsx,pptx}/public/private/<fmt>/`
- `packages/legacy-office/public/private/{doc,xls,ppt}/` (with paired OOXML
  controls)
- `packages/opendocument/public/private/{odt,ods,odp}/` (with their PDFs)
- `packages/*/src/*privateDemo.stories.ts`

Generated `packages/*/src/wasm/`, VRT screenshots and diffs stay local too.
A worktree may copy private stories/samples from the main checkout, without
`.DS_Store`, Office lock files, screenshots, VRT diffs, or private
reference images.

Storybook is unified at the repository root and serves
`packages/<fmt>/public/` at `/<fmt>/`; use those prefixes when fetching
samples. In sandboxed environments its port probe may fail with `EPERM`;
running it outside the sandbox is acceptable when the user asks to view it.
