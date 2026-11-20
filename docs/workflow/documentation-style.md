# Documentation style

Public docs and commit messages are written in English, in a plain, factual
tone: descriptive headings that name the content, no slogans or emotive copy.

README, docs, and Storybook public-facing code examples:

- use `as Type` assertions instead of postfix non-null assertions;
- use `canvas` for canvas-backed viewers (`DocxViewer`, `PptxViewer`) and
  `container` for container-backed viewers (`XlsxViewer`);
- do not document private/local-only sample paths as public API.

Announcements are concise and user-facing:

- lead with the visible outcome, grouping low-level fixes by the user problem
  they solve;
- state whether migration is required; when it is, give the minimum steps and
  name any changed default or behavior;
- omit implementation detail that does not change a user's decision; put a
  useful technical qualification in a final `Technical note` section;
- keep bundle measurements on the `/bundle-size` page, and do not link API
  reference details directly from release notes.
