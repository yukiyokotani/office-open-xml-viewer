# Office sample and test files

Office fixtures are organized by redistribution policy first, then by purpose.
The same layout is used by the DOCX, XLSX, and PPTX packages:

```text
packages/<format>/public/
├── demo/
│   ├── sample-1.<format>
│   ├── sample-1.pdf
│   ├── test-1.<format>
│   └── test-1.pdf
└── private/
    ├── sample-1.<format>
    └── sample-1.pdf
```

## Public website demos

`packages/<format>/public/demo/sample-N.<format>` is an official website demo.
These files should be concise product showcases and must be safe to
redistribute. The numbering is independent for each format and existing numbers
must not be reused.

The site currently copies `sample-1.docx`, `sample-1.xlsx`, and
`sample-1.pptx` into its public sample directory.

## Public regression tests

`packages/<format>/public/demo/test-N.<format>` is a redistributable regression
fixture. Public tests may be broad edge-case corpora or focused reproductions,
but every included asset and text fragment must be synthetic, freely licensed,
or otherwise safe to publish.

Use the next available `test-N` number within that format. Do not rename an
existing test when adding another one because test paths are durable identifiers
for bug reports and visual references.

## Private regression samples

`packages/<format>/public/private/sample-N.<format>` is a local-only fixture
that cannot be redistributed. Use the next available `sample-N` number within
that format. Private files are diagnostic evidence; durable regressions should
be reduced to a public synthetic `test-N` fixture whenever practical.

The entire `public/private/` directory is ignored. Never force-add files from
it, and never mention private filenames or document contents in public commit
messages or pull requests.

## Adjacent PDF references

Place a reference PDF next to its Office source using the same basename:

- `demo/sample-1.docx` → `demo/sample-1.pdf`
- `demo/test-1.pptx` → `demo/test-1.pdf`
- `private/sample-1.xlsx` → `private/sample-1.pdf`

Public PDFs under `demo/` are committed. Private PDFs under `private/` remain
local. Export references from the corresponding Microsoft Office application
at standard quality without manual scaling. Viewer- or LibreOffice-generated
PDFs are useful diagnostics but must not replace Microsoft Office ground truth.

Before committing a public PDF, verify every page or slide, confirm its basename
matches the source, and record the Office version, operating system, export
date, source SHA-256, and PDF SHA-256 in the pull request.
