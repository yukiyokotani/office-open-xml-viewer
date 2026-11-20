---
name: docx-architecture-audit
description: Audit the office-open-xml-viewer DOCX layout architecture for a single immutable layout-to-paint pipeline. Use after major DOCX layout, pagination, measurement, paint, parser-model, worker, or compatibility changes, and before declaring Issue #1037-style architecture work complete.
---

# DOCX Architecture Audit

Audit semantics that static identifier checks cannot prove. Report only
evidence-backed findings with an exact code reference or a reproducible command.

## Workflow

1. Read `AGENTS.md` and `docs/docx-layout-engine-redesign.md` completely.
2. Inspect the complete change against `main`; do not limit review to the latest
   commit. Read the owning modules and focused tests, not only the diff.
3. Consult the relevant local `spec/` material before judging OOXML behavior.
   Prefer ECMA-376 / ISO/IEC 29500, then the applicable Microsoft implementation
   note. Do not disclose private sample contents or paths.
4. Run the deterministic gates:

   ```bash
   mise exec -- node scripts/check-docx-layout-boundaries.mjs --final
   mise exec -- pnpm test:docx-boundaries
   mise exec -- pnpm test:docx-compatibility
   mise exec -- pnpm test:docx-public-api
   mise exec -- pnpm test:docx-package-build
   mise exec -- pnpm typecheck
   git diff --check
   ```

5. Audit the final dependency direction and runtime behavior:

   - one production layout algorithm exists for each supported feature class;
   - parser facts flow into immutable point-space layout and then
     measurement-free paint;
   - paint performs no measurement, shaping, style resolution, pagination, or
     parser-model access, including through transitive imports;
   - parser objects and retained layouts are not mutated or runtime-stamped;
   - no migration flag, silent legacy fallback, dry-render geometry collection,
     or transitional allowance remains;
   - body, stories, tables, text boxes, and supported floating content participate
     in `DocumentLayout` and preserve page/layer ownership;
   - main and worker use the same contracts and stable layout fingerprints;
   - compatibility decisions are isolated with exact evidence, while normative
     rules and unsupported diagnostics are not mislabeled as observations;
   - `renderer.ts` remains a thin adapter with no hidden layout or paint algorithm;
   - public declarations remain compatible;
   - shared OOXML primitives live in `core` or `ooxml-common` only when the
     abstraction is genuinely cross-format.

6. Run focused behavior tests for every suspicious path. Run the full repository
   test suite and browser conformance checks when declaring the architecture
   complete or when the touched surface crosses package/runtime boundaries.
7. Classify findings by severity. For each finding, state the violated invariant,
   exact evidence, affected behavior, and smallest coherent fix. Distinguish a
   proven defect from a question or missing evidence.
8. End with an explicit `APPROVE` only when no unresolved finding remains and
   every required command passes. Otherwise use `REQUEST_CHANGES` and list the
   blockers. Never approve based only on green tests.
