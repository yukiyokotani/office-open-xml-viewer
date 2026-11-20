---
name: adversarial-review
description: Adversarially review an office-open-xml-viewer OOXML feature or bug-fix diff before it is declared complete or a draft PR is marked ready — specification fidelity, heuristics, bounded inference, performance, module boundaries, cross-format wiring and regression evidence. Use for the mandatory review in AGENTS.md, whether reviewing your own diff with a fresh reviewer or someone else's.
---

# Adversarial review

`AGENTS.md` ("Mandatory Adversarial Review") requires this review for every
OOXML feature or bug fix. Review the complete change against `main`, not only
the latest commit, and read the owning modules and tests, not only the diff.
Report only evidence-backed findings, each with an exact code reference or a
reproducible command.

## Checklist

- **Specification fidelity:** verify against the applicable ECMA-376 /
  ISO-29500 rule and Microsoft extension notes. Distinguish normative
  behavior, documented Office extensions, and implementation policy. Do not
  claim specification support from sample output alone.
- **No unexplained heuristics:** search for sample-specific branches, magic
  thresholds, empirical scale factors, font-name or path checks, and renderer
  guesses. Remove them, or justify them under "Compatibility rules for implicit
  Office behavior" in `AGENTS.md`. A passing VRT is not evidence that a heuristic is correct.
- **Bounded inference:** state the observed Office behavior, the tested range,
  and the limits of the compatibility claim; keep the implementation no
  broader than the evidence.
- **Performance and resource safety:** review asymptotic work, allocation
  size, decoded and retained memory, cache ownership/eviction, concurrency,
  worker transfer cost, repeated parsing or decoding, hot paths, and failure
  cleanup. Reject unbounded fan-out or whole-document work when page-local
  work is sufficient. Add budget, concurrency, cache, or stress tests where
  the risk is material.
- **Module boundaries:** shared OOXML concepts and pure rendering/resource
  primitives live in `ooxml-common` or `core`; DOCX, XLSX, and PPTX keep only
  format-specific parsing, layout, orchestration, and painting. Check for
  duplicated policy, false abstractions, package cycles, and functions that
  mix unrelated responsibilities.
- **Cross-format wiring:** for every shared capability, inspect the DOCX,
  XLSX, and PPTX parser/model, public options, viewer, renderer, cache/loader,
  worker protocol, error propagation, documentation, and tests. Document the
  reason when one format is intentionally excluded.
- **Regression evidence:** run focused behavioral tests plus the broadest
  proportionate cross-package checks. For renderer changes, use the previous
  renderer as the self-VRT oracle and keep Office-reference fidelity
  evaluation separate. Investigate every unexplained difference.
- **DOCX architecture:** after major DOCX layout, pagination, measurement,
  paint, parser-model, worker, or compatibility changes, also run the
  `docx-architecture-audit` skill.

## Verdict

Classify findings by severity. For each one, state the violated rule, the
evidence, the affected behavior, and the smallest coherent fix; separate a
proven defect from a question or missing evidence. End with `PASS` only when
no unresolved finding remains and every required command passes; otherwise
`FAIL` with the blockers. Never pass a change on green tests alone. Fixes made
in response go back through another round until `PASS`.
