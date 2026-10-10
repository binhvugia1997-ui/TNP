# TNP Constitution

## Core Principles

### I. Source Data Integrity

AI MAY help locate information in a report. Application output MUST preserve the actual
source PPTX text and image evidence, with its provenance. Generated AI text MUST NOT
silently replace source evidence written to Excel. Finding information and authoring
replacement evidence are different operations; only the former is permitted in extraction.

### II. Image Evidence Semantics

Changes MUST preserve the existing production-improvement evidence rules. Selection MUST
exclude Before images, inspection/control-only images, temporary countermeasures,
verification/monitoring evidence, decorative images, and evidence owned by another
`ImprovementItem`. One slide MAY contain multiple independent `ImprovementItem`s;
implementations MUST NOT restore `one slide == one improvement`. Every item MUST retain
its own After-region ownership through selection, rendering, review, and export.

### III. Authored After Region

Rendered After evidence MAY include authored captions, borders, arrows, circles,
annotations, and whitespace when they belong to that After region. Changes MUST preserve
this authored context rather than reducing every region to isolated picture bytes.
Regions MUST NOT accidentally absorb unrelated items or Before content. Ownership and
region boundaries, not the mere presence of an annotation, determine inclusion.

### IV. Report Isolation

State from report A MUST NOT leak into report B. This applies to images, candidate
ownership, render caches, evidence bytes, and learning state wherever report isolation is
expected. Deliberately shared learning behavior MUST remain compatible with existing
contracts; it MUST NOT become a path for leaking report-specific evidence or ownership.
Changes affecting state MUST include cross-report regression coverage.

### V. Excel Transaction Safety

Changes MUST preserve atomic/report-level commit behavior and existing lock/retry
semantics. Where the transaction contract requires an all-or-safe result, processing MUST
NOT leave partial report output. Success and committed-output counts MUST describe what
was actually committed, not merely prepared in memory. Failure or cancellation MUST
preserve the contract's safe workbook state.

### VI. Cancellation Contract

The distinct meanings of `Dừng sau file hiện tại`, `Dừng tất cả`, `cancelled`, and `failed`
MUST be preserved. Stop-after-current MUST respect the current report's transaction
boundary; stop-all MUST retain cooperative cancellation and safe cleanup. Cooperative
cancellation MUST NOT be classified as a processing failure. UI, service, batch, and
reported-result behavior MUST remain consistent with existing cancellation tests.

### VII. Windows Acceptance

Linux sandbox success is development evidence only and MUST NOT be presented as Windows
acceptance. PowerPoint COM validation requires real Windows PowerPoint. Portable/
PyInstaller validation requires a real Windows build. Reports MUST state the environment,
checks actually performed, and missing acceptance evidence. No agent may infer these
acceptances from mocks, static inspection, a Linux run, or the existence of a commit.

### VIII. Test Before Claim

Bug work MUST follow `reproduce -> root cause -> fix -> regression test -> acceptance`.
Assessment MUST capture reproduction evidence and distinguish a demonstrated cause from a
hypothesis. Fixes MUST address the demonstrated cause without unrelated changes. Testing
MUST exercise the original case and relevant regression coverage before a verified
verdict. Missing reproduction or acceptance MUST remain explicitly pending, partial, or
not-run. A commit is never proof that a bug is fixed. Tests MUST NOT be weakened, skipped,
or have assertions removed merely to obtain green output.

### IX. Release Safety

Coding agents MUST NOT automatically run `BUILD_AND_PUBLISH.bat`, publish a LAN update,
modify production `version.json`, bump the application version/build, create a release
(including Portable), or merge an acceptance PR. Each action requires explicit user
authorization for that action. A request to implement, test, commit, push, or open a PR
is not release or merge authorization.

### X. Immutable Backup

`backup/ReportExtractor_v1.0.4_Build004_STABLE/` MUST NOT be modified, renamed, moved, or
deleted. It is immutable historical recovery evidence, not a development target.

### XI. Git Safety

Normal development MUST follow:

`authoritative base -> Arena session branch -> spec/assessment -> plan -> implementation
-> tests -> commit -> push -> PR -> acceptance -> authorized merge`.

Agents MUST confirm the requested authoritative base and work only on the branch assigned
to the current Arena session. They MUST push only that branch and target the requested PR
base, not implicitly `main`. Unrelated branches and PRs MUST remain untouched. Agents MUST
NOT automatically force-push, reset, or rebase shared history, merge `main`, or merge
acceptance PRs. Stage only reviewed, intentional paths; do not use `git add .`.

## Brownfield Scope and Compatibility

TNP is an existing application. `app/`, `frontend/`, `tests/`, and `tools/` remain the real
source; Spec Kit MUST NOT create another copy of the application. Existing source, tests,
Git history, and `PROGRESS_LOG.md` remain historical/brownfield evidence. Existing tests
and TNP compatibility behavior are authoritative constraints on new work; any intended
contract change MUST be explicit, justified, and reviewed, not an incidental regression.

Spec Kit governs NEW bounded work after PROMPT-031S. Agents MUST NOT retroactively create
specs for PROMPT-001 through PROMPT-030 or reverse-engineer the entire application into a
fake umbrella spec. PROMPT-030 / PR #7 is independent existing work, not part of this
adoption. This adoption MUST NOT import or validate its Portable/pythonnet runtime fix,
change application behavior, or change application version/build metadata.

## Workflow and Review Gates

- Before implementation, new bounded features MUST have a specification and plan; defects
  MUST have an assessment and remediation plan. Only an explicit user request for an
  emergency minimal change permits bypassing that order. Record the exception, bounded
  scope, and outstanding evidence; other safety and acceptance rules still apply.
- Feature workflow: requirement -> specify -> clarify when needed -> plan -> tasks ->
  analyze -> implement -> converge -> tests -> commit/push/PR. Convergence findings MUST
  return to implementation and verification; a task checkbox alone is not test evidence.
- Bug workflow: assess -> fix -> test -> verdict -> commit/push/PR. Assessment records the
  reproduction, evidence, and root cause; fix records the minimal changes; test records
  the original-case rerun, regression results, environment, and honest verdict.
- Plans MUST check every applicable principle and define test/acceptance gates. Existing
  tests take precedence over generic template suggestions that testing is optional.
- Commit the bounded requirements/specification, plan, task status, and verification
  evidence with the implementation. Feature artifacts belong in `specs/<feature>/`;
  bug reports belong in `.specify/bugs/<slug>/`. Evidence MUST be recoverable without prior
  chat history, and sensitive source reports/credentials MUST NOT be committed.
- Follow `docs/SPECKIT_ARENA.md` for native skills or the explicit checked-in-instructions
  fallback. Reports MUST distinguish reading/following instructions from native skill
  invocation; missing tools or unperformed checks MUST NOT be presented as successful.

## Governance

This constitution governs TNP Spec Kit work and its generated plans/tasks. No template,
optional extension, hook, or automatic workflow authorizes violating these principles.
Every PR MUST include a constitution compliance review, scoped changed-file evidence,
test results, and acceptance limitations. Conflicts MUST be raised before implementation
or release activity; agents MUST NOT silently weaken the constitution to pass a gate.

Amendments require a documented rationale, explicit user approval, review of affected
workflows, and a migration plan when existing commitments change. Constitution versions
use semantic versioning: MAJOR for incompatible principle removals/redefinitions, MINOR
for new or materially expanded principles, PATCH for non-semantic clarifications.
Amendment records MUST include ISO dates and any follow-up work. Upstream skills and
templates MUST NOT be hand-rewritten as a substitute for project governance.

Version 1.0.0 is the initial TNP constitution, not an application or Spec Kit CLI version.

**Version**: 1.0.0 | **Ratified**: 2026-10-10 | **Last Amended**: 2026-10-10
