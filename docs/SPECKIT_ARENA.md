# Spec Kit in Arena Web

TNP adopts official [GitHub Spec Kit v1.1.3](https://github.com/github/spec-kit/tree/v1.1.3)
with the **Codex skills integration**, **PowerShell scripts**, and bundled **bug 1.0.1**
extension. This is development workflow tooling, not an application dependency.
`app/`, `frontend/`, `tests/`, and `tools/` remain the real source. Spec Kit does **not**
create another copy of the application.

## Start or resume a session

1. Read [the constitution](../.specify/memory/constitution.md), then this guide.
2. Inspect Git status, the assigned Arena session branch, and the user-specified
   authoritative base. Stay on the session branch; never let a workflow switch branches.
3. Read the bounded work's committed artifacts before editing code. Existing source,
   tests, Git history, and `PROGRESS_LOG.md` remain brownfield evidence.
4. Spec Kit governs **new bounded work after PROMPT-031S**. Do not backfill PROMPT-001
   through PROMPT-030, invent historical specs, or produce a giant application spec.
   PROMPT-030 / PR #7 remains independent of this adoption.

For a fresh environment, first inspect `specify --version` if installed. If absent,
install the pinned development tool (Python 3.11+ and `uv` required):

```sh
uv tool install specify-cli --from git+https://github.com/github/spec-kit.git@v1.1.3
specify --version
specify check
specify integration list
specify extension list
```

Do not silently downgrade a newer compatible CLI. Stop and report network/install
failures; do not fabricate replacement scaffolding. The checked-in infrastructure
already exists: **do not routinely re-run `init --force`**. For reconciliation/upgrades,
inspect `specify integration upgrade --help` and `specify extension update --help`, use
those official flows, and review the entire diff. Do not hand-edit upstream skills.

The original brownfield installation commands were:

```sh
specify init --here --force --integration codex --script ps --ignore-agent-tools --non-interactive
specify extension add bug
```

`--force` permitted initialization in a non-empty directory, not unrelated overwrites.
`--ignore-agent-tools` skipped only the coding-agent executable check. Required script
interpreters still matter: use PowerShell for `.specify/scripts/powershell/*.ps1` and
Python 3/PyYAML when the resolver requires them. Report missing prerequisites instead of
claiming execution. `specify check` reports available agent tools; exit 0 does not prove
native skills or PowerShell are available. See the [adoption audit](SPECKIT_ADOPTION_031S.md)
for this sandbox's limitations.

## Two execution modes

### Mode A — native skills available

Use the exact Codex invocation names emitted by this installation. These are **agent
prompt invocations**, not shell commands or `specify` CLI subcommands:

| Purpose | Native Codex invocation |
| --- | --- |
| Project governance | `$speckit-constitution` |
| Bounded specification | `$speckit-specify` |
| Resolve ambiguity | `$speckit-clarify` |
| Technical plan | `$speckit-plan` |
| Task breakdown | `$speckit-tasks` |
| Cross-artifact analysis | `$speckit-analyze` |
| Implementation | `$speckit-implement` |
| Remaining-work assessment | `$speckit-converge` |
| Requirements quality checklist | `$speckit-checklist` |
| Bug assessment | `$speckit-bug-assess` |
| Bug fix | `$speckit-bug-fix` |
| Bug verification | `$speckit-bug-test` |

Pass the bounded requirement to specify and the defect evidence plus `slug=<bug-slug>`
to bug assess. Pass that same slug to fix and test. The generated upstream
`speckit-taskstoissues` skill is retained unmodified but is **not** part of TNP's workflow;
TNP must not depend on that deprecated workflow.

### Mode B — native invocation unavailable

Arena Web may read skills without exposing native Codex invocation. The agent MUST:

1. Locate the appropriate checked-in `.agents/skills/speckit-*/SKILL.md` (including
   `speckit-bug-assess`, `speckit-bug-fix`, or `speckit-bug-test` for defects).
2. Read the complete skill instructions, prerequisites, and configured hooks.
3. Follow those instructions explicitly, including scope, evidence, and write limits.
   If a required step cannot run, identify the blocker; never simulate its success.
4. Create/update the **same Spec Kit artifacts** as native execution. Do not replace
   the workflow with a chat-only summary or a new parallel documentation system.
5. Include this exact disclosure in the report:

> Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

Never pretend a native skill or an unavailable script was executed. The constitution
continues to govern both modes; hooks are not permission to publish or change branches.

## Feature workflow

**requirement -> specify -> clarify when needed -> plan -> tasks -> analyze -> implement
-> converge -> tests -> commit/push/PR**

Use one bounded directory under `specs/<feature>/`. Keep the input requirement, `spec.md`,
`plan.md`, and checked task status in `tasks.md` in Git. Preserve generated supporting
artifacts (`research.md`, `data-model.md`, `contracts/`, `quickstart.md`, `checklists/`)
where the skills require them; planning examples are not a second application tree.

Analyze before implementation. Converge runs only after tasks and implementation: it
appends remaining work to `tasks.md`, not code or revised requirements. Complete those
tasks via implement, rerun converge as needed, and run the actual tests. Record commands,
results, environment, pending Windows acceptance, and convergence outcome in a bounded
verification note (e.g. `specs/<feature>/verification.md`), separate from converge's
append-only operation. Commit that evidence too. Neither checklists nor completed tasks
replace tests.

To resume an **existing** feature on a fresh checkout, explicitly select its committed
directory; do not infer it from the Arena branch name or invent a new feature:

```powershell
$env:SPECIFY_FEATURE_DIRECTORY = 'specs/<existing-feature-directory>'
# Read-only path discovery; this does not validate or implement the feature:
& .specify/scripts/powershell/check-prerequisites.ps1 -Json -PathsOnly
```

Replace the placeholder with a real directory. In nonpersistent tool shells, supply the
environment setting for each relevant call. `.specify/feature.json` is intentionally
ignored per-checkout state; the official scripts can recreate it. The spec directory
name is independent of the Git branch. No branch-changing git extension is installed.

## Bug workflow

**assess -> fix -> test -> verdict -> commit/push/PR**

- **Assess:** reproduce the defect, capture the source report/evidence, and identify the
  root cause. Write `.specify/bugs/<slug>/assessment.md`; mark hypotheses and blocked
  reproduction honestly. Do not edit runtime code during assessment.
- **Fix:** read that assessment, change the demonstrated cause only, add the planned
  regression tests, and record changes/deviations in the same directory's `fix.md`.
- **Test:** rerun the original case and relevant regression coverage; record commands,
  output, environment, verdict, and residual risks in `test.md`. The test skill writes
  the report, not source changes. Use `partial` (with not-run checks) when reproduction
  or critical acceptance is unavailable; never label it `verified` from a commit alone.

Commit all three reports with the bounded fix. In a new session, choose the existing
slug explicitly and read its reports; never overwrite or guess between multiple bugs.

## Finish safely

Review the constitution gates and `git diff --check`. Stage explicit, reviewed paths
(not `git add .`), commit the artifacts and authorized changes, push only the session
branch, and open a PR against the requested authoritative base. Leave acceptance and
merge to explicit authorization. Never present Linux evidence as Windows/PowerPoint COM/
Portable acceptance; never automatically publish, bump versions, or touch the immutable
stable backup. Do not commit credentials, local agent state, or sensitive source reports.
