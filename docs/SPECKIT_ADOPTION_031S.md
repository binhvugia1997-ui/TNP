# PROMPT-031S adoption audit

Workflow/tooling only; no application implementation or release activity.

## Baseline and pre-install audit

- Audit date: **2026-10-10** (UTC).
- Authoritative base: `origin/arena/01a10c02-tnp` at
  `b3a054e5181cb0643ef454acd47b2421fcd21388`.
- Working branch: `arena/17ed8515-tnp`, initially exactly equal to that base.
  The shallow checkout initially lacked the remote-base ref; an explicit fetch of
  `refs/heads/arena/01a10c02-tnp:refs/remotes/origin/arena/01a10c02-tnp` resolved it.
- Pre-install `git status --short`, `git diff --binary`, and
  `git diff --cached --binary`: **empty (0 bytes)**. No user edits were overwritten.
- `.specify/`, `.agents/`, `AGENTS.md`, `specs/`, and `.codex/` were absent. Searches of
  tracked content and untracked workflow names found no prior Spec Kit metadata.
  Neither `specify` nor `uv` was installed; no CLI was downgraded.
- PR #7 was inspected read-only: OPEN, base `arena/01a10c02-tnp`, head
  `arena/28cc318f-tnp` at `45a59d776ccd7226f5c26e9e358426023e7e2c5b`.
  No PROMPT-030 commit was imported and no PR #7 action was performed.

**Version discrepancy:** the requested “Version 1.3.4 / Build 017 unchanged” cannot be
asserted as the version of this authoritative base. Its canonical `app/__init__.py`
contains **Version 1.3.2 / Build 015**. That file and all version/build metadata are
preserved byte-for-byte; this task does not correct the discrepancy, bump versions,
or import a different branch's application changes.

## Official installation and provenance

Upstream: <https://github.com/github/spec-kit>, tag **v1.1.3**, resolved commit
`0282d03eb6fc3091b3deaa3a855d4030dc40cd6c`. CLI result: `specify 1.1.3`.
The CLI package's VCS provenance confirms both the requested tag and resolved commit.

```sh
uv tool install specify-cli --from git+https://github.com/github/spec-kit.git@v1.1.3
specify --version
specify init --here --force --integration codex --script ps --ignore-agent-tools --non-interactive
specify extension --help
specify extension add --help
specify extension add bug
```

Initialization used the official existing-directory flow and version-matched bundled
assets, with **Codex skills** and **PowerShell** scripts. Immediate post-init and
post-extension `git diff` / `git diff --stat` were empty for existing tracked files;
only new workflow directories appeared. No runtime changes were accepted.

The bundled **Bug Triage Workflow 1.0.1** was installed by the v1.1.3 CLI, exposing
`speckit.bug.assess`, `speckit.bug.fix`, and `speckit.bug.test` as Codex skills
`$speckit-bug-assess`, `$speckit-bug-fix`, and `$speckit-bug-test`. Its registry labels
the bundled install `source: local`; the payload matches the official CLI's bundled
extension, not a hand-authored local substitute. It installs no hooks.

All upstream skills/templates/scripts are unmodified. The extra generated
`speckit-taskstoissues` skill is retained for upstream fidelity, but TNP does not use
or depend on it. The upstream MIT notice is copied verbatim from the pinned tag to
`.specify/UPSTREAM-LICENSE`.

## Constitution execution record

Read the full checked-in `speckit-constitution` skill and the initialized scaffold.
The provided TNP principles were the constitution input; other installation/documentation
work was performed separately, not as application work within that skill's scope.
Before/after hook inspection found `hooks: {}`.

The official `specify preset resolve constitution-template` command succeeded and
selected `.specify/templates/constitution-template.md` from **core** (no contributing
preset/project override; the bug extension supplies no templates). The resolved scaffold's
heading structure and governance/version/date fields were used for the TNP constitution.
Initial version **1.0.0** replaces the unratified example scaffold with eleven principles,
brownfield scope, workflow gates, and amendment rules. There are no deferred placeholders.
The skill's temporary Sync Impact Report was reviewed and removed before staging;
no versioned template or skill was rewritten.

> Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

**Explicit execution deviation:** `pwsh` / PowerShell was unavailable. An optional
PowerShell download from the official GitHub release was blocked at the sandbox's
release-asset host boundary. Therefore the skill's exact
`.specify/scripts/powershell/resolve-template.ps1 constitution-template -Json` step was
**not executed**; the official CLI resolver above was used instead. Full native skill
execution and the PowerShell wrapper remain unvalidated. No replacement script was
fabricated, and this is not a claim of exact end-to-end skill execution.

## Validation

| Check | Observed result |
| --- | --- |
| `specify --version` | Exit 0; `specify 1.1.3` |
| `specify check` | Exit 0; “Specify CLI is ready to use!”; Codex and other agent executables not found (not proof of native execution) |
| `specify integration list` | Exit 0; `codex` installed and default; no other integration installed |
| `specify extension list` | Exit 0; bug 1.0.1 enabled; 3 commands, 0 hooks, priority 10 |
| `specify integration status` | Exit 0; OK; 0 modified/missing managed files, 0 invalid paths/unchecked manifests |
| `specify artifact list --json` | Exit 0; official discovery includes the core commands and all three installed bug commands |
| `specify preset resolve constitution-template` | Exit 0; official core constitution template selected |
| Read-only workflow/integrity smoke checks | PASS; 9 required core + 3 bug skills discovered, all 6 PowerShell scripts present, JSON/YAML and guide links valid, all 22 managed hashes match; 39/40 generated files untouched (only constitution customized); bundled bug payload identical |
| Git whitespace and scope checks | PASS; `git diff --check` and `git diff --cached --check` exit 0; all 44 staged paths are additions; authoritative-base comparison shows zero changes in `app/`, `frontend/`, `tests/`, `tools/`, `backup/`, or any existing tracked file |

The smoke test performed discovery and integrity checking only. No fake feature,
historical spec, or bug case was created; `specs/`, `.specify/bugs/`, and
`.specify/feature.json` are absent. No temporary smoke artifacts remain in the repository.

To reproduce the committed scope/whitespace proof from the adoption branch:

```sh
git diff --name-only origin/arena/01a10c02-tnp...HEAD
git diff --exit-code origin/arena/01a10c02-tnp...HEAD -- app/ frontend/ tests/ tools/ backup/
git diff --check origin/arena/01a10c02-tnp...HEAD
```

The first command must list only the classified workflow files below; the other two
must exit 0 with no output. Application extraction, renderer, Excel, Portable runtime,
and the immutable backup are outside this change.

### Not validated / not performed

- Native Arena/Codex invocation, PowerShell script execution, or a real feature/bug
  workflow from start to finish. Script presence and references are statically checked.
- Application regression suite or frontend build (no application changes in this task).
- Windows acceptance, real PowerPoint COM, a Windows Portable/PyInstaller build, or
  PROMPT-030 acceptance. Linux CLI checks establish development-tooling evidence only.
- No `BUILD_AND_PUBLISH.bat`, LAN publish, Portable release, production `version.json`
  update, application version/build bump, stable-backup change, or PR merge.

## Complete changed-file classification

All paths below are new development-workflow files. The only project customization of
an initialized upstream file is `.specify/memory/constitution.md`. No existing TNP file
is modified. The commit/PR delivery report provides the final SHA and remote PR link.

**44 files added; zero existing files changed.**

| File | Classification |
| --- | --- |
| `.agents/skills/speckit-analyze/SKILL.md` | Official generated Codex core skill |
| `.agents/skills/speckit-bug-assess/SKILL.md` | Official generated Codex bug skill |
| `.agents/skills/speckit-bug-fix/SKILL.md` | Official generated Codex bug skill |
| `.agents/skills/speckit-bug-test/SKILL.md` | Official generated Codex bug skill |
| `.agents/skills/speckit-checklist/SKILL.md` | Official generated Codex core skill |
| `.agents/skills/speckit-clarify/SKILL.md` | Official generated Codex core skill |
| `.agents/skills/speckit-constitution/SKILL.md` | Official generated Codex core skill |
| `.agents/skills/speckit-converge/SKILL.md` | Official generated Codex core skill |
| `.agents/skills/speckit-implement/SKILL.md` | Official generated Codex core skill |
| `.agents/skills/speckit-plan/SKILL.md` | Official generated Codex core skill |
| `.agents/skills/speckit-specify/SKILL.md` | Official generated Codex core skill |
| `.agents/skills/speckit-tasks/SKILL.md` | Official generated Codex core skill |
| `.agents/skills/speckit-taskstoissues/SKILL.md` | Official generated extra skill; retained, not used by TNP |
| `.specify/.gitignore` | Official generated configuration, registry, or integrity metadata |
| `.specify/UPSTREAM-LICENSE` | Verbatim pinned upstream MIT notice |
| `.specify/extensions.yml` | Official generated configuration, registry, or integrity metadata |
| `.specify/extensions/.registry` | Official generated configuration, registry, or integrity metadata |
| `.specify/extensions/bug/README.md` | Official bundled bug extension payload |
| `.specify/extensions/bug/commands/speckit.bug.assess.md` | Official bundled bug extension payload |
| `.specify/extensions/bug/commands/speckit.bug.fix.md` | Official bundled bug extension payload |
| `.specify/extensions/bug/commands/speckit.bug.test.md` | Official bundled bug extension payload |
| `.specify/extensions/bug/extension.yml` | Official bundled bug extension payload |
| `.specify/init-options.json` | Official generated configuration, registry, or integrity metadata |
| `.specify/integration.json` | Official generated configuration, registry, or integrity metadata |
| `.specify/integrations/codex.manifest.json` | Official generated configuration, registry, or integrity metadata |
| `.specify/integrations/speckit.manifest.json` | Official generated configuration, registry, or integrity metadata |
| `.specify/memory/.constitution-template.json` | Official generated configuration, registry, or integrity metadata |
| `.specify/memory/constitution.md` | TNP constitution, completed from official scaffold |
| `.specify/scripts/powershell/check-prerequisites.ps1` | Official PowerShell workflow script |
| `.specify/scripts/powershell/common.ps1` | Official PowerShell workflow script |
| `.specify/scripts/powershell/create-new-feature.ps1` | Official PowerShell workflow script |
| `.specify/scripts/powershell/resolve-template.ps1` | Official PowerShell workflow script |
| `.specify/scripts/powershell/setup-plan.ps1` | Official PowerShell workflow script |
| `.specify/scripts/powershell/setup-tasks.ps1` | Official PowerShell workflow script |
| `.specify/templates/checklist-template.md` | Official unmodified template |
| `.specify/templates/constitution-template.md` | Official unmodified template |
| `.specify/templates/plan-template.md` | Official unmodified template |
| `.specify/templates/spec-template.md` | Official unmodified template |
| `.specify/templates/tasks-template.md` | Official unmodified template |
| `.specify/workflows/speckit/workflow.yml` | Official bundled SDD workflow |
| `.specify/workflows/workflow-registry.json` | Official generated configuration, registry, or integrity metadata |
| `AGENTS.md` | Short TNP agent entry point |
| `docs/SPECKIT_ADOPTION_031S.md` | Adoption provenance, audit, limitations, and file classification |
| `docs/SPECKIT_ARENA.md` | TNP Arena workflow, recovery, and fallback guide |
