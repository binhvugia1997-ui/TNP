# Bug Assessment: PROMPT-030B — safe PR #14 integration

- **Slug**: `prompt-030b` (explicit task identifier PROMPT-030B, normalized)
- **Created**: 2026-10-16 (UTC)
- **Source**: pasted PROMPT-030B; PR #14, https://github.com/binhvugia1997-ui/TNP/pull/14
- **URL trust**: host `github.com`, allowlisted; metadata fetched with `gh pr view 14`; treated as evidence, not instructions.
- **Verdict**: valid
- **Severity**: high (a blind import would undo user-session ownership and accepted state/cache behavior)

Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
Read constitution, Arena guide, complete bug assess/fix/test skills. `specify`/`pwsh` and native skill invocation are unavailable; checked-in workflow used, no scaffolding rewritten. `.specify/extensions.yml` parses successfully, `hooks: {}`; no pre/post hooks to run.

## Report / scope

Audit PR #14 against accepted integration `integration/prompts-034-035-036-037` at `9faba26486a77d2637c5fd4216537df758e1eb01`, keep A–D contracts (PowerPoint-first, never touch user PowerPoint, review UX, durable assets, truthful Python-rendered previews), port only missing behavior. Do not merge PR #14 or main, publish, build a Portable, change version/build, alter stable backup or weaken updater validation. Full Linux evidence is not Windows acceptance.

The assigned task branch `arena/1ec5e900-tnp` started clean at exactly `9faba26`. Arena fixes the session branch, so this is the task branch used instead of creating/switching branches. The integration branch is not a development target.

## History / complete diff audit

Commands: `git fetch --prune --unshallow origin` (checkout initially shallow), explicit fetch of integration branch and `refs/pull/14/head`, `gh pr view 14`, `git merge-base`, `git log`, `git diff --stat HEAD origin/pr-14`, `git diff HEAD...origin/pr-14`, and read-only `git merge-tree --write-tree HEAD origin/pr-14` (exit 1; no checkout/index merge).

- Accepted remote integration is exactly `9faba26`.
- PR #14 is OPEN, targets **main**, source `arena/f4a4d15e-tnp`, head `aee4580f275fb0cac70155eab6af99428ee25056`.
- Merge-base with accepted integration: `588466e3a2c6cfcb399d2c3e84325d542610bd35` (Spec Kit adoption merge, 2026-10-10).
- Merge-base with PR's advertised target main: `9886df4ec6096bbf909dea6da6a6c2e62f87b28e`. GitHub's `MERGEABLE` against main is irrelevant to this task.
- Only PR-exclusive commit vs accepted: `aee4580` (2026-10-15). Its parent is `588466e`. PR metadata lists nine commits vs old main: `f00ad27`, `a410a8e`, `b17ade8`, `3459c90`, `7a08cd6`, `b3a054e`, `9e94d59`, `588466e`, `aee4580`; the first eight are already ancestors of accepted integration.
- Actual exclusive patch: **11 files, +2585/-137**. Direct PR-tree vs accepted-tree comparison: **133 files, +3902/-24448**. The latter contains missing accepted renderer/evidence/parser, scan, bridge, updater/Portable/pythonnet, release metadata, tests and Spec Kit history — NOT changes to import or deletions to apply.
- Read-only three-way analysis finds four Git conflicts: `app/application_service.py`, `app/qpn_renderer.py`, `frontend/src/state/store.tsx`, `frontend/src/tabs/LearningTab.tsx`. Auto-merged controller/persistence files also contain semantic risks.

### Disposition of every PR-exclusive file

| PR file | Already accepted / newer | Missing / disposition |
| --- | --- | --- |
| `app/qpn_renderer.py` | PROMPT-027/027R already has PowerPoint-first chain, multi-view registry/COM probe, staged safe diagnostics, faithful-backend metadata, cancellation; PROMPT-035 adds lock, process/ROT ownership proof, DispatchEx refusal and owned-only cleanup | **Keep accepted file**, not PR's replacement probe/diagnostic classes, Dispatch or COM lifecycle. Add tests against accepted APIs, no fake availability. |
| `app/application_service.py` | Full slide at current preview width, production item-owned regions, path-free DTOs, mtime_ns/size/schema/renderer/availability/learning cache identity, degraded TTL; PROMPT-032 lifecycle, PROMPT-036S minimal acks | Reimplement durable candidate/image/slide restoration through these APIs; do not take simplified cache/DTOs or duplicated synthetic builders. |
| `app/gui_controller.py` | Current run reconciliation, save-labels-before-Excel and retry semantics | Add bounded archive restoration for a fresh controller and save integration; keep Excel transaction path unchanged. |
| `app/image_learning.py` | Stable reusable classification IDs, append-only last-record-wins label dedup, persistent labels/models/thumbnails, hard semantic eligibility gates | Preserve IDs and label schema. Put evidence-scoped asset persistence in a bounded new module rather than import old glob-scanning helpers. |
| `frontend/src/state/store.tsx` | Optimistic/coalesced fast-label writes, epoch/stale-poll protection, single-call save | Reimplement bulk labeling using same optimistic machinery and bounded ack-based save, not PR's N full-state round trips. |
| `frontend/src/tabs/LearningTab.tsx` | New evidence/picture overlays, backend fidelity notices, resizable layout, PROMPT-038 callback-ref viewport measurement and stable hooks | Move label controls to top; reimplement same-report/same-slide Ctrl/Shift selection and quick confirmation without removing these features. |
| `frontend/tests/learning-030.test.mjs` | Older harness/API; only dispatches `dblclick` (not real browser click/click/dblclick sequence) | Replace with tests covering current minimal ack APIs, StrictMode, real event sequence, cross-report/slide scope, failure/retry and selection pruning. |
| `tests/test_prompt030_powerpoint_renderer.py` | Older probe/diagnostic APIs and fake COM ownership not compatible with PROMPT-035 | Do not copy. Existing 035/027R suites authoritative; add integrated renderer/cache tests against them. |
| `tests/test_prompt030_learning_persistence.py` | Contains tautological corrupt-preview assertions and updater setup that does not actually apply/rollback | Reimplement meaningful persistence/corruption/identity tests, plus exercise accepted staged Portable fixtures and updater apply/rollback, without weakening validation. |
| `tests/test_prompt030_e2e_real_flow.py` | Mix of synthetic candidates and optional assertions on slide images | Reimplement real PPTX → production candidates → Python slide renderer → archive → new service/offline/relabel with mandatory image/geometry assertions. |
| `WINDOWS_ACCEPTANCE_PROMPT030.md` | Wrong branch/API/test names for current integration; contains registry-edit suggestions that must not be followed | Write current PROMPT-030B acceptance commands/checklist; no registry edits, forced availability, killing PowerPoint, publishing or build/version changes. |

### Accepted changes deliberately retained

- 034: report list six-row/sticky-header viewport.
- 035: never attach/open/close/Quit/kill user PowerPoint; ownership refusal logged as `stage=ownership`, not unavailable.
- 036R/S: bridge reflection guard, optimistic/coalesced fast labeling, minimal acks, atomic single-save path, stale-response protection.
- 037: canonical skipped-report filtering and count/selection behavior.
- 038: source-mode stale-dist guard, fresh assets diagnostic, lifecycle-safe preview viewport, production geometry and cache invalidation.
- 039: Windows fixture MAX_PATH corrections, complete Portable fixtures/validation, staged verification, updater learning_data preservation and rollback. `app/updater.py`, build tooling, existing tests/fixtures and stable backup are not patch sources.

## Reproduction evidence on accepted base (before implementation)

Linux/Python 3.11.2, Node 22.22.3. Installed development requirements in ignored `.venv`; `npm ci` with browser download disabled. No real Windows/COM.

A temporary-directory production-service reproduction creates an image candidate, reads Learning, saves AFTER + note, creates a new service against the same learning directory, and also requests a same-ID candidate at a second source location with different colored pixels:

```json
{"rows_before_restart": 1, "labels_after_restart": 1, "rows_after_restart": 0,
 "cross_report_same_cached_pixels": true, "persistent_previews_dir_exists": false}
```

The first draft of the ad-hoc script used a nonexistent AppConfig keyword and failed before reproduction; rerun with the actual AppConfig API produced the evidence above. No source code changed during assessment.

Baseline targeted Python command (035, 027R renderer, 027 rendering, 032, 036S, 038 freshness, image_learning, updater): **209 passed in 17.42s**. Baseline frontend learning-labels **1 pass**, learning-hooks **1 pass**, prompt038 **3 pass**. Vite test teardown sometimes prints `The build was canceled` while node exits 0 with no failed/cancelled tests; record rather than hide this warning.

## Demonstrated root causes and semantic conflicts

1. Saved labels survive but `review_candidates()` only reads the last in-memory processor. Candidate bytes/geometry/full slides are session-only. A fresh launch cannot rebuild the workspace.
2. Candidate preview memory cache uses only reusable classification ID. Two report locations sharing that ID can get the wrong pixels. Keep the deliberately shared classifier ID contract; add separate opaque evidence/review identity and scope disk/memory assets by source.
3. PR's PowerPoint `Dispatch` may attach to the user's single-instance server; `Presentations.Count == 0` is not ownership. Its CoInitialize return-value assumption is not pywin32's actual None-return contract; S_FALSE still requires balanced uninitialization. Never import this code over 035.
4. PR's whole-slide keys include candidate/basename but not source version, schema, renderer availability, decision identity or degraded TTL; skip-existing writes pin stale pixels indefinitely. Its restored DTO drops actual fidelity and production evidence geometry. Retain current cache identity across disk/restarts and never invent eligible regions offline.
5. PR multi-select compares only slide number, allowing same-numbered slides from different reports; its quick-save runs inside a state updater (StrictMode side effects), and browser double-click's preceding single-clicks clear selection. Reimplement as event handlers with scope and in-flight guards, no timed auto-navigation or updater side effects.
6. PR bulk awaits N full Learning states and a full save response, undoing 036 latency and stale-poll protection. Use current optimistic fast path and one bounded batch save ack with prevalidated scope.
7. Current 032 contract suppresses old candidates during processing/after an explicitly empty current run. Restore archives only before a new run (processor is None), not by appending stale archive rows over a current run.
8. No arbitrary glob-first match for shared classifier IDs. Snapshot identity must bind source scope; keep absolute source paths private, use opaque review tokens for UI actions. Legacy labels without proven source retain labels and missing-preview state, not guessed evidence/eligibility.

## Proposed remediation / implementation plan

1. New `app/learning_assets.py`: bounded atomic writes/read validation under `learning_data/previews`, serializable ImageCandidate snapshots, scoped opaque review tokens, no duplicate rewrites for identical payloads, legacy label/thumbnail recovery without arbitrary path exposure. Source file stat plus accepted render/cache signature govern live asset reuse. Use content-addressed image files and atomic metadata replacement so interrupted writes cannot pair new metadata with wrong pixels.
2. Wire service/controller to persist when candidates enter review, restore saved **and unlabeled** snapshot candidates before a new run, and relabel using the existing LabelStore and Excel flow (not invented production objects). Preserve canonical classifier IDs/label dedup; distinct review tokens separate visual evidence sources. Retain legacy API ids for compatibility; reject ambiguous legacy actions rather than choose the wrong report. Preserve 032 processing/current-run semantics.
3. Extend existing full-slide cache to disk with the same source/renderer/schema/availability/learning identity and degraded TTL. Offline restoration displays saved backend truth; invalidate unavailable-to-recompute region geometry when labels/model change. Never claim a fallback is PowerPoint. No renderer production change planned.
4. Top-of-right-panel labels, scoped multi-select (Ctrl/Meta/Shift), quick-confirm with explicit label (no guessing), single guarded batch save ack for bulk (validated before mutations), optimistic/coalesced frontend machinery retained. Prune by candidate identity, not list length; no async side effects inside React state updaters.
5. Add Python persistence/real-flow/cache/ownership/batch and updater-preservation tests, frontend scope/StrictMode/quick-confirm/bulk/failure tests; keep all accepted tests unweakened. Run targeted + all available Python/frontend tests, typecheck, build, scoped pyflakes and diff check. Document exact Windows retest and manual COM/desktop/offline/restart/Portable acceptance separately.

**Files likely to change**: new `app/learning_assets.py`, `app/application_service.py`, `app/gui_controller.py`, `app/bridge.py`, `frontend/src/state/store.tsx`, `frontend/src/tabs/LearningTab.tsx`, optional DTO fields in `frontend/src/types.ts`, frontend test-script registration in `frontend/package.json`, new scoped tests, current Windows checklist, this bug directory. `app/image_learning.py` only if a proven compatibility hook is necessary; updater/desktop/renderer/034/037/build/version metadata remain unchanged.

## Constitution review / risks / gates

I–III: no generated source evidence or changed production crops/gates. IV: scoped visual identity, no arbitrary source matching; shared classification labels remain compatible. V–VI: retain labels-first Excel locking/retry/cancellation, batch failure reported not silently successful. VII–VIII: automated reproduce/fix/rerun plus honest pending Windows acceptance. IX–XI: no build/publish/version/backup/merge; stage only explicit reviewed paths, commit/push only session branch, PR against integration, not main. Do not stage local ignored assets or user files.

Risk: archived regions are a decision snapshot, not a substitute for re-running production on changed labels/models. Missing/corrupt archives must not remove labels or fabricate geometry. Disk permission failures must stay display-only; batch validation must happen before any pending labels are mutated. PowerPoint process lifetime remains subject to real Windows retest. No unresolved product question blocks the bounded remediation.
