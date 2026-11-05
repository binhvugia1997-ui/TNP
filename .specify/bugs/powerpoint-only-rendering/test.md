# Bug Verification: PowerPoint-only rendering, fail-closed (powerpoint-only-rendering)

- **Slug**: powerpoint-only-rendering
- **Tested**: 2026-11-04
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: **partial** — automated verification passes on Linux; real Windows PowerPoint-COM acceptance NOT performed and NOT claimed. The agent runs on Linux with no PowerPoint and no pywin32, so PowerPoint behavior is exercised through COM stubs (the repo's established pattern). Windows verification via the cloud workflow build is still required (see the fix record's checklist).
- **Spec Kit execution mode**: Executed from the checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
- **Environment**: Linux sandbox, `.venv-build` in the snapshot tree, pytest; Node + `node --test` for the frontend.

## Checks performed

| # | Check | Result |
|---|---|---|
| 1 | New regressions `tests/test_powerpoint_only_rendering.py` | **7 passed** (unavailable / import / ownership-refusal / generic COM failure / success / default-chain control / remedy coverage) |
| 2 | Renderer suites (`test_prompt027_rendering`, `test_prompt027r_renderer`, `test_prompt028_portable`, `test_prompt035*`, `test_prompt041*`) | **248 passed, 8 skipped, 1 failed** — the failure is the pre-existing `pythonnet` environment probe (see #4) |
| 3 | Full Python suite | **1404 passed, 1 failed, 13 skipped** in 377.64 s (baseline before this batch: 1385 passed / 4 failed; the 3 LibreOffice-selection environment failures now pass, the `pythonnet` one remains) |
| 4 | The one failure | `tests/test_prompt028_portable.py::test_live_import_probe_reports_every_module_in_this_venv` — pre-existing, Linux-only (`ModuleNotFoundError: pythonnet`, Windows-only module), unchanged from baseline; not caused by this fix |
| 5 | Frontend typecheck | `npx tsc --noEmit` exit 0 |
| 6 | Pyflakes on changed Python files | clean except pre-existing unused imports |

## Contract tests updated (not weakened)

Tests pinning the old silent-fallback contract were updated to the new fail-closed contract with intent preserved; each update is documented in the fix record. Batch-pipeline tests that exercise production render paths now run behind the shared COM stub (`tests/powerpoint_render_stub.py`: stubbed export draws each picture shape as its average-colour rectangle at authored geometry, so geometry and colour-presence assertions still bind to real slide content). No assertion was weakened: counts, colours, geometry, and status assertions are unchanged.

## Not verified here

PowerPoint-COM behavior on a real Windows machine (closed/open/missing PowerPoint; batch + re-apply renders) — see the fix record.
