# Bug Verification: "Học cải tiến" click/input lag

- **Slug**: prompt036-learning-ui-lag
- **Tested**: 2026-10-13
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: partial
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
- **Extension hooks**: `.specify/extensions.yml` parsed successfully; `hooks: {}` — no `before_bug_test` / `after_bug_test` hooks registered, none executed.
- **Slug resolution**: from conversation context (this session ran assess then fix for `prompt036-learning-ui-lag`; `.specify/bugs/prompt036-learning-ui-lag/fix.md` exists).

## Summary

The defect no longer reproduces on Linux: a Trước/Sau selection is visible inside the click event
(22.3 ms vs 418.6 ms before, with a 400 ms bridge), rapid switching loses no click, a stale in-flight
snapshot can no longer revert an accepted selection, and one click's Python cost fell from 1203.7 ms to
13.4 ms. All 15 new regression tests fail without the fix and pass with it, the full Python failure set is
byte-identical to the pre-fix baseline, and frontend typecheck is clean. The result is **partial**, not
**verified**, only because Windows pywebview / PowerPoint COM / Portable acceptance could not be performed
in this Linux sandbox (Constitution VII) — no functional check is outstanding or failing.

## Environment

| Item | Value |
|---|---|
| OS | Linux (Arena sandbox, Debian) |
| Python | 3.11.2 (project venv `.venv`) |
| Node / npm | v22.22.3 / 10.9.8 |
| Renderer backends available | `builtin` only — no PowerPoint, no LibreOffice, no PyMuPDF |
| Base commit | `502968c0837f52af03da180a56b2073a5e8f83fa` (v1.3.6 Build 019) |
| Branch | `arena/417148a4-tnp` |

Because only the `builtin` renderer ran, every render-related timing here is a **lower bound** for Windows,
where PowerPoint COM or LibreOffice would be used.

## Checks Performed

| Check | Command / Action | Result | Notes |
|-------|------------------|--------|-------|
| Reproduction (pre-fix, backend) | `/tmp/probe_backend.py 60 60 400` on base | pass (reproduced) | one click = 1203.7 ms; 301 label-file reads per store; 89.3 % of the click |
| Reproduction (pre-fix, frontend) | `frontend/probe-p036.scratch.mjs` on base | pass (reproduced) | selection visible after 418.6 ms; 5/5 rapid clicks not shown; stale poll reverted an accepted click |
| Reproduction (post-fix, backend) | same probe with the fix | pass (gone) | one click = 13.4 ms; 0 label-file reads |
| Reproduction (post-fix, frontend) | same probe with the fix | pass (gone) | selection visible after 22.3 ms; 0/5 rapid clicks missed; no revert |
| New Python tests | `.venv/bin/python -m pytest tests/test_prompt036_learning_ui_lag.py -q -p no:randomly` | pass | **8 passed** in 3.42 s |
| New Python tests fail without the fix | `git stash push app/image_learning.py` + same command | pass | **3 failed, 5 passed** — the 5 are contract-preservation tests that must pass either way |
| New frontend tests | `cd frontend && node --test tests/prompt036-learning-click.test.mjs` | pass | **7 passed** |
| New frontend tests fail without the fix | `git stash push frontend/src/state/store.tsx` + same command | pass | **0 passed, 7 failed**, each with the intended assertion message |
| Learning regression suites | `pytest tests/test_prompt024r_learning_workspace.py tests/test_prompt032_learning_state.py tests/test_prompt027r_regions.py tests/test_improvement_visual_regions.py tests/test_after_evidence.py tests/test_prompt025_multi_item.py -q` | pass | **75 passed** — evidence-region, multi-item and PROMPT-032 lifecycle contracts intact |
| Full frontend suite | `cd frontend && npm run test:frontend` | pass (1 pre-existing failure) | **43 tests, 42 pass, 1 fail**; baseline was 36/35/1 — the same single failure |
| Frontend typecheck | `cd frontend && npm run typecheck` (`tsc --noEmit`) | pass | clean, no output |
| Full Python regression suite | `.venv/bin/python -m pytest tests/ -q -p no:randomly` | pass (36 pre-existing failures) | **36 failed, 1189 passed, 18 skipped**; baseline **36 failed, 1181 passed, 18 skipped** |
| Failure-set equality | `diff` of sorted `FAILED` lists, baseline vs post-fix | pass | **empty diff — identical 36 failures**; +8 passed = the new tests |
| Lint | `.venv/bin/python -m pyflakes app/image_learning.py app/application_service.py tests/test_prompt036_learning_ui_lag.py` | pass | clean |
| Whitespace / conflict markers | `git diff --check` | pass | clean |
| Windows PowerPoint COM acceptance | — | **not-run** | requires real Windows PowerPoint (Constitution VII) |
| Windows pywebview WebView2 transport timing | — | **not-run** | requires the real desktop host; the jsdom harness uses an instrumented bridge mock |
| Portable / PyInstaller acceptance | — | **not-run** | requires a real Windows build; out of scope for PROMPT-036 |
| Real user-perceived latency in the shipped window | — | **not-run** | measured with an instrumented bridge, not a human click on Windows |

## Output Excerpts

New Python tests:

```text
$ .venv/bin/python -m pytest tests/test_prompt036_learning_ui_lag.py -q -p no:randomly
........                                                                 [100%]
8 passed in 3.42s
```

Those same tests with the backend fix stashed (proves they pin the demonstrated cause):

```text
$ git stash push app/image_learning.py && .venv/bin/python -m pytest tests/test_prompt036_learning_ui_lag.py -q
E   AssertionError: one Before/After click took 398 ms of Python work; measured 1203.7 ms before the
    PROMPT-036 fix with the same 40+40 candidates and 200 saved labels per store
FAILED ...::test_label_store_serves_repeated_reads_from_one_disk_read
FAILED ...::test_one_label_click_reads_the_label_files_a_bounded_number_of_times
FAILED ...::test_one_label_click_is_fast_with_a_large_saved_label_history
3 failed, 5 passed in 4.62s
```

New frontend tests, and the same file with the frontend fix stashed:

```text
$ node --test tests/prompt036-learning-click.test.mjs
# tests 7   # pass 7   # fail 0

$ git stash push frontend/src/state/store.tsx && node --test tests/prompt036-learning-click.test.mjs
# pass 0   # fail 7
AssertionError: the selection must be visible inside the click, not after the round trip
AssertionError: Sau cải tiến must be visible the moment it is clicked
AssertionError: a stale snapshot reverted an accepted selection (PROMPT-036 M5 regression)
AssertionError: optimistic feedback is shown first
AssertionError: a poll erased an unsaved note (PROMPT-036 input lag)
AssertionError: a heavy preview payload must not delay the visible selection
AssertionError: the content selection must be visible immediately
```

Before/after timing, identical harnesses (frontend probe, real `StoreProvider` + `LearningTab` in jsdom):

```text
                                              BEFORE                       AFTER
A. single click -> selection visible     418.6 ms                       22.3 ms
B. rapid clicks not shown immediately    5 / 5                          0 / 5   (4.9-7.0 ms each)
B. final UI vs backend truth             agreed (AFTER)                 agreed (AFTER)
D. click while a poll is in flight       shown=false (ack 514 ms)       shown=true in 5.2 ms (ack 507 ms)
E. stale poll reverts accepted click     true  (UI ended unlabeled)     false (UI ended BEFORE)
   learning polls during 3 s settle      2                              1
```

Backend probe, 60 image + 60 content candidates, 400 saved labels per store:

```text
                                              BEFORE                       AFTER
set_learning_label  (ONE click)          1203.7 ms                      13.4 ms
label-file reads per click (per store)   301                            0
label-store I/O share of the click       89.3 %                         0.0 %
learning_state() warm, labels populated  2058.0 ms                      770.7 ms
```

Full-suite equality:

```text
$ pytest tests/ -q -p no:randomly            # post-fix
36 failed, 1189 passed, 18 skipped in 91.09s
$ diff <(sort baseline FAILED) <(sort post-fix FAILED)
IDENTICAL (36 pre-existing failures, unchanged)
```

## Pre-existing failures (NOT caused by this fix)

The 36 Python failures and 1 frontend failure are present on the untouched base commit and are identical
before and after. They fall into two groups:

- **Version/build metadata drift (majority).** Tests still pin `1.3.5` / `Build 018` while the base commit is
  `1.3.6` / `Build 019` — e.g. `test_prompt001.py::test_single_canonical_version`,
  `test_prompt027_rendering.py::test_backend_version_is_the_authoritative_1_3_5_build_018`,
  `test_gui_responsive.py::test_version_and_build_release018`,
  `frontend/tests/prompt027-overlays.test.mjs` → `the frontend package metadata agrees with the
  authoritative backend version` (its fixture hardcodes `APP_VERSION = '1.3.5'`, `APP_BUILD = '018'`).
  Fixing these would require touching version metadata, which PROMPT-036 explicitly forbids.
- **Windows-only updater paths.** `test_updater.py` (13), `test_auto_update_check.py` (3),
  `test_runtime_paths.py` (2), `test_prompt028_portable.py` (2) and the two learning-file updater tests fail
  with `RuntimeError: Không có bản cập nhật mới hợp lệ để cài.` from `app/updater.py:1010` — the same
  version-comparison drift, on Linux.

Two of these sit next to learning data and were checked individually to confirm the failure *reason* is
unchanged, not merely the name:

```text
tests/test_image_learning.py::test_28_apply_update_keeps_learning_data_and_rollback_restores_it
tests/test_content_learning.py::test_updater_preserves_content_learning_files
  -> E RuntimeError: Không có bản cập nhật mới hợp lệ để cài.   (app/updater.py:1010, both runs)
```

This matters because the fix adds a label-store cache that must be invalidated when the updater restores
learning data. That behaviour is covered directly and passes:
`test_label_store_cache_is_invalidated_by_an_external_file_change`.

## Residual Risks

- **Windows acceptance outstanding.** Real pywebview WebView2 transport, PowerPoint COM preview rendering
  and Portable behaviour were not exercised. The frontend evidence comes from a jsdom harness with an
  instrumented bridge mock: it faithfully reproduces the ordering race and the round-trip coupling, but it
  is not a measurement of the shipped Windows window.
- **Renderer timings are a lower bound.** Only `builtin` rendered here. On Windows the §25 post-save slide
  re-render (~770 ms measured) will be slower with PowerPoint/LibreOffice, though it stays off the click
  path.
- **Payload size unchanged by design.** Every candidate still ships its `src` and `slidePreview` data URI
  (4.6 MiB for 60 synthetic candidates, ~18 MiB with realistic bitmaps). Changing that is a DTO contract
  change rejected in the assessment because authoritative tests pin `slidePreview` on every candidate.
- **Optimistic label is provisional.** Between a click and its reply the UI shows the label with
  `labelPending` ("chưa lưu"), matching Python's own pending semantics. If the backend rejects, the label
  rolls back and the real error is surfaced — covered by a test, but a reviewer clicking very fast could
  briefly see a provisional state.
- **Poll hold-off window.** `LEARNING_MUTATION_QUIET_MS = 600` delays (never cancels) the poll-driven
  Learning snapshot after a mutation. Idle polling cadence is unchanged (measured 3 learning + 6 dashboard
  fetches per 4 s idle, before and after).
- **Cache invalidation depends on `(st_size, st_mtime_ns)`.** Sound for the append-only writer and for an
  external replace or delete (both tested). A hypothetical external rewrite to the *identical* size within
  the *identical* nanosecond mtime would not be noticed; no such writer exists in this codebase.

## Recommendation

**Accept on Linux; hold Windows acceptance as explicitly pending.** The original symptom is gone in every
measured dimension, the fix is confined to the two demonstrated causes, no authoritative test was weakened,
and the regression suites are unchanged apart from the 15 new tests. Before release, run on real Windows:
`pytest tests/` plus `npm run test:frontend` and `npm run typecheck`, then manually confirm in the desktop
app that (a) Trước/Sau selection is instant with a large `learning_data` history, (b) rapid switching keeps
the last choice, and (c) "Lưu nhãn & ghi chú" still updates Excel and the preview region. The 36 + 1
pre-existing version-metadata failures should be triaged separately — they are outside PROMPT-036's scope
and fixing them here would require the forbidden version/build bump.
