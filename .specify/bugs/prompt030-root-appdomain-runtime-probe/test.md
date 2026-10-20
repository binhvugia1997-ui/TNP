# Bug Verification: PROMPT-030S-R — root AppDomain semantics in the portable runtime probe + diagnostics

- **Slug**: prompt030-root-appdomain-runtime-probe
- **Tested**: 2026-10-10
- **Assessment**: ./assessment.md
- **Fix**: ./fix.md
- **Result**: partial

## Summary

The demonstrated defect — the packaged-runtime probe failing a healthy folder solely because
`pyclr_create_appdomain` returned its root-domain NULL sentinel — is fixed and pinned by new regression
tests that verifiably fail on the pre-fix code and pass on the fixed code, together with the hardened
fail-closed evidence contract and corrected desktop diagnostics. The FULL Linux suite passes. The
original Windows reproduction (`test_validate_only_passes_a_healthy_folder` with real CLR/CFFI loads)
is a Windows-only test: it was NOT run here (Linux skips it) and remains the acceptance case, hence
`partial`, not `verified`.

## Checks Performed

| Check | Command / Action | Result | Notes |
|-------|------------------|--------|-------|
| Reproduction (Linux decision-level, post-fix) | `python -m pytest tests/test_prompt030_runtime.py -q` | pass | 74 passed, 1 skipped (`test_validate_only_passes_a_healthy_folder`, Windows-only by design) |
| Regression pins vs pre-fix code | new/updated tests run against `git checkout HEAD -- tools/build_portable.py app/desktop.py` while keeping the new tests | pass (as in: they correctly FAIL pre-fix) | 16 failed on the unfixed sources — incl. the root-domain-NULL sentinel accept test and every fail-closed matrix case — and all pass after the fix |
| Full gate shape | `test_gate_accepts_the_healthy_root_domain_probe_end_to_end` | pass | structural proofs still run; healthy root-domain probe yields zero problems |
| Targeted regression set | `python -m pytest tests/test_prompt028_portable.py tests/test_prompt027r_renderer.py tests/test_prompt027r_regions.py -q` | pass | 124 passed, 7 skipped (PowerPoint-COM/Windows cases skip on Linux as before) |
| Combined focused+regression | same four files in one run | pass | 198 passed, 8 skipped |
| FULL Linux suite (once) | `python -m pytest -q` | pass | **1171 passed, 8 skipped** in 104.63s; skips are the pre-existing Windows/COM-gated tests |
| compileall | `python -m compileall -q` on the three changed files | pass | clean |
| pyflakes | `python -m pyflakes` on the three changed files | pass (with 1 pre-existing finding) | only `app/desktop.py:422 'clr' imported but unused` — the deliberate `import clr` startup probe, present at baseline commit `76ce2c4`; no new findings |
| `git diff --check` | whitespace check | pass | clean |
| Frontend | not run | not-applicable | changes are build-gate + startup-diagnostic strings only; no API/payload touched — frontend was not affected, so no frontend test/typecheck applies |
| Windows live reproduction | — | **not-run** | requires a real Windows machine with installed pythonnet 3.0.5 / clr-loader 0.2.10 DLLs; Linux sandbox cannot execute .NET Framework loading. This is the acceptance gate: `.\.venv-build\Scripts\python.exe -m pytest tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder -vv -s` |
| Mark-of-the-Web A/B on H:\ | — | not-run / out of scope | explicitly NOT claimed: MOTW remains an unproven hypothesis; this bug changed only the honesty of that wording |

## Output Excerpts

- Focused: `74 passed, 1 skipped in 0.50s`; full: `1171 passed, 8 skipped in 104.63s (0:01:44)`.
- Pre-fix pin evidence (new tests against old sources): `16 failed, 3 passed, 56 deselected in 0.51s` —
  e.g. `test_probe_accepts_the_root_appdomain_null_sentinel`,
  `test_probe_fails_closed_on_incomplete_or_malformed_evidence[initialize_rc-missing-fails-closed]`,
  `test_gate_accepts_the_healthy_root_domain_probe_end_to_end` all failed pre-fix and pass post-fix.
- Healthy Windows evidence replayed through the fixed gate (unit-level): `probe appdomain NULL`,
  `probe resolved True`, `probe initialize_rc 0` → `[]` problems, plus the printed line describing
  "root/current AppDomain … KHÔNG phải lỗi tạo AppDomain".

## Verdict

`partial` — the demonstrated root cause is fixed with targeted, fail-closed, and full-suite Linux
evidence, and the regression tests verifiably encode the corrected contract. `verified` is withheld
because the original reproduction ran on real Windows and the definitive rerun (the Windows-only
`test_validate_only_passes_a_healthy_folder` live-load test) has not been executed on Windows in this
session. Linux results are development evidence only (constitution §VII).

## Residual Risks

- The live Windows rerun must confirm the gate passes with genuine DLLs (expected healthy output keeps
  the `probe appdomain NULL / resolved True / initialize_rc 0` lines; the added descriptive line is
  informational and does not alter the pinned output contract).
- The fail-closed contract rejects probe payloads missing fields the probe always emits; a future edit
  to `PROBE_PACKAGED_RUNTIME` must keep emitting `arch/appdomain/resolved/initialize_rc` (or
  `error`) — the source-content tests partially guard this.
- The broader Windows suite (beyond the targeted test) is intentionally deferred per the task: it is
  considered only after the targeted test passes on Windows.
- MOTW unblock behavior ships unchanged; if the Windows rerun still fails after this fix, the blocked
  `Zone.Identifier` path (already reported/unblocked at startup and build) plus the `error`/`resolved`
  probe fields are the next evidence sources — not the domain handle.

## Recommendation

Hold for the scheduled Windows acceptance: push, then rerun
`.\.venv-build\Scripts\python.exe -m pytest tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder -vv -s`
on the target machine. A pass there converts this report to verified for the original symptom; do not
merge PR #7 or build/publish anything before that, and do not reinterpret this as MOTW proof.
