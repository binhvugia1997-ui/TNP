# Bug Fix: PROMPT-030S-R — root AppDomain semantics in the portable runtime probe + diagnostics

- **Slug**: prompt030-root-appdomain-runtime-probe
- **Fixed**: 2026-10-10
- **Assessment**: ./assessment.md
- **Status**: applied

## Summary

`tools/build_portable.py::probe_packaged_runtime()` no longer treats the probe's `appdomain NULL` as an
AppDomain creation failure: on clr-loader 0.2.10's default netfx path (pythonnet 3.0.5, `domain=None`)
`ClrLoader.cs::CreateAppdomain` deliberately answers the unnamed request with the root/current-domain
index-0 sentinel. A PASS now requires every applicable positive evidence (subprocess exit 0, well-formed
JSON object, no `error`, requested arch echoed, `resolved is True`, integer `initialize_rc == 0`) and the
gate FAILS CLOSED on missing/malformed/incomplete output — several fail-open holes found during assessment
are now explicit failures. `app/desktop.py` diagnostics report the NULL handle as the root/current
AppDomain instead of "the CLR itself would not start", the `_CLR_HINTS` reader guide was corrected, and
Mark-of-the-Web wording was downgraded from asserted cause to an explicitly unproven hypothesis (the safe
unblock feature is kept unchanged). The live probe itself (`PROBE_PACKAGED_RUNTIME`) was NOT weakened:
it still dlopens the packaged `ClrLoader.dll` and resolves `Python.Runtime.Loader.Initialize` from the
packaged `Python.Runtime.dll`; no static-only or mocked substitute was introduced.

## Changes

| File | Change | Notes |
|------|--------|-------|
| `tools/build_portable.py` | modified | `probe_packaged_runtime()`: removed the false `appdomain == "NULL"` failure; added returncode-0 requirement, JSON-must-be-object validation (previously crashed the gate with uncaught `TypeError` on e.g. `42`), `resolved is True` (missing key no longer passes), `initialize_rc` must be a non-bool int equal to 0 (missing / JSON `false` no longer pass), architecture echo check, `appdomain` field must be `NULL`/`created` (malformed fails closed); printed diagnostics describe `NULL` as the root/current AppDomain sentinel; `PROBE_PACKAGED_RUNTIME` comment block records the verified semantics; probe script body byte-identical (`probe appdomain NULL` output preserved for the Windows acceptance evidence format). `mark_of_the_web()` docstring and the blocked-file problem message now grade MOTW as an unproven candidate mechanism, and the `pyclr_get_function` NULL problem message labels Zone.Identifier the leading unproven candidate. |
| `app/desktop.py` | modified | `appdomain_report()`: distinguishes `no-runtime-selected` / `not-a-netfx-runtime(…)` / `root(AppDomain.CurrentDomain — clr-loader's unnamed default; a normal, successful value)` / `created(named AppDomain …)` / `NULL(pyclr_create_appdomain failed for named domain …)` (the only NULL-as-failure case, matching the upstream impossibility profile) / `unknown`. `_CLR_HINTS[0]`: item (2) rewritten — NULL/`root(...)` is NOT a startup failure and says nothing about the .NET Framework version; only `no-runtime-selected`/`unknown`/`not-a-netfx-runtime` mark a runtime that never started. Module comment corrected to state the designed NULL meaning. MOTW comment block and the `runtime_diagnostics` unblock comment now state hypothesis grade ("no A/B evidence ever captured"); the unblock feature, its call site, and `dotnet_appdomain` field placement are unchanged. |
| `tests/test_prompt030_runtime.py` | modified | Replaced the bug-encoding `test_probe_turns_a_null_appdomain_into_a_build_failure` with the positive root-domain tests and a fail-closed matrix (see below); updated `appdomain_report` tests to the root-vs-named-failure distinction; added an end-to-end gate test for the healthy root-domain probe shape; module/probe docstrings audited to state root-domain semantics and MOTW-as-hypothesis. No unrelated test was weakened; all pre-existing structural-gate tests are retained. |

## Diff Highlights

Gate interpretation (before → after):

```python
# before: every NULL meant failure
if out.get("appdomain") == "NULL":
    problems.append("pyclr_create_appdomain trả về NULL: clr_loader không tạo được AppDomain …")

# after: NULL is the root-domain sentinel; success must POSITIVELY prove itself
if out.get("appdomain") == "NULL":
    print("   probe domain     root/current AppDomain — … KHÔNG phải lỗi tạo AppDomain")
…
if proc.returncode != 0: problems.append("… kết thúc với mã lỗi …")
if out.get("error"): problems.append("… thất bại: …"); return problems
if out.get("arch") != arch: problems.append("… kiến trúc …")
if out.get("appdomain") not in ("NULL", "created"): problems.append("… output hỏng")
if resolved is False: problems.append("pyclr_get_function trả về NULL …")   # exact real crash case
elif resolved is not True: problems.append("… không xuất ra bằng chứng phân giải …")  # missing/!True
else: rc must be a non-bool int and == 0, else fail closed
```

Diagnostics (before → after):

```python
# before
if is_null:
    return "NULL(pyclr_create_appdomain failed – the CLR itself would not start)"
return "created"
# after
if is_null:
    if name:
        return f"NULL(pyclr_create_appdomain failed for named domain {name!r})"
    return "root(AppDomain.CurrentDomain — clr-loader's unnamed default; a normal, successful value)"
return f"created(named AppDomain {name!r})" if name else "created(named AppDomain)"
```

## Tests Added or Updated

- `test_probe_accepts_the_root_appdomain_null_sentinel` — the reproduced Windows evidence
  (`NULL`+`resolved True`+`initialize_rc 0`) now yields zero problems and is described as root/current.
- `test_gate_accepts_the_healthy_root_domain_probe_end_to_end` — full `validate_pythonnet_runtime`
  gate on a faithful package + that probe output passes with every structural check still running.
- `test_probe_turns_null_appdomain_plus_failed_resolution_into_a_build_failure` — NULL domain does not
  excuse a failed `pyclr_get_function` (the real PROMPT-030 crash still fails).
- `test_probe_fails_closed_on_incomplete_or_malformed_evidence` (7 parametrizations) — `resolved`
  missing, `initialize_rc` missing / JSON false / JSON true, wrong arch echo, garbage `appdomain`
  value, probe `error` field.
- `test_probe_fails_closed_on_nonzero_exit_with_parseable_json` — healthy-looking JSON with exit 1 fails.
- `test_probe_fails_closed_on_unparseable_output` (4 parametrizations) — non-JSON / JSON int / JSON
  null / empty stdout: named failure, and the validator must not raise.
- `test_probe_reports_a_crash_instead_of_a_silent_pass`, `test_probe_reports_its_own_exception`,
  `test_probe_rejects_a_nonzero_initialize_result`, `test_probe_accepts_a_runtime_that_resolves`,
  `test_probe_turns_a_null_resolution_into_a_build_failure`, `test_probe_reports_a_timeout_as_a_failure`
  — retained (helper-normalized, assertions kept or strengthened: exact rc in the message, error-only
  payload must produce exactly one honest problem).
- `test_appdomain_report_reads_the_handle_instead_of_trusting_runtime_info` — NULL/unnamed → `root(`,
  named+NULL → `NULL(` failure, non-NULL → `created(named AppDomain …)` with the requested name; the
  root string must not contain "failed"/"would not start"/"NULL(".
- `test_appdomain_report_falls_back_when_cffi_is_unavailable` — fallback path reports `root(`.
- `test_hint_does_not_misread_the_root_domain_sentinel_as_a_CLR_failure` — new; pins the corrected
  `_CLR_HINTS` guidance (old false inference must be absent, discriminators named).
- `test_probe_checks_the_domain_handle_that_clr_loader_never_reads` — additionally pins that the probe
  requests the UNNAMED domain (`pyclr_create_appdomain(ffi.NULL, ffi.NULL)`), i.e. models production.
- `test_validate_only_passes_a_healthy_folder` (Windows-only, real DLLs) — UNCHANGED; it is the
  acceptance case to rerun on Windows after push.

## Local Verification (see also ../prompt030-windows-healthy-fixture for the prior round)

- `python -m pytest tests/test_prompt030_runtime.py -q` → 74 passed, 1 skipped (Windows-only live test).
- Regression set `tests/test_prompt028_portable.py tests/test_prompt027r_renderer.py
  tests/test_prompt027r_regions.py` → 124 passed, 7 skipped.
- Pre-fix execution evidence on this checkout (assessment, unchanged code): root-domain NULL healthy
  payload → 1 problem (false negative); `initialize_rc` missing / JSON-false / wrong-arch → 0 problems
  (fail-open); JSON `42` → uncaught `TypeError`. Post-fix all six behaviors flip to the required contract
  (pinned by the new tests, which fail on the pre-fix code).
- `compileall` on the three changed files → clean. `pyflakes` → only the pre-existing intentional
  `import clr` probe-import finding (present at baseline `76ce2c4` too).
- Frontend: NOT affected — changes are backend build-gate/startup-diagnostics strings only; no API,
  payload, or React-visible behavior changed, so no frontend test/typecheck run was applicable.

## Deviations from Assessment

- None functional. One refinement while implementing: the assessment anticipated "gate PASSES" for
  non-object JSON; actual pre-fix behavior was an uncaught `TypeError` from `if key in out` — strictly
  worse, and covered by the same fix (JSON must parse to an object). `assessment.md` was written with
  the executed result; this paragraph records the correction path.
- Error-path early return: when the probe reports `error`, the gate reports the error (+ nonzero exit if
  any) and stops, rather than also complaining about evidence that a thrown chain could not have produced.

## Follow-ups

- Rerun the targeted Windows acceptance on the real machine (see test.md / task §9); broader Windows
  suite only after that passes.
- If a future named-domain (domain=...) configuration is ever adopted, the `created`/`NULL(named)`
  branches already distinguish it; no further change needed.
