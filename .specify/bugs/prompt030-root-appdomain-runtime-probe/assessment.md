# Bug Assessment: PROMPT-030S-R — root AppDomain NULL misread as a failed AppDomain

- **Slug**: prompt030-root-appdomain-runtime-probe
- **Created**: 2026-10-10
- **Source**: Pasted text — PROMPT-030S-R task report with real Windows reproduction evidence
- **Verdict**: valid (root cause DEMONSTRATED against upstream sources, not a hypothesis)
- **Severity**: high — a healthy, verified packaged runtime is rejected by the build/validate gate
  (blocks the PR #7 Windows acceptance and every Portable build), and the shipped diagnostics
  actively misdirect support by naming a working root-domain sentinel "the CLR itself would not
  start". The same function also fails OPEN on incomplete probe output (see Suspected Code Paths).

## Report (verbatim or summarized)

Real Windows targeted test `tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder`
against real installed runtime DLLs produced:

```
.NET Framework Release = 533509
Python.Runtime.dll: Python.Runtime 3.0.5.0, .NETStandard,Version=v2.0, IL-only/any-cpu,
                    expected Loader.Initialize entry point present
ClrLoader.dll:      clr_loader 0.2.10 payload, amd64, .NETFramework,Version=v4.7.2,
                    all required pyclr_* exports present
LIVE PROBE:
probe arch          amd64
probe appdomain     NULL
probe resolved      True
probe initialize_rc 0
```

Every structural and live check succeeded — the function pointer resolved and
`Python.Runtime.Loader.Initialize` returned 0 — yet `probe_packaged_runtime()` fails the package
solely because `pyclr_create_appdomain` returned NULL. The expected behaviour is PASS: in
clr-loader 0.2.10 the NULL result of an *unnamed* `pyclr_create_appdomain` request is the designed
sentinel for the ROOT/CURRENT AppDomain, not a creation failure.

## Symptom

The Windows live probe reports a fully successful runtime startup chain (dlopen → pyclr_initialize →
pyclr_create_appdomain(NULL) → pyclr_get_function → Initialize()==0), but `tools/build_portable.py`
turns the `appdomain NULL` line into a build failure, so `--validate-only <healthy folder>` exits 1 and
`build_portable` refuses to publish a package whose runtime provably works. Secondary symptom:
`app/desktop.py` diagnostics describe the same NULL handle as "pyclr_create_appdomain failed – the CLR
itself would not start", which is false for clr-loader's default root-domain path.

## Reproduction

1. On Windows, with the build venv containing pythonnet 3.0.5 + clr-loader 0.2.10, run
   `.\.venv-build\Scripts\python.exe -m pytest tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder -vv -s`.
   The probe prints `probe appdomain NULL`, `probe resolved True`, `probe initialize_rc 0` and the
   validator still returns a problem starting with `pyclr_create_appdomain trả về NULL` → test fails.
   (This exact output was captured on the user's Windows machine — authoritative supplied evidence.)
2. Linux proxy of the decision logic (this sandbox): the checked-in test
   `tests/test_prompt030_runtime.py::test_probe_turns_a_null_appdomain_into_a_build_failure` feeds
   `{"arch": "amd64", "appdomain": "NULL", ...}` to `probe_packaged_runtime()` and currently PASSES —
   i.e. the codebase actively encodes the wrong expectation. Feeding the healthy Windows evidence
   `{"appdomain": "NULL", "resolved": true, "initialize_rc": 0}` to the same function also yields a
   problem (verified by execution during assessment of the unchanged gate; see Suspected Code Paths).

## Suspected Code Paths — with verified root cause

- `tools/build_portable.py:561-601` `probe_packaged_runtime()` — the defect:
  - `:589-593` treats `out.get("appdomain") == "NULL"` as unconditional failure. WRONG per the
    upstream evidence below: the probe intentionally requests the unnamed root domain
    (`:547` `domain = fw.pyclr_create_appdomain(ffi.NULL, ffi.NULL)`), mirroring production.
  - `:594` `if out.get("resolved") is False` — a probe payload with **no** `resolved` key (crash
    before resolution, tampered output) is only caught if some OTHER rule fires. Executed proof:
    `{"arch":"amd64","appdomain":"created","resolved":true}` with NO `initialize_rc` returns zero
    problems (gate PASSES); `initialize_rc: false` passes because JSON `false` is Python `False`,
    `isinstance(False, int)` is True and `False == 0`.
  - No check of `proc.returncode` after a parseable last stdout line: the exact healthy payload with
    `returncode=1` produces problems ONLY from the NULL-appdomain rule (so it would PASS once that
    rule is correctly removed unless a returncode check is added). Wrong `arch` echo (`"x86"` for an
    `amd64` request) passes. Non-dict JSON (`42`) does not even fail — `if key in out` raises an
    **uncaught TypeError inside the build gate itself**. Fail-open/crash confirmed by execution on
    this checkout. Timeout/launch failure is covered (`:573-576`).
- `tools/build_portable.py:522-559` `PROBE_PACKAGED_RUNTIME` — live probe itself is CORRECT and
  models production (`dlopen` of packaged `ClrLoader.dll`, packaged `Python.Runtime.dll` path, root
  domain request, real `func(buffer, 0)` call). Only its explanatory comment (`:516-521`) asserts
  the wrong meaning of the NULL handle.
- `app/desktop.py:322-351` `appdomain_report()` — `:347` returns
  `"NULL(pyclr_create_appdomain failed – the CLR itself would not start)"` for a NULL `_domain`,
  which is precisely the root-domain sentinel in clr-loader 0.2.10. Needs to distinguish:
  no-runtime-selected / not-netfx / netfx root-current domain (normal) / named created domain /
  genuinely failed named-domain / unknown.
- `app/desktop.py:409-441` `_CLR_HINTS` — first hint item (2) (`:419-420`) states
  "dotnet_appdomain=NULL — the CLR itself would not start, i.e. .NET Framework older than 4.7.2":
  an invalid inference; a NULL root-domain handle says nothing about the .NET Framework version.
- `app/desktop.py:230-237` comment — correctly notes `NetFx.__init__` stores the un-checked handle,
  but must add that NULL is the DESIGNED root-domain value for the unnamed path.
- Mark-of-the-Web wording (see separate section): `app/desktop.py:244-252` ("the one that fits every
  observed fact is a blocked assembly") and `tools/build_portable.py:414-426` ("the one mechanism
  which lets clr_loader create its app domain and still fail to resolve") present MOTW as the
  demonstrated cause. It is an unproven hypothesis: no H:-drive A/B (block vs unblock) evidence
  exists in this repository. The PROMPT-030S-R reproduction demonstrates only the root-AppDomain
  false negative.

## Verified upstream source evidence (retrieved 2026-10-10, matching the pinned versions)

Sources fetched in this sandbox and quoted exactly: clr-loader **0.2.10** sdist from
PyPI (`files.pythonhosted.org/.../clr_loader-0.2.10.tar.gz`), clr-loader git tag **v0.2.10**
(`codeload.github.com/pythonnet/clr-loader/tar.gz/refs/tags/v0.2.10`), and pythonnet **3.0.5**
wheel from PyPI. `requirements-webview.txt` pins exactly `pythonnet==3.0.5` / `clr-loader==0.2.10`.

1. `clr_loader/__init__.py` (0.2.10), `get_netfx`:
   `def get_netfx(*, domain: Optional[str] = None, config_file: Optional[Path] = None)` —
   docstring: "domain: Name of the domain to create. **If no value is passed, assemblies will be
   loaded into the root domain.**" and "config_file: ... will only be used for non-root-domains as we
   can not control the configuration of the implicitly loaded root domain."
2. `clr_loader/netfx.py` (0.2.10), `NetFx.__init__`:
   `domain_s = domain.encode("utf8") if domain else ffi.NULL` →
   `self._domain = _FW.pyclr_create_appdomain(domain_s, config_file_s)` — the result is stored with
   **no NULL check**; `NetFx.info()` hardcodes `initialized=True`; `_get_callable` passes the
   stored `self._domain` straight into `pyclr_get_function`.
3. `netfx_loader/ClrLoader.cs` (tag v0.2.10):
   - `Initialize()`: `_domains.Add(new DomainData(AppDomain.CurrentDomain));` — the CURRENT domain is
     registered as `_domains[0]`.
   - `CreateAppDomain(name, configFile)`: `if (!string.IsNullOrEmpty(name)) { ... return new
     IntPtr(_domains.Count - 1); } else { return IntPtr.Zero; }` — returning `IntPtr.Zero` for an
     empty/NULL name is the **designed** root-domain answer, not an error path (there is no failure
     return at all: a named domain either succeeds or throws).
   - `GetFunction(domain, ...)`: `var domainData = _domains[(int)domain];` — `IntPtr.Zero` indexes
     `_domains[0]`, i.e. `AppDomain.CurrentDomain`. Named domains can only get index ≥ 1, so a NULL
     handle is *equivalent to* "the root domain was requested".
   - `CloseAppDomain` likewise guards `if (domain != IntPtr.Zero)` — the root domain is intentionally
     not closable.
4. pythonnet `3.0.5` `pythonnet/__init__.py`:
   - `_create_runtime_from_spec`: on `win32`, `"default"` → `"netfx"` → `clr_loader.get_netfx(**params)`
     with params from `PYTHONNET_NETFX_*` env (empty by default) → `domain=None` → root-domain path.
   - `load()`: `func = assembly.get_function("Python.Runtime.Loader.Initialize")` then
     `if func(b"") != 0: raise RuntimeError("Failed to initialize Python.Runtime.dll")` — so
     "resolved is a real pointer AND initialize returned 0" is EXACTLY pythonnet's own success
     condition, and `NetFx._get_callable` raises "Failed to resolve ..." whenever
     `pyclr_get_function` returns NULL. The gate must therefore demand these positive results, not a
     non-NULL domain handle.

Conclusion (demonstrated): `appdomain NULL` + `resolved True` + `initialize_rc 0` is a **valid,
fully successful** runtime probe on clr-loader 0.2.10's default netfx path. The build gate's NULL
check is a false negative; the missing-field / nonzero-exit paths are false-positive risks. No code
change should switch to a named AppDomain to make the handle non-NULL — that would make the probe
diverge from the production startup path it exists to model.

## Merit and severity

Valid. The defect reproduces on real Windows with real DLLs (supplied authoritative evidence), the
cause is closed by upstream source at the exact pinned versions, and the fix scope is bounded to the
probe gate + desktop diagnostics + their tests. Severity high: it hard-blocks the PR #7 Portable
acceptance pipeline and misleads field diagnosis; it does NOT weaken runtime safety of anything that
ships (it fails closed in the wrong direction today for this case).

## Root Cause (demonstrated — not a hypothesis)

`tools/build_portable.py` `probe_packaged_runtime()` equates `pyclr_create_appdomain` returning NULL
with AppDomain creation failure. Under clr-loader 0.2.10 + pythonnet 3.0.5 the probe (like production)
requests the UNNAMED root domain; `ClrLoader.cs::CreateAppDomain` deliberately returns `IntPtr.Zero`
for that request, meaning "root/current AppDomain (index 0, `AppDomain.CurrentDomain`)". The gate's
interpretation of that sentinel — mirrored in `app/desktop.py` diagnostics — is simply wrong.
Confidence: high (source-verified at the pinned versions; corroborated by the Windows run where
function resolution and Initialize succeeded from the very same call chain).

## Mark-of-the-Web wording (audit result)

The MOTW mechanism (`LoadLibrary` ignores `Zone.Identifier`, `Assembly.LoadFrom` refuses blocked
assemblies with 0x80131515 unless `loadFromRemoteSources`) is a genuine .NET Framework behaviour and
remains the leading *candidate* for the original H:-drive report. It is NOT proven to be that
report's cause — no A/B evidence is in the repository — and PROMPT-030S-R does not prove it. Wording
that presents it as "the one mechanism / the one that fits every observed fact" must be downgraded to
an evidence-based hypothesis. The packaged-runtime unblock feature and the build-time
`mark_of_the_web` reporting are safe, stay in place, and are not redesigned here.

## Proposed Remediation

**Preferred** (matches PROMPT-030S-R §4–§7; smallest correct change):

1. `tools/build_portable.py::probe_packaged_runtime`:
   - remove the NULL-appdomain problem; validate `appdomain ∈ {"NULL", "created"}` as well-formed
     output only (`"NULL"` = root/current-domain sentinel, described accurately in the printed
     diagnostics); keep `PROBE_PACKAGED_RUNTIME` byte-for-byte (Windows acceptance expects
     `probe appdomain NULL`).
   - success contract (all required): subprocess completed with returncode 0; parseable JSON object;
     no `error` field; `arch` echoes the requested architecture; `resolved is True`; `initialize_rc`
     is an int (non-bool) equal to 0. Any missing/malformed/nonzero → failure (fail closed). The
     structural checks (paths, identity, architecture, exports, signature, framework release,
     duplicate copy, SHA256, MOTW reporting) are untouched.
   - keep the real live load (`ffi.dlopen` of the packaged `ClrLoader.dll` + packaged
     `Python.Runtime.dll`); no static-only or mocked substitution.
2. `app/desktop.py`:
   - `appdomain_report()`: NULL `_domain` with no requested name → `root(...)` (root/current AppDomain,
     normal for netfx default); NULL with a named domain → explicit failure string; non-NULL →
     `created(...)` named AppDomain; keep `no-runtime-selected`, `not-a-netfx-runtime(...)`,
     `unknown`, `not-applicable`.
   - `_CLR_HINTS`: stop inferring "CLR would not start / .NET too old" from a NULL domain; teach the
     reader to discriminate via `no-runtime-selected` (real startup failure) vs `root(...)`/`created(...)`
     (host alive → look at blocked_runtime / dotnet_framework / dll_exists instead).
   - adjust comments and MOTW wording to hypothesis grade without removing the unblock feature.
3. `tests/test_prompt030_runtime.py`: delete/replace
   `test_probe_turns_a_null_appdomain_into_a_build_failure` (it encodes the bug) with
   root-domain-NULL-PASS tests; add the full fail-closed matrix (resolved missing/false,
   initialize_rc missing/nonzero, error field, nonzero exit with valid JSON, malformed/non-JSON,
   timeout, arch echo); update `appdomain_report` tests to the root-vs-failure distinction; keep all
   structural-gate tests.

**Alternatives**: switch the probe to a named AppDomain so the handle is non-NULL — rejected: it
diverges from the pythonnet production startup path (§4F) and masks rather than fixes the semantics;
drop the appdomain field entirely — rejected: it is useful positive evidence (root vs named) and its
absence must remain a malformed-output failure.

**Files likely to change**:
- `tools/build_portable.py`
- `app/desktop.py`
- `tests/test_prompt030_runtime.py`

**Tests to add or update** (all in `tests/test_prompt030_runtime.py`):
- root-NULL + resolved True + rc 0 → no problems (the reproduced Windows evidence as a unit test).
- resolved False / missing → problems; initialize_rc missing / nonzero / non-int / JSON bool → problems.
- `error` field → problem; subprocess returncode ≠ 0 with parseable JSON → problem; non-JSON and
  non-object JSON → problem; TimeoutExpired/OSError → problem; wrong `arch` echo → problem.
- `appdomain_report`: NULL+no name → `root(...)` string (NOT containing "failed"/"would not start");
  NULL+named → failure string; non-NULL → `created(...)`; absent runtime → `no-runtime-selected`;
  non-netfx → `not-a-netfx-runtime(...)`; exception → `unknown`.
- `_CLR_HINTS`: hint must not claim NULL domain means CLR startup failure; must name the discriminators.
- Windows-only live test `test_validate_only_passes_a_healthy_folder` stays the acceptance case
  (unmodified; it will exercise the fixed gate on real Windows).

## Risks & Considerations

- Weakening the gate: mitigated by demanding positive evidence (`resolved is True`, int `initialize_rc
  == 0`, returncode 0, arch echo) instead of merely deleting the NULL rule; fail-closed on missing or
  malformed fields.
- String-compatibility: several tests and the printed contract (probe keys, `OK:` line,
  `pyclr_get_function trả về NULL` wording) are pinned by existing tests; messages are rewritten to
  keep those exact substrings where they remain accurate.
- Diagnostics log format (`dotnet_appdomain=...`) changes value strings; no consumer parses them
  programmatically (grep verified: `runtime_diagnostics` output feeds only logs and
  `explain_clr_failure` keyword matching).
- Frontend/React: unaffected — no API or payload change; validation/diagnostics remain
  `logs/app.log`-only.
- Version/build metadata untouched (v1.3.4 / Build 017); no packaging, publishing, or merge action.
- Windows acceptance stays pending: Linux can prove the decision logic only; the real-load verdict
  requires the targeted Windows rerun after push.

## Open Questions

- None blocking. `[NEEDS CLARIFICATION]` items: none — the task statement plus upstream sources close
  the root cause. The only outstanding evidence is the post-fix Windows rerun (tracked in §9 of the
  task, not a code question).
