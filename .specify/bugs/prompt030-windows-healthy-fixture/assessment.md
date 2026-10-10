# Bug Assessment: PROMPT-030 Windows healthy-fixture regression

- **Slug**: prompt030-windows-healthy-fixture
- **Created**: 2026-10-10
- **Source**: Pasted text — PROMPT-030R and supplied Windows acceptance evidence
- **Verdict**: valid
- **Severity**: medium — one Windows test regresses; production validation is correctly enforcing the stronger runtime contract

## Report (verbatim or summarized)

The supplied Windows full-suite result is `1 failed, 1161 passed`; the failure is `tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder`. The synthetic package passes structural PE/export inspection but the live CFFI/Windows loader probe reports `function/symbol 'pyclr_initialize' not found`.

Control evidence supplied for the real development payload:

- `.venv-build\Lib\site-packages\clr_loader\ffi\dlls\amd64\ClrLoader.dll` exists and is 10,240 bytes.
- Windows dynamic loading resolves all five expected exports: `pyclr_initialize`, `pyclr_create_appdomain`, `pyclr_get_function`, `pyclr_close_appdomain`, and `pyclr_finalize`.

This is authoritative user-provided Windows evidence; it was not independently rerun in this Linux Agent environment.

## Symptom

On Windows, validate-only runs the new live packaged-runtime probe against the test's synthetic `ClrLoader.dll` and fails at `pyclr_initialize`, even though the structural parser finds all five expected export names. A real installed `ClrLoader.dll` resolves those exports, so the regression is in the test fixture/contract boundary, not evidence that the production probe is wrong.

## Reproduction

1. On Windows, run `python -m pytest tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder` (the supplied full-suite run reports this as the single failure).
2. The test calls `_package()` at `tests/test_prompt030_runtime.py:68-82`, which writes `python_runtime_image()` and synthetic `clr_loader_image()` into the temporary package.
3. `validate_existing_package()` calls the same `validate_pythonnet_runtime()` gate used by production and passes the interpreter into `probe_packaged_runtime()` (`tools/build_portable.py:754-777`).
4. On Windows, `probe_packaged_runtime()` runs `ffi.dlopen()` and then looks up/calls the packaged `pyclr_*` functions (`tools/build_portable.py:522-601`). The synthetic DLL fails the live lookup despite its export-name metadata.
5. Local Linux check: `.venv/bin/python -m pytest -q -vv tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder` passed `1/1`. This is not a Windows reproduction: `probe_packaged_runtime()` explicitly prints `BỎ QUA` and does not execute .NET Framework/CFFI native loading on Linux (`tools/build_portable.py:568-571`).

## Suspected Code Paths

- `tests/test_prompt030_runtime.py:68-82` — `_package()` builds the temporary package from synthetic runtime images.
- `tests/test_prompt030_runtime.py:686-690` — the failing validate-only test currently labels the synthetic package healthy and passes a Python interpreter, causing the live probe on Windows.
- `tests/dotnet_image_fixtures.py:131-337` — emits a synthetic ECMA-335/PE image; its export construction writes export-name metadata and assigns each function entry the same `base_rva` (`:267-281`), without implementing native callable functions.
- `tests/dotnet_image_fixtures.py:347-354` — `clr_loader_image()` returns that structural synthetic PE.
- `tools/dotnet_pe.py` — `_exports()` parses export-name metadata for the structural check; this is not a Windows `LoadLibrary`/`GetProcAddress` execution.
- `tools/build_portable.py:491-515` — `prove_clr_loader()` checks architecture, mixed-mode metadata, and names found by the PE parser.
- `tools/build_portable.py:522-601` — `probe_packaged_runtime()` performs the separate live runtime chain on Windows.
- `tests/test_prompt028_portable.py:925-997` — PROMPT-028/PROMPT-029 package helpers also use generated images for structural package-layout and validator tests, not live DLL execution.

## Root Cause Hypothesis

**Confidence: high.** The test uses a structurally parseable PE/ECMA-335 fixture designed to exercise metadata and export-table parsing. Its export strings satisfy the parser's structural `image.exports` check, but its synthetic function table points at a shared placeholder RVA and the fixture contains no genuine native/mixed-mode `ClrLoader` implementation. Consequently, it is not a loadable/callable replacement for the installed `clr_loader` C++/CLI DLL. PROMPT-030 intentionally changed “healthy package” to require a live runtime resolution on Windows, and the old synthetic test fixture does not meet that stronger contract. The supplied Windows error plus successful resolution from the real DLL distinguish the fixture defect from a production export/API defect.

Keep these concepts separate:

- **Structurally valid test image** — suitable for parser, metadata, export-name, identity, architecture, and malformed-image tests.
- **Live-loadable runtime DLL** — genuine installed runtime payload whose functions resolve and execute through the Windows runtime chain.

## Proposed Remediation

**Preferred**: keep the synthetic images for structural tests, but do not use them as a live-loadable package. Make `test_validate_only_passes_a_healthy_folder` a Windows-only integration test that discovers the actual `Python.Runtime.dll` and architecture-matched `ClrLoader.dll` from the same interpreter validate-only will use, copies their bytes into the temporary package, and executes the real production probe. Make missing Windows integration prerequisites fail with an explicit message rather than silently skipping the loadability check. On Linux, retain coverage of the validator decision contract with a narrowly scoped injected probe result and clearly label the synthetic package structural-only; do not claim native execution.

Update the synthetic fixture documentation to state that its `ClrLoader.dll` image models parseable export metadata only and must not be used as evidence of live loadability. No runtime/application production edits are indicated by the evidence.

**Alternatives**:
- Separate the live Windows validate-only integration into a differently named test. This is semantically clear, but retaining the established failing test name for the real integration makes it obvious that the original regression is now covered by the correct payload.

**Files likely to change**:
- `tests/test_prompt030_runtime.py`
- `tests/dotnet_image_fixtures.py`
- `.specify/bugs/prompt030-windows-healthy-fixture/fix.md`
- `.specify/bugs/prompt030-windows-healthy-fixture/test.md`

**Tests to add or update**:
- Windows integration: `test_validate_only_passes_a_healthy_folder` with the real installed payloads and production live probe.
- Cross-platform validator decision test using a synthetic structural package and a narrowly injected probe success/failure result; it must not claim Windows native loading.
- Retain structural PE/export tests and run PROMPT-028/PROMPT-027 regression coverage.

## Risks & Considerations

- The Agent environment is Linux; it cannot verify actual Windows DLL loading or .NET Framework execution. A real Windows rerun remains mandatory.
- The Windows integration test must source both DLLs from the same interpreter that runs the production probe and must report missing prerequisites explicitly.
- Do not remove or bypass `probe_packaged_runtime()`, change its Windows behavior, or weaken identity/export/runtime checks.
- The old `H:\ReportExtractor_v1.3.4_Portable` folder and any `Zone.Identifier` evidence are out of scope and must remain unmodified. MOTW remains unproven.

## Open Questions

- None blocking. The actual Windows rerun after the fix remains pending and cannot be answered from Linux evidence.
