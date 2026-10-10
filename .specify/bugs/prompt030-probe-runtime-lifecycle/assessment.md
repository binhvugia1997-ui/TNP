# Bug Assessment: PROMPT-030T — live probe models only the initialization half of the pythonnet lifecycle

- **Slug**: prompt030-probe-runtime-lifecycle
- **Created**: 2026-10-10
- **Source**: Pasted text — PROMPT-030T task report carrying real Windows re-run evidence of
  `tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder` after PROMPT-030S-R was
  integrated. No URL was supplied, so the URL Trust Policy had nothing to classify
  (`auto-refused: no URL present`).
- **Verdict**: valid — root cause DEMONSTRATED from the pinned upstream sources, not inferred from the
  Windows screenshot alone.
- **Severity**: high — the Windows targeted acceptance test for PR #7 still fails
  (`assert 1 == 0`), so the live probe still rejects a package whose runtime provably initializes. It is
  not a false rejection this time: the probe subprocess really does die, and the probe's own evidence
  hides that it never performed the shutdown half of the contract it claims to model.

## Report (verbatim or summarized)

PROMPT-030S-R is confirmed integrated; the earlier false failure is gone. On real Windows the live probe
now reaches and reports:

```
probe arch          amd64
probe appdomain     NULL
probe resolved      True
probe initialize_rc 0
```

and correctly describes the handle as `root/current AppDomain`. So: NULL root AppDomain is accepted,
`Python.Runtime.Loader.Initialize` resolves from the packaged `Python.Runtime.dll`, and it returns 0.

But the probe subprocess then terminates abnormally. The observed Windows output includes:

```
[validate-only] KẾT LUẬN: gói này KHÔNG thể khởi động, vì:
  → probe nạp runtime đã đóng gói kết thúc với mã lỗi ...
```

and .NET reports:

```
System.InvalidOperationException:
GIL must always be released, and it must be released from the same thread that acquired it.
```

with a stack involving Python.Runtime GIL / shutdown / finalization. The targeted test therefore ends
`FAILED  assert 1 == 0`. The subprocess-return-code gate is correct and MUST stay: a probe that crashes
after printing apparently healthy JSON has not proven anything.

## Symptom

The probe proves initialization and then abandons the runtime. `PROBE_PACKAGED_RUNTIME`
(`tools/build_portable.py`) performs only:

```
ffi.dlopen(ClrLoader.dll) → pyclr_initialize() → pyclr_create_appdomain(ffi.NULL, ffi.NULL)
→ pyclr_get_function(…, "Python.Runtime.Loader", "Initialize") → func(b"")
```

and then falls off the end of the script. It never resolves/calls `Python.Runtime.Loader.Shutdown` and
never calls `pyclr_finalize()`. The managed runtime is therefore torn down by the CLR's own
process-teardown path instead of by pythonnet's ordered shutdown, an `InvalidOperationException` escapes
a finalizer, and the interpreter exits nonzero — after the healthy JSON line was already on stdout.

## Reproduction

1. Windows, build venv with pythonnet 3.0.5 + clr-loader 0.2.10:
   `.\.venv-build\Scripts\python.exe -m pytest tests/test_prompt030_runtime.py::test_validate_only_passes_a_healthy_folder -vv -s`
   → probe prints the four healthy lines above, `.NET` raises `GIL must always be released…`,
   `validate_existing_package` reports `probe nạp runtime đã đóng gói kết thúc với mã lỗi …`, exit 1,
   test `FAILED assert 1 == 0`. (Reported evidence; **not re-executed here** — Linux cannot run the
   netfx native chain, see Residual Risks.)
2. Host-independent reproduction of the same contract gap (run on Linux at base commit
   `0502855316d0d7e0cc2632ff94f35fd531dd82ad`, before any production change):

   ```
   $ python -m pytest tests/test_prompt030_runtime.py -q
   70 passed, 5 skipped in 0.72s
   ```

   Every existing probe test passes while `PROBE_PACKAGED_RUNTIME` contains no `Shutdown` call and no
   `pyclr_finalize` call:

   ```
   $ grep -c 'Loader", b"Shutdown"\|full_shutdown\|fw.pyclr_finalize' tools/build_portable.py
   0        # (0 matches: the shutdown half of the lifecycle is absent from the probe body)
   ```

   The new lifecycle regression tests (Fix section, `tests/test_prompt030_runtime.py`) fail against
   `0502855`: a payload that carries complete lifecycle evidence is indistinguishable from one that
   proves initialization only, and a payload missing shutdown evidence is accepted as a PASS. That is
   the reproduced failure boundary expressed at the decision level.

## Suspected Code Paths

| Location | Why it is implicated |
| --- | --- |
| `tools/build_portable.py` → `PROBE_PACKAGED_RUNTIME` (probe script body) | Executes `pyclr_initialize` / `pyclr_create_appdomain` / `pyclr_get_function` / `Loader.Initialize` and then ends. No `Loader.Shutdown`, no `pyclr_finalize`, no `try/finally` around the native state it created. |
| `tools/build_portable.py` → `probe_packaged_runtime()` | Demands positive evidence for `arch`, `appdomain`, `resolved`, `initialize_rc`, exit code — i.e. it accepts a PASS on initialization evidence alone, so it cannot distinguish a probe that completed the lifecycle from one that abandoned it. |
| `pythonnet/__init__.py::load` / `unload` (pinned 3.0.5) | Defines the contract the probe is supposed to model, including `atexit.register(unload)`. |
| `src/runtime/Loader.cs::Shutdown` (pinned 3.0.5) | The `full_shutdown` entry point the probe never calls. |
| `clr_loader/netfx.py::initialize` / `_release` (pinned 0.2.10) | `atexit.register(_release)` → `pyclr_finalize()`; the probe never reaches it. |

`app/desktop.py` is **not** implicated: its `appdomain_report()` reads `pythonnet.get_runtime_info()`
inside the frozen app and has no relationship to the build-time probe subprocess. PROMPT-030S-R's
diagnostics stay untouched.

## Root Cause (demonstrated)

The probe executes the **initialization half** of pythonnet's lifecycle and then exits, leaving the
managed runtime to be torn down by the CLR's process-teardown path instead of by pythonnet's ordered
shutdown. The causal chain, each step quoted from the pinned sources this project actually uses:

### 1. pythonnet 3.0.5 — `pythonnet/__init__.py` (sha256 of sdist `pythonnet-3.0.5.tar.gz`
###    `48e43ca463941b3608b32b4e236db92d8d40db4c58a75ace902985f76dac21cf`)

`load()` — lines 142–152:

```python
_LOADER_ASSEMBLY = assembly = _RUNTIME.get_assembly(str(dll_path))
func = assembly.get_function("Python.Runtime.Loader.Initialize")

if func(b"") != 0:
    raise RuntimeError("Failed to initialize Python.Runtime.dll")

_LOADED = True

import atexit

atexit.register(unload)
```

`unload()` — lines 160–167:

```python
func = _LOADER_ASSEMBLY.get_function("Python.Runtime.Loader.Shutdown")
if func(b"full_shutdown") != 0:
    raise RuntimeError("Failed to call Python.NET shutdown")

_LOADER_ASSEMBLY = None

if _RUNTIME is not None:
    _RUNTIME.shutdown()
```

So pythonnet treats a *successful initialize* as an obligation to run
`Loader.Shutdown(b"full_shutdown")` and only then `_RUNTIME.shutdown()`, and it enforces that ordering
with `atexit.register(unload)` at line 152. Initialization without shutdown is not a supported end
state.

### 2. pythonnet 3.0.5 — `src/runtime/Loader.cs`

`Shutdown` — lines 40–60:

```csharp
public unsafe static int Shutdown(IntPtr data, int size)
{
    try
    {
        var command = Encodings.UTF8.GetString((byte*)data.ToPointer(), size);

        if (command == "full_shutdown")
        {
            using var _ = Py.GIL();
            PythonEngine.Shutdown();
        }
    }
    catch (Exception exc) { …; return 1; }

    return 0;
}
```

`full_shutdown` therefore performs the real Python.Runtime shutdown path and reports 0/1 — a
positive-evidence rc, exactly like `Initialize` (lines 11–37, `PythonEngine.InitExt()`, `return 0`).

### 3. pythonnet 3.0.5 — `src/runtime/PythonEngine.cs`: what `Initialize` leaves subscribed

`Initialize` (reached through `Loader.Initialize` → `InitExt()`) at lines 218–219:

```csharp
AppDomain.CurrentDomain.DomainUnload += OnDomainUnload;
AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
```

and `InitExt()` at line 343:

```csharp
Finalizer.Instance.ErrorHandler += AllowLeaksDuringShutdown;
```

`Shutdown()` (lines 368–397) is what removes them and drains the managed state:

```csharp
AppDomain.CurrentDomain.DomainUnload -= OnDomainUnload;
AppDomain.CurrentDomain.ProcessExit -= OnProcessExit;

ExecuteShutdownHandlers();
// Remember to shut down the runtime.
Runtime.Shutdown();

initialized = false;
```

`Runtime.Shutdown()` (`src/runtime/Runtime.cs` 257–…) drains the pending-finalization queue through
`Finalizer.Shutdown()` while still holding `PyGILState_Ensure()`, with `AllowLeaksDuringShutdown`
installed so leak errors during shutdown are *handled* rather than thrown.

**Consequence of skipping it:** `initialized` stays `true`, both AppDomain handlers stay subscribed,
and the `Finalizer` queue keeps live handles. Teardown is then driven by
`OnProcessExit` → `Runtime.ProcessIsTerminating = true; Shutdown();` (lines 315–319) from the CLR's own
process-exit event, i.e. uncoordinated with CPython's `Py_FinalizeEx` for a `python -c` probe.

### 4. pythonnet 3.0.5 — `src/runtime/Py.cs`: the unique source of the observed exception

```csharp
public class GILState : IDisposable
{
    private readonly PyGILState state;
    internal GILState() { state = PythonEngine.AcquireLock(); }

    public virtual void Dispose()
    {
        if (this.isDisposed) return;
        PythonEngine.ReleaseLock(state);
        GC.SuppressFinalize(this);          // line 35 — only Dispose suppresses the finalizer
        this.isDisposed = true;
    }

    ~GILState()                             // line 39
    {
        throw new InvalidOperationException("GIL must always be released, and it must be released from the same thread that acquired it.");
    }
}
```

`grep -rn "GIL must always be released" src/` over the 3.0.5 tree returns exactly this line (41) plus
the `DebugGILState` variant (55, different wording, and `DebugGIL` is off by default). So the observed
message can only come from `~GILState()` on the GC finalizer thread. `Dispose()` is the only thing that
calls `GC.SuppressFinalize`, so the finalizer runs precisely when `Dispose()` never did — which is what
happens when the `GILState` constructor's `PythonEngine.AcquireLock()` (→
`Runtime.PyGILState_Ensure()`) throws against an interpreter that is already finalizing, leaving the
allocated-but-unconstructed object finalizable. On .NET Framework an exception escaping a finalizer is
unhandled and fatal, which is the abnormal subprocess termination with a nonzero return code.

### 5. clr-loader 0.2.10 — `clr_loader/netfx.py` + `netfx_loader/ClrLoader.cs` (sha256 of sdist
###    `clr_loader-0.2.10.tar.gz` `81f114afbc5005bafc5efe5af1341d400e22137e275b042a8979f3feb9fc9446`)

`clr_loader/netfx.py` — root-domain request and the finalize obligation:

```python
domain_s = domain.encode("utf8") if domain else ffi.NULL          # line 23
self._domain = _FW.pyclr_create_appdomain(domain_s, config_file_s)  # line 27
…
def shutdown(self):
    if self._domain and _FW:                                       # line 56
        _FW.pyclr_close_appdomain(self._domain)

def initialize():
    global _FW
    if _FW is not None: return
    _FW = load_netfx()
    _FW.pyclr_initialize()                                         # line 66
    atexit.register(_release)                                      # line 68

def _release():
    global _FW
    if _FW is not None:
        _FW.pyclr_finalize()                                       # line 74
        _FW = None
```

`_get_callable` (lines 40–50) is also the origin of the PROMPT-030 string
`f"Failed to resolve {typename}.{function} from {assembly_path}"` on a NULL functor.

`netfx_loader/ClrLoader.cs`:

- `pyclr_initialize` → `Initialize()` lines 15–24: `_domains.Add(new DomainData(AppDomain.CurrentDomain));
  _initialized = true;` — the root/current domain is registered at **index 0**.
- `pyclr_create_appdomain` → `CreateAppDomain` lines 38–66: a non-empty name creates a domain and
  returns `new IntPtr(_domains.Count - 1)` (line 60); the **unnamed** request falls to `else { return
  IntPtr.Zero; }` (line 64). This is the PROMPT-030S-R sentinel and it stays valid.
- `pyclr_get_function` → line 74: `_domains[(int)domain]` — `IntPtr.Zero` indexes the root/current
  domain registered by `Initialize()`.
- `pyclr_close_appdomain` → lines 89–104: `if (domain != IntPtr.Zero)` — a NULL handle is a deliberate
  no-op, matching `NetFx.shutdown()`'s truthiness guard at line 56.
- `pyclr_finalize` → `Close()` lines 106–116: `foreach (var domainData in _domains) domainData.Dispose();
  _domains.Clear(); _initialized = false;`

`netfx_loader/DomainData.cs` `Dispose()` lines 108–118:

```csharp
public void Dispose()
{
    if (!_disposed)
    {
        _functors.Clear();
        if (Domain != AppDomain.CurrentDomain)
            AppDomain.Unload(Domain);
        _disposed = true;
    }
}
```

For the root domain `Domain == AppDomain.CurrentDomain`, so `pyclr_finalize()` only clears the functor
table — it does **not** unload the current domain. It is idempotent (`if (!_disposed)`). Finalizing after
a root-domain probe is therefore safe and is exactly what production does.

### 6. The probe as it stands (base commit `0502855`)

`PROBE_PACKAGED_RUNTIME` body, verbatim tail:

```python
fw = ffi.dlopen(str(host))
fw.pyclr_initialize()
domain = fw.pyclr_create_appdomain(ffi.NULL, ffi.NULL)
out["appdomain"] = "NULL" if domain == ffi.NULL else "created"
func = fw.pyclr_get_function(domain, str(assembly).encode("utf8"),
                             b"Python.Runtime.Loader", b"Initialize")
out["resolved"] = func != ffi.NULL
if out["resolved"]:
    buffer = ffi.from_buffer("char[]", b"")
    out["initialize_rc"] = int(func(ffi.cast("void*", buffer), 0))
```

Confirmed present and confirmed absent:

- present: `pyclr_initialize`, `pyclr_create_appdomain`, `pyclr_get_function`, `Loader.Initialize`
- absent: `Loader.Shutdown`, `b"full_shutdown"`, `fw.pyclr_finalize()`, any `try/finally` protecting the
  native state it created (`grep` over the probe body returns no match for any of them)

### Verified lifecycle vs. probe lifecycle

| Stage | Production (pythonnet 3.0.5 + clr-loader 0.2.10) | Probe at `0502855` |
| --- | --- | --- |
| 1 | `pyclr_initialize()` — `netfx.py:66`, registers `_release` at `:68` | done |
| 2 | `pyclr_create_appdomain(ffi.NULL, …)` → `IntPtr.Zero` = root/current domain — `netfx.py:23,27`, `ClrLoader.cs:64` | done |
| 3 | `Loader.Initialize(b"")` must return 0 — `__init__.py:143-146` | done |
| 4 | `Loader.Shutdown(b"full_shutdown")` must return 0 — `__init__.py:160-162`, `Loader.cs:46-49` | **missing** |
| 5 | `_RUNTIME.shutdown()` → `NetFx.shutdown()`; `pyclr_close_appdomain` **skipped** because the root handle is falsy — `netfx.py:56-57` | n/a (correctly skipped) |
| 6 | `pyclr_finalize()` via `atexit._release` — `netfx.py:74`, `ClrLoader.cs:106-116` | **missing** |
| 7 | process exits 0 | **abnormal exit / nonzero** |

atexit ordering confirms 4 → 6: `_release` is registered first (`netfx.py:68`, during
`NetFx.__init__`), `unload` second (`__init__.py:152`), and `atexit` runs LIFO — so `Loader.Shutdown`
always precedes `pyclr_finalize`. This matches the lifecycle PROMPT-030T proposed, so the proposal is
confirmed rather than invented; the only refinement is that stage 5 stays skipped for the root domain.

### What is demonstrated vs. reconstructed

- **Demonstrated from source**: stages 1–7 above; that `Loader.Initialize` subscribes AppDomain
  `ProcessExit`/`DomainUnload` handlers and leaves `initialized == true`; that only
  `PythonEngine.Shutdown()` removes them and drains the `Finalizer` queue; that the observed message
  exists at exactly one place, `Py.cs:41`, inside `~GILState()`, reachable only when `Dispose()` never
  ran; that the probe performs no shutdown and no finalize.
- **Reconstructed, not instrumented**: the exact interleaving of CPython's finalization with the CLR's
  `ProcessExit` on the Windows machine, and therefore which specific `GILState` allocation was collected
  instead of disposed. No Windows debugger was attached in this sandbox and Linux cannot execute the
  netfx native chain. This step is graded as a source-derived reconstruction, not an observation.
- **Not re-run here**: the Windows targeted test itself. Linux results are development evidence only.

## Merit and Severity

Valid, high. The gate is doing its job (the subprocess really dies), but the probe is the defect: it
models half of the contract and then reports success-shaped JSON from a process that is about to crash.
Beyond PR #7 acceptance, the same half-lifecycle would misreport any package that initializes but cannot
shut down cleanly.

## Proposed Remediation (preferred)

Make `PROBE_PACKAGED_RUNTIME` model the verified lifecycle, with structured cleanup:

1. Track the lifecycle stage reached (`start` → `clr-initialized` → `appdomain` →
   `pythonnet-initialized` → `pythonnet-shutdown`).
2. Resolve `Python.Runtime.Loader.Shutdown` and call `func(b"full_shutdown")` **only when
   `initialize_rc == 0`** — never shut down a runtime that never initialized.
3. Wrap the native state in `try/except/finally`: record the first exception in `error` and never let a
   secondary cleanup exception overwrite it; run `pyclr_finalize()` in `finally`, guarded so it runs
   only when `pyclr_initialize()` actually succeeded, exactly once, and record `finalized`
   (`True`/`False`) plus a separate `finalize_error` when it throws.
4. Keep `pyclr_close_appdomain` uncalled for the root/NULL handle — upstream `NetFx.shutdown()` skips it
   (`netfx.py:56`) and `ClrLoader.CloseAppDomain` no-ops on `IntPtr.Zero` (`ClrLoader.cs:91`);
   `pyclr_finalize()` disposes the registered `DomainData` either way. Do **not** introduce a named
   AppDomain to dodge the NULL sentinel.
5. Emit positive evidence per stage: `shutdown_resolved`, `shutdown_rc`, `finalized`, and print them.
6. `probe_packaged_runtime()`: a PASS must require every stage's positive evidence — exit 0, JSON
   object, no `error`, arch echo, `appdomain ∈ {NULL, created}`, `resolved is True`,
   integer `initialize_rc == 0`, and then (only when initialization succeeded) `shutdown_resolved is
   True`, integer `shutdown_rc == 0`, `finalized is True`, no `finalize_error`. Anything missing,
   malformed, non-True or nonzero fails CLOSED. The subprocess exit-code gate is unchanged.

### Alternatives considered and rejected

- **Drop the subprocess exit-code gate.** Explicitly out of scope and wrong: a crashing probe proves
  nothing, and the constitution forbids weakening tests to reach green.
- **Run the probe through real `pythonnet.load()`/`unload()` instead of raw cffi.** It would hide the
  packaged-byte proof: `pythonnet.load()` resolves `Python.Runtime.dll` from the *interpreter's*
  site-packages, not from `_internal/pythonnet/runtime/…`, and the PROMPT-030 gate must prove the
  shipped bytes load. The raw cffi chain is what failed on `H:\` and must stay.
- **Call `pyclr_close_appdomain` on the root handle to "clean up".** Upstream deliberately skips it for
  a falsy handle and `ClrLoader.CloseAppDomain` no-ops on `IntPtr.Zero`; calling it adds no evidence and
  diverges from the pinned lifecycle.
- **Redefine MOTW/Zone.Identifier handling.** Out of scope: still an unproven hypothesis for the
  original `H:\` failure, with no new reproduction evidence.

## Scope Boundaries

`tools/build_portable.py` (probe + validator) and `tests/test_prompt030_runtime.py`. No change to
`app/desktop.py`, no version bump (`v1.3.4 / Build 017`), no Portable build, no publish, no MOTW
redesign, no `TAT_QPN_PHASE6/`, no stable-backup changes.

## Residual Risks

- Linux cannot execute the netfx native chain, so the *native* outcome (subprocess rc 0, no
  `GIL must always be released`) is pending real Windows acceptance.
- If Windows still exits nonzero after an ordered shutdown, the residual cause is a different defect
  (e.g. a second undisposed finalizable object or a `Python error indicator is set` condition inside
  `PythonEngine.Shutdown()`, `PythonEngine.cs:375-383`), and this assessment should be re-run with that
  new evidence rather than patched around.

## Unverified

- The exact Windows subprocess exit code value and full .NET stack text were summarized in the task
  report, not captured in this sandbox.
- No URL was supplied with the report; nothing was fetched.
