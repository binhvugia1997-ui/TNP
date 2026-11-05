# Version bump: Build 021 -> Build 022 (v1.4.0 unchanged)

## Exception record (constitution: Workflow and Review Gates)

- User explicitly requested this minimal change on 2026-11-05 ("sửa ver đi để dễ nhận biết"):
  bump the build number so the fixed build is distinguishable from the old Build 021
  artifact. The fix batch (PR #19) intentionally did not bump the version, so old and new
  artifacts shared the name `ReportExtractor_v1.4.0_Build021_portable`.
- This bypasses the normal spec-first order as an explicit emergency minimal change.
  Bounded scope below; no behavior change.
- Outstanding evidence: Linux-only verification (see Test). Windows / PowerPoint-COM /
  Portable acceptance still pending, as for the fix batch.

## Scope

- `app/__init__.py`: `BUILD_NUMBER` 21 -> 22 only. `__version__` stays "1.4.0"
  (`BUILD_NUMBER` is the authoritative numeric ordering, PROMPT-005). All derived
  labels follow automatically (`BUILD_ID` "022", `BUILD_LABEL` "Build 022",
  `VERSION_LABEL` "1.4.0 — Build 022").
- Test expected values updated to the new build (21 / "021" / "Build 021" ->
  22 / "022" / "Build 022") where they assert the CURRENT build. No assertion
  weakened; UpdateInfo fixtures follow the existing convention (fixture == current
  build). Backup-frozen assertions (1.0.4 / 004) untouched.
- Workflow artifact name `ReportExtractor_v1.4.0_Build021_portable` ->
  `..._Build022_portable` is a SEPARATE file (`.github/workflows/build-windows-portable.yml`
  + `HUONG_DAN.txt`) delivered for web upload by the user; it is NOT part of this PR
  because the connector cannot push workflow files (403, App lacks Workflows permission).

## Not authorized by this change

- Merge of this PR and any release / publish action require separate explicit user
  authorization (constitution IX / XI).
