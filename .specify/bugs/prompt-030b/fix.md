# Bug Fix: PROMPT-030B — selective integration of PR #14 contracts

- **Slug**: `prompt-030b` (continued from the explicit task identifier / assessment)
- **Fixed**: 2026-10-16 (UTC)
- **Assessment**: ./assessment.md (unchanged after assessment)
- **Status**: applied; real Windows acceptance remains pending
- **Base / branch**: `9faba26` → `arena/1ec5e900-tnp`

Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
Before implementation: assessment + remediation plan completed; parsed extension hooks are empty. No native CLI/PowerShell invocation claimed.

## Summary / plan applied

- Kept accepted PowerPoint renderer/ownership, production evidence regions, cancellation and updater validation byte-for-byte; did not merge/cherry-pick PR #14's old commit.
- Reimplemented only missing durable review evidence and review UX on current service/controller/036S fast paths.
- Kept canonical classification IDs and append-only LabelStore dedup; added independent scoped opaque review identity to stop visual cross-report collisions.
- Extended accepted cache identity and degraded TTL to disk/restarts instead of importing skip-existing, basename/candidate-only slide cache code.
- Added regression tests and exact Windows acceptance instructions; existing tests/assertions remain unchanged.

## Changes / complete file list

| File | Change |
| --- | --- |
| `app/learning_assets.py` | New bounded asset adapter under learning_data/previews. Private ImageCandidate snapshots, source mtime_ns/size, stable report-scoped review tokens, content-addressed validated images, atomic fsync+replace metadata, no-op identical writes, safe missing/corrupt/legacy handling. |
| `app/application_service.py` | Persist candidates on review entry (including unlabeled candidates); use report-scoped/versioned candidate memory cache; reuse durable full slides with existing cache identity/TTL; restore offline pixel/backend truth and clear obsolete decision regions; expose optional reviewId/reportScope/persisted; accept scoped actions while retaining unambiguous legacy IDs; prevalidate bounded bulk save and return minimal ack. |
| `app/gui_controller.py` | Restore snapshots/legacy label rows before a new run only, dedup current evidence by scoped review identity, keep the selected source alongside pending canonical labels for correct Excel reapply provenance. Existing labels-first/Excel retry flow remains. |
| `app/bridge.py` | Add bounded `save_learning_images` operation; no arbitrary filesystem-read API or public Window exposure. |
| `frontend/src/state/store.tsx` | Normalize optional reviewId into UI action identity; bulk optimistic overlay/save reuses label workers, epoch and stale-poll guards; one batch ack, failure remains visible, superseded writes persist latest choice; single save returns success for quick-confirm feedback. |
| `frontend/src/tabs/LearningTab.tsx` | Existing labels/actions moved above metadata; Ctrl/Meta/Shift same-report+slide selection, selection identity pruning, guarded quick-confirm including real double-click's preceding clicks, bulk failure/retry feedback. Current overlay geometry, viewport callback ref, zoom and layout remain intact. |
| `frontend/src/types.ts` | Optional reviewId/reportScope/persisted fields only; canonical report visibility rules unchanged. |
| `frontend/package.json` | Register `test:prompt030b` and include it in full frontend suite; version/dependencies unchanged. |
| `tests/test_prompt030b_integration.py` | 23 meaningful integration/regression cases (parametrized total), real PPTX production flow, durable pixels/geometry/labels, scoped identity, seven cache identity dimensions, ownership diagnostics/TTL recovery, corruption/atomic failure, current-run lifecycle, bulk validation/partial disk failure/retry, real updater apply+rollback with accepted strict fixtures. |
| `frontend/tests/prompt030b-learning.test.mjs` | Seven StrictMode/production-store UI scenarios: controls first, no guessing, native double-click exactly once, same-report+slide boundaries, optimistic bulk/deferred polling/failure/retry, multi quick-confirm, identity pruning and current-label Save behavior. |
| `docs/WINDOWS_ACCEPTANCE_PROMPT030B.md` | Current branch/API, exact targeted/full commands and manual real COM/desktop/restart/offline/Portable acceptance boundaries. |
| `.specify/bugs/prompt-030b/assessment.md` | Preimplementation reproduction, complete history/diff/disposition matrix and plan. |
| `.specify/bugs/prompt-030b/fix.md` | This implementation/disposition/compliance report. |
| `.specify/bugs/prompt-030b/test.md` | Verification report written in the subsequent bug-test step. |

## Semantics retained / deliberately not imported

The assessment's eleven-file PR disposition matrix is authoritative. In particular:

- **PowerPoint**: retain accepted tuple probe, multi-view ProgID lookup, SLIDE_RENDER stage/HRESULT/path-safe logs, PowerPointStageError ownership refusal, owned-only cleanup, balanced COM, cooperative cancellation. Do not take PR #14's Dispatch, Count-based Quit, global last diagnostics or return-value/apartment assumptions.
- **Preview**: retain 1600px Python complete-slide rendering, production item-owned evidence geometry and separate picture/evidence DTO fields. Disk identity carries the same source location/mtime_ns/size, renderer preference/width/schema, availability and learning-model/label identity as memory. Degraded entries use elapsed wall time on disk, reconstructed into session monotonic age; an expired memory entry cannot be resurrected by rereading the disk entry. Actual backend/faithfulness are retained, never forged as PowerPoint.
- **Offline/changed source**: preserved pixels are historical rendered evidence; if decisions change without PPTX, old production regions clear rather than inventing an updated crop. A changed existing source retains the row/label but suppresses mismatched images/regions and requires processing again. No guessed default crop or invented eligible ownership.
- **Identity**: canonical `candidate_id` still means reusable classification data (same as accepted LabelStore/overrides). New `reviewId` and asset key separate visual evidence by source; `reportScope` is opaque. Dataset counts remain canonical, while evidence rows can include multiple distinct report locations. Ambiguous legacy-ID actions fail closed; explicit scoped actions retain selected-source provenance. No arbitrary sidecar glob-first match.
- **036S / Excel**: label buttons remain optimistic; single save remains the existing one-ack labels-first path. Bulk prevalidates all IDs, notes, label and same-report+slide before any mutation, then uses the same pending-label → JSONL → Excel retry contract as single save (including committing existing pending labels of that kind). A disk failure after some JSONL appends is reported as a failure with pending choices retained, not falsely called an all-or-none transaction. Retry dedups already written labels. Excel locks do not roll back confirmed labels.
- **032 / 038**: an active/current run, even an explicitly empty one, remains authoritative; archives are not appended over it. Preview initialization still works without a learning/controller storage owner (rendering cannot depend on persistence availability). Every new hook precedes the candidate early return; the 038 callback ref remains unchanged.
- **Legacy data**: current accepted JSONL labels/models/thumbnails remain readable and unmodified. Source-less legacy labels retain rows with missing/zero geometry and conservative eligibility; only their exact legacy thumbnail key may be used. Unaccepted PR #14 sidecar/cache formats are not guessed or imported as authoritative geometry/backend truth. Existing files are not deleted.
- **Not taken**: old full-state bulk bridge loop, simplified geometry DTOs, candidate/basename-only skip-existing cache, duplicated synthetic candidates, tautological persistence tests, updater tests with fake incomplete Portable layouts, old Windows checklist/registry-edit suggestions, old versions/build/runtime/updater files from PR tree.

## Local verification during fixing

- Original accepted-base reproduction: saved label survives, image rows do not; same canonical ID at two locations shares wrong memory pixels.
- Baseline targeted Python: 209 passed; frontend 036S/hooks/038: 1+1+3 passed.
- New tests: Python 23 passed; frontend initially 6, then 7 passed including current-label Save regression.
- Final targeted Python: **289 passed**; full available Python: **1313 passed, 1 Windows-only skipped**.
- Full frontend: **79 passed**, zero skipped/failed; includes actual Linux Chromium report-list measurements. `tsc --noEmit`, frontend build (1588 modules), scoped pyflakes, pip check and diff whitespace check pass.
- First full Python attempt had two dependency-environment failures (webview/pythonnet absent), 1299 passed/13 skipped. Installed existing development/runtime dependencies in ignored .venv, reran without weakening any assertions; final result above.
- Intermediate regression run caught eight preview tests because an initial persistence wiring assumed all preview-only consumers had a controller; fixed by keeping storage optional, not by changing tests. Fixed an initial new-test misuse of strict `release_entries` on an installed app: use the validated staged release instead, retain real apply/rollback and byte-for-byte preservation assertions. No updater validation changed.

## Deviations from Assessment

No change to the planned runtime boundary or protected files. Reimplementation is intentionally not source-format compatibility with unaccepted PR #14. It additionally persists unlabeled candidates (the contract says candidates survive restart), while preserving the accepted current-run lifecycle instead of blindly appending every historical row to every current run.

## Follow-ups / safety / constitution review

Real Windows Office/COM ownership/fidelity/cleanup, EdgeChromium timing/DPI/zoom, machine restart and any separately authorized installed Portable update remain NOT RUN here. See Windows checklist and test report; do not merge on a claim of Linux == Windows acceptance.

Source evidence/production crops/semantic gates/cancellation unchanged (I–III, V–VI). Report-scoped visual assets and opaque DTO identities enforce IV without breaking shared learning labels. Reproduction and existing tests preserved; environment limitations disclosed (VII–VIII). No merge of main/PR #14/integration, LAN publish, BUILD_AND_PUBLISH, Portable build/release, version/build bump, stable-backup modification, updater weakening, `git add .`, user-file staging or branch switch (IX–XI). Frontend dist/venv/node_modules/test artifacts remain ignored and uncommitted. All work stays on the assigned Arena task branch; review PR targets the accepted integration, not main.
