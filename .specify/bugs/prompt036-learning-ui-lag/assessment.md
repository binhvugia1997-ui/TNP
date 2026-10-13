# Bug Assessment: "Học cải tiến" tab click/input lag — selections need repeated clicks or a long wait

- **Slug**: prompt036-learning-ui-lag
- **Created**: 2026-10-13
- **Source**: pasted text (PROMPT-036)
- **Verdict**: valid
- **Severity**: high
- **Authoritative base**: `arena/c27c3203-tnp` @ `502968c0837f52af03da180a56b2073a5e8f83fa` (v1.3.6 Build 019)
- **Working branch**: `arena/417148a4-tnp`
- **Implementation status**: assessment complete; no production or test source modified before this artifact
- **Spec Kit execution**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.
- **Extension hooks**: `.specify/extensions.yml` parsed successfully; `hooks: {}` — no `before_bug_assess` hooks registered, none executed.

## Report (verbatim summary)

> Tab "Học cải tiến" phản hồi rất chậm. Các thao tác chọn/gán như: Trước cải tiến, Sau cải tiến, các
> lựa chọn/candidate tương tự — bấm nhiều lần hoặc phải chờ lâu UI mới nhận.
> MỤC TIÊU: Click phải phản hồi gần như ngay lập tức. Không được block UI trong lúc
> backend/preview/ảnh/learning data đang xử lý.

Expected: a Before/After (or candidate) selection is visible immediately. Observed: the selection only
appears after a long wait, and users click repeatedly because the first click looks swallowed.

## Reproduction

Reproduced on Linux with two independent instrumented harnesses (scratch, not committed):

1. `/tmp/probe_backend.py` — builds a real `GuiController` + `ApplicationService` over a real
   `make_report` deck, injects N synthetic reviewable image/content candidates into
   `controller.processor.results`, then times the actual service methods and counts the real work
   each one performs (`parse_pptx`, `SlideRenderer.availability`, `SlideRenderer.render`,
   `LabelStore.records`, JSON payload size).
2. `frontend/probe-p036.scratch.mjs` — jsdom + Vite SSR harness rendering the real `StoreProvider`
   and `LearningTab` against an instrumented `window.pywebview.api` mock with configurable bridge
   latency; measures click → visible-active latency, bridge call counts, rerenders, and the
   poll/click ordering race.

Environment: Linux (Debian sandbox), Python 3.11.2 venv, Node v22.22.3, npm 10.9.8. No PowerPoint,
no LibreOffice, no PyMuPDF — the `builtin` renderer backend is used, so all render timings here are
a **lower bound** for Windows.

## Measured evidence

### M1 — one click costs ~1.2 s of Python work once real learning data exists

60 image + 60 content candidates, 400 already-saved labels per store
(`image_labels.jsonl` = 385 KiB, `content_labels.jsonl` = 419 KiB):

| Call | Empty label stores | 400 saved labels per store |
|---|---|---|
| `learning_state()` (warm) | 6.5 ms | **2058.0 ms** |
| `set_learning_label(...)` — ONE click | 6.4 ms | **1203.7 ms** |

### M2 — decomposition of one click (measured, 1319.7 ms total)

```text
image_labels.jsonl   re-read+parse :  713.4 ms  (301 reads)
content_labels.jsonl re-read+parse :  464.5 ms  (301 reads)
SlideRenderer.availability()       :    4.3 ms  (60 probes)
=> repeated label-store I/O share  :   89.3% of the click
```

**89.3 % of a single click's backend time is re-reading and re-parsing the two JSONL label files.**
301 reads × ~400 KiB ≈ 236 MiB of disk I/O + JSON parsing to acknowledge one Before/After selection.

### M3 — the repeated reads scale O(candidates × saved labels)

`img_store_records` / `content_store_records` = full `LabelStore.records()` (file read + per-line
`json.loads`) calls per `learning_state()`:

| image+content candidates | store reads each | payload per call |
|---|---|---|
| 8 + 8 | 42 | 619.0 KiB |
| 24 + 24 | 122 | 1855.6 KiB |
| 60 + 60 | 302 | 4632.6 KiB |

≈ 5 full file re-reads **per candidate per call**, and the payload grows linearly with candidates.

### M4 — the UI shows nothing until the whole round trip finishes

`learning-hooks`-style harness, 40 candidates, 400 ms simulated bridge:

```text
A. click handler returned after 2.4 ms; label visibly active after 413.3 ms
A. bridge calls: {"getLearning":6,"setLabel":1,"saveLabels":0,"getDashboard":12,"renders":21}
```

Perceived latency is 1:1 with the round trip. There is no optimistic local update, so with the real
M1 backend cost the button stays visually unselected for **> 1.2 s** — exactly "phải chờ lâu UI mới nhận".

### M5 — a stale in-flight poll overwrites an accepted click (lost click, demonstrated)

```text
E. backend truth = BEFORE; UI active during flight = true;
   UI REVERTED to unlabeled after stale poll = true; UI active at end = false
```

The poll's `get_learning_state` snapshot was taken **before** the click was applied; because pywebview
serializes each reply after Python returns, the large stale reply can be delivered after the small
`set_learning_label` reply. `refreshLearning` and `labelImage` both call `setLearningState` with no
ordering guard, so the stale snapshot reverts the UI to "Chưa xác nhận" while the backend already holds
`BEFORE`. The user sees their click undone and clicks again — exactly "bấm nhiều lần".

### M6 — the Learning tab re-fetches the full multi-MiB state forever

```text
C. idle 4 s on Learning tab: store renders +1, get_learning_state +3, get_dashboard_state +6
D. click during in-flight poll: backend ack after 514 ms; get_learning_state total=7, setLabel=1
```

~1.3 full learning-state fetches/second, each paying M1–M3 and each holding `ApplicationService._lock`
— the same `RLock` `set_learning_label` needs, so clicks queue behind polls (M6/D: 514 ms ack).

### M7 — Save re-renders every slide preview

```text
saveImageCandidate (2 bridge calls)  818.4 ms  {'parse_pptx': 8, 'renderer_render': 8, ...}
```

`saveImageCandidate` issues two sequential full-state round trips. Persisting labels changes
`_learning_identity()` (all saved overrides), which is part of the `_slide_preview_cache` key, so every
cached slide preview misses and the whole deck is re-parsed and re-rendered (8 × `parse_pptx`,
8 × `render`) to acknowledge one label.

## Click-path trace

```text
UI event      LearningTab.tsx  <Button onClick={() => s.labelImage(cand.id, l.key)}>   (IMAGE_LABELS:
                'AFTER' = "Sau cải tiến", 'BEFORE' = "Trước cải tiến"); candidate list
                onClick={() => s.setImageReviewIndex(idx)}; ContentReview uses s.labelContent
React state   store.tsx labelImage/labelContent — async, NO local write before await:
                  await callNative('set_learning_label', kind, id, label, note)
                  setLearningState(normalizeLearning(next))     <- first visible change
bridge        lib/bridge.ts callBridgeWithOptions -> window.pywebview.api.set_learning_label
                (app/bridge.py:165 -> _respond -> service.set_learning_label)
backend       application_service.py:1012 set_learning_label
                  with self._lock:                 <- shared with every poll (M6)
                    self._assert_idle(...)
                    c.review_candidates()          <- rebuild list to find ONE id
                    c.set_pending_label(...)
                    return self.learning_state()   <- FULL rebuild (application_service.py:912)
disk/image      learning_state() -> c.pump(); c.reconcile(); c.pump()
                  -> _image_candidate_dto() per candidate:
                       learning.store.current_label(id)   -> latest() -> records()  [read #1]
                       learning.store.latest()            -> records()              [read #2]
                       _slide_preview_for(candidate, learning)
                         -> _learning_identity(learning) -> overrides() -> records() [read #3]
                         -> SlideRenderer(...).availability()  (3x shutil.which + COM probe)
                         -> on cache miss: parse_pptx() + renderer.render() + PIL + base64
                  -> _content_candidate_dto() per candidate: 2 more records() reads
learning persistence  save_learning_labels -> LabelStore.label() append -> changes
                  _learning_identity -> invalidates EVERY slide preview (M7)
rerender        setLearningState(new object) -> store object rebuilt (not memoized) ->
                  every useStore() consumer rerenders -> ImageReview `groups` useMemo recomputed
                  (list identity changed) -> whole candidate list reconciles; repeats ~1.3x/sec (M6)
```

## Root cause (demonstrated, not hypothesised)

Four causes compound; all four were measured above.

- **RC-1 (dominant, backend, M1/M2/M3)** — `LabelStore.records()` re-reads and re-parses the entire
  append-only JSONL file on every call, and one `learning_state()` build calls it ~5× per candidate
  (`current_label` + `latest` in `_image_candidate_dto`, `overrides` via `_learning_identity` inside
  `_slide_preview_for`, plus the two `set(lrn.store.latest())` calls, and 2× per content candidate).
  Cost is O(candidates × saved labels): 89.3 % of a click. `set_learning_label` returns
  `self.learning_state()`, so **every single selection click pays the full rebuild.**
- **RC-2 (dominant, frontend, M4)** — `labelImage`/`labelContent` perform no optimistic local state
  write; the UI cannot reflect a selection until the whole RC-1 round trip returns.
- **RC-3 (frontend race, M5)** — no ordering guard between poll-driven `refreshLearning` and
  mutation-driven `setLearningState`; a stale snapshot reverts an accepted click. `labelImage` also
  does not participate in `learningInFlight`, so it can overlap a poll.
- **RC-4 (amplifier, M6/M7)** — the Learning tab refetches the full multi-MiB state ~1.3×/s while
  idle, holding the same `RLock` the click needs; `saveImageCandidate` makes two sequential
  full-state calls, and saving invalidates every slide preview so the whole deck re-renders.

## Checklist from the request

| Suspect | Finding |
|---|---|
| synchronous bridge calls | No. `callBridgeWithOptions` is async; pywebview proxies return promises. Not the cause. |
| backend call on click | **Yes — RC-1.** One click = one full `learning_state()` rebuild (~1.2 s with real data). |
| preview/image regeneration | **Yes on Save — RC-4/M7.** `_learning_identity` in the cache key re-renders the whole deck. Not per label click. |
| disk I/O | **Yes — RC-1, the largest single term.** 301 full JSONL re-reads per click (89.3 %). |
| full candidate reload | **Yes — RC-1/RC-2.** `set_learning_label` returns and the store replaces the entire `LearningState`. |
| unnecessary React rerenders | **Yes — RC-4.** Store object is not memoized; `labelImage`/`labelContent`/`saveLearningLabels` depend on `learningState.images`/`contents`, so their identity changes on every poll and every consumer rerenders ~1.3×/s. |
| repeated polling/refresh | **Yes — RC-4/M6.** 3 full learning fetches + 6 dashboard fetches per 4 s idle. |
| duplicate calls | **Partly.** `learningInFlight` dedupes polls but not click-vs-poll overlap; `saveImageCandidate` intentionally issues 2 calls. No duplicated `set_learning_label` per click was observed (M4: `setLabel:1`). |
| disabled/loading state nuốt click | Checked, **not** the primary cause: `learningLoad === 'loading'` does not disable the label buttons. `disabled={s.locked}` does swallow clicks while `running`/`scanBusy`/`training`/`diagRunning`/`serverApplyRunning`; that matches the backend `_assert_idle` contract for `running`/`scanBusy` and is left unchanged (semantics preserved). |

## Severity rationale

**High.** The Learning review workspace is the primary human labelling surface; with a realistic label
history every selection costs > 1.2 s and can be silently reverted, so labellers click repeatedly and
can believe labels were lost. Blast radius is confined to the Learning tab (reports/scan/Excel paths
are unaffected), and no data corruption was observed — the persisted label is correct (M5: backend
truth `BEFORE`), only the UI is wrong. That keeps it below `critical`.

## Proposed remediation

**Preferred (minimal, both sides of the demonstrated cause):**

1. `app/image_learning.py` — memoize `LabelStore.records()`/`latest()` keyed on the file's identity
   (`exists`, `st_size`, `st_mtime_ns`), invalidated by `label()` and by any external file change.
   Pure memoization of an append-only read; no semantic change. Removes ~89 % of click cost (RC-1).
2. `app/application_service.py` — compute the renderer availability identity **once per
   `learning_state()` build** instead of once per candidate (60 probes → 1), preserving PROMPT-027R
   §30 semantics across builds.
3. `frontend/src/state/store.tsx` — apply an **optimistic local label** immediately in
   `labelImage`/`labelContent` (marked `labelPending`, exactly the backend's pending semantics), then
   let the authoritative reply reconcile; revert and surface the error if the backend rejects (RC-2,
   preserving learning feedback).
4. `frontend/src/state/store.tsx` — add a monotonic **learning epoch** so a full-state snapshot
   requested before a local mutation is discarded instead of reverting it; include click mutations in
   the in-flight guard so a click cannot overlap a poll (RC-3).
5. `frontend/src/state/store.tsx` — suppress **poll-driven** learning refreshes while a mutation is in
   flight or was just applied; explicit user refresh stays unconditional (RC-4). Memoize the store
   value object and stabilize the learning callbacks so a poll no longer re-identities every consumer.

**Alternatives considered and rejected:**

- *Return a small ack from `set_learning_label` instead of the full state.* Rejected: existing
  authoritative tests pin the contract (`tests/test_application_service.py:670-677` asserts
  `labeled_image["counts"]` and `["images"][0]["labelPending"]`; the constitution forbids weakening
  tests). Fix 1 makes the full state cheap enough that the contract can stay.
- *Send `slidePreview` only for the selected candidate.* Rejected: `tests/test_prompt024r_learning_workspace.py`
  asserts `slidePreview` on every candidate DTO (PROMPT-027R §54 cross-report isolation), and this is
  a contract change, not a lag fix.
- *Drop `_learning_identity` from the preview cache key to avoid M7 re-render.* Rejected: PROMPT-027R
  §25 requires that a label change invalidate the region decision; removing it would be a semantics
  regression in evidence geometry.
- *Debounce/throttle the label click.* Rejected: it would delay or drop real user intent (the request
  explicitly requires no lost clicks under rapid Before ↔ After switching).

## Out of scope / untouched

No version or build bump, no Portable build, no LAN publish, no PR merge. PROMPT-035 and
TAT_QPN_PHASE6 artifacts are untouched. Before/After label semantics, `select_after_pictures`,
evidence-region geometry, report isolation, cancellation and Excel transaction contracts are unchanged.

## Not-run / pending acceptance

- Windows PowerPoint COM acceptance, pywebview transport timing on Windows, and Portable/PyInstaller
  behaviour are **not-run** (Linux sandbox only, per Constitution VII). The `builtin` renderer was the
  only backend exercised, so Windows preview-render timings will differ.
- Real user-perceived latency inside the actual pywebview WebView2 window was **not-run**; M4–M6 use
  an instrumented bridge mock with configurable latency, which reproduces the ordering and the
  1:1 round-trip coupling but is not a Windows UI measurement.
