# Bug Assessment: Last slide "Hiệu quả cải tiến" (effectiveness) still contributes its content to the results

- **Slug**: last-slide-effectiveness-leak
- **Created**: 2026-11-04
- **Source**: pasted text (user report in chat, 2026-11-04 — added to the same fix batch as Bugs A and B)
- **Verdict**: valid
- **Severity**: high (effectiveness/follow-up content is defined to contribute ZERO text to the improvement report; leaked content corrupts the Excel "Nội dung đối sách cải tiến" field and the After-evidence picture selection)

## Report (verbatim or summarized)

User (translated): "There was already exclusion logic: do NOT take the content of the last slide about improvement effectiveness — but the tool is still taking it (into the results/report)."

## Symptom

A final slide titled "Hiệu quả cải tiến" (or equivalent effectiveness/follow-up heading) has its body text included in the extracted improvement text (and its pictures are eligible as After evidence) when the slide's body restates the countermeasure with an improvement-like heading line.

## Reproduction

Code-level reproduction with simulated slide data (no real PPTX; probe script run 2026-11-04 against `heuristic_classify` + `collect_sections` + the Excel improvement-text filter `kind in ("improvement", "standard")`, `app/extractor.py:783`). Six simulated shapes were probed; five are correctly excluded, one leaks:

| # | Last-slide shape | Result |
|---|---|---|
| S1 | Title placeholder "5. Hiệu quả cải tiến" + plain body | OK — verify, excluded |
| S2 | Title typed as first line of a content textbox (bold) + plain body | OK — verify, excluded |
| S3 | **Title placeholder "5. Hiệu quả cải tiến" + body whose first line is the improvement-like heading "Đối sách cải tiến:"** | **LEAK — whole body classified `improvement`, enters the improvement text** |
| S4 | Heading-less body-only last slide after an improvement slide | OK — verify (via "theo dõi" marker), excluded |
| S5 | Title shape ordered after the body shape in reading order | OK — verify, excluded |
| S6 | Trailing "Hiệu quả cải tiến:" block on the last improvement slide | OK — mid-slide heading terminates the improvement section, excluded |

S3 probe output (verbatim): `improvement_slides=[3, 4] verify_slides=[4]`; sections `S4 kind=improvement heading='Đối sách cải tiến:'` and `S4 kind=improvement heading='Sau cải tiến: không phát sinh lỗi.'`; improvement text contains the full effectiveness body. The title's `verify` kind appears nowhere in the sections.

[NEEDS CLARIFICATION: no real user PPTX was available in this session; the S3 shape is a simulated fixture shaped like real countermeasure decks (title placeholder + body restating the countermeasure). Windows/PPTX acceptance with a real deck remains pending.]

## Suspected Code Paths

Exclusion logic EXISTS (it is not a missing feature — it is bypassed on one path):
- `app/classifier.py:44-46` `VERIFY_MARKERS` includes "hieu qua cai tien", "hieu qua doi sach", "hieu qua sau", "xac nhan hieu qua"; `:58-63` `FOLLOWUP_MARKERS` includes the same; `:54` comment: "follow-up / effectiveness blocks contribute ZERO text."
- `app/classifier.py:103-131` `section_kind_of_heading` checks VERIFY before IMPROVE, so "Hiệu quả cải tiến" → `verify`.
- `app/extractor.py:199` `EXCLUDED_SECTION_KINDS = ("inspection", "followup", "verify", "temporary")`.
- `app/extractor.py:780-805` `extract_record`: `improvement_sections = [s for s in sections if s.kind in ("improvement", "standard")]`; `rec.improvement = join_sections(...)`. Both Excel paths use this filter: `app/image_review.py:191-196` (re-apply path) and `app/extractor.py:807-812` (content-learning path via `select_content_with_learning`).
- `app/extractor.py:202` `excluded_bands` — the geometric twin of the segmentation, used by picture selection so excluded blocks contribute zero pictures.

Bypass:
- `app/extractor.py:150-175` `split_sections`: a `ROLE_TITLE` block sets the current section kind (`cur = Section(kind=verify, heading=title)`), but the title section has NO lines yet; when the first content line is itself a heading (`is_heading_like` + `_semantic_kind` → `improvement`), the code replaces `cur` because `cur.lines` is empty (`if cur.lines: sections.append(cur)` is skipped) — the title's `verify` kind is silently discarded and the slide's content becomes `improvement`. The docstring even states the intent being violated: "the slide title still decides the section kind."
- `app/extractor.py:252-283` `collect_sections`: default-kind precedence checks `in_imp` before `in_ver` (`elif in_imp and not in_cause: default = "improvement"`), so a slide in BOTH `improvement_slides` and `verify_slides` (S3: `improvement_slides=[3, 4]`, `verify_slides=[4]`, because `_slide_kinds` tags per heading-like line) defaults to `improvement` for any lines before the first heading — a second, compounding leak path for the same slide shape.
- `app/extractor.py:202-245` `excluded_bands`: the title opens an excluded band, but the first non-excluded heading line closes it (`elif open_kind: bands.append(...)`) even though no content was consumed under the title kind — the same discard shape in geometry form, so pictures on a verify-led slide can be selected as After evidence.

## Root Cause Hypothesis

Confidence: high (reproduced with the probe above; mechanism traced to the exact branches). The exclusion vocabulary and section-kind filter are correct; the segmentation layer loses the slide title's excluded kind when the slide's first content line is an improvement-like heading. The title kind is treated as a placeholder that the first content heading overwrites, instead of the slide-level section decision the design intends. The same discard exists in the geometric band computation used for picture selection.

## Proposed Remediation

**Preferred** (minimal, root-cause):
1. `split_sections`: an empty title-led section whose kind is in `EXCLUDED_SECTION_KINDS` is sticky for the slide's first content block — the first content heading does not overwrite it; the heading line becomes content of the excluded section. Later headings (after content lines exist) still terminate sections exactly as today (S6 behavior preserved; PROMPT-001 tests must keep passing).
2. `excluded_bands`: mirror the same rule — a title-opened excluded band is not closed by the first non-excluded heading when no content was consumed under it (or, equivalently, the band extends to the next heading that follows consumed content).
3. `collect_sections`: make the default kind respect a verify-led slide — when the slide is in `verify_slides` and its first heading kind is `verify`/`followup`, the default is that excluded kind, not `improvement`. Only if needed to close the pre-heading-lines leak for the same shape; keep the change minimal and prove it with the S3 fixture (lines before any heading must also be excluded).

**Alternatives**:
- Add "Đối sách cải tiến" to the verify markers. Rejected: that phrase is genuinely an improvement heading on production slides; the distinction is the slide-level (title) context, not the phrase.
- Drop the slide from `improvement_slides` in the classifier when it is also in `verify_slides`. Rejected: slide-level membership is used for picture/learning paths too; the section layer is where the exclusion is defined and tested.

**Files likely to change**:
- `app/extractor.py` (`split_sections`, `excluded_bands`, possibly `collect_sections`)
- `tests/test_prompt001.py` — read-only reference; new regression tests in a NEW file (e.g. `tests/test_effectiveness_slide_exclusion.py`) using the same fake-`SlideData` construction, so pinned tests are untouched

**Tests to add or update**:
- S3 shape (title-placeholder effectiveness slide + improvement-like first body heading): zero effectiveness-slide lines in the improvement text; the slide's sections are `verify`/`followup`; the slide appears in `rec.excluded_sections`.
- Same shape through `extract_record` with a `Classification` in both slide lists (LLM-style input, mirroring `test_qwen_selecting_mixed_slide_cannot_reintroduce_inspection`): effectiveness text excluded even when the classifier puts the slide in `improvement_slides`.
- Picture path: `excluded_bands` covers the effectiveness slide's body (no After picture selected from it) for the S3 shape.
- Existing PROMPT-001 tests (mid-slide termination, S6 shape) keep passing unchanged.

## Risks & Considerations

- The "title decides the section kind" rule must not swallow a genuine new section that starts a slide whose title is a generic/section heading and whose first content line is a different section heading with real content following it. The sticky rule applies ONLY to an EMPTY title-led section with an EXCLUDED kind, and ONLY for the first content block — after content lines exist, normal heading termination resumes. This bounds the behavior change to exactly the leak shape.
- `collect_sections` default-kind precedence change, if made, affects slides in multiple lists generally; scope it to the verify-led first-heading case and verify against the full suite.
- The S3 fixture is simulated; a real deck's title/body geometry (placeholder vs. textbox, reading order) may differ. The fix is at the segmentation layer, which is geometry-independent once blocks are classified — but Windows acceptance with a real effectiveness-last deck is still required.

## Open Questions

- [NEEDS CLARIFICATION: none blocking. Whether the user's real deck matches the S3 shape (title placeholder + countermeasure-restating body) vs. another shape is unconfirmed; the S3 fix also covers S1/S2/S4/S5/S6, so all probed shapes end excluded either way.]
