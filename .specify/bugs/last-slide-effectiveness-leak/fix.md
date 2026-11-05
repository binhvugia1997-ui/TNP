# Fix: Last slide "Hiệu quả cải tiến" still contributes its content (last-slide-effectiveness-leak)

- **Slug**: last-slide-effectiveness-leak
- **Assessment**: `.specify/bugs/last-slide-effectiveness-leak/assessment.md`
- **Status**: fixed (Linux dev evidence only; Windows/PPTX acceptance still required)
- **Mode B disclosure**: Executed from checked-in Spec Kit skill instructions; native Arena skill invocation was not available.

## Root cause (confirmed by reproduction)

The exclusion vocabulary existed (`VERIFY_MARKERS` / `FOLLOWUP_MARKERS` in `app/classifier.py`, `EXCLUDED_SECTION_KINDS` + the kind filter in `app/extractor.py:780-805`), but the segmentation layer lost the slide title's excluded kind in two places:

1. `app/extractor.py::split_sections` — a `ROLE_TITLE` block set the current section kind, but the section had no lines yet, so the slide's first content heading silently replaced it. A final slide titled "Hiệu quả cải tiến" whose body restates the countermeasure ("Đối sách cải tiến:", "Sau cải tiến: ...") therefore became two `improvement` sections. `classify_blocks` also preserves document order, so a title shape ordered after a body shape arrived too late to protect already-consumed lines.
2. `app/extractor.py::collect_sections` — default-kind precedence checked `improvement_slides` membership before `verify_slides`, so a slide in both lists defaulted to `improvement` for lines before the first heading.
3. `app/extractor.py::excluded_bands` — the geometric twin: a title-opened excluded band was closed by the first non-excluded heading, so pictures below it on an effectiveness slide could be selected as After evidence.

Reproduced with simulated slide data (probe, 6 shapes): S1/S2/S4/S5/S6 excluded correctly, S3 (title placeholder + improvement-like first body heading) leaked the whole effectiveness body into the improvement text.

## Changes

- `app/extractor.py`:
  - `split_sections`: the first title block decides the slide before the loop runs (pre-scan of `classify_blocks` roles); when the title declares an excluded kind, the whole slide starts in that kind and included-kind body headings join the excluded section instead of opening new sections. A later title block with a non-excluded kind clears the stickiness; excluded-kind headings still open their own excluded sections. Mid-slide headings on non-title-led slides split exactly as before (PROMPT-001 S6 contract preserved).
  - `excluded_bands`: a band opened by an excluded title covers the whole slide (never closed by body headings); bands opened by mid-slide excluded headings keep the old close-on-next-heading rule.
  - `collect_sections`: a verify-led slide (first heading is `verify`/`followup`) defaults to that excluded kind even when also listed as an improvement slide. Requires the new `first_heading_kind` import from `.classifier`.
- `tests/test_effectiveness_slide_exclusion.py` (new, 8 tests): end-to-end heuristic classification; LLM-style dual-listed classification through `extract_record`; body-before-title shape; trailing effectiveness block still splits; picture bands cover the whole effectiveness slide and are empty for production slides; mid-slide excluded block still closes at the next heading; `first_heading_kind` unit check; header-band effectiveness heading does not swallow a later long-term heading (A1285 shape, see Correction above). Fixtures are simulated `SlideData`/`Block` objects, clearly shaped like real decks.

## Behavior deliberately NOT changed

- The exclusion vocabulary, `improvement_subkind`, `_resolve_long_term`, and the `kind in ("improvement", "standard")` Excel filters are untouched.
- No version/build change, no test weakening: `tests/test_prompt001.py` (PROMPT-001 segmentation contracts) passes unchanged.

## Verification (Linux)

- New tests: 8/8 pass.
- Probe (6 shapes): all excluded after the fix.
- `tests/test_prompt001.py`, `test_content_learning.py`, `test_content_region.py`, `test_content_group_candidates.py`, `test_core.py`, `test_improvement_visual_regions.py`, `test_after_evidence.py`, `test_image_learning.py`: 140 passed / 4 failed — all 4 failures are the Bug A PowerPoint-only gate mechanism (embedded-picture fallback instead of builtin renders through the batch pipeline), not this change; see the Bug A fix record. No failure asserts segmentation text.

## Correction (2026-11-04): title-led stickiness applies ONLY to real title placeholders

The first version of the `split_sections` / `excluded_bands` change treated ANY `ROLE_TITLE` block as the slide title. `classify_blocks` also labels header-band textboxes `ROLE_TITLE`, so a slide with NO title placeholder — the real A1285 deck's slide 8, whose header-band "7. HIỆU QUẢ SAU CẢI TIẾN" heading is followed by a later "ĐỐI SÁCH LÂU DÀI" heading — had its long-term countermeasure body swallowed into the excluded section. `tests/test_real_a1285.py::test_a1285_record_contents` caught it ("ĐỐI SÁCH LÂU DÀI" / "Cập nhật SOP Rev.03" missing from the improvement text).

Fix: stickiness (`title_excluded_kind` in `split_sections`, `band_title_led` in `excluded_bands`) is set only when the title block is a real title placeholder (`block.kind == "title"`). A header-band textbox title opens its section as before but imposes no stickiness, so later headings on the same slide split normally. New regression test `test_header_band_effectiveness_heading_does_not_swallow_later_long_term_heading` pins the A1285 shape (header-band effectiveness heading + later long-term heading: sections `[verify, improvement]`, long-term text kept, effectiveness body excluded). The S3 title-placeholder shape is unchanged — all 8 Bug C tests pass, and `test_real_a1285.py` passes again.

## Still needs Windows verification

- A real countermeasure deck whose last slide is "Hiệu quả cải tiến" (title placeholder + countermeasure-restating body, as in the S3 fixture) run through the cloud Windows build: the Excel "Nội dung đối sách cải tiến" field must contain zero effectiveness-slide lines, and no After picture may come from that slide.
