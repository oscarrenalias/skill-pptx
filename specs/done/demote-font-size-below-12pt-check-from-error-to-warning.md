---
name: Demote font-size-below-12pt check from error to warning
id: spec-ac19111d
description: Reclassify the verify font-size-below-12pt check from a hard error to a warning
dependencies: null
priority: low
complexity: null
status: done
tags: []
scope:
  in: null
  out: null
feature_root_id: B-8b50ace0
---
# Demote font-size-below-12pt check from error to warning

## Objective

The `verify` command currently treats any run with a font size below 12pt as a hard **error**, causing `verify_pptx()` to report it in the `errors` list and the CLI to exit non-zero. This is too strict — small fonts are sometimes intentional (footnotes, captions, watermarks). Reclassify this check as a **warning** so the deck still passes verification while the issue is surfaced for human review.

## Changes

### `_check_font_sizes` in `verify.py`

- Change the function signature: replace the `errors: list[str]` parameter with `warnings: list[str]`.
- Change the `errors.append(...)` call inside to `warnings.append(...)`. The message text remains unchanged: `"Slide N: 'ShapeName' font X.Xpt is below minimum (12pt)"`.
- Update the docstring to reflect it now produces a warning, not an error.

### Call site in `verify_pptx()`

- Update the call `_check_font_sizes(slide_index, slide, slide_height, errors)` to pass `warnings` instead of `errors`.

### Tests

- In `TestCheck2FontSize` in `tests/test_verify.py`, update all assertions that currently check `result["errors"]` for font-size violations to check `result["warnings"]` instead.
- Ensure `result["errors"]` is empty for these cases (font-size below 12pt must no longer populate the errors list).

## Files to Modify

| File | Change |
|---|---|
| `.apm/skills/pypptx/pypptx/ops/verify.py` | Change `_check_font_sizes` to append to `warnings`; update call site |
| `tests/test_verify.py` | Update `TestCheck2FontSize` assertions from `errors` to `warnings` |

## Acceptance Criteria

- `verify_pptx()` returns an empty `errors` list for a deck whose only issue is a font below 12pt.
- `verify_pptx()` returns a non-empty `warnings` list containing the font-size message for that same deck.
- `uv run pytest tests/test_verify.py` passes with no failures.
- `uv run pytest` (full suite) passes.
