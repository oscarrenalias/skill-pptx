---
name: Add pypptx verify command for post-generation validation
id: spec-50e77785
description: "Add a generic `pypptx verify` CLI command that programmatically checks a .pptx file for common quality issues after generation"
dependencies: null
priority: high
complexity: medium
status: done
tags:
- cli
- qa
- verification
scope:
  in: null
  out: null
feature_root_id: null
---

## Background

After generating or editing a `.pptx` file, agents currently rely solely on
visual thumbnail inspection to catch quality issues. This requires LibreOffice
and Poppler, and depends on the agent's vision to spot problems. A programmatic
verifier catches a class of structural errors reliably and cheaply, without any
system dependencies.

Inspired by the verification approach in the PowerPoint Creator skill, but
reimplemented from scratch with only generic, template-agnostic checks.

## New CLI command

```bash
pypptx verify presentation.pptx
```

JSON output (default):

```json
{
  "errors": [
    "Slide 2: 'TextBox 3' has unfilled placeholder text — \"Click to add title\"",
    "Slide 4: 'TextBox 1' font 8.0pt is below minimum (12pt)"
  ],
  "warnings": [
    "Slide 3: 'Content Placeholder 2' text may be clipped (est=2.10\" vs box=1.80\", +17%)"
  ],
  "passed": false,
  "slide_count": 5,
  "error_count": 2,
  "warning_count": 1
}
```

`--plain` output (one issue per line, empty when clean):

```
FAIL  Slide 2: 'TextBox 3' has unfilled placeholder text — "Click to add title"
FAIL  Slide 4: 'TextBox 1' font 8.0pt is below minimum (12pt)
WARN  Slide 3: 'Content Placeholder 2' text may be clipped (est=2.10" vs box=1.80", +17%)
```

Exit code `0` when no errors (warnings are acceptable). Exit code `1` on any error.

## Checks to implement

All checks are generic — no hardcoded fonts, dimensions, locales, or template assumptions.

### 1. Unfilled placeholder hint text (ERROR)

Shapes whose text matches common default placeholder strings left by PowerPoint
or python-pptx when a placeholder was not filled in:

Patterns to match (case-insensitive):
- `"Click to add title"`
- `"Click to add text"`
- `"Click to add subtitle"`
- `"Click to edit Master title style"`
- `"Click to edit Master text styles"`
- `"Add title"`
- `"Add text"`
- `"Enter text here"`
- `"Place subtitle here"`

Skip shapes with no text frame. Skip shapes where `text.strip()` is empty
(blank placeholders are fine — they were intentionally left empty).

### 2. Font size below minimum (ERROR)

Walk all `a:rPr[@sz]` elements across all shapes. Report any font below 12pt
(1200 in hundredths-of-a-point units) as an error. This threshold catches
text that will be illegible when printed or presented.

Exclude shapes at the very bottom of the slide that are likely footers
(top > 90% of slide height) to avoid false positives on page-number fields.

### 3. Shape overflow (ERROR)

Compare each shape's bounding box against the slide dimensions:
- `left + width > slide_width` → overflows right
- `top + height > slide_height` → overflows bottom

Apply a tolerance of 50,000 EMU (~0.05") to ignore sub-pixel rounding.
Negative `left` or `top` values also count as overflow.

### 4. Text clipping (WARNING or ERROR)

Estimate the rendered height of each text frame by summing across paragraphs:
- Font size in EMU × 1.35 line-spacing multiplier × estimated line count
- Line count: 1 if `word_wrap=False`; otherwise `ceil(text_length / chars_per_line)`
  where `chars_per_line = box_width / avg_char_width` (avg char width ≈ 0.55× font size)

If estimated height exceeds shape height by more than 8%: WARNING.
If estimated height exceeds shape height by more than 20%: ERROR.

Skip shapes with no text, table shapes, and footer-region shapes.

### 5. Significant shape overlap (WARNING)

For each pair of shapes on a slide (excluding tiny decorative shapes smaller
than 50,000 EMU in either dimension), compute the intersection area. If the
intersection area is:
- Greater than 0.3 sq inches, AND
- Greater than 20% of the smaller shape's area

AND neither shape fully contains the other: report as a WARNING.

## Files to create/modify

| File | Change |
|---|---|
| `.apm/skills/pypptx/pypptx/ops/verify.py` | New — implements `verify_pptx()` returning `{errors, warnings, slide_count}` |
| `.apm/skills/pypptx/pypptx/cli.py` | Add `verify` command using `verify_pptx()` |
| `tests/test_verify.py` | New — tests for each check |
| `.apm/skills/pypptx/SKILL.md` | Add `pypptx verify` to QA checklist as step 1 (before thumbnails) |

## Acceptance criteria

- `pypptx verify clean.pptx` exits 0 and returns `{"passed": true, "errors": [], "warnings": [], ...}`.
- `pypptx verify bad.pptx` exits 1 when errors are present.
- Each of the 5 checks is covered by at least one test asserting it triggers correctly.
- Each check is covered by at least one test asserting it does NOT trigger on clean input.
- `--plain` output is tested.
- SKILL.md QA checklist updated: `pypptx verify` appears as step 1, before the thumbnail check.
- No Accenture-specific logic: no hardcoded font names, no template dimension constants,
  no locale-specific placeholder text patterns.
