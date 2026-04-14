---
name: Improve slide layout selection guidance for agents
id: spec-f8f5c045
description: Add layout names to slide layouts CLI output and strengthen SKILL.md guidance to prevent agents defaulting to the cover slide layout
dependencies: null
priority: high
complexity: small
status: done
tags:
- cli
- skill
- ux
scope:
  in: null
  out: null
feature_root_id: null
---

## Background

Agents using the skill default to `slide_layouts[0]`, which is almost always
the cover/title slide layout. This produces decks where every slide looks like
a cover page. Additionally, agents use `add_textbox()` instead of writing into
layout placeholders, placing unstyled text boxes on top of branded backgrounds.

## Changes required

### 1. CLI — add `name` to `slide layouts` output

`list_layouts()` in `.apm/skills/pypptx/pypptx/ops/slides.py` currently returns:

```python
[{"index": 1, "file": "slideLayout1.xml"}, ...]
```

Add the layout name via `layout.name` from python-pptx:

```python
[{"index": 1, "file": "slideLayout1.xml", "name": "Title Slide"}, ...]
```

The `--plain` output should become one entry per line in the form:

```
slideLayout1.xml  Title Slide
slideLayout2.xml  Title and Content
slideLayout3.xml  Section Header
```

### 2. SKILL.md — layout selection and placeholder guidance

Add to the **Creating a presentation from scratch** section:

#### Always inspect layouts before choosing one

Never hardcode a layout index. Always run `slide layouts` on the target deck
first to see available layout names, then choose the appropriate one:

```bash
python3 pypptx.py slide layouts presentation.pptx
```

**Layout index 0 / entry 1 is almost always the cover slide** ("Title Slide").
Never use it for regular content slides.

Common layout names and when to use them:

| Name | Use for |
|---|---|
| Title Slide | Cover slide only (first slide) |
| Title and Content | Standard body slides |
| Section Header | Section dividers |
| Two Content | Side-by-side content |
| Title Only | Slides with custom content below the title |
| Blank | Fully custom slides (no placeholders) |

#### Always use placeholders, never `add_textbox()`

Before writing content, inspect what placeholders the chosen layout provides:

```python
for ph in slide.placeholders:
    print(ph.placeholder_format.idx, ph.name)
```

Write into them by index:

```python
slide.placeholders[0].text = "My Title"
slide.placeholders[1].text = "Body text"
```

**Never use `add_textbox()`** on a slide that has a layout with placeholders.
It places an unstyled text box on top of the master design — wrong font, wrong
colour, wrong position, often unreadable against the background.
`add_textbox()` is only appropriate for fully blank slides (layout "Blank")
where no placeholders exist.

## Acceptance criteria

- `pypptx slide layouts <file>` JSON output includes a `name` field for every layout entry.
- `pypptx slide layouts <file> --plain` shows filename and name on each line.
- Existing layout tests updated; new test asserts `name` field is a non-empty string.
- SKILL.md updated with layout selection table, placeholder inspection snippet,
  and strong prohibition on `add_textbox()` for slides with layouts.
