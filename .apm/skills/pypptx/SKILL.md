---
name: pypptx
description: Python CLI for reading, editing, and creating PowerPoint .pptx files
author: Renalias, Oscar
tags:
  - powerpoint
  - pptx
  - presentations
  - office
entry_point: pypptx.py
requires:
  optional:
    - LibreOffice (soffice) — needed for thumbnails command
    - Poppler (pdftoppm) — needed for thumbnails command
    - Pillow — needed for thumbnails command (installed automatically via pip)
---

# pypptx

A Python CLI for reading, editing, and creating PowerPoint `.pptx` files.

## Quick reference

| Task | How |
|---|---|
| Read slide text | `python3 pypptx.py extract-text <file>` |
| List slides | `python3 pypptx.py slide list <file>` |
| Add / delete / move slides | `python3 pypptx.py slide add/delete/move <file> ...` |
| Visual overview | `python3 pypptx.py thumbnails <file>` |
| Structural edit (XML) | unpack → edit → clean → pack |
| Create from scratch | Write a python-pptx script, run it via `.venv/bin/python3` |

The entry point is `python3 pypptx.py` at the repo root. It self-bootstraps a
`.venv` on first run with no external tooling required.

---

## Reading content

Extract all text from a presentation:

```bash
python3 pypptx.py extract-text presentation.pptx
```

Limit to specific slides with `--slides 1,3`. Output goes to stdout (no JSON
wrapper) unless `--output <file>` is given, in which case command metadata is
emitted as JSON.

---

## Visual inspection

Generate a labeled thumbnail grid to see slide layout at a glance:

```bash
python3 pypptx.py thumbnails presentation.pptx
```

Requires LibreOffice and Poppler — see README for installation. Use this to
verify edits look correct before delivering or committing a file. Hidden slides
appear as a hatched grey placeholder so grid index always matches slide number.

---

## Editing an existing presentation

Use the unpack → edit → clean → pack workflow for structural changes:

```bash
python3 pypptx.py unpack presentation.pptx      # expand to directory
# manipulate slides, edit XML, or run a python-pptx script against the directory
python3 pypptx.py clean presentation/           # remove orphans
python3 pypptx.py pack presentation/ output.pptx
```

Most slide commands also accept a `.pptx` file directly and handle
unpack/clean/repack internally:

```bash
python3 pypptx.py slide add presentation.pptx --duplicate 2
python3 pypptx.py slide delete presentation.pptx 3
python3 pypptx.py slide move presentation.pptx 3 1
```

Use `slide list` to confirm slide order and `slide layouts` to find layout
indices for `slide add --layout`.

---

## Creating a presentation from scratch

Write a python-pptx script, then run it using the skill's virtual environment.
If this is the first time running the skill, bootstrap the venv first:

```bash
python3 pypptx.py --version   # triggers first-run bootstrap
```

Then write your script and run it:

```bash
.venv/bin/python3 create_deck.py
```

Example `create_deck.py`:

```python
from pptx import Presentation
from pptx.util import Inches, Pt

# If a template file is provided, pass it here — it loads the slide master,
# layouts, theme, and color palette. If no template is provided, omit the
# argument and python-pptx uses its built-in blank default.
prs = Presentation('template.pptx')  # or: Presentation()

slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
slide.shapes.title.text = "Introduction"
slide.placeholders[1].text = "Key points go here"
prs.save("output.pptx")
```

Use `python3 pypptx.py slide layouts <file>` on the template (or any existing
deck) to find the layout indices available in that theme.

### Design guidance

**Colors** — use the theme palette from an existing file where possible.
Extract it with `python3 pypptx.py unpack` and inspect
`ppt/theme/theme1.xml`. Avoid free-form hex values unrelated to the deck's
palette.

**Typography** — respect the slide master's font stack. Prefer placeholder
text frames over adding raw text boxes; placeholders inherit master styles.

**Spacing** — use `Inches()` and `Pt()` from `pptx.util` for all measurements.
Leave slide margins of at least 0.5 in on all sides.

**Common mistakes to avoid**
- Adding a text box when a placeholder would do — breaks theme inheritance
- Hard-coding RGB colors not present in the deck's theme
- Forgetting to call `prs.save()` at the end of a creation script

---

## Output contract

All commands write a single JSON object to stdout by default.
Pass `--plain` for human-readable text. Errors go to stderr. Exit code 0 on
success, 1 on any error.

---

## QA checklist

Before delivering or committing a modified or newly created `.pptx`:

1. **Re-open with python-pptx** — must not raise:
   ```bash
   .venv/bin/python3 -c "from pptx import Presentation; Presentation('output.pptx')"
   ```
2. **Slide count** — `python3 pypptx.py slide list output.pptx` matches expectation.
3. **Text check** — `python3 pypptx.py extract-text output.pptx` to verify content landed in the right slides.
4. **No orphans** — `python3 pypptx.py clean output.pptx` returns `{"removed": []}`.
5. **Visual check** (strongly recommended) — `python3 pypptx.py thumbnails output.pptx`
   and inspect the grid image. Skip this step only if LibreOffice or Poppler
   is not available in the environment.

---

## Dependencies

The skill self-bootstraps on first run. The following are installed automatically
into `.venv/`:

- `python-pptx` — core presentation manipulation
- `click` — CLI framework
- `defusedxml` — safe XML parsing

Optional (for `thumbnails`): `Pillow`, LibreOffice (`soffice`), Poppler (`pdftoppm`).
See README for system installation instructions.
