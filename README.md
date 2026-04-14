# pypptx — PowerPoint skill for AI agents

`pypptx` is an AI agent skill for creating and editing `.pptx` files. It exposes
PowerPoint operations as a CLI with structured JSON output, designed to be called
by agents running in Claude Code or similar environments. It can also be used directly
from the terminal by humans.

## Installing the skill

```
apm install oscarrenalias/skill-pptx
```

After installation the skill is available via `python3 pypptx.py` inside the skill
directory, or as `uv run pypptx` if the package is installed in the project environment.

---

## Agent scenarios

### Scenario 1 — Create a presentation from scratch

Ask Claude Code to build a deck from a template or from nothing:

```
Create a 6-slide project status presentation using the template in template.pptx.
Slide 1 should be the cover slide with the project name and today's date.
Slides 2–5 should cover: overview, current status, risks, and next steps.
Slide 6 is a closing/Q&A slide.

Use the pypptx skill to inspect the available layouts first, then write a
python-pptx script to build the deck and run it.
After generating the file, run `pypptx verify` to check for quality issues,
then generate thumbnails so you can visually confirm the output looks correct.
```

Claude will typically:
1. Run `pypptx slide layouts template.pptx` to discover available layout names
2. Write a `create_deck.py` script using python-pptx and the skill's `.venv`
3. Run `pypptx verify output.pptx` to catch structural issues
4. Run `pypptx thumbnails output.pptx` and read the image to visually confirm the result

### Scenario 2 — Modify an existing deck

Ask Claude Code to make targeted edits to an existing presentation:

```
I have a deck in quarterly_review.pptx. Please:
- Move the "Risks" slide (currently slide 5) to be slide 3
- Delete the blank slide at position 7
- Extract all the text so I can review what's there

Use the pypptx skill. After making changes, run verify and generate thumbnails
so we can confirm everything looks right.
```

Claude will typically:
1. Run `pypptx slide list quarterly_review.pptx` to see the current structure
2. Run `pypptx extract-text quarterly_review.pptx` to read the content
3. Run `pypptx slide move` and `pypptx slide delete` to make the structural changes
4. Run `pypptx verify quarterly_review.pptx` to validate the result
5. Run `pypptx thumbnails quarterly_review.pptx` and read the image for visual QA

---

## Human usage

The CLI can also be called directly from the terminal.

```bash
uv run pypptx --help
uv run pypptx slide list deck.pptx
uv run pypptx extract-text deck.pptx
```

Without `uv`, use the self-bootstrapping entry point at the skill root:

```bash
python3 pypptx.py --help
python3 pypptx.py slide list deck.pptx
```

`pypptx.py` creates its own `.venv` on first run — no manual setup required.

---

## Installation (package)

```
pip install -e .
```

Or with [uv](https://github.com/astral-sh/uv):

```
uv pip install -e .
```

### Optional: thumbnails support

The `thumbnails` command requires [Pillow](https://pillow.readthedocs.io/) and two system tools.

```
pip install 'pypptx[thumbnails]'
```

| Tool | macOS | Debian/Ubuntu |
|---|---|---|
| LibreOffice (`soffice`) | `brew install --cask libreoffice` | `sudo apt-get install libreoffice` |
| Poppler (`pdftoppm`) | `brew install poppler` | `sudo apt-get install poppler-utils` |

---

## Output contract

- **Default output**: every command writes a single JSON object to stdout.
- **`--plain` flag**: pass `--plain` to receive human-readable text instead.
- **Errors**: all error messages are written to stderr, never stdout.
- **Exit codes**: `0` on success, `1` on any error.

---

## Commands

Run `pypptx --help` or `pypptx <command> --help` for the full option reference.

---

### `verify`

Run quality checks on a `.pptx` file. Catches unfilled placeholder text, font sizes
below 12pt, shapes that overflow the slide boundary, likely text clipping, and
significant shape overlaps. Run this after every generation or edit.

```
pypptx verify presentation.pptx
```

```json
{
  "errors": ["Slide 2: 'TextBox 3' has unfilled placeholder text — \"Click to add title\""],
  "warnings": ["Slide 3: 'Content Placeholder 2' text may be clipped (est=2.10\" vs box=1.80\", +17%)"],
  "passed": false,
  "slide_count": 5,
  "error_count": 1,
  "warning_count": 1
}
```

With `--plain`:

```
pypptx verify presentation.pptx --plain
```

```
FAIL  Slide 2: 'TextBox 3' has unfilled placeholder text — "Click to add title"
WARN  Slide 3: 'Content Placeholder 2' text may be clipped (est=2.10" vs box=1.80", +17%)
```

Exit code `0` when no errors (warnings are acceptable). Exit code `1` on any error.

---

### `extract-text`

Extract all text from a `.pptx` file.

```
pypptx extract-text presentation.pptx
```

```
--- Slide 1 ---
Welcome to pypptx
A PowerPoint manipulation toolkit
--- Slide 2 ---
Installation
pip install -e .
```

Limit to specific slides with `--slides`:

```
pypptx extract-text presentation.pptx --slides 1,3
```

With `--output`, text is written to the given file and command metadata is emitted as JSON:

```
pypptx extract-text presentation.pptx --output extracted.txt
```

```json
{"output_file": "extracted.txt", "slide_count": 2}
```

---

### `thumbnails`

Generate labeled thumbnail grid images from a `.pptx` file.
Requires LibreOffice, Poppler/pdftoppm, and Pillow — see [thumbnails support](#optional-thumbnails-support).

```
pypptx thumbnails presentation.pptx
```

```json
{"files": ["thumbnails.jpg"]}
```

**Options:**

| Option | Default | Description |
|---|---|---|
| `--output PREFIX` | `thumbnails` | Output filename prefix; `.jpg` is appended automatically. |
| `--cols N` | `3` (max `6`) | Number of columns in the thumbnail grid. |
| `--plain` | off | Emit one file path per line instead of JSON. |

Hidden slides appear as a hatched grey placeholder so the grid index always matches
the slide number. Large decks are split across multiple JPEG files automatically.

---

### `slide list`

List slides in presentation order.

```
pypptx slide list presentation.pptx
```

```json
{"slides": [
  {"index": 1, "file": "slide1.xml", "hidden": false},
  {"index": 2, "file": "slide2.xml", "hidden": true}
]}
```

---

### `slide layouts`

List all slide layouts with their index and name.

```
pypptx slide layouts presentation.pptx
```

```json
{"layouts": [
  {"index": 1, "file": "slideLayout1.xml", "name": "Title Slide"},
  {"index": 2, "file": "slideLayout2.xml", "name": "Title and Content"},
  {"index": 3, "file": "slideLayout3.xml", "name": "Section Header"}
]}
```

With `--plain`:

```
slideLayout1.xml  Title Slide
slideLayout2.xml  Title and Content
slideLayout3.xml  Section Header
```

Always run this before choosing a layout index for `slide add`. Layout index 1 is
almost always the cover slide — never use it for regular content slides.

---

### `slide add`

Add a slide to a `.pptx` file. Use `slide layouts` to find the layout index.

```
pypptx slide add presentation.pptx --layout 2
pypptx slide add presentation.pptx --duplicate 2
pypptx slide add presentation.pptx --duplicate 1 --position 2
```

---

### `slide delete`

Delete a slide by its 1-based index.

```
pypptx slide delete presentation.pptx 3
```

---

### `slide move`

Move a slide from one 1-based position to another.

```
pypptx slide move presentation.pptx 3 1
```

---

### `unpack` / `clean` / `pack`

Structural editing workflow for XML-level changes:

```
pypptx unpack presentation.pptx      # expand to directory
# edit XML files directly
pypptx clean presentation/           # remove orphaned parts
pypptx pack presentation/ output.pptx
```

Most slide commands accept a `.pptx` file directly and handle unpack/clean/repack
internally. The explicit workflow is only needed when editing XML by hand.
