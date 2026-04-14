---
name: pypptx CLI — PowerPoint manipulation toolkit
id: spec-b1c47a0f
description: "A Python CLI package for manipulating .pptx files: pack/unpack, slide management, text extraction, thumbnails"
dependencies: null
priority: high
complexity: high
status: done
tags:
- cli
- pptx
- python
scope:
  in: pypptx package and all subcommands
  out: "SKILL.md authoring, GUI, DOCX/XLSX support"
feature_root_id: B-f993e1f7
---
# pypptx CLI — PowerPoint manipulation toolkit

## Objective

Create a `pypptx` Python CLI package that exposes a clean, composable interface for manipulating `.pptx` files. A `.pptx` is a ZIP archive of XML files (Office Open XML / OPC format). The CLI should make common operations — unpacking for editing, slide management, text extraction, and visual inspection — first-class commands.

## Background

The Anthropic `pptx` skill uses a set of Python helper scripts to manipulate `.pptx` files: unpack/pack (zip/unzip), clean orphaned XML files after editing, add/delete slides, and generate thumbnail grids for visual QA. This project recreates that functionality as a standalone, installable CLI — independently implemented, with no code copied from Anthropic.

The core editing workflow is: `unpack` → edit XML → `clean` → `pack`. Slide subcommands shortcut this by auto-unpacking to a temp dir, operating, and repacking atomically.

## Output Contract

All commands write structured data to **stdout** and errors/warnings to **stderr**. Exit code is `0` on success, non-zero on failure. This makes every command composable and agent-parseable without screen-scraping.

By default output is **JSON**. Pass `--plain` to any command for a human-readable plain-text alternative (useful when invoking from a shell script or inspecting manually).

Error output (stderr) is always plain text regardless of `--plain`.

## Commands

### `pypptx unpack <file.pptx> [output_dir]`

Extracts the archive to a directory (defaults to filename without extension). Files are extracted as-is with no post-processing.

**stdout (JSON):**
```json
{ "unpacked_dir": "deck" }
```

**stdout (--plain):**
```
deck
```

---

### `pypptx pack <unpacked_dir> <output.pptx>`

Rezips an unpacked directory into a valid `.pptx`. `[Content_Types].xml` must be the first ZIP entry per the OPC spec. Writes atomically (temp file → rename).

**stdout (JSON):**
```json
{ "output_file": "deck.pptx" }
```

**stdout (--plain):**
```
deck.pptx
```

---

### `pypptx clean <file_or_dir>`

Removes orphaned files. If given a `.pptx`, unpacks to temp, cleans, repacks in-place. If given a directory, operates directly.

Orphaned files include: slides not listed in `sldIdLst`, their `.rels` files, and anything not reachable by a transitive walk of all `.rels` files (media, embeddings, charts, diagrams, drawings, tags, ink, notesSlides). Updates `[Content_Types].xml` and `presentation.xml.rels` accordingly. The walk is iterative — repeat until no more removals.

**stdout (JSON):**
```json
{ "removed": ["ppt/slides/slide3.xml", "ppt/media/image2.png"] }
```
`removed` is an empty array `[]` when nothing was removed.

**stdout (--plain):**
```
ppt/slides/slide3.xml
ppt/media/image2.png
```
Empty output when nothing was removed.

---

### `pypptx extract-text <file.pptx> [--slides 1,3,5] [--output file.txt]`

Extracts text from all slides (or a subset) in reading order (top-to-bottom, left-to-right by shape position). Requires `python-pptx` (optional dependency).

When `--output` is given, writes to that file and emits a JSON confirmation to stdout. When omitted, writes content to stdout (always plain text regardless of `--plain`, since the content itself is the output).

**stdout (JSON, with --output):**
```json
{ "output_file": "deck.txt", "slide_count": 12 }
```

**stdout (plain text, without --output):**
```
--- Slide 1 ---
Title text here
Body paragraph one
Body paragraph two

--- Slide 2 ---
Another title
...
```

---

### `pypptx slide list <file_or_dir>`

Lists slides in presentation order.

**stdout (JSON):**
```json
{
  "slides": [
    { "index": 1, "file": "slide1.xml", "hidden": false },
    { "index": 2, "file": "slide2.xml", "hidden": false },
    { "index": 3, "file": "slide4.xml", "hidden": true }
  ]
}
```

**stdout (--plain):**
```
1  slide1.xml
2  slide2.xml
3  slide4.xml  [hidden]
```

---

### `pypptx slide add <file.pptx> [--duplicate N] [--layout N] [--position N]`

Adds a slide by duplicating an existing one (`--duplicate`) or creating a blank from a layout (`--layout`). Inserts at `--position` (1-based, default: end). Auto-unpacks, operates, cleans, repacks. Exactly one of `--duplicate` or `--layout` must be provided.

When duplicating, the notes slide relationship is stripped from the new slide's `.rels` so the copy starts without speaker notes.

**stdout (JSON):**
```json
{ "added_file": "slide5.xml", "position": 3 }
```

**stdout (--plain):**
```
slide5.xml at position 3
```

---

### `pypptx slide delete <file.pptx> <index>`

Deletes the slide at 1-based index. Auto-unpacks, operates, cleans, repacks.

**stdout (JSON):**
```json
{ "deleted_file": "slide2.xml", "deleted_index": 2 }
```

**stdout (--plain):**
```
deleted slide2.xml (was at index 2)
```

---

### `pypptx slide move <file.pptx> <from> <to>`

Moves a slide from one 1-based position to another by reordering entries in `sldIdLst`. Auto-unpacks, operates, repacks.

**stdout (JSON):**
```json
{ "file": "slide3.xml", "from": 3, "to": 1 }
```

**stdout (--plain):**
```
moved slide3.xml from 3 to 1
```

---

### `pypptx slide layouts <file_or_dir>`

Lists available slide layouts from `ppt/slideLayouts/`.

**stdout (JSON):**
```json
{
  "layouts": [
    { "index": 1, "file": "slideLayout1.xml" },
    { "index": 2, "file": "slideLayout2.xml" }
  ]
}
```

**stdout (--plain):**
```
1  slideLayout1.xml
2  slideLayout2.xml
```

## Implementation Notes

### Library strategy

This project uses a hybrid approach based on what each library is good at:

| Operation | Library | Reason |
|---|---|---|
| `unpack` / `pack` | `zipfile` + `defusedxml` | Need raw ZIP control; OPC entry ordering matters |
| `clean` | `zipfile` + `defusedxml` | Orphan detection requires walking all `.rels` files — not modelled by python-pptx |
| `slide list/add/delete/move` | `python-pptx` | High-level API handles `sldIdLst`, relationship IDs, and Content_Types correctly |
| `slide layouts` | `python-pptx` | `prs.slide_layouts` exposes this directly |
| `extract-text` | `python-pptx` | Clean iteration over shapes and paragraphs; handles reading order via shape positions |

`python-pptx` is a **required** dependency since it covers slide management and text extraction, both core functionality. No optional dependencies in this spec — thumbnails are tracked separately.

### Key design constraints

- `defusedxml` is used for all XML parsing in `pack.py` and `clean.py` to safely handle untrusted files. `xml.etree.ElementTree` must not be used for writing — it mangles namespace prefixes.
- The `pptx_edit` context manager handles auto-unpack/clean/repack for slide commands: on success it cleans and repacks; on exception it does **not** repack, preserving the original file.
- Slide indexing is always 1-based in the CLI interface; internally `python-pptx` uses 0-based list indices.
- `python-pptx` slide objects expose `._element` for direct lxml access when the high-level API doesn't reach far enough.

## Package Structure

```
pyproject.toml
pypptx/
  __init__.py        # __version__
  cli.py             # click group + all command definitions
  ops/
    __init__.py
    pack.py          # unpack(), pack() — zipfile; defusedxml for XML parsing in clean
    clean.py         # clean_unused_files() — zipfile + defusedxml
    slides.py        # list_slides(), list_layouts(), add_slide(),
                     # add_slide_from_layout(), delete_slide(), move_slide()
                     # — python-pptx
    extract.py       # extract_text() — python-pptx
```

## Files to Modify

| File | Change |
|---|---|
| `pyproject.toml` | New — hatchling build, entry point `pypptx`, deps: click, defusedxml, python-pptx; optional: Pillow |
| `pypptx/__init__.py` | New |
| `pypptx/cli.py` | New — click group and all commands |
| `pypptx/ops/pack.py` | New — zipfile + defusedxml |
| `pypptx/ops/clean.py` | New — zipfile + defusedxml |
| `pypptx/ops/slides.py` | New — python-pptx |
| `pypptx/ops/extract.py` | New — python-pptx |

## Acceptance Criteria

- `pip install -e .` succeeds and `pypptx --version` prints the version.
- `pypptx unpack deck.pptx` outputs valid JSON `{"unpacked_dir": "..."}` to stdout and produces a directory with the extracted files unchanged.
- `pypptx pack unpacked/ out.pptx` outputs valid JSON `{"output_file": "..."}` and produces a file that PowerPoint/LibreOffice can open.
- `pypptx clean deck.pptx` outputs valid JSON `{"removed": [...]}` with a list of removed paths; `[]` when nothing removed.
- `pypptx slide list deck.pptx` outputs valid JSON with a `slides` array; each entry has `index`, `file`, `hidden`.
- `pypptx slide add deck.pptx --duplicate 1` outputs valid JSON `{"added_file": "...", "position": N}`; the resulting file opens correctly.
- `pypptx slide delete deck.pptx 2` outputs valid JSON `{"deleted_file": "...", "deleted_index": 2}`; the resulting file opens correctly.
- `pypptx slide move deck.pptx 3 1` outputs valid JSON `{"file": "...", "from": 3, "to": 1}`; slide order in the file is correct.
- `pypptx slide layouts deck.pptx` outputs valid JSON with a `layouts` array; each entry has `index` and `file`.
- `pypptx extract-text deck.pptx` writes plain text with `--- Slide N ---` delimiters to stdout.
- `pypptx extract-text deck.pptx --output out.txt` writes to file and outputs JSON `{"output_file": "...", "slide_count": N}`.
- All commands: `--plain` flag switches stdout to human-readable text; errors always go to stderr; exit code is non-zero on failure.
- `pypptx --help` and each subcommand `--help` print usage.

## Test Fixtures

All tests use programmatically generated PPTX files — no binary fixture files are committed to the repo. `conftest.py` uses `python-pptx` to create minimal valid decks as pytest fixtures scoped to `tmp_path`, so each test gets an isolated copy. Different fixtures cover different shapes (single slide, multi-slide, deck with hidden slide, deck with known text content). A committed `.pptx` fixture is only warranted if a test requires complex embedded content (charts, media) that is impractical to generate — none of the commands in this spec require that.

## Decisions

- `unpack` does not post-process files in any way — it is a pure unzip operation.
- `slide add --duplicate` always strips the notes slide relationship from the copy.
- JSON is the default output format; `--plain` is the opt-out for human use.

