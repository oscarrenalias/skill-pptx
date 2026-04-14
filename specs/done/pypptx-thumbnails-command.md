---
name: pypptx thumbnails command
id: spec-80d46efe
description: Adds a thumbnails command to pypptx that renders each slide as a JPEG and stitches them into a labeled grid image for visual inspection.
dependencies: spec-b1c47a0f
priority: low
complexity: medium
status: done
tags:
- cli
- pptx
- python
scope:
  in: pypptx thumbnails command and ops/thumbnails.py
  out: "slide editing, text extraction, any other command"
feature_root_id: null
---
# pypptx thumbnails command

## Objective

Add a `pypptx thumbnails` command that converts each slide of a `.pptx` file into a JPEG and stitches them into a labeled grid image, primarily for visual inspection by agents doing QA on generated or edited presentations.

## Background

This is a follow-on to `spec-b1c47a0f` (core pypptx CLI). The thumbnails command requires system-level dependencies (LibreOffice, Poppler) and an optional Python dependency (Pillow) that are out of scope for the core CLI.

## Command

### `pypptx thumbnails <file.pptx> [--output PREFIX] [--cols N]`

Converts each slide to a JPEG and stitches them into a labeled grid image. Pipeline: LibreOffice (`soffice`) renders the `.pptx` to PDF → `pdftoppm` converts each PDF page to a JPEG → `Pillow` stitches them into a grid with slide-number labels.

For large decks, multiple grid files are created, each capped at `cols × (cols + 1)` slides, suffixed `-1`, `-2`, etc.

Options:
- `--output PREFIX` — output filename prefix (default: `thumbnails`)
- `--cols N` — number of columns in the grid (default: 3, max: 6)

**stdout (JSON):**
```json
{ "files": ["thumbnails.jpg"] }
```
Multi-file: `{ "files": ["thumbnails-1.jpg", "thumbnails-2.jpg"] }`.

**stdout (--plain):**
```
thumbnails.jpg
```

## Implementation Notes

- Pipeline: `soffice --headless --convert-to pdf` → `pdftoppm -jpeg -r 150` → `Pillow` grid assembly.
- Each slide image is labeled with its slide number.
- Hidden slides are shown as a hatched grey placeholder so slide numbering stays consistent with the file.
- If `Pillow` is not installed or `soffice`/`pdftoppm` are not on PATH, print a clear install hint to stderr and exit non-zero.
- `Pillow` is an optional dependency: `pip install 'pypptx[thumbnails]'`.

## Files to Modify

| File | Change |
|---|---|
| `pyproject.toml` | Add `thumbnails` optional dependency group: `Pillow>=10.0` |
| `pypptx/ops/thumbnails.py` | New — full pipeline implementation |
| `pypptx/cli.py` | Add `thumbnails` command |

## Acceptance Criteria

- `pypptx thumbnails deck.pptx` produces `thumbnails.jpg` containing a grid of slide images labeled by slide number.
- `pypptx thumbnails deck.pptx --cols 4 --output out` produces `out.jpg` with 4 columns.
- For a deck with more than `cols × (cols + 1)` slides, multiple files are produced (`thumbnails-1.jpg`, `thumbnails-2.jpg`, …).
- Output is JSON `{"files": [...]}` listing all produced files.
- `--plain` prints one file path per line.
- If `Pillow` is not installed, prints `pip install 'pypptx[thumbnails]'` to stderr and exits non-zero.
- If `soffice` or `pdftoppm` are not on PATH, prints install instructions to stderr and exits non-zero.

## Pending Decisions

- Should hidden slides be shown as placeholders (current proposal: yes, to preserve index alignment) or skipped entirely?
