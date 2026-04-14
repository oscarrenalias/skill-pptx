# pypptx — CLAUDE.md

## What this project is

`pypptx` is a Python CLI toolkit for manipulating `.pptx` files, intended primarily as a skill for AI agents. It exposes structural operations (unpack/pack/clean, slide reordering, text extraction, thumbnail generation) as a CLI with JSON output by default.

## Running the CLI

```bash
uv run pypptx --help
uv run pypptx slide list deck.pptx
uv run pypptx extract-text deck.pptx
```

For environments without `uv`, use the standalone entry point at the repo root:

```bash
python3 pypptx.py --help
python3 pypptx.py slide list deck.pptx
python3 pypptx.py extract-text deck.pptx
```

`pypptx.py` bootstraps its own `.venv` on first run, then re-execs inside it — no manual setup required.

## Running tests

```bash
uv run pytest              # all 159 tests
uv run pytest tests/test_slides.py   # single file
```

Always run tests with `uv run pytest`, not bare `pytest` — the package is not installed system-wide.

## Project structure

```
pypptx/
  cli.py          # click group + all commands; output_result() helper
  ops/
    pack.py       # unpack() and pack() — raw ZIP/OPC, no python-pptx
    clean.py      # clean_unused_files() — transitive .rels walk for orphan detection
    slides.py     # list_slides, list_layouts, add_slide, delete_slide, move_slide, pptx_edit
    extract.py    # extract_text()
    thumbnails.py # render pipeline: soffice → pdftoppm → Pillow grid
tests/
  conftest.py     # minimal_pptx and unpacked_pptx fixtures (generated via python-pptx)
  test_pack.py / test_clean.py / test_slides.py / test_extract.py / test_cli.py / test_thumbnails.py
specs/
  drafts/         # specs being written
  planned/        # specs with beads created, implementation in progress
  done/           # shipped specs
```

## Output contract (do not break)

- Every command writes a single JSON object to **stdout** by default.
- `--plain` switches to human-readable text.
- All errors go to **stderr**; stdout is always machine-parseable on success.
- Exit code `0` on success, `1` on any error.

## Library choices

| Layer | Library | Why |
|---|---|---|
| Slide management, text extraction | `python-pptx` | High-level API |
| Pack/unpack/clean | `zipfile` + `defusedxml` | python-pptx doesn't expose raw ZIP ops |
| XML parsing | `defusedxml` (always) | Security — never use `xml.etree.ElementTree` to parse untrusted XML |
| XML serialisation | `xml.etree.ElementTree` (stdlib) | Writing only; `register_namespace` before every write |
| Thumbnails | `soffice` + `pdftoppm` + `Pillow` | System tools for rendering; Pillow is optional |

**Critical:** `defusedxml` for all XML parsing. `xml.etree.ElementTree` is imported only for `register_namespace` and serialisation of trees we already own.

## Known quirks

- **`clean.py` path normalisation**: OPC relationship targets often contain `..` (e.g. `../slideLayouts/slideLayout1.xml`). `_normalize()` uses `os.path.normpath()` to resolve these before comparing against the reachable set — without this, all slide layouts are incorrectly treated as orphans and deleted.
- **`move_slide` and python-pptx renaming**: python-pptx calls `rename_slide_parts()` whenever `prs.slides` is accessed, renumbering slide XML files in `sldIdLst` order. This means slide filenames are not stable across open/move/re-open cycles. Test slide ordering via content attributes (e.g. `hidden`) not filenames.
- **`pptx_edit` context manager**: used by all write ops on `.pptx` files. Unpacks to a temp dir, yields it, then calls `clean_unused_files` + `pack` on clean exit. On exception, the original file is left untouched.

## Spec-driven development

New features are developed via specs. Use the spec management skill:

```bash
python3 .claude/skills/skill-spec-management/spec.py list
python3 .claude/skills/skill-spec-management/spec.py show spec-5636a620
```

Plan a spec with takt:

```bash
uv run takt plan specs/drafts/my-spec.md        # dry run
uv run takt plan --write specs/drafts/my-spec.md # persist beads
uv run takt --runner claude run --max-workers 4  # run scheduler
uv run takt merge <feature-root-bead-id>         # merge when done
```

Test gate for merges uses `uv run pytest` (configured in `.takt/config.yaml`).

## What NOT to do

- Do not use `xml.etree.ElementTree` to **parse** XML — use `defusedxml` instead.
- Do not run bare `pytest` — use `uv run pytest`.
- Do not commit `__pycache__/` or `.pyc` files — they block takt merges.
- Do not add content-authoring commands to this CLI (e.g. `set-text`). For content creation, agents write python-pptx scripts directly.
