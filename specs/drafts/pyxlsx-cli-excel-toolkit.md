---
name: pyxlsx CLI — Excel manipulation toolkit
id: spec-c4e8b12f
description: "A Python CLI package for reading and writing .xlsx files: workbook inspection, sheet management, table read/write, cell operations, and raw pack/unpack"
dependencies: null
priority: high
complexity: medium
status: draft
tags:
  - cli
  - xlsx
  - excel
  - python
scope:
  in: ".apm/skills/pyxlsx/ — all source, SKILL.md, apm.yml, pyproject.toml, self-bootstrap entry point; tests/test_xlsx_*.py"
  out: "formula recalculation, chart operations, conditional formatting, pivot tables, LibreOffice integration, DOCX/PPTX support, coupling to pypptx"
feature_root_id: null
---

# pyxlsx CLI — Excel manipulation toolkit

## Objective

Create a `pyxlsx` Python CLI package that exposes a clean, agent-friendly interface for reading and writing `.xlsx` files. The primary use cases are:

1. **Extracting structured data** from Excel workplans (e.g. task lists, schedules) so an agent can process them and produce output in other formats (PowerPoint roadmaps, reports).
2. **Creating and maintaining simple tables** — updating individual cells, adding or removing sheets, and writing table data via short `openpyxl` scripts.

For table creation and writes, agents write self-contained Python scripts using `openpyxl` directly and run them via the skill's `.venv`. The CLI provides read-oriented and structural commands; `SKILL.md` provides ready-made openpyxl patterns for writes.

## Background

`pyxlsx` is an independent skill. It lives at `.apm/skills/pyxlsx/` and has no coupling to `pypptx`. The two skills may be used together in an agent workflow (extract from Excel, write to PowerPoint) but they do not share code or configuration.

The design deliberately omits formula recalculation (which requires LibreOffice), complex formatting commands, and content-authoring operations. Agents handle those by writing short `openpyxl` scripts and running them through `.venv/bin/python3` — a pattern documented in `SKILL.md`.

The `table read` command returns an array-of-objects format that agents can process and hand directly to an openpyxl write script.

## Output Contract

All commands write a single JSON object to **stdout** by default. Pass `--plain` to any command for human-readable text. Errors go to **stderr**. Exit code `0` on success, `1` on any error.

## Commands

### `pyxlsx info <file>`

Workbook-level metadata: sheet names and named ranges.

**stdout (JSON):**
```json
{
  "file": "workplan.xlsx",
  "sheets": ["Q1 Plan", "Q2 Plan", "Summary"],
  "named_ranges": ["Budget", "Timeline"]
}
```

**stdout (--plain):**
```
workplan.xlsx
Sheets: Q1 Plan, Q2 Plan, Summary
Named ranges: Budget, Timeline
```

---

### `pyxlsx sheet list <file>`

Lists all sheets with row/column counts and visibility.

**stdout (JSON):**
```json
{
  "sheets": [
    { "name": "Q1 Plan", "rows": 45, "cols": 8, "visible": true },
    { "name": "Q2 Plan", "rows": 32, "cols": 8, "visible": true },
    { "name": "Summary", "rows": 10, "cols": 5, "visible": false }
  ]
}
```

**stdout (--plain):**
```
Q1 Plan    45 rows  8 cols
Q2 Plan    32 rows  8 cols
Summary    10 rows  5 cols  [hidden]
```

`rows` and `cols` reflect the actual data extent (`max_row`, `max_column` from openpyxl), not the full sheet dimensions.

---

### `pyxlsx sheet read <file> <sheet> [--range A1:H50]`

Reads a sheet (or a sub-range) as a raw 2D array. Rows are top-to-bottom; within each row, values are left-to-right. Cell values are returned as their Python type (int, float, str, bool, datetime as ISO-8601 string, None for empty cells).

**stdout (JSON):**
```json
{
  "sheet": "Q1 Plan",
  "range": "A1:H3",
  "rows": [
    ["Task", "Owner", "Start", "End", "Status", "Priority", "Notes", "Budget"],
    ["Design phase", "Alice", "2025-01-06", "2025-02-28", "Done", "High", "", 50000],
    ["Build phase", "Bob", "2025-03-01", null, "In Progress", "High", "Started", 120000]
  ]
}
```

**stdout (--plain):**
```
Task          Owner  Start       End         Status       Priority  Notes    Budget
Design phase  Alice  2025-01-06  2025-02-28  Done         High               50000
Build phase   Bob    2025-03-01              In Progress  High      Started  120000
```

If `--range` is omitted, the entire used range is returned.

---

### `pyxlsx table read <file> <sheet> [--range A1:H50] [--header-row 1]`

Reads a sheet (or range) as an array of objects, using the header row to key each record. This is the primary command for extracting structured data from workplans.

`--header-row N` (default: `1`) specifies which row within the data (1-based) contains column headers. The header row is consumed and does not appear in `rows`.

**stdout (JSON):**
```json
{
  "sheet": "Q1 Plan",
  "range": "A1:H45",
  "header_row": 1,
  "headers": ["Task", "Owner", "Start", "End", "Status", "Priority", "Notes", "Budget"],
  "rows": [
    { "Task": "Design phase", "Owner": "Alice", "Start": "2025-01-06", "End": "2025-02-28", "Status": "Done", "Priority": "High", "Notes": "", "Budget": 50000 },
    { "Task": "Build phase",  "Owner": "Bob",   "Start": "2025-03-01", "End": null,          "Status": "In Progress", "Priority": "High", "Notes": "Started", "Budget": 120000 }
  ]
}
```

**stdout (--plain):**
```
Task          Owner  Start       End         Status       Priority  Notes    Budget
Design phase  Alice  2025-01-06  2025-02-28  Done         High               50000
Build phase   Bob    2025-03-01              In Progress  High      Started  120000
```

If two header cells have the same text, the later column is suffixed with its column letter (e.g. `"Status"` and `"Status_C"`).

---

### `pyxlsx cell get <file> <sheet> <cell>`

Returns the value of a single cell. Cell address is in A1 notation (e.g. `B3`). Value is typed as for `sheet read`.

**stdout (JSON):**
```json
{ "sheet": "Q1 Plan", "cell": "B3", "value": "Alice" }
```

**stdout (--plain):**
```
Alice
```

---

### `pyxlsx cell set <file> <sheet> <cell> <value>`

Sets the value of a single cell. Type inference rules:

- Value starting with `=` → written as a formula string (e.g. `=SUM(A2:A10)`)
- Value parseable as integer → stored as int
- Value parseable as float → stored as float
- Otherwise → stored as string

**stdout (JSON):**
```json
{ "sheet": "Q1 Plan", "cell": "B3", "value": "Bob" }
```

**stdout (--plain):**
```
set Q1 Plan!B3 = Bob
```

---

### `pyxlsx sheet add <file> <name> [--position N]`

Adds a new blank sheet. `--position` (1-based, default: end). Errors if a sheet with that name already exists.

**stdout (JSON):**
```json
{ "name": "Q3 Plan", "position": 3 }
```

**stdout (--plain):**
```
added sheet Q3 Plan at position 3
```

---

### `pyxlsx sheet delete <file> <name>`

Deletes a sheet by name. Errors if the workbook would be left with zero sheets.

**stdout (JSON):**
```json
{ "deleted": "Q3 Plan" }
```

**stdout (--plain):**
```
deleted sheet Q3 Plan
```

---

### `pyxlsx sheet rename <file> <old-name> <new-name>`

Renames a sheet. Errors if `old-name` does not exist or `new-name` already exists.

**stdout (JSON):**
```json
{ "old_name": "Q3 Plan", "new_name": "Q3 Roadmap" }
```

**stdout (--plain):**
```
renamed Q3 Plan → Q3 Roadmap
```

---

### `pyxlsx unpack <file.xlsx> [output_dir]`

Extracts the ZIP archive to a directory (defaults to the filename without extension). Files are extracted as-is with no post-processing.

**stdout (JSON):**
```json
{ "unpacked_dir": "workplan" }
```

**stdout (--plain):**
```
workplan
```

---

### `pyxlsx pack <unpacked_dir> <output.xlsx>`

Rezips an unpacked directory into a valid `.xlsx`. Writes atomically (temp file → rename).

**stdout (JSON):**
```json
{ "output_file": "workplan.xlsx" }
```

**stdout (--plain):**
```
workplan.xlsx
```

---

## Implementation Notes

### Library strategy

| Operation | Library | Reason |
|---|---|---|
| `info`, `sheet list`, `sheet read`, `table read`, `cell get` | `openpyxl` | High-level API, typed values, named range access |
| `cell set`, `sheet add/delete/rename` | `openpyxl` | Full read/write API; no LibreOffice required |
| `unpack`, `pack` | `zipfile` + `defusedxml` | Raw ZIP control; openpyxl does not expose unpack/repack |

`openpyxl` is a required dependency. `defusedxml` is required for safe XML parsing in `unpack`. There is no optional dependency for LibreOffice — formula recalculation is out of scope.

### Key design constraints

- `openpyxl` is opened with `data_only=False` for all read operations so formula strings are preserved. If a file has been saved by Excel, openpyxl reads cached values; if saved only by openpyxl, formula cells return `None` for the value. The CLI documents this in `SKILL.md` — agents that need evaluated formula values should open the file in Excel/LibreOffice first.
- `defusedxml` is used for all XML parsing in `pack.py`. `xml.etree.ElementTree` must not be used for parsing.
- All write commands (`cell set`, `sheet add/delete/rename`) load the workbook, modify it, and save to the same path atomically (write to a temp file in the same directory, then `os.replace()`).
- Datetime values from openpyxl are serialised as ISO-8601 strings (`YYYY-MM-DD` for date-only, `YYYY-MM-DDTHH:MM:SS` for datetime) in JSON output.

### Type inference for `cell set`

```
value.startswith("=")  →  str (formula)
int(value) succeeds    →  int
float(value) succeeds  →  float
else                   →  str
```

### Row/column counting for `sheet list`

Uses `ws.max_row` and `ws.max_column` from openpyxl. These reflect the extent of cells that have been written; entirely empty sheets return `(0, 0)`.

## Package Structure

```
.apm/skills/pyxlsx/
  pyxlsx.py              # self-bootstrapping entry point
  SKILL.md               # agent documentation
  apm.yml                # skill metadata
  pyproject.toml         # hatchling build, entry point, deps
  pyxlsx/
    __init__.py          # __version__
    cli.py               # click group + all commands; output_result() helper
    ops/
      __init__.py
      inspect.py         # info(), list_sheets(), read_sheet(), read_table(), get_cell()
      write.py           # set_cell(), add_sheet(), delete_sheet(), rename_sheet()
      pack.py            # unpack(), pack() — zipfile + defusedxml
tests/
  test_xlsx_inspect.py
  test_xlsx_write.py
  test_xlsx_pack.py
  test_xlsx_cli.py
```

Tests live in the repo-root `tests/` directory alongside the pypptx tests. A `conftest.py` addition (or a separate `conftest_xlsx.py` included via `conftest.py`) provides the `minimal_xlsx` fixture.

## Files to Create

| File | Description |
|---|---|
| `.apm/skills/pyxlsx/pyxlsx.py` | Self-bootstrapping entry point; deps: `click`, `openpyxl`, `defusedxml` |
| `.apm/skills/pyxlsx/SKILL.md` | Agent documentation: quick reference, workflows, openpyxl escape hatch, output contract |
| `.apm/skills/pyxlsx/apm.yml` | Skill metadata |
| `.apm/skills/pyxlsx/pyproject.toml` | hatchling build; entry point `pyxlsx`; deps: click, openpyxl, defusedxml |
| `.apm/skills/pyxlsx/pyxlsx/__init__.py` | `__version__ = "0.1.0"` |
| `.apm/skills/pyxlsx/pyxlsx/cli.py` | Click group + all command definitions; `output_result()` helper |
| `.apm/skills/pyxlsx/pyxlsx/ops/__init__.py` | Empty |
| `.apm/skills/pyxlsx/pyxlsx/ops/inspect.py` | `info()`, `list_sheets()`, `read_sheet()`, `read_table()`, `get_cell()` |
| `.apm/skills/pyxlsx/pyxlsx/ops/write.py` | `set_cell()`, `add_sheet()`, `delete_sheet()`, `rename_sheet()` |
| `.apm/skills/pyxlsx/pyxlsx/ops/pack.py` | `unpack()`, `pack()` |
| `tests/test_xlsx_inspect.py` | Tests for all inspect operations |
| `tests/test_xlsx_write.py` | Tests for all write operations |
| `tests/test_xlsx_pack.py` | Tests for unpack/pack |
| `tests/test_xlsx_cli.py` | CLI integration tests via `click.testing.CliRunner` |

## Acceptance Criteria

- `python3 .apm/skills/pyxlsx/pyxlsx.py --version` triggers first-run bootstrap, installs deps, and prints the version.
- `pyxlsx info workplan.xlsx` outputs valid JSON with `file`, `sheets`, and `named_ranges`.
- `pyxlsx sheet list workplan.xlsx` outputs valid JSON with a `sheets` array; each entry has `name`, `rows`, `cols`, `visible`.
- `pyxlsx sheet read workplan.xlsx "Sheet1"` outputs valid JSON with `sheet`, `range`, and `rows` (2D array).
- `pyxlsx sheet read workplan.xlsx "Sheet1" --range A1:C3` returns only the specified sub-range.
- `pyxlsx table read workplan.xlsx "Sheet1" --header-row 1` outputs valid JSON with `headers` and `rows` as array-of-objects; header row is not repeated in `rows`.
- `pyxlsx cell get plan.xlsx Sheet1 B2` outputs `{"sheet": "Sheet1", "cell": "B2", "value": <typed>}`.
- `pyxlsx cell set plan.xlsx Sheet1 B2 "Alice"` writes the value and outputs confirmation JSON; re-running `cell get` returns `"Alice"`.
- `pyxlsx cell set plan.xlsx Sheet1 C2 "=SUM(A2:A10)"` writes the formula string; openpyxl reads it back as `"=SUM(A2:A10)"`.
- `pyxlsx sheet add`, `sheet delete`, `sheet rename` operate correctly and error cleanly on invalid inputs (duplicate name, last sheet deletion, missing sheet).
- `pyxlsx unpack workplan.xlsx` extracts to a directory and outputs `{"unpacked_dir": "workplan"}`.
- `pyxlsx pack workplan/ workplan_repacked.xlsx` produces a file openpyxl can open without errors.
- All commands: `--plain` flag switches stdout to human-readable text; errors always go to stderr; exit code is `1` on failure.
- `pyxlsx --help` and each subcommand `--help` print usage.
- All 4 test files pass under `uv run pytest tests/test_xlsx_*.py`.

## Test Fixtures

All tests use programmatically generated `.xlsx` files via `openpyxl` — no binary fixtures committed to the repo. A `minimal_xlsx` pytest fixture (in `tests/conftest.py` or a dedicated `conftest_xlsx.py`) creates a workbook with at least two sheets, a header row, and several data rows in `tmp_path`, returned as a `Path`. Tests that verify mutation use `copy.copy` of the fixture path or write to a separate `tmp_path` file.

## Decisions

- Formula recalculation (LibreOffice) is explicitly out of scope. Agents that need evaluated formula values are directed in `SKILL.md` to open the file in Excel or LibreOffice manually.
- Table creation and bulk writes are handled by openpyxl scripts, not the CLI. `SKILL.md` provides ready-made patterns for the common cases.
- JSON is the default output format; `--plain` is the opt-out for human use.
- The skill is fully independent of `pypptx` — no shared code, no shared config, no cross-imports.
