# pyxlsx — Excel Manipulation Skill

`pyxlsx` is an agent-friendly CLI for reading and writing `.xlsx` files. It covers workbook inspection, sheet management, table and cell reads, single-cell writes, and raw ZIP pack/unpack.

## Running the CLI

```bash
# Standalone entry point — bootstraps its own .venv on first run
python3 .claude/skills/pyxlsx/pyxlsx.py --help
python3 .claude/skills/pyxlsx/pyxlsx.py sheet list workplan.xlsx
```

---

Please make sure to run the script from the right folder, and manage file paths appropriately when running.

## Output Contract

- Every command writes a **single JSON object** to **stdout** by default.
- Pass `--plain` to any command to switch stdout to human-readable text.
- All errors go to **stderr**; stdout is always machine-parseable on success.
- Exit code `0` on success, `1` on any error.

```bash
pyxlsx info workplan.xlsx               # JSON to stdout
pyxlsx --plain info workplan.xlsx       # plain text to stdout
pyxlsx sheet list missing.xlsx 2>err.txt  # error message in err.txt, exit 1
```

---

## Command Reference

### `pyxlsx info <file>`

Workbook-level metadata: file path, sheet names, and named ranges.

```bash
pyxlsx info workplan.xlsx
```

```json
{
  "file": "workplan.xlsx",
  "sheets": ["Q1 Plan", "Q2 Plan", "Summary"],
  "named_ranges": ["Budget", "Timeline"]
}
```

`--plain`:
```
workplan.xlsx
Sheets: Q1 Plan, Q2 Plan, Summary
Named ranges: Budget, Timeline
```

---

### `pyxlsx sheet list <file>`

All sheets with row/column counts and visibility.

```bash
pyxlsx sheet list workplan.xlsx
```

```json
{
  "sheets": [
    { "name": "Q1 Plan", "rows": 45, "cols": 8, "visible": true },
    { "name": "Q2 Plan", "rows": 32, "cols": 8, "visible": true },
    { "name": "Summary", "rows": 10, "cols": 5, "visible": false }
  ]
}
```

`rows` and `cols` reflect the data extent (`max_row`, `max_column`). Entirely empty sheets return `0` for both.

---

### `pyxlsx sheet read <file> <sheet> [--range A1:H50]`

Sheet data as a raw 2D array. Cell values are typed (int, float, str, bool, ISO-8601 string for dates/datetimes, null for empty cells).

```bash
pyxlsx sheet read workplan.xlsx "Q1 Plan"
pyxlsx sheet read workplan.xlsx "Q1 Plan" --range A1:H3
```

```json
{
  "sheet": "Q1 Plan",
  "range": "A1:H3",
  "rows": [
    ["Task", "Owner", "Start", "End", "Status", "Priority", "Notes", "Budget"],
    ["Design phase", "Alice", "2025-01-06", "2025-02-28", "Done", "High", null, 50000],
    ["Build phase", "Bob", "2025-03-01", null, "In Progress", "High", "Started", 120000]
  ]
}
```

If `--range` is omitted, the entire used range is returned.

---

### `pyxlsx table read <file> <sheet> [--range A1:H50] [--header-row 1]`

Sheet data as an array-of-objects keyed by the header row. The primary command for extracting structured data from workplans.

```bash
pyxlsx table read workplan.xlsx "Q1 Plan"
pyxlsx table read workplan.xlsx "Q1 Plan" --range A1:H45 --header-row 1
```

```json
{
  "sheet": "Q1 Plan",
  "range": "A1:H45",
  "header_row": 1,
  "headers": ["Task", "Owner", "Start", "End", "Status", "Priority", "Notes", "Budget"],
  "rows": [
    { "Task": "Design phase", "Owner": "Alice", "Start": "2025-01-06", "End": "2025-02-28", "Status": "Done", "Priority": "High", "Notes": null, "Budget": 50000 },
    { "Task": "Build phase",  "Owner": "Bob",   "Start": "2025-03-01", "End": null,          "Status": "In Progress", "Priority": "High", "Notes": "Started", "Budget": 120000 }
  ]
}
```

If two header cells have the same text, the later column is suffixed with its column letter (e.g. `"Status_C"`). The header row itself does not appear in `rows`.

---

### `pyxlsx cell get <file> <sheet> <cell>`

Value of a single A1-addressed cell, typed as for `sheet read`.

```bash
pyxlsx cell get workplan.xlsx "Q1 Plan" B2
```

```json
{ "sheet": "Q1 Plan", "cell": "B2", "value": "Alice" }
```

---

### `pyxlsx cell set <file> <sheet> <cell> <value>`

Set a single cell value. Type is inferred from the string:

| Input string | Stored as |
|---|---|
| Starts with `=` | formula string (e.g. `=SUM(A2:A10)`) |
| Parseable as integer | int |
| Parseable as float | float |
| Otherwise | str |

```bash
pyxlsx cell set workplan.xlsx "Q1 Plan" B2 "Bob"
pyxlsx cell set workplan.xlsx "Q1 Plan" C2 42
pyxlsx cell set workplan.xlsx "Q1 Plan" D2 "=SUM(A2:A10)"
```

```json
{ "sheet": "Q1 Plan", "cell": "B2", "value": "Bob" }
```

Writes atomically: saves to a temp file in the same directory, then renames it into place.

---

### `pyxlsx sheet add <file> <name> [--position N]`

Add a new blank sheet. `--position` is 1-based; omit to append at the end. Errors if a sheet with that name already exists.

```bash
pyxlsx sheet add workplan.xlsx "Q3 Plan"
pyxlsx sheet add workplan.xlsx "Q3 Plan" --position 3
```

```json
{ "name": "Q3 Plan", "position": 3 }
```

---

### `pyxlsx sheet delete <file> <name>`

Delete a sheet by name. Errors if the sheet does not exist or if deleting it would leave the workbook with zero sheets.

```bash
pyxlsx sheet delete workplan.xlsx "Q3 Plan"
```

```json
{ "deleted": "Q3 Plan" }
```

---

### `pyxlsx sheet rename <file> <old-name> <new-name>`

Rename a sheet. Errors if `old-name` does not exist or `new-name` already exists.

```bash
pyxlsx sheet rename workplan.xlsx "Q3 Plan" "Q3 Roadmap"
```

```json
{ "old_name": "Q3 Plan", "new_name": "Q3 Roadmap" }
```

---

### `pyxlsx unpack <file.xlsx> [output_dir]`

Extract the `.xlsx` ZIP archive to a directory. Defaults to the filename stem (`workplan.xlsx` → `workplan/`).

```bash
pyxlsx unpack workplan.xlsx
pyxlsx unpack workplan.xlsx extracted/
```

```json
{ "unpacked_dir": "workplan" }
```

Useful for inspecting or directly editing the underlying XML before repacking.

---

### `pyxlsx pack <source_dir> <output.xlsx>`

Rezip an unpacked directory into a valid `.xlsx`. Writes atomically (temp file → rename).

```bash
pyxlsx pack workplan/ workplan_modified.xlsx
```

```json
{ "output_file": "workplan_modified.xlsx" }
```

---

## Formula Evaluation Caveat

`pyxlsx` reads all files with `data_only=True`, which means:

- **Excel-saved files**: formula cells return the **cached calculated value** stored by Excel on the last save.
- **openpyxl-only files**: formula cells return `null` because no cached value exists.

The CLI does **not** evaluate or recalculate formulas. If you need evaluated formula results from a file that was only written by openpyxl (or by `pyxlsx cell set`), **open the file in Excel or LibreOffice Calc first and save it**, then re-run the read command.

```bash
# After editing formulas with cell set, open and save in Excel/LibreOffice,
# then read back the recalculated values:
pyxlsx cell get workplan.xlsx "Summary" D10
```

---

## openpyxl Escape Hatch

For table creation, bulk writes, and operations not covered by the CLI, write a self-contained Python script using `openpyxl` and run it inside the skill's `.venv`.

### Create a table from scratch

```python
#!/usr/bin/env python3
"""Create a task table in a new workbook."""
import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Tasks"

headers = ["Task", "Owner", "Start", "End", "Status"]
ws.append(headers)

rows = [
    ["Design phase", "Alice", "2025-01-06", "2025-02-28", "Done"],
    ["Build phase",  "Bob",   "2025-03-01", "2025-06-30", "In Progress"],
    ["Testing",      "Carol", "2025-07-01", None,         "Not Started"],
]
for row in rows:
    ws.append(row)

wb.save("tasks.xlsx")
print("saved tasks.xlsx")
```

Run with the skill's venv:
```bash
python3 .apm/skills/pyxlsx/.venv/bin/python3 create_table.py
# or, after activating:
source .apm/skills/pyxlsx/.venv/bin/activate
python3 create_table.py
```

### Bulk-write cells from a list of dicts

```python
#!/usr/bin/env python3
"""Write rows from a dict list into an existing sheet, starting at a given row."""
import openpyxl

FILE = "workplan.xlsx"
SHEET = "Q1 Plan"
START_ROW = 2  # row 1 is the header

# Data from a previous pyxlsx table read, processed by the agent
updates = [
    {"B": "Alice", "E": "Done"},
    {"B": "Bob",   "E": "In Progress"},
]

wb = openpyxl.load_workbook(FILE)
ws = wb[SHEET]

for i, row_updates in enumerate(updates):
    for col_letter, value in row_updates.items():
        ws[f"{col_letter}{START_ROW + i}"] = value

wb.save(FILE)
print(f"wrote {len(updates)} rows to {SHEET}")
```

---

## Agent Workflow: Extract from Excel, Write to PowerPoint

This example shows a complete extract-process-write pipeline using both `pyxlsx` and `pypptx`.

### 1. Read the task table from Excel

```bash
pyxlsx table read workplan.xlsx "Q1 Plan" --range A1:E20 > tasks.json
```

`tasks.json` contains:
```json
{
  "sheet": "Q1 Plan",
  "range": "A1:E20",
  "header_row": 1,
  "headers": ["Task", "Owner", "Start", "End", "Status"],
  "rows": [
    { "Task": "Design phase", "Owner": "Alice", "Start": "2025-01-06", "End": "2025-02-28", "Status": "Done" },
    { "Task": "Build phase",  "Owner": "Bob",   "Start": "2025-03-01", "End": "2025-06-30", "Status": "In Progress" }
  ]
}
```

### 2. Process the data and write a PowerPoint slide

Write a short `python-pptx` script that reads `tasks.json` and creates a slide:

```python
#!/usr/bin/env python3
"""Build a status slide from tasks.json."""
import json
from pptx import Presentation
from pptx.util import Inches, Pt

with open("tasks.json") as f:
    data = json.load(f)

prs = Presentation()
slide_layout = prs.slide_layouts[1]  # Title and Content
slide = prs.slides.add_slide(slide_layout)

slide.shapes.title.text = "Q1 Plan — Status Update"

tf = slide.placeholders[1].text_frame
tf.text = ""
for row in data["rows"]:
    p = tf.add_paragraph()
    p.text = f"{row['Task']}  ({row['Owner']})  {row['Status']}"
    p.level = 0

prs.save("roadmap.pptx")
print("saved roadmap.pptx")
```

### 3. Verify the result with pypptx

```bash
pypptx slide list roadmap.pptx
pypptx extract-text roadmap.pptx
```

### One-liner pipeline (shell)

```bash
pyxlsx table read workplan.xlsx "Q1 Plan" | python3 make_slide.py && pypptx slide list roadmap.pptx
```
