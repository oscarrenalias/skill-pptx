import json
import sys
from typing import Callable

import click

import pyxlsx
from pyxlsx.ops.inspect import (
    get_cell as _get_cell,
    info as _info,
    list_sheets as _list_sheets,
    read_sheet as _read_sheet,
    read_table as _read_table,
)
from pyxlsx.ops.write import (
    add_sheet as _add_sheet,
    delete_sheet as _delete_sheet,
    rename_sheet as _rename_sheet,
    set_cell as _set_cell,
)


def output_result(data: dict, plain: bool, plain_fn: Callable[[dict], str]) -> None:
    """Write data to stdout as JSON or as plain text via plain_fn."""
    if plain:
        sys.stdout.write(plain_fn(data) + "\n")
    else:
        sys.stdout.write(json.dumps(data) + "\n")


@click.group()
@click.version_option(version=pyxlsx.__version__, prog_name="pyxlsx")
@click.option("--plain", is_flag=True, default=False, help="Output plain text instead of JSON.")
@click.pass_context
def cli(ctx: click.Context, plain: bool) -> None:
    """pyxlsx — Excel manipulation toolkit."""
    ctx.ensure_object(dict)
    ctx.obj["plain"] = plain


@cli.command("info")
@click.argument("file")
@click.pass_context
def info_cmd(ctx: click.Context, file: str) -> None:
    """Show workbook-level metadata: sheet names and named ranges."""
    plain: bool = ctx.obj["plain"]

    def plain_fn(data: dict) -> str:
        lines = [data["file"]]
        lines.append("Sheets: " + ", ".join(data["sheets"]))
        nr = data["named_ranges"]
        lines.append("Named ranges: " + (", ".join(nr) if nr else "(none)"))
        return "\n".join(lines)

    try:
        result = _info(file)
    except SystemExit:
        sys.exit(1)
    output_result(result, plain, plain_fn)


@cli.group()
def sheet() -> None:
    """Commands for working with sheets."""


@sheet.command("list")
@click.argument("file")
@click.pass_context
def sheet_list_cmd(ctx: click.Context, file: str) -> None:
    """List all sheets with row/column counts and visibility."""
    plain: bool = ctx.obj["plain"]

    def plain_fn(data: dict) -> str:
        lines = []
        for s in data["sheets"]:
            line = f"{s['name']:<20} {s['rows']:>5} rows  {s['cols']:>3} cols"
            if not s["visible"]:
                line += "  [hidden]"
            lines.append(line)
        return "\n".join(lines)

    try:
        result = _list_sheets(file)
    except SystemExit:
        sys.exit(1)
    output_result(result, plain, plain_fn)


@sheet.command("read")
@click.argument("file")
@click.argument("sheet")
@click.option("--range", "range_str", default=None, help="Cell range in A1:H50 notation.")
@click.pass_context
def sheet_read_cmd(ctx: click.Context, file: str, sheet: str, range_str: str | None) -> None:
    """Read a sheet as a 2D array of typed cell values."""
    plain: bool = ctx.obj["plain"]

    def plain_fn(data: dict) -> str:
        rows = data.get("rows", [])
        if not rows:
            return ""
        col_count = max(len(row) for row in rows)
        # Convert all values to strings for display
        str_rows = []
        for row in rows:
            str_row = [str(v) if v is not None else "" for v in row]
            # Pad short rows to col_count
            while len(str_row) < col_count:
                str_row.append("")
            str_rows.append(str_row)
        # Compute per-column widths
        widths = [0] * col_count
        for str_row in str_rows:
            for i, v in enumerate(str_row):
                widths[i] = max(widths[i], len(v))
        lines = []
        for str_row in str_rows:
            padded = [str_row[i].ljust(widths[i]) for i in range(col_count)]
            lines.append("  ".join(padded).rstrip())
        return "\n".join(lines)

    try:
        result = _read_sheet(file, sheet, range_str)
    except SystemExit:
        sys.exit(1)
    output_result(result, plain, plain_fn)


@sheet.command("add")
@click.argument("file")
@click.argument("name")
@click.option(
    "--position",
    "position",
    default=None,
    type=int,
    help="1-based position to insert the new sheet (default: end).",
)
@click.pass_context
def sheet_add_cmd(ctx: click.Context, file: str, name: str, position: int | None) -> None:
    """Add a new blank sheet at the given position (default: end)."""
    plain: bool = ctx.obj["plain"]

    def plain_fn(data: dict) -> str:
        return f"added sheet {data['name']} at position {data['position']}"

    try:
        result = _add_sheet(file, name, position)
    except SystemExit:
        sys.exit(1)
    output_result(result, plain, plain_fn)


@sheet.command("delete")
@click.argument("file")
@click.argument("name")
@click.pass_context
def sheet_delete_cmd(ctx: click.Context, file: str, name: str) -> None:
    """Delete a sheet by name."""
    plain: bool = ctx.obj["plain"]

    def plain_fn(data: dict) -> str:
        return f"deleted sheet {data['deleted']}"

    try:
        result = _delete_sheet(file, name)
    except SystemExit:
        sys.exit(1)
    output_result(result, plain, plain_fn)


@sheet.command("rename")
@click.argument("file")
@click.argument("old_name")
@click.argument("new_name")
@click.pass_context
def sheet_rename_cmd(ctx: click.Context, file: str, old_name: str, new_name: str) -> None:
    """Rename a sheet from OLD_NAME to NEW_NAME."""
    plain: bool = ctx.obj["plain"]

    def plain_fn(data: dict) -> str:
        return f"renamed {data['old_name']} \u2192 {data['new_name']}"

    try:
        result = _rename_sheet(file, old_name, new_name)
    except SystemExit:
        sys.exit(1)
    output_result(result, plain, plain_fn)


@cli.group()
def table() -> None:
    """Commands for working with tables."""


@table.command("read")
@click.argument("file")
@click.argument("sheet")
@click.option(
    "--header-row",
    "header_row",
    default=1,
    show_default=True,
    type=int,
    help="1-based row number to use as the header.",
)
@click.option("--range", "range_str", default=None, help="Cell range in A1:H50 notation.")
@click.pass_context
def table_read_cmd(
    ctx: click.Context, file: str, sheet: str, header_row: int, range_str: str | None
) -> None:
    """Read a sheet as an array-of-objects keyed by the header row."""
    plain: bool = ctx.obj["plain"]

    def plain_fn(data: dict) -> str:
        headers = data.get("headers", [])
        rows = data.get("rows", [])
        if not rows and not headers:
            return ""
        # Compute column widths from headers and values
        widths = {h: len(h) for h in headers}
        for row in rows:
            for h in headers:
                v = row.get(h)
                widths[h] = max(widths[h], len(str(v) if v is not None else ""))
        # Header line
        header_line = "  ".join(h.ljust(widths[h]) for h in headers).rstrip()
        sep_line = "  ".join("-" * widths[h] for h in headers).rstrip()
        lines = [header_line, sep_line]
        for row in rows:
            parts = []
            for h in headers:
                v = row.get(h)
                parts.append((str(v) if v is not None else "").ljust(widths[h]))
            lines.append("  ".join(parts).rstrip())
        return "\n".join(lines)

    try:
        result = _read_table(file, sheet, header_row, range_str)
    except SystemExit:
        sys.exit(1)
    output_result(result, plain, plain_fn)


@cli.group()
def cell() -> None:
    """Commands for working with cells."""


@cell.command("get")
@click.argument("file")
@click.argument("sheet")
@click.argument("cell")
@click.pass_context
def cell_get_cmd(ctx: click.Context, file: str, sheet: str, cell: str) -> None:
    """Return the typed value of a single A1-addressed cell."""
    plain: bool = ctx.obj["plain"]

    def plain_fn(data: dict) -> str:
        v = data.get("value")
        return str(v) if v is not None else ""

    try:
        result = _get_cell(file, sheet, cell)
    except SystemExit:
        sys.exit(1)
    output_result(result, plain, plain_fn)


@cell.command("set")
@click.argument("file")
@click.argument("sheet")
@click.argument("cell")
@click.argument("value")
@click.pass_context
def cell_set_cmd(ctx: click.Context, file: str, sheet: str, cell: str, value: str) -> None:
    """Set the value of a single A1-addressed cell (type-inferred)."""
    plain: bool = ctx.obj["plain"]

    def plain_fn(data: dict) -> str:
        return f"set {data['sheet']}!{data['cell']} = {data['value']}"

    try:
        result = _set_cell(file, sheet, cell, value)
    except SystemExit:
        sys.exit(1)
    output_result(result, plain, plain_fn)
