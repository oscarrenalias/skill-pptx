"""Pack operations: unpack and pack .xlsx ZIP archives."""

from __future__ import annotations

import os
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict


def unpack(file: str | Path, output_dir: str | Path | None = None) -> Dict[str, Any]:
    """Extract the .xlsx ZIP archive to a directory.

    Args:
        file: Path to the .xlsx file.
        output_dir: Directory to extract into. Defaults to the filename stem
                    (e.g. ``workplan.xlsx`` → ``workplan/``).

    Returns:
        dict with key:
            unpacked_dir (str): Path of the directory the archive was extracted to.
    """
    path = Path(file)
    if not path.exists():
        print(f"error: file not found: {file}", file=sys.stderr)
        sys.exit(1)

    if output_dir is None:
        dest = path.parent / path.stem
    else:
        dest = Path(output_dir)

    try:
        with zipfile.ZipFile(str(path), "r") as zf:
            zf.extractall(str(dest))
    except zipfile.BadZipFile as exc:
        print(f"error: not a valid xlsx/zip file: {exc}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        sys.exit(1)

    return {"unpacked_dir": str(dest)}


def pack(source_dir: str | Path, output_file: str | Path) -> Dict[str, Any]:
    """Rezip an unpacked directory into a valid .xlsx file.

    Writes atomically: builds the archive into a temporary file in the same
    directory as ``output_file``, then renames it into place via ``os.replace()``.

    Args:
        source_dir: Directory produced by :func:`unpack`.
        output_file: Destination .xlsx path.

    Returns:
        dict with key:
            output_file (str): Path of the produced .xlsx file.
    """
    src = Path(source_dir)
    if not src.is_dir():
        print(f"error: not a directory: {source_dir}", file=sys.stderr)
        sys.exit(1)

    out = Path(output_file)
    out_dir = out.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        fd, tmp_path = tempfile.mkstemp(dir=str(out_dir), suffix=".xlsx.tmp")
        os.close(fd)
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for entry in sorted(src.rglob("*")):
                if entry.is_file():
                    arcname = entry.relative_to(src)
                    zf.write(str(entry), str(arcname))
        os.replace(tmp_path, str(out))
    except Exception as exc:
        # Clean up the temp file if something went wrong before the rename
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        print(f"error: {exc}", file=sys.stderr)
        sys.exit(1)

    return {"output_file": str(out)}
