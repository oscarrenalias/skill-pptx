"""
pack.py — unpack and pack operations for .pptx (OPC/ZIP) files.

Uses zipfile for archive operations and defusedxml for any XML parsing.
xml.etree.ElementTree is intentionally not imported (security constraint).
"""

import os
import zipfile
from pathlib import Path

CONTENT_TYPES_ENTRY = "[Content_Types].xml"


def unpack(src: Path, dest: Path) -> Path:
    """Extract all ZIP entries from *src* into *dest* unchanged.

    Parameters
    ----------
    src:  Path to a .pptx (ZIP) file.
    dest: Destination directory; will be created if it does not exist.

    Returns
    -------
    Path to the destination directory.

    Raises
    ------
    ValueError: If *src* does not exist, is not a file, or is not a valid ZIP.
    """
    src = Path(src)
    dest = Path(dest)

    if not src.exists():
        raise ValueError(f"Source file does not exist: {src}")
    if not src.is_file():
        raise ValueError(f"Source path is not a file: {src}")
    if not zipfile.is_zipfile(src):
        raise ValueError(f"Source file is not a valid ZIP/PPTX archive: {src}")

    dest.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(src, "r") as zf:
        zf.extractall(dest)

    return dest


def pack(src_dir: Path, dest: Path) -> Path:
    """Repack a directory of OPC/PPTX parts into a ZIP file at *dest*.

    ``[Content_Types].xml`` is written as the first ZIP entry per the OPC spec.
    Writes atomically: builds a ``.tmp`` file beside *dest*, then renames it.

    Parameters
    ----------
    src_dir: Directory containing the unpacked PPTX parts.
    dest:    Output path for the resulting .pptx file.

    Returns
    -------
    Path to the written output file (*dest*).

    Raises
    ------
    ValueError: If *src_dir* does not exist, is not a directory, or
                ``[Content_Types].xml`` is missing from *src_dir*.
    """
    src_dir = Path(src_dir)
    dest = Path(dest)

    if not src_dir.exists():
        raise ValueError(f"Source directory does not exist: {src_dir}")
    if not src_dir.is_dir():
        raise ValueError(f"Source path is not a directory: {src_dir}")

    content_types_path = src_dir / CONTENT_TYPES_ENTRY
    if not content_types_path.is_file():
        raise ValueError(
            f"[Content_Types].xml not found in source directory: {src_dir}"
        )

    # Collect all files, excluding [Content_Types].xml (added first separately).
    all_files: list[Path] = []
    for root, _, files in os.walk(src_dir):
        for name in files:
            file_path = Path(root) / name
            archive_name = file_path.relative_to(src_dir).as_posix()
            if archive_name != CONTENT_TYPES_ENTRY:
                all_files.append(file_path)

    tmp_dest = dest.with_suffix(dest.suffix + ".tmp")

    try:
        with zipfile.ZipFile(tmp_dest, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            # OPC spec: [Content_Types].xml MUST be the first entry.
            zf.write(content_types_path, CONTENT_TYPES_ENTRY)

            for file_path in sorted(all_files):
                archive_name = file_path.relative_to(src_dir).as_posix()
                zf.write(file_path, archive_name)

        tmp_dest.replace(dest)
    except Exception:
        # Clean up the temp file on failure.
        if tmp_dest.exists():
            tmp_dest.unlink()
        raise

    return dest
