"""Read operations for slides: list_slides and list_layouts."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

from pptx import Presentation


def _open_presentation(path: Path) -> Presentation:
    """Open a Presentation from a .pptx file or an unpacked directory."""
    if path.is_dir():
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            # OPC spec: [Content_Types].xml must be the first ZIP entry.
            ct = path / "[Content_Types].xml"
            if ct.exists():
                zf.write(ct, "[Content_Types].xml")
            for f in sorted(path.rglob("*")):
                if not f.is_file():
                    continue
                arcname = str(f.relative_to(path))
                if arcname == "[Content_Types].xml":
                    continue  # already written first
                zf.write(f, arcname)
        buf.seek(0)
        return Presentation(buf)
    return Presentation(path)


def list_slides(path: Path) -> list[dict]:
    """Return slides in presentation order.

    Each dict contains:
        index  -- 1-based position in the presentation
        file   -- bare filename, e.g. "slide1.xml"
        hidden -- True only when the slide's show attribute is explicitly 'false' or '0'
    """
    prs = _open_presentation(Path(path))
    result = []
    for i, slide in enumerate(prs.slides, start=1):
        filename = slide.part.partname.rsplit("/", 1)[-1]
        show_attr = slide._element.get("show")
        hidden = show_attr in ("0", "false") if show_attr is not None else False
        result.append({"index": i, "file": filename, "hidden": hidden})
    return result


def list_layouts(path: Path) -> list[dict]:
    """Return all slide layouts from ppt/slideLayouts/ in filename order.

    Each dict contains:
        index -- 1-based, assigned after sorting by filename
        file  -- bare filename, e.g. "slideLayout1.xml"
    """
    prs = _open_presentation(Path(path))
    seen: set[str] = set()
    filenames: list[str] = []
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            filename = layout.part.partname.rsplit("/", 1)[-1]
            if filename not in seen:
                seen.add(filename)
                filenames.append(filename)
    filenames.sort()
    return [{"index": i, "file": f} for i, f in enumerate(filenames, start=1)]
