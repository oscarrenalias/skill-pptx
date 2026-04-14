"""
thumbnails.py — dependency checking and conversion pipeline for thumbnail generation.

The thumbnail workflow requires three external dependencies:
  - soffice (LibreOffice) for PPTX → PDF conversion
  - pdftoppm (poppler-utils) for PDF → image rasterisation
  - Pillow for image post-processing

Call check_dependencies() early in any entry point that uses this module.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from PIL import Image as PILImage


def check_dependencies() -> None:
    """Verify that soffice, pdftoppm, and Pillow are available.

    Prints a clear, actionable install hint to stderr for each missing
    dependency, then raises SystemExit(1) if any are absent.
    """
    missing = False

    if shutil.which("soffice") is None:
        print(
            "Error: 'soffice' (LibreOffice) not found in PATH.\n"
            "  Install on macOS:  brew install --cask libreoffice\n"
            "  Install on Debian: sudo apt-get install libreoffice",
            file=sys.stderr,
        )
        missing = True

    if shutil.which("pdftoppm") is None:
        print(
            "Error: 'pdftoppm' (poppler-utils) not found in PATH.\n"
            "  Install on macOS:  brew install poppler\n"
            "  Install on Debian: sudo apt-get install poppler-utils",
            file=sys.stderr,
        )
        missing = True

    try:
        import PIL  # noqa: F401
    except ImportError:
        print(
            "Error: Pillow is not installed.\n"
            "  pip install 'pypptx[thumbnails]'",
            file=sys.stderr,
        )
        missing = True

    if missing:
        raise SystemExit(1)


def pptx_to_jpegs(pptx_path: Path | str, temp_dir: Path | str) -> list[Path]:
    """Convert a .pptx file to a list of per-page JPEG images.

    Runs a two-step subprocess pipeline:
      1. ``soffice --headless --convert-to pdf``  →  intermediate PDF in *temp_dir*
      2. ``pdftoppm -jpeg -r 150``                →  per-page JPEGs in *temp_dir*

    Args:
        pptx_path: Path to the source .pptx file.
        temp_dir:  Directory for intermediate and output files (managed by caller).

    Returns:
        An ordered list of :class:`~pathlib.Path` objects, one JPEG per slide,
        sorted by page number.

    Raises:
        RuntimeError: If either subprocess exits with a non-zero return code.
                      The message includes the captured stderr output.
    """
    pptx_path = Path(pptx_path)
    temp_dir = Path(temp_dir)

    # Step 1: .pptx → PDF via LibreOffice
    soffice_result = subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(temp_dir),
            str(pptx_path),
        ],
        capture_output=True,
        text=True,
    )
    if soffice_result.returncode != 0:
        raise RuntimeError(
            f"soffice failed (exit {soffice_result.returncode}):\n{soffice_result.stderr}"
        )

    pdf_path = temp_dir / (pptx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError(
            f"soffice did not produce expected PDF at {pdf_path}"
        )

    # Step 2: PDF → per-page JPEGs via pdftoppm
    jpeg_prefix = str(temp_dir / pptx_path.stem)
    pdftoppm_result = subprocess.run(
        [
            "pdftoppm",
            "-jpeg",
            "-r", "150",
            str(pdf_path),
            jpeg_prefix,
        ],
        capture_output=True,
        text=True,
    )
    if pdftoppm_result.returncode != 0:
        raise RuntimeError(
            f"pdftoppm failed (exit {pdftoppm_result.returncode}):\n{pdftoppm_result.stderr}"
        )

    # Collect output JPEGs; pdftoppm zero-pads page numbers so lexicographic
    # sort matches page order.
    jpegs = sorted(temp_dir.glob(f"{pptx_path.stem}-*.jpg"))
    return jpegs


def _make_hatched_placeholder(width: int, height: int) -> "PILImage.Image":
    """Generate a hatched grey placeholder image for a hidden slide.

    Args:
        width:  Image width in pixels.
        height: Image height in pixels.

    Returns:
        A PIL Image filled with a light grey background and diagonal hatch lines.
    """
    from PIL import Image, ImageDraw

    img = Image.new("RGB", (width, height), color=(200, 200, 200))
    draw = ImageDraw.Draw(img)

    spacing = 20
    line_color = (160, 160, 160)
    # Diagonal lines from top-left toward bottom-right, covering the full image.
    for offset in range(-height, width + height, spacing):
        draw.line([(offset, 0), (offset + height, height)], fill=line_color, width=1)

    return img


def assemble_grid(images: "list[PILImage.Image]", cols: int) -> "PILImage.Image":
    """Arrange a list of PIL Images into a labelled grid.

    Images are placed in row-major order (left-to-right, top-to-bottom).
    Each cell is annotated with its 1-based slide number in the bottom-left
    corner using a white label with a dark drop-shadow for readability.

    Args:
        images: Ordered list of PIL Images, one per slide.
        cols:   Number of columns in the grid (must be >= 1).

    Returns:
        A single PIL Image containing all thumbnails arranged as a grid.
    """
    import math
    from PIL import Image, ImageDraw, ImageFont

    if not images:
        return Image.new("RGB", (1, 1), color=(255, 255, 255))

    # Use first image's dimensions as the canonical cell size.
    cell_w, cell_h = images[0].size

    rows = math.ceil(len(images) / cols)
    grid = Image.new("RGB", (cols * cell_w, rows * cell_h), color=(255, 255, 255))
    draw = ImageDraw.Draw(grid)

    # Load a font scaled to ~5 % of cell height; fall back to the built-in default.
    font_size = max(12, cell_h // 20)
    try:
        # Pillow >= 10.0 supports a size argument on load_default.
        font = ImageFont.load_default(size=font_size)
    except TypeError:
        font = ImageFont.load_default()

    for idx, img in enumerate(images):
        row, col = divmod(idx, cols)
        x = col * cell_w
        y = row * cell_h

        grid.paste(img, (x, y))

        label = str(idx + 1)
        text_x = x + 4
        text_y = y + cell_h - font_size - 6
        # Dark drop-shadow for contrast against any background colour.
        draw.text((text_x + 1, text_y + 1), label, fill=(0, 0, 0), font=font)
        draw.text((text_x, text_y), label, fill=(255, 255, 255), font=font)

    return grid


def generate_thumbnails(
    pptx_path: Path | str,
    temp_dir: Path | str,
) -> "list[PILImage.Image]":
    """Generate thumbnail images for every slide in a PPTX file.

    Hidden slides (``slide.show is False``) are represented by a hatched grey
    placeholder image at the same pixel dimensions as the rendered thumbnails.
    The returned list always has the same length as the total slide count,
    preserving index alignment.

    LibreOffice may include or exclude hidden slides when producing the
    intermediate PDF.  This function detects both cases by comparing the JPEG
    count against the total and visible slide counts and routes the mapping
    accordingly.

    Args:
        pptx_path: Path to the source .pptx file.
        temp_dir:  Directory for intermediate files (managed by caller).

    Returns:
        An ordered list of :class:`~PIL.Image.Image` objects, one per slide.
        Hidden slide positions contain a hatched grey placeholder; visible
        slide positions contain the rendered thumbnail.

    Raises:
        RuntimeError: If the conversion pipeline fails or the JPEG count is
                      inconsistent with the slide metadata.
    """
    from PIL import Image
    from pptx import Presentation

    pptx_path = Path(pptx_path)

    # Determine hidden status for each slide.
    # slide.show is None → visible (attribute not explicitly set); False → hidden.
    prs = Presentation(pptx_path)
    hidden_flags: list[bool] = [slide.show is False for slide in prs.slides]
    total_count = len(hidden_flags)
    visible_count = sum(1 for h in hidden_flags if not h)

    jpeg_paths = pptx_to_jpegs(pptx_path, temp_dir)
    jpeg_count = len(jpeg_paths)

    # Determine reference dimensions from the first available JPEG.
    if jpeg_paths:
        with Image.open(jpeg_paths[0]) as ref:
            ref_width, ref_height = ref.size
    else:
        ref_width, ref_height = 960, 540  # sensible fallback

    thumbnails: list[PILImage.Image] = []

    if jpeg_count == total_count:
        # LibreOffice rendered all slides (including hidden ones).
        for idx, is_hidden in enumerate(hidden_flags):
            if is_hidden:
                thumbnails.append(_make_hatched_placeholder(ref_width, ref_height))
            else:
                thumbnails.append(Image.open(jpeg_paths[idx]))

    elif jpeg_count == visible_count:
        # LibreOffice skipped hidden slides; map JPEGs to visible positions only.
        jpeg_iter = iter(jpeg_paths)
        for is_hidden in hidden_flags:
            if is_hidden:
                thumbnails.append(_make_hatched_placeholder(ref_width, ref_height))
            else:
                thumbnails.append(Image.open(next(jpeg_iter)))

    else:
        raise RuntimeError(
            f"Unexpected JPEG count {jpeg_count} for '{pptx_path.name}': "
            f"expected {total_count} (all slides) or {visible_count} (visible slides only)."
        )

    return thumbnails
