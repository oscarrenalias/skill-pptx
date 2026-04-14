"""Verification checks for .pptx files."""

from __future__ import annotations

import math
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

# ── Namespaces ────────────────────────────────────────────────────────────────

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

# ── Check 1: unfilled placeholder hint text ───────────────────────────────────

_PLACEHOLDER_PATTERNS: tuple[str, ...] = (
    "click to add title",
    "click to add text",
    "click to add subtitle",
    "click to edit master title style",
    "click to edit master text styles",
    "add title",
    "add text",
    "enter text here",
    "place subtitle here",
)


def _check_unfilled_placeholders(
    slide_index: int,
    slide,
    errors: list[str],
) -> None:
    """Check 1: detect shapes whose text matches an unfilled placeholder pattern."""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text
        if not text.strip():
            continue
        if text.strip().lower() in _PLACEHOLDER_PATTERNS:
            errors.append(
                f"Slide {slide_index}: '{shape.name}' has unfilled placeholder"
                f' text \u2014 "{text.strip()}"'
            )


# ── Check 2: font size below minimum ─────────────────────────────────────────

_MIN_FONT_UNITS = 1200  # 12pt in hundredths-of-a-point


def _check_font_sizes(
    slide_index: int,
    slide,
    slide_height: int,
    errors: list[str],
) -> None:
    """Check 2: report fonts below 12pt; skip footer-region shapes."""
    footer_threshold = int(slide_height * 0.9)

    for shape in slide.shapes:
        # Skip shapes in the footer region (top > 90% of slide height)
        top = shape.top
        if top is not None and top > footer_threshold:
            continue

        # Walk a:rPr[@sz] elements in this shape's XML
        sp_elem = shape._element
        for rpr in sp_elem.iter(f"{{{NS_A}}}rPr"):
            sz = rpr.get("sz")
            if sz is None:
                continue
            try:
                sz_int = int(sz)
            except ValueError:
                continue
            if sz_int < _MIN_FONT_UNITS:
                # Convert from hundredths-of-a-point to pt for the message
                pt = sz_int / 100.0
                errors.append(
                    f"Slide {slide_index}: '{shape.name}' font {pt:.1f}pt"
                    f" is below minimum (12pt)"
                )


# ── Check 3: shape overflow ───────────────────────────────────────────────────

_OVERFLOW_TOLERANCE = 50_000  # EMU — ignore sub-pixel rounding
_EMU_PER_INCH = 914_400


def _check_shape_overflow(
    slide_index: int,
    slide,
    slide_width: int,
    slide_height: int,
    errors: list[str],
) -> None:
    """Check 3: detect shapes whose bounding box overflows slide dimensions."""
    footer_threshold = int(slide_height * 0.9)

    for shape in slide.shapes:
        # Skip table shapes
        if shape.has_table:
            continue

        left = shape.left
        top = shape.top
        width = shape.width
        height = shape.height

        if left is None or top is None or width is None or height is None:
            continue

        # Skip footer-region shapes (top > 90% of slide height)
        if top > footer_threshold:
            continue

        left = int(left)
        top = int(top)
        width = int(width)
        height = int(height)

        # Negative position → overflows that edge
        if left < 0:
            errors.append(
                f"Slide {slide_index}: '{shape.name}' has negative left"
                f" position ({left} EMU) — shape overflows left edge"
            )
        if top < 0:
            errors.append(
                f"Slide {slide_index}: '{shape.name}' has negative top"
                f" position ({top} EMU) — shape overflows top edge"
            )

        # Right overflow: allow up to _OVERFLOW_TOLERANCE past slide boundary
        if left + width > slide_width + _OVERFLOW_TOLERANCE:
            errors.append(
                f"Slide {slide_index}: '{shape.name}' overflows right edge"
                f" (left+width={left + width}, slide_width={slide_width})"
            )

        # Bottom overflow
        if top + height > slide_height + _OVERFLOW_TOLERANCE:
            errors.append(
                f"Slide {slide_index}: '{shape.name}' overflows bottom edge"
                f" (top+height={top + height}, slide_height={slide_height})"
            )


# ── Check 4: text clipping ────────────────────────────────────────────────────

_DEFAULT_FONT_EMU = 18 * 12700  # 18pt in EMU (1pt = 12700 EMU)
_LINE_SPACING = 1.35
_AVG_CHAR_WIDTH_FACTOR = 0.55


def _para_font_size_emu(para) -> int:
    """Return effective font size for a paragraph in EMU; falls back to 18pt."""
    for run in para.runs:
        if run.font.size is not None:
            return int(run.font.size)
    return _DEFAULT_FONT_EMU


def _check_text_clipping(
    slide_index: int,
    slide,
    slide_height: int,
    errors: list[str],
    warnings: list[str],
) -> None:
    """Check 4: estimate rendered text height and warn/error on likely clipping."""
    footer_threshold = int(slide_height * 0.9)

    for shape in slide.shapes:
        # Skip table shapes and non-text shapes
        if shape.has_table or not shape.has_text_frame:
            continue

        top = shape.top
        if top is None:
            continue

        # Skip footer-region shapes (top > 90% of slide height)
        if int(top) > footer_threshold:
            continue

        tf = shape.text_frame

        # Skip shapes with no text
        if not tf.text.strip():
            continue

        box_width = shape.width
        box_height = shape.height
        if box_width is None or box_height is None or box_width <= 0 or box_height <= 0:
            continue

        box_width = int(box_width)
        box_height = int(box_height)
        word_wrap = tf.word_wrap  # True, False, or None (None treated as True)

        estimated_height = 0.0
        for para in tf.paragraphs:
            font_size_emu = _para_font_size_emu(para)

            if word_wrap is False:
                line_count = 1
            else:
                avg_char_width = _AVG_CHAR_WIDTH_FACTOR * font_size_emu
                chars_per_line = box_width / avg_char_width if avg_char_width > 0 else 1
                text_len = len(para.text)
                line_count = max(1, math.ceil(text_len / chars_per_line)) if chars_per_line > 0 else 1

            estimated_height += font_size_emu * _LINE_SPACING * line_count

        if estimated_height <= 0:
            continue

        overage = (estimated_height - box_height) / box_height

        if overage > 0.20:
            est_in = estimated_height / _EMU_PER_INCH
            box_in = box_height / _EMU_PER_INCH
            pct = round(overage * 100)
            errors.append(
                f"Slide {slide_index}: '{shape.name}' text may be clipped"
                f" (est={est_in:.2f}\" vs box={box_in:.2f}\", +{pct}%)"
            )
        elif overage > 0.08:
            est_in = estimated_height / _EMU_PER_INCH
            box_in = box_height / _EMU_PER_INCH
            pct = round(overage * 100)
            warnings.append(
                f"Slide {slide_index}: '{shape.name}' text may be clipped"
                f" (est={est_in:.2f}\" vs box={box_in:.2f}\", +{pct}%)"
            )


# ── Public API ────────────────────────────────────────────────────────────────


def verify_pptx(path: Path) -> dict:
    """Run quality checks on a .pptx file.

    Returns a dict with keys:
        errors      -- list of error message strings
        warnings    -- list of warning message strings
        slide_count -- total number of slides
    """
    prs = Presentation(Path(path))
    slide_width: int = int(prs.slide_width or Emu(9144000))   # default 10in
    slide_height: int = int(prs.slide_height or Emu(6858000))  # default 7.5in

    errors: list[str] = []
    warnings: list[str] = []

    for slide_index, slide in enumerate(prs.slides, start=1):
        _check_unfilled_placeholders(slide_index, slide, errors)
        _check_font_sizes(slide_index, slide, slide_height, errors)
        _check_shape_overflow(slide_index, slide, slide_width, slide_height, errors)
        _check_text_clipping(slide_index, slide, slide_height, errors, warnings)

    return {
        "errors": errors,
        "warnings": warnings,
        "slide_count": len(prs.slides),
    }
