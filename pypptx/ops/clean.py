"""
clean.py — orphan file detection and removal for .pptx files.

All XML parsing uses defusedxml (security constraint).
xml.etree.ElementTree is imported only for namespace registration and
serialisation — never for parsing untrusted input.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import xml.etree.ElementTree as StdET
from pathlib import Path, PurePosixPath

import defusedxml.ElementTree as ET

from pypptx.ops.pack import pack, unpack

# ── Namespaces ────────────────────────────────────────────────────────────────

NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_PML = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_REL_TYPE_SLIDE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
)


# ── Path helpers ──────────────────────────────────────────────────────────────


def _rels_path_for(part: PurePosixPath) -> PurePosixPath:
    """Return the .rels file path for a given OPC part path.

    e.g. ppt/slides/slide1.xml  →  ppt/slides/_rels/slide1.xml.rels
         .  (package root)      →  _rels/.rels
    """
    return part.parent / "_rels" / (part.name + ".rels")


def _resolve_target(base_part: PurePosixPath, target: str) -> PurePosixPath | None:
    """Resolve a relationship *Target* relative to *base_part*.

    Returns ``None`` for external (``://``) targets that live outside the
    package.
    """
    if "://" in target:
        return None
    if target.startswith("/"):
        # Absolute part URI — strip leading slash and treat as package-relative.
        return PurePosixPath(target.lstrip("/"))
    return base_part.parent / target


def _normalize(p: PurePosixPath) -> str:
    """Canonical string key for a package-relative path.

    Resolves ``..`` components so that paths like
    ``ppt/slideMasters/../slideLayouts/slideLayout1.xml``
    normalise to ``ppt/slideLayouts/slideLayout1.xml``.
    """
    return os.path.normpath(p.as_posix())


# ── XML helpers ───────────────────────────────────────────────────────────────


def _parse_rels_file(path: Path) -> list[tuple[str, str, str]]:
    """Parse a ``.rels`` file; return list of ``(Id, Type, Target)`` tuples."""
    try:
        tree = ET.parse(str(path))
    except Exception:
        return []
    root = tree.getroot()
    return [
        (rel.get("Id", ""), rel.get("Type", ""), rel.get("Target", ""))
        for rel in root.findall(f"{{{NS_RELS}}}Relationship")
    ]


def _get_allowed_slide_rids(pptx_dir: Path) -> set[str]:
    """Return the rIds of slides listed in ``sldIdLst`` in presentation.xml.

    Slides present in ``ppt/_rels/presentation.xml.rels`` but *absent* from
    ``sldIdLst`` are orphaned and must not be treated as reachable.
    """
    prs = pptx_dir / "ppt" / "presentation.xml"
    if not prs.exists():
        return set()
    try:
        tree = ET.parse(str(prs))
    except Exception:
        return set()
    root = tree.getroot()
    sld_id_lst = root.find(f"{{{NS_PML}}}sldIdLst")
    if sld_id_lst is None:
        return set()
    return {
        el.get(f"{{{NS_R}}}id")
        for el in sld_id_lst
        if el.get(f"{{{NS_R}}}id")
    }


# ── Reachability walk ─────────────────────────────────────────────────────────


def _build_reachable(pptx_dir: Path, allowed_slide_rids: set[str]) -> set[str]:
    """Transitive ``.rels`` walk; return reachable posix paths relative to *pptx_dir*.

    Slide relationships whose ``Id`` is absent from *allowed_slide_rids* are
    excluded from the walk so their targets are not considered reachable.
    """
    reachable: set[str] = set()
    processed: set[str] = set()

    pkg_rels_posix = PurePosixPath("_rels/.rels")
    if not (pptx_dir / str(pkg_rels_posix)).exists():
        return reachable

    reachable.add(_normalize(pkg_rels_posix))

    # Queue of (rels_file_posix, owning_part_posix) pairs.
    # The owning part is used to resolve relative Target URIs.
    queue: list[tuple[PurePosixPath, PurePosixPath]] = [
        (pkg_rels_posix, PurePosixPath("."))
    ]

    while queue:
        rels_posix, base_part = queue.pop()
        rels_key = _normalize(rels_posix)
        if rels_key in processed:
            continue
        processed.add(rels_key)

        for rid, rtype, target in _parse_rels_file(pptx_dir / str(rels_posix)):
            # Slides missing from sldIdLst must not be walked.
            if rtype == _REL_TYPE_SLIDE and rid not in allowed_slide_rids:
                continue

            resolved = _resolve_target(base_part, target)
            if resolved is None:
                continue

            part_key = _normalize(resolved)
            if part_key not in reachable:
                reachable.add(part_key)

            # If this part has a .rels file, schedule it for processing.
            part_rels = _rels_path_for(resolved)
            part_rels_key = _normalize(part_rels)
            if (pptx_dir / part_rels_key).exists() and part_rels_key not in reachable:
                reachable.add(part_rels_key)
                queue.append((part_rels, resolved))

    return reachable


# ── Index-file updaters ───────────────────────────────────────────────────────


def _update_content_types(pptx_dir: Path, removed: set[str]) -> None:
    """Remove ``<Override>`` entries for *removed* files from ``[Content_Types].xml``."""
    ct_path = pptx_dir / "[Content_Types].xml"
    if not ct_path.exists():
        return

    try:
        tree = ET.parse(str(ct_path))
    except Exception:
        return

    root = tree.getroot()
    to_drop = [
        el
        for el in root.findall(f"{{{NS_CT}}}Override")
        if el.get("PartName", "").lstrip("/") in removed
    ]
    if not to_drop:
        return
    for el in to_drop:
        root.remove(el)

    StdET.register_namespace("", NS_CT)
    StdET.ElementTree(root).write(str(ct_path), xml_declaration=True, encoding="UTF-8")


def _update_presentation_rels(pptx_dir: Path, removed: set[str]) -> None:
    """Remove ``<Relationship>`` entries pointing to *removed* files from
    ``ppt/_rels/presentation.xml.rels``.
    """
    prs_rels = pptx_dir / "ppt" / "_rels" / "presentation.xml.rels"
    if not prs_rels.exists():
        return

    try:
        tree = ET.parse(str(prs_rels))
    except Exception:
        return

    root = tree.getroot()
    base = PurePosixPath("ppt/presentation.xml")
    to_drop = []
    for rel in root.findall(f"{{{NS_RELS}}}Relationship"):
        resolved = _resolve_target(base, rel.get("Target", ""))
        if resolved is not None and _normalize(resolved) in removed:
            to_drop.append(rel)
    if not to_drop:
        return
    for el in to_drop:
        root.remove(el)

    StdET.register_namespace("", NS_RELS)
    StdET.ElementTree(root).write(str(prs_rels), xml_declaration=True, encoding="UTF-8")


# ── Core pass ─────────────────────────────────────────────────────────────────


def _one_pass(pptx_dir: Path) -> list[str]:
    """Detect and remove one round of orphans; return removed posix paths.

    Returns an empty list when the package is already clean.
    """
    allowed_slide_rids = _get_allowed_slide_rids(pptx_dir)
    reachable = _build_reachable(pptx_dir, allowed_slide_rids)

    # The package index is never a removable part.
    reachable.add("[Content_Types].xml")

    all_files: set[str] = {
        f.relative_to(pptx_dir).as_posix()
        for f in pptx_dir.rglob("*")
        if f.is_file()
    }
    orphans = all_files - reachable
    if not orphans:
        return []

    for key in orphans:
        (pptx_dir / key).unlink(missing_ok=True)

    _update_content_types(pptx_dir, orphans)
    _update_presentation_rels(pptx_dir, orphans)

    return sorted(orphans)


def _clean_dir(pptx_dir: Path) -> list[str]:
    """Iteratively remove orphans from *pptx_dir* until convergence."""
    all_removed: list[str] = []
    while True:
        removed = _one_pass(pptx_dir)
        if not removed:
            break
        all_removed.extend(removed)
    return all_removed


# ── Public API ────────────────────────────────────────────────────────────────


def clean_unused_files(path: Path) -> list[str]:
    """Remove orphaned files from a ``.pptx`` file or unpacked directory.

    Parameters
    ----------
    path:
        A ``.pptx`` archive or a directory produced by :func:`~pypptx.ops.pack.unpack`.

    Returns
    -------
    Sorted list of removed file paths (posix, relative to the package root).
    Returns ``[]`` when nothing was removed.

    When *path* is a ``.pptx`` file the archive is unpacked to a temporary
    directory, cleaned, and repacked in-place.  The original file is left
    untouched if an error occurs during repacking.
    """
    path = Path(path)

    if path.is_file() and path.suffix.lower() == ".pptx":
        tmp_dir = Path(tempfile.mkdtemp(prefix="pypptx_clean_"))
        try:
            unpack(path, tmp_dir)
            removed = _clean_dir(tmp_dir)
            if removed:
                pack(tmp_dir, path)
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return removed

    return _clean_dir(path)
