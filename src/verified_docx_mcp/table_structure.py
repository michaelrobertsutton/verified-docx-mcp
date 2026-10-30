"""Shared table-structure helpers for formatting-preserving cell edits.

Property-only helpers: IDs, revision markers, comments, bookmarks, fields and
relationships in source content are never duplicated. Copied from the #106
work so a later #106 port extends this module rather than conflicting with it.
"""

from __future__ import annotations

import copy
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from . import mutations, projection, tables
from .errors import ErrorCode, _make_error

w = tables._w


def positive(value: Any, name: str) -> int:
    if type(value) is not int or value < 1:
        raise _make_error(ErrorCode.INVALID_INPUT, f"{name} must be a positive integer")
    return value


def simple_table(root: Any, table_id: int) -> Any:
    positive(table_id, "table_id")
    table = tables._find_table_element(root, table_id)
    if tables._has_merged_cells(table) or tables._has_nested_table(table):
        raise _make_error(
            ErrorCode.MERGED_OR_NESTED_TABLE,
            "Structural edits require an unmerged table without nested tables",
        )
    return table


def inherit_properties(source: Any, dest: Any, name: str) -> None:
    """Overlay destination properties on source properties, preserving unknown tags."""
    original = source.find(w(name))
    overrides = dest.find(w(name))
    if original is None:
        return
    merged = copy.deepcopy(original)
    # Property-change histories belong to the original object, not the copy.
    for child in list(merged):
        if projection._ln(child).endswith("Change") or projection._ln(child) in {"ins", "del"}:
            merged.remove(child)
    if overrides is not None:
        for child in overrides:
            for old in list(merged):
                if old.tag == child.tag:
                    merged.remove(old)
            merged.append(copy.deepcopy(child))
        dest.remove(overrides)
    dest.insert(0, merged)


def paragraph_text(p: Any) -> str:
    return "".join(
        (node.text or "")
        if node.tag == w("t")
        else "\t"
        if node.tag == w("tab")
        else "\n"
        if node.tag in {w("br"), w("cr")}
        else ""
        for node in p.iter()
    )


def plain_row(row: Any) -> str:
    return "\t".join(
        "\n".join(paragraph_text(p) for p in cell.findall(w("p")))
        for cell in tables._row_cells(row)
    )


def property_signature(element: Any) -> tuple:
    """Prefix/order-independent direct formatting, ignoring revision history."""

    def node_key(node):
        return (
            node.tag,
            tuple(sorted(node.attrib.items())),
            node.text or "",
            tuple(
                sorted(
                    node_key(c)
                    for c in node
                    if not projection._ln(c).endswith("Change")
                    and projection._ln(c) not in {"ins", "del"}
                )
            ),
        )

    return tuple(
        node_key(n)
        for n in element.iter()
        if projection._ln(n) in {"tblPr", "trPr", "tcPr", "pPr", "rPr"}
    )


def flat_package(path: Path, root: Any, raw: bytes, ctx: Any) -> str:
    """Build a Flat OPC import preserving styles/relationships of rendered cells."""
    import base64
    import zipfile

    overrides = mutations._build_overrides(path, root, raw, ctx)
    pkg = "http://schemas.microsoft.com/office/2006/xmlPackage"
    ct = "http://schemas.openxmlformats.org/package/2006/content-types"
    package = ET.Element(f"{{{pkg}}}package")
    with zipfile.ZipFile(path) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}
    parts.update(overrides)
    types = ET.fromstring(parts["[Content_Types].xml"])
    defaults = {n.get("Extension"): n.get("ContentType") for n in types.findall(f"{{{ct}}}Default")}
    specific = {n.get("PartName"): n.get("ContentType") for n in types.findall(f"{{{ct}}}Override")}
    for name, data in parts.items():
        if name == "[Content_Types].xml":
            continue
        content_type = specific.get("/" + name) or defaults.get(name.rsplit(".", 1)[-1])
        part = ET.SubElement(
            package,
            f"{{{pkg}}}part",
            {
                f"{{{pkg}}}name": "/" + name,
                f"{{{pkg}}}contentType": content_type or "application/octet-stream",
            },
        )
        if name.endswith((".xml", ".rels")):
            ET.SubElement(part, f"{{{pkg}}}xmlData").append(ET.fromstring(data))
        else:
            ET.SubElement(part, f"{{{pkg}}}binaryData").text = base64.b64encode(data).decode()
    return ET.tostring(package, encoding="unicode")
