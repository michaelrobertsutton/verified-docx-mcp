"""Live reads for the file-read tools (issue #33).

``read_document`` / ``find_sections`` / ``list_tables`` / ``get_table`` read a
local ``.docx``. On a co-authored SharePoint document the live Word body is the
source of truth, and the local file is at best a same-named placeholder (live
writes need one in an allowed root). Before this module those tools read the
placeholder and returned ``warnings: []`` -- a placeholder read could pass as
the live body.

``read_source`` is the single entry point all four tools go through. It picks
between the connected pane (``body_ooxml`` op -> Flat OPC -> a temporary
``.docx`` the existing projection code reads unchanged) and the file, and
always reports which one it used:

- ``source="live"``: needs a connected pane that reports the ``body_ooxml``
  capability, and the default part (``getOoxml`` covers the body only).
- ``source="file"``: the file, as before.
- ``source="auto"`` (the tools' default): live when a pane session exists for
  the document's basename AND the read can actually be served live AND no
  other document has ever connected under the same basename (SharePoint
  sessions are matched by basename alone, so a collision means the pane may
  be editing a different document). Otherwise the file.

Two rules keep a placeholder from passing as the live body:

1. A file read while ANY session is registered for the basename carries the
   ``live_session_ignored`` warning plus a ``live_session`` object saying why.
2. Once live is chosen it is final: a timeout, disconnect, ``too_large`` or
   conversion failure raises, it never falls back to the file.

``flat_opc_to_docx`` is an adapter for projection only. Its output is not a
valid export: binary parts are empty and namespace prefixes are re-generated.
"""

from __future__ import annotations

import base64
import binascii
import tempfile
import zipfile
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .. import paths
from ..errors import ErrorCode, VerifyError, _make_error
from ..projection import DEFAULT_PART
from . import bridge as live_bridge
from . import write_mode
from .protocol import OP_ERROR_TOO_LARGE
from .session import LiveDisconnected, LiveOpFailed, LiveSession, LiveStale, LiveUnavailable

VALID_SOURCES = ("auto", "file", "live")
BODY_OOXML_CAPABILITY = "body_ooxml"

WARNING_LIVE_SESSION_IGNORED = "live_session_ignored"
WARNING_LIVE_STYLES_MISSING = "live_styles_missing"
WARNING_LIVE_NUMBERING_MISSING = "live_numbering_missing"

REASON_REQUESTED_FILE = "requested_file"
REASON_PART_NOT_IN_LIVE = "part_not_in_live"
REASON_PANE_MISSING_CAPABILITY = "pane_missing_capability"
REASON_BASENAME_COLLISION = "basename_collision"
REASON_SESSION_MISMATCH = "session_mismatch"

LIVE_REVISION_PREFIX = "live:sha256:"

_PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage"
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_DOCUMENT_PART = "word/document.xml"
_MAX_PARTS = 5000
_MAX_TOTAL_BYTES = 256 * 2**20

_REASON_ACTION = {
    REASON_REQUESTED_FILE: 'pass source="live" (or "auto") to read the connected live document instead',
    REASON_PART_NOT_IN_LIVE: 'live reads cover the document body only; read the body with source="live" to see live content',
    REASON_PANE_MISSING_CAPABILITY: "the connected pane is an older build; close and reopen the Live pane in Word, then retry",
    REASON_BASENAME_COLLISION: (
        "another document connected under this same file name; confirm which one is open in Word, "
        'then pass source="live" explicitly'
    ),
    REASON_SESSION_MISMATCH: (
        "the pane's document_url names a different local file than this path; this read is of the file you named"
    ),
}


def _flat_opc_error(message: str, **diagnostics: Any) -> VerifyError:
    return _make_error(ErrorCode.LIVE_OP_FAILED, message, {"stage": "flat_opc", **diagnostics})


def _q(local: str, ns: str = _PKG_NS) -> str:
    return f"{{{ns}}}{local}"


@dataclass
class FlatOpcResult:
    part_names: list[str]
    warnings: list[str] = field(default_factory=list)


def flat_opc_to_docx(ooxml: str, dest: Path) -> FlatOpcResult:
    """Write the Flat OPC ``pkg:package`` string *ooxml* to *dest* as a zip
    package projection can read. Raises ``LIVE_OP_FAILED`` (diagnostics
    ``stage: "flat_opc"``) on anything that is not a well-formed package
    with a body part. Namespace-aware, so a non-``pkg`` prefix is fine.
    """
    # stdlib ElementTree expands entities; a Flat OPC export never has a DTD.
    if "<!DOCTYPE" in ooxml or "<!ENTITY" in ooxml:
        raise _flat_opc_error("Flat OPC export contains a DTD/entity declaration; refusing to parse it")
    try:
        root = ET.fromstring(ooxml)
    except ET.ParseError as exc:
        raise _flat_opc_error(f"Flat OPC export is not well-formed XML: {exc}") from exc
    if root.tag != _q("package"):
        raise _flat_opc_error(f"Flat OPC export root is {root.tag!r}, expected pkg:package")

    entries: dict[str, tuple[str, bytes]] = {}
    total = 0
    for part in root.findall(_q("part")):
        name = part.get(_q("name"))
        content_type = part.get(_q("contentType"))
        if not name or not name.startswith("/"):
            raise _flat_opc_error("Flat OPC part has a missing or non-absolute pkg:name", part_name=name)
        if not content_type:
            raise _flat_opc_error("Flat OPC part has no pkg:contentType", part_name=name)
        zip_name = name.lstrip("/")
        if zip_name in entries:
            raise _flat_opc_error("Flat OPC export names the same part twice", part_name=name)
        xml_data = part.find(_q("xmlData"))
        binary_data = part.find(_q("binaryData"))
        if xml_data is not None:
            children = list(xml_data)
            if len(children) != 1:
                raise _flat_opc_error("pkg:xmlData must hold exactly one root element", part_name=name)
            data = ET.tostring(children[0], encoding="utf-8", xml_declaration=True)
        elif binary_data is not None:
            try:
                data = base64.b64decode((binary_data.text or "").strip())
            except (binascii.Error, ValueError) as exc:
                raise _flat_opc_error("pkg:binaryData is not valid base64", part_name=name) from exc
        else:
            raise _flat_opc_error("Flat OPC part has neither pkg:xmlData nor pkg:binaryData", part_name=name)
        total += len(data)
        if len(entries) + 1 > _MAX_PARTS or total > _MAX_TOTAL_BYTES:
            raise _flat_opc_error(
                "Flat OPC export exceeds the part-count or size limit", parts=len(entries) + 1, bytes=total
            )
        entries[zip_name] = (content_type, data)

    if _DOCUMENT_PART not in entries:
        raise _flat_opc_error("Flat OPC export has no /word/document.xml part", parts=sorted(entries))
    document_root = ET.fromstring(entries[_DOCUMENT_PART][1])
    if document_root.find(_q("body", _W_NS)) is None:
        raise _flat_opc_error("/word/document.xml has no w:body")

    types = ET.Element("Types", xmlns="http://schemas.openxmlformats.org/package/2006/content-types")
    for zip_name, (content_type, _data) in entries.items():
        ET.SubElement(types, "Override", PartName=f"/{zip_name}", ContentType=content_type)
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ET.tostring(types, encoding="utf-8", xml_declaration=True))
        for zip_name, (_content_type, data) in entries.items():
            zf.writestr(zip_name, data)

    return FlatOpcResult(part_names=sorted(entries), warnings=_fidelity_warnings(document_root, entries))


def _fidelity_warnings(document_root: ET.Element, entries: dict[str, Any]) -> list[str]:
    """The document references styles/numbering but the export omitted the
    part: projection would silently fall back to no styles / no numbering and
    could mis-render headings and lists. Report it; never fill the gap from
    the local file."""
    warnings: list[str] = []
    local_names = {el.tag.rsplit("}", 1)[-1] for el in document_root.iter()}
    if local_names & {"pStyle", "rStyle", "tblStyle"} and "word/styles.xml" not in entries:
        warnings.append(WARNING_LIVE_STYLES_MISSING)
    if "numPr" in local_names and "word/numbering.xml" not in entries:
        warnings.append(WARNING_LIVE_NUMBERING_MISSING)
    return warnings


@dataclass
class ReadSource:
    """What ``read_source`` yields: the path projection should read, plus
    the provenance every tool response carries."""

    local_path: Path
    resolved: Path
    source: str
    warnings: list[str] = field(default_factory=list)
    live: dict[str, Any] | None = None
    live_session: dict[str, Any] | None = None
    revision: str | None = None  # "live:sha256:<hex>" for a live read, else None

    def annotate(self, result: dict[str, Any]) -> dict[str, Any]:
        """Add provenance to a tool response in place: ``source``, ``live`` /
        ``live_session``, and any warnings merged into an existing
        ``warnings`` list (created only when there is something to say). A
        live read also replaces the file revision keys the tool computed
        from the temporary package (they would describe it, not the live
        body)."""
        result["source"] = self.source
        if self.live is not None:
            result["live"] = self.live
        if self.live_session is not None:
            result["live_session"] = self.live_session
        if self.warnings:
            existing = result.setdefault("warnings", [])
            existing.extend(w for w in self.warnings if w not in existing)
        if self.revision is not None and "revision" in result:
            result["revision"] = self.revision
            result["revision_detail"] = None
        return result


def _collided(registry: Any, document_name: str) -> bool:
    collisions = getattr(registry, "collisions", None)
    if collisions is None:
        return False
    return any(c.get("document_name") == document_name for c in collisions())


def _session_if_any(document_name: str) -> tuple[Any, LiveSession | None]:
    # Never start the bridge here: a file read must not bind the live ports.
    registry = live_bridge.current_registry()
    return registry, (registry.get(document_name) if registry is not None else None)


def _fetch_body_ooxml(session: LiveSession) -> dict[str, Any]:
    from .comments_live import _raise_from_live_error

    try:
        result = session.request_threadsafe("body_ooxml", {})
    except LiveOpFailed as exc:
        if exc.code == OP_ERROR_TOO_LARGE:
            raise _make_error(
                ErrorCode.LIVE_OP_FAILED,
                f"{exc.message}. The live body is too large to read through the pane; "
                'use source="file" to read the local copy instead (it may not match the live body).',
                {"pane_error_code": exc.code},
            ) from exc
        _raise_from_live_error(exc)
        raise
    except (LiveUnavailable, LiveDisconnected, LiveStale) as exc:
        _raise_from_live_error(exc)
        raise
    ooxml = result.get("ooxml")
    body_sha = result.get("bodySha256")
    if not isinstance(ooxml, str) or not isinstance(body_sha, str):
        raise _make_error(
            ErrorCode.LIVE_OP_FAILED,
            "the pane's body_ooxml reply is missing ooxml/bodySha256",
            {"stage": "body_ooxml_reply"},
        )
    return result


@contextmanager
def read_source(path: str, source: str, part: str = DEFAULT_PART) -> Iterator[ReadSource]:
    """Resolve *source* for a read of *path* and yield a ``ReadSource``;
    temporary files are removed on exit. See the module docstring."""
    from .. import server as _server

    if source not in VALID_SOURCES:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"source must be one of {VALID_SOURCES}, got {source!r}",
            {"source": source},
        )
    # must_exist=False: a SharePoint-only document has no local file (issue #22 B3).
    resolved = paths.resolve_allowed_docx_path(path, must_exist=False)
    document_name = resolved.name
    registry, session = _session_if_any(document_name)

    # Why a connected session could NOT serve this read live (None = it can).
    blocker: str | None = None
    if session is not None:
        if part != DEFAULT_PART:
            blocker = REASON_PART_NOT_IN_LIVE
        elif BODY_OOXML_CAPABILITY not in session.hello.capabilities:
            blocker = REASON_PANE_MISSING_CAPABILITY
        elif _collided(registry, document_name):
            blocker = REASON_BASENAME_COLLISION
        else:
            try:
                write_mode._check_session_identity(path, session)
            except VerifyError as exc:
                if exc.envelope.error_code != ErrorCode.LIVE_SESSION_MISMATCH:
                    raise
                blocker = REASON_SESSION_MISMATCH

    if source == "live":
        if session is None:
            raise _make_error(
                ErrorCode.LIVE_UNAVAILABLE,
                f"source='live' requested but no connected pane session for document {document_name!r}. "
                "Open the Live pane in Word (docs/live-mode.md) and retry, or call live_status.",
                {"document_name": document_name},
            )
        if blocker == REASON_PART_NOT_IN_LIVE:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f"source='live' reads the document body only (part {DEFAULT_PART!r}); got part {part!r}.",
                {"part": part},
            )
        write_mode.require_capability(session, BODY_OOXML_CAPABILITY, feature_description="live reads (source='live')")
        # An explicit live read goes through a basename collision (the caller
        # asked for it by name), but a document_url naming a different local
        # file still refuses.
        write_mode._check_session_identity(path, session)
        use_live = True
    elif source == "auto":
        use_live = session is not None and blocker is None
    else:
        use_live = False
        if session is not None:
            blocker = REASON_REQUESTED_FILE

    if use_live:
        assert session is not None
        reply = _fetch_body_ooxml(session)
        with tempfile.TemporaryDirectory(prefix="verified-docx-live-read-") as tmp:
            temp_docx = Path(tmp) / document_name
            built = flat_opc_to_docx(reply["ooxml"], temp_docx)
            yield ReadSource(
                local_path=temp_docx,
                resolved=resolved,
                source="live",
                warnings=list(built.warnings),
                live={
                    "document_url": session.document_url,
                    "body_sha256": reply["bodySha256"],
                    "stripped_parts": reply.get("strippedParts") or [],
                },
                revision=f"{LIVE_REVISION_PREFIX}{reply['bodySha256']}",
            )
        return

    # File read: the file must exist (a live-only document needs source="live").
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    local_path, is_temp = _server._read_local_copy(resolved)
    try:
        warnings: list[str] = []
        live_session: dict[str, Any] | None = None
        if session is not None:
            warnings.append(WARNING_LIVE_SESSION_IGNORED)
            live_session = {
                "document_url": session.document_url,
                "reason": blocker,
                "action": _REASON_ACTION.get(blocker or "", ""),
            }
        yield ReadSource(
            local_path=local_path, resolved=resolved, source="file", warnings=warnings, live_session=live_session
        )
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


def file_read_warnings(path: str) -> tuple[list[str], dict[str, Any] | None]:
    """For tools that stay file-only (``list_parts``): the
    ``live_session_ignored`` warning and ``live_session`` object when a pane
    session is registered for *path*'s basename, else ``([], None)``."""
    resolved = paths.resolve_allowed_docx_path(path, must_exist=False)
    _registry, session = _session_if_any(resolved.name)
    if session is None:
        return [], None
    return [WARNING_LIVE_SESSION_IGNORED], {
        "document_url": session.document_url,
        "reason": REASON_REQUESTED_FILE,
        "action": 'this tool reads the local file only; use read_document(source="live") for the live body',
    }
