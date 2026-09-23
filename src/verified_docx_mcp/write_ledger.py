# New for issue #27. No lifted analogue: this server had no memory of what it
# last wrote, so a co-author's overwrite (Office Online's autosave writing
# its stale in-memory copy back over the file) was indistinguishable from
# the state this server itself left behind.
"""Per-document write ledger: what this server last wrote, and whether the
file on disk still matches it.

Why this exists (issue #27). A file-mode write can be silently reverted by
an editor this server has no other signal for -- Office Online has no
Live pane session, writes no ``~$`` owner file, and leaves no conflict
copy. The one thing that IS observable is that the package on disk no
longer matches the bytes this server last put there. The ledger records a
full-package fingerprint (``projection.compute_package_fingerprint``) after
every file-mode write, so a later call can tell "someone else changed this
file" from "nothing happened since my write".

What the signal does and does NOT establish (the ``note`` field of
``external_activity`` says this too, verbatim):

- Divergence shows another writer changed the bytes. It does not show an
  editor is still open, who it was, or that the change was a co-author's
  rather than a completed save, a restore, or this server's own live save
  (live writes never reach the ledger).
- It is detection after the fact. A file this server never wrote has no
  ledger entry, and an external A -> B -> A sequence between observations
  is invisible.
- An expired window is not proof it is safe to write.

Storage: one JSON record per document under ``<state dir>/ledger/`` (never
a shared map, so writes to different documents cannot lose each other's
entries), plus a separate ``.obs`` sidecar for the first-observed time of
a divergence. The sidecar is a different file from the main record on
purpose: ``external_activity`` runs from read tools with no write claim,
and folding its observation into the main record could lose a concurrent
``record_write``'s update.
"""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
import time
from pathlib import Path
from typing import Any

from . import audit, projection

_EXTERNAL_EDIT_WINDOW_S = 600.0
_WINDOW_ENV = "VERIFIED_DOCX_MCP_EXTERNAL_EDIT_WINDOW_S"

LEDGER_OK = "ok"
LEDGER_NONE = "none"
LEDGER_UNREADABLE = "unreadable"

NOTE = (
    "divergent means another writer changed this package after this server's last file-mode write; "
    "it does not show an editor is still open, or who changed it. No divergence is not proof it is "
    "safe to write (a file this server never wrote has no ledger entry), and an elapsed window is "
    "not proof either."
)


def window_s() -> float:
    """The observation window in seconds: the env var when set to a valid
    non-negative number, else the module default."""
    raw = os.environ.get(_WINDOW_ENV)
    if raw is not None:
        try:
            value = float(raw)
        except ValueError:
            return _EXTERNAL_EDIT_WINDOW_S
        if value >= 0:
            return value
    return _EXTERNAL_EDIT_WINDOW_S


def _ledger_dir() -> Path:
    return audit._state_dir() / "ledger"


def _key(resolved: Path) -> str:
    return hashlib.sha256(str(resolved).encode("utf-8")).hexdigest()[:32]


def _record_path(resolved: Path) -> Path:
    return _ledger_dir() / f"{_key(resolved)}.json"


def _obs_path(resolved: Path) -> Path:
    return _ledger_dir() / f"{_key(resolved)}.obs"


def _atomic_write_json(target: Path, payload: dict[str, Any]) -> None:
    target.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    fd, tmp_name = tempfile.mkstemp(prefix=target.name + ".", suffix=".tmp", dir=str(target.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump(payload, fh)
        os.chmod(tmp_name, 0o600)
        os.replace(tmp_name, target)
    except BaseException:
        Path(tmp_name).unlink(missing_ok=True)
        raise


def _read_record(resolved: Path) -> tuple[str, dict[str, Any] | None, str]:
    """(status, record, reason). A record file that exists but cannot be
    read, parsed, or does not carry the required fields is ``unreadable``
    -- never silently ``none``, because ``none`` permits a write."""
    path = _record_path(resolved)
    try:
        raw = path.read_text(encoding="utf-8")
    except (FileNotFoundError, NotADirectoryError):
        # NotADirectoryError: a path component of the state dir is a plain
        # file, so no record can exist and none can ever be written. That
        # is "no ledger", not a damaged one -- refusing every write on a
        # machine whose state dir is unusable would make the server
        # unusable there. Each write reports ledger_logged=false instead.
        return LEDGER_NONE, None, ""
    except OSError as exc:
        return LEDGER_UNREADABLE, None, f"{type(exc).__name__}: {exc}"
    try:
        record = json.loads(raw)
    except ValueError as exc:
        return LEDGER_UNREADABLE, None, f"corrupt JSON: {exc}"
    if (
        not isinstance(record, dict)
        or not isinstance(record.get("fingerprint"), str)
        or not isinstance(record.get("token"), str)
        or not isinstance(record.get("written_at"), (int, float))
    ):
        return LEDGER_UNREADABLE, None, "record is missing required fields"
    return LEDGER_OK, record, ""


def record_write(resolved: Path, fingerprint: str, token: str) -> tuple[bool, str]:
    """Record the state this server just left on disk. Never raises;
    returns ``(ok, reason)`` like ``audit.append_audit`` so a caller can
    surface a failed record instead of losing the protection silently.

    On failure the previous record (if any) is removed: leaving it would
    make this server's own next state look like an external change.
    """
    try:
        _atomic_write_json(
            _record_path(resolved),
            {
                "path": str(resolved),
                "fingerprint": fingerprint,
                "token": token,
                "written_at": time.time(),
            },
        )
        _obs_path(resolved).unlink(missing_ok=True)
        return True, ""
    except Exception as exc:  # noqa: BLE001
        try:
            _record_path(resolved).unlink(missing_ok=True)
            _obs_path(resolved).unlink(missing_ok=True)
        except OSError:
            pass
        return False, f"{type(exc).__name__}: {exc}"


def _observed_at(resolved: Path, ledger_fp: str, current_fp: str, now: float) -> float:
    """First time this server SAW *current_fp* diverge from *ledger_fp*.
    Persisted best-effort in the ``.obs`` sidecar; on any failure the
    observation is simply "now" (a fresh divergence), which errs toward
    refusing rather than toward silence."""
    path = _obs_path(resolved)
    try:
        obs = json.loads(path.read_text(encoding="utf-8"))
        if (
            isinstance(obs, dict)
            and obs.get("ledger_fingerprint") == ledger_fp
            and obs.get("observed_fingerprint") == current_fp
            and isinstance(obs.get("observed_at"), (int, float))
        ):
            return float(obs["observed_at"])
    except (OSError, ValueError):
        pass
    try:
        _atomic_write_json(
            path,
            {"ledger_fingerprint": ledger_fp, "observed_fingerprint": current_fp, "observed_at": now},
        )
    except Exception:  # noqa: BLE001 -- best-effort; unpersisted, the next look re-observes as "now"
        return now
    return now


def external_activity(
    resolved: Path,
    *,
    window: float | None = None,
    now: float | None = None,
    current_fingerprint: str | None = None,
) -> dict[str, Any]:
    """Compare the file on disk against this server's last recorded write.
    Data only -- never raises for a missing/corrupt ledger (that is
    reported in ``ledger``), never refuses. The guard in
    ``mutations._guard_before_write`` decides what to do with it.

    Age is measured from when this server first OBSERVED the divergence
    (``divergence_first_observed_at``), not from the file's mtime: a
    downloaded file can carry an old mtime, a clock skew can carry a
    future one, and a touch of an already-divergent file would otherwise
    make an old change look recent.
    """
    win = window_s() if window is None else window
    ts = time.time() if now is None else now
    status, record, reason = _read_record(resolved)

    result: dict[str, Any] = {
        "ledger": status,
        "ledger_reason": reason,
        "last_write": None,
        "still_current": None,
        "divergent": False,
        "divergence_first_observed_at": None,
        "divergence_age_s": None,
        "recent": False,
        "unattributed_recent_mtime": False,
        "window_s": win,
        "note": NOTE,
    }

    if status == LEDGER_UNREADABLE:
        return result

    if status == LEDGER_NONE:
        try:
            age = ts - resolved.stat().st_mtime
        except OSError:
            return result
        result["unattributed_recent_mtime"] = age <= win
        return result

    assert record is not None
    result["last_write"] = {"token": record["token"], "written_at": record["written_at"]}
    try:
        current_fp = (
            current_fingerprint
            if current_fingerprint is not None
            else projection.compute_package_fingerprint(resolved)
        )
    except Exception as exc:  # noqa: BLE001 -- mid-sync partial zip, missing file, ...
        # "unknown", not "unchanged": a failed read must not read as a match.
        result["ledger_reason"] = f"could not fingerprint the current file: {type(exc).__name__}: {exc}"
        return result

    result["still_current"] = current_fp == record["fingerprint"]
    if result["still_current"]:
        _obs_path(resolved).unlink(missing_ok=True)
        return result

    result["divergent"] = True
    observed = _observed_at(resolved, record["fingerprint"], current_fp, ts)
    result["divergence_first_observed_at"] = observed
    result["divergence_age_s"] = max(0.0, ts - observed)
    result["recent"] = result["divergence_age_s"] <= win
    return result
