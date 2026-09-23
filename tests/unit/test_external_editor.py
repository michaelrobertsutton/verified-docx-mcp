"""Unit tests for issue #27: file-mode writes silently reverted by a
co-authoring editor (Office Online) that this server has no other signal for.

Covers the write ledger (write_ledger.py), the EXTERNAL_EDITOR_ACTIVE guard
and its allow_concurrent_editor override, lock_status's external_activity
block, and the write-window fixes in mutations.atomic_replace_docx_parts
(staged revision_after, recheck under the claim, guarded rollback).

An "external editor" is simulated by rewriting the .docx out of band -- the
server cannot tell a co-author's autosave from any other writer, which is
exactly the property under test.
"""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
import threading
import time
import unittest
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import mutations, paths, projection, server, tables, write_ledger
from verified_docx_mcp.errors import ErrorCode, VerifyError

FIXTURES = REPO / "tests" / "fixtures"

mutations._QUIESCE_INTERVAL_SECONDS = 0.02


def _external_overwrite(path: Path, *, old: str = "R2C2", new: str = "EXTERNAL") -> None:
    """Rewrite *path* the way a co-authoring editor's autosave would: same
    package, different content, written from another process. Defaults to
    a cell the tests never write (row 2, col 2), so the change always lands."""
    tmp = path.with_name(path.name + ".ext")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item)
            if item.filename == projection.DEFAULT_PART:
                data = data.replace(old.encode(), new.encode())
            dst.writestr(item, data)
    os.replace(tmp, path)


def _add_comments_part(path: Path) -> None:
    """A comments-only external change: the document body is untouched."""
    tmp = path.with_name(path.name + ".ext")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            dst.writestr(item, src.read(item))
        dst.writestr(
            "word/comments.xml",
            b'<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w='
            b'"http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        )
    os.replace(tmp, path)


class _Base(unittest.TestCase):
    fixture_name = "tables.docx"

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self._old_xdg = os.environ.get("XDG_STATE_HOME")
        self.state = Path(self._tmp.name) / "state"
        os.environ["XDG_STATE_HOME"] = str(self.state)
        self._old_window = write_ledger._EXTERNAL_EDIT_WINDOW_S
        self._old_env_window = os.environ.pop(write_ledger._WINDOW_ENV, None)
        self.target = Path(self._tmp.name) / self.fixture_name
        shutil.copyfile(FIXTURES / self.fixture_name, self.target)

    def tearDown(self):
        write_ledger._EXTERNAL_EDIT_WINDOW_S = self._old_window
        if self._old_env_window is not None:
            os.environ[write_ledger._WINDOW_ENV] = self._old_env_window
        for key, old in (
            (paths._ALLOWED_FILE_ROOTS_ENV, self._old_allowed),
            ("XDG_STATE_HOME", self._old_xdg),
        ):
            if old is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = old
        self._tmp.cleanup()

    def write_cell(self, text: str, *, cell: int = 1, **kwargs):
        return tables.execute_replace_cell_markdown(str(self.target), 1, 1, cell, text, **kwargs)

    def activity(self):
        return server.execute_lock_status(str(self.target), quiesce_interval=0.01)["external_activity"]

    def assertRefused(self, code: ErrorCode, fn, *args, **kwargs):
        with self.assertRaises(VerifyError) as ctx:
            fn(*args, **kwargs)
        self.assertEqual(ctx.exception.envelope.error_code, code)
        return ctx.exception.envelope


class LedgerAndGuardTests(_Base):
    def test_write_records_ledger_and_reports_still_current(self):
        evidence = self.write_cell("first")
        self.assertTrue(evidence["applied"])
        self.assertTrue(evidence["ledger_logged"])
        activity = self.activity()
        self.assertEqual(activity["ledger"], "ok")
        self.assertTrue(activity["still_current"])
        self.assertFalse(activity["divergent"])
        self.assertFalse(activity["recent"])

    def test_body_overwrite_is_reported_and_refuses_next_write(self):
        self.write_cell("first")
        _external_overwrite(self.target)
        activity = self.activity()
        self.assertFalse(activity["still_current"])
        self.assertTrue(activity["divergent"])
        self.assertTrue(activity["recent"])
        envelope = self.assertRefused(ErrorCode.EXTERNAL_EDITOR_ACTIVE, self.write_cell, "second")
        self.assertIn("allow_concurrent_editor=True", envelope.message)
        # replace_cell_markdown has a live route, so its message may name it...
        self.assertIn('write_mode="live"', envelope.message)
        # ...but a file-mode tool with no live route must not recommend it.
        envelope = self.assertRefused(
            ErrorCode.EXTERNAL_EDITOR_ACTIVE,
            tables.execute_replace_table_row,
            str(self.target),
            1,
            1,
            ["a", "b"],
        )
        self.assertNotIn('write_mode="live"', envelope.message)
        self.assertIn("allow_concurrent_editor=True", envelope.message)

    def test_comments_only_external_change_is_also_detected(self):
        self.write_cell("first")
        _add_comments_part(self.target)
        self.assertTrue(self.activity()["divergent"])
        self.assertRefused(ErrorCode.EXTERNAL_EDITOR_ACTIVE, self.write_cell, "second")

    def test_live_capable_tool_message_recommends_live_mode(self):
        self.write_cell("first")
        _external_overwrite(self.target)
        from verified_docx_mcp import text_edit

        envelope = self.assertRefused(
            ErrorCode.EXTERNAL_EDITOR_ACTIVE,
            text_edit.execute_replace_text,
            str(self.target),
            "R1C2",
            "x",
            1,
            write_mode="file",
        )
        self.assertIn('write_mode="live"', envelope.message)

    def test_override_writes_and_records_it(self):
        self.write_cell("first")
        _external_overwrite(self.target)
        evidence = self.write_cell("second", allow_concurrent_editor=True)
        self.assertTrue(evidence["applied"])
        self.assertIn("concurrent_editor_override", evidence)
        self.assertIsNotNone(evidence["concurrent_editor_override"]["age_s"])
        # The write is the new baseline: nothing left to refuse on.
        self.assertTrue(self.activity()["still_current"])
        self.assertTrue(self.write_cell("third")["applied"])

    def test_override_is_absent_when_no_activity_was_present(self):
        self.write_cell("first")
        evidence = self.write_cell("second", allow_concurrent_editor=True)
        self.assertNotIn("concurrent_editor_override", evidence)

    def test_stale_revision_wins_over_external_editor_check_even_with_override(self):
        first = self.write_cell("first")
        _external_overwrite(self.target)
        for override in (False, True):
            envelope = self.assertRefused(
                ErrorCode.REVISION_CONFLICT,
                self.write_cell,
                "second",
                revision_before=first["revision_after"],
                allow_concurrent_editor=override,
            )
            self.assertIn("current_revision", envelope.diagnostics)

    def test_first_write_to_a_file_never_written_by_this_server_is_not_refused(self):
        activity = self.activity()
        self.assertEqual(activity["ledger"], "none")
        self.assertIsNone(activity["still_current"])
        # Fresh copy => mtime is inside the window: reported, never refused.
        self.assertTrue(activity["unattributed_recent_mtime"])
        self.assertTrue(self.write_cell("first")["applied"])

    def test_no_ledger_and_old_mtime_is_not_flagged(self):
        os.utime(self.target, (1_000_000_000, 1_000_000_000))
        self.assertFalse(self.activity()["unattributed_recent_mtime"])

    def test_mtime_only_touch_does_not_flag(self):
        self.write_cell("first")
        os.utime(self.target, (1_000_000_000, 1_000_000_000))
        activity = self.activity()
        self.assertTrue(activity["still_current"])
        self.assertFalse(activity["divergent"])
        self.assertTrue(self.write_cell("second")["applied"])

    def test_old_mtime_does_not_shorten_the_observation_window(self):
        self.write_cell("first")
        _external_overwrite(self.target)
        os.utime(self.target, (1_000_000_000, 1_000_000_000))  # a downloaded copy with an old stamp
        self.assertTrue(self.activity()["recent"])
        self.assertRefused(ErrorCode.EXTERNAL_EDITOR_ACTIVE, self.write_cell, "second")

    def test_future_mtime_does_not_extend_the_observation_window(self):
        write_ledger._EXTERNAL_EDIT_WINDOW_S = 0.05
        self.write_cell("first")
        _external_overwrite(self.target)
        future = time.time() + 86_400
        os.utime(self.target, (future, future))
        self.assertTrue(self.activity()["recent"])  # first observation
        time.sleep(0.12)
        self.assertFalse(self.activity()["recent"])
        self.assertTrue(self.write_cell("second")["applied"])

    def test_window_lapse_stops_refusing(self):
        write_ledger._EXTERNAL_EDIT_WINDOW_S = 0.05
        self.write_cell("first")
        _external_overwrite(self.target)
        self.assertTrue(self.activity()["recent"])
        time.sleep(0.12)
        activity = self.activity()
        self.assertTrue(activity["divergent"])
        self.assertFalse(activity["recent"])
        self.assertIn("not proof", activity["note"])
        self.assertTrue(self.write_cell("second")["applied"])

    def test_a_further_external_change_restarts_the_window(self):
        write_ledger._EXTERNAL_EDIT_WINDOW_S = 0.05
        self.write_cell("first")
        _external_overwrite(self.target)
        self.activity()
        time.sleep(0.12)
        self.assertFalse(self.activity()["recent"])
        _external_overwrite(self.target, old="EXTERNAL", new="EXTERNAL2")
        self.assertTrue(self.activity()["recent"])

    def test_window_env_var_is_honored(self):
        os.environ[write_ledger._WINDOW_ENV] = "1234"
        self.assertEqual(write_ledger.window_s(), 1234.0)
        os.environ[write_ledger._WINDOW_ENV] = "not a number"
        self.assertEqual(write_ledger.window_s(), write_ledger._EXTERNAL_EDIT_WINDOW_S)


class LedgerFailureTests(_Base):
    def _record_file(self) -> Path:
        return write_ledger._record_path(self.target.resolve())

    def test_corrupt_record_fails_closed_and_override_recovers(self):
        self.write_cell("first")
        self._record_file().write_text("{not json")
        activity = self.activity()
        self.assertEqual(activity["ledger"], "unreadable")
        self.assertIsNone(activity["still_current"])
        envelope = self.assertRefused(ErrorCode.EXTERNAL_EDITOR_ACTIVE, self.write_cell, "second")
        self.assertEqual(envelope.diagnostics["ledger"], "unreadable")
        evidence = self.write_cell("second", allow_concurrent_editor=True)
        self.assertTrue(evidence["applied"])
        self.assertEqual(self.activity()["ledger"], "ok")

    def test_record_missing_required_fields_is_unreadable(self):
        self.write_cell("first")
        self._record_file().write_text('{"fingerprint": "x"}')
        self.assertEqual(self.activity()["ledger"], "unreadable")

    def test_unwritable_state_dir_reports_ledger_logged_false_and_still_writes(self):
        blocker = Path(self._tmp.name) / "blocked"
        blocker.write_bytes(b"not a directory")
        os.environ["XDG_STATE_HOME"] = str(blocker)
        evidence = self.write_cell("first")
        self.assertTrue(evidence["applied"])
        self.assertFalse(evidence["ledger_logged"])
        self.assertTrue(evidence["ledger_reason"])

    def test_failed_record_removes_the_previous_record(self):
        # A stale record left behind would label this server's OWN next state external.
        self.write_cell("first")
        self.assertTrue(self._record_file().exists())
        real = write_ledger._atomic_write_json

        def failing(*args, **kwargs):
            raise OSError("disk full")

        write_ledger._atomic_write_json = failing
        try:
            ok, reason = write_ledger.record_write(self.target.resolve(), "fp", "tok")
        finally:
            write_ledger._atomic_write_json = real
        self.assertFalse(ok)
        self.assertIn("disk full", reason)
        self.assertFalse(self._record_file().exists())
        self.assertEqual(self.activity()["ledger"], "none")

    def test_records_for_different_documents_do_not_clobber_each_other(self):
        paths_ = [Path(self._tmp.name) / f"doc-{i}.docx" for i in range(24)]
        errors: list[str] = []

        def worker(p: Path, i: int):
            ok, reason = write_ledger.record_write(p, f"fp-{i}", f"tok-{i}")
            if not ok:
                errors.append(reason)

        threads = [threading.Thread(target=worker, args=(p, i)) for i, p in enumerate(paths_)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        self.assertEqual(errors, [])
        for i, p in enumerate(paths_):
            status, record, _ = write_ledger._read_record(p)
            self.assertEqual(status, write_ledger.LEDGER_OK)
            self.assertEqual(record["fingerprint"], f"fp-{i}")


class IncidentReplayTests(_Base):
    def test_34_writes_chain_revisions_then_a_revert_is_caught(self):
        """Issue #27's sequence: many file-mode writes each returning
        applied:true, then an external editor writes its stale copy back."""
        original = self.target.read_bytes()
        revision = None
        for i in range(34):
            evidence = self.write_cell(f"value {i}", cell=1 + (i % 2), revision_before=revision)
            self.assertTrue(evidence["applied"])
            self.assertTrue(evidence["ledger_logged"])
            revision = evidence["revision_after"]
        self.assertTrue(self.activity()["still_current"])

        self.target.write_bytes(original)  # Office Online's autosave: the stale copy lands

        activity = self.activity()
        self.assertFalse(activity["still_current"])
        self.assertTrue(activity["recent"])
        self.assertRefused(ErrorCode.EXTERNAL_EDITOR_ACTIVE, self.write_cell, "value 34")
        # ...and with the caller's last token, a revert is a plain revision conflict.
        self.assertRefused(
            ErrorCode.REVISION_CONFLICT,
            self.write_cell,
            "value 34",
            revision_before=revision,
            allow_concurrent_editor=True,
        )


class WriteWindowTests(_Base):
    """Interleavings inside atomic_replace_docx_parts, forced deterministically."""

    def _overrides(self, marker: str = "WRITTEN") -> dict[str, bytes]:
        with zipfile.ZipFile(self.target) as zf:
            xml = zf.read(projection.DEFAULT_PART)
        return {projection.DEFAULT_PART: xml.replace(b"R2C2", marker.encode())}

    def _leftovers(self) -> list[str]:
        return sorted(
            p.name
            for p in self.target.parent.iterdir()
            if p.name.endswith((".jsbak", ".jsclaim")) or ".tmp-" in p.name
        )

    def test_external_change_between_guard_and_claim_refuses_and_writes_nothing(self):
        real_acquire = mutations.acquire_lock

        def acquire_after_external_write(resolved):
            _external_overwrite(self.target)
            return real_acquire(resolved)

        mutations.acquire_lock = acquire_after_external_write
        try:
            envelope = self.assertRefused(ErrorCode.REVISION_CONFLICT, self.write_cell, "mine")
        finally:
            mutations.acquire_lock = real_acquire
        self.assertEqual(envelope.diagnostics["detail"], "source changed while the edit was being prepared")
        with zipfile.ZipFile(self.target) as zf:
            body = zf.read(projection.DEFAULT_PART)
        self.assertIn(b"EXTERNAL", body)
        self.assertNotIn(b"mine", body)
        self.assertEqual(self._leftovers(), [])

    def test_external_change_after_replace_is_flagged_and_not_absorbed_into_revision_after(self):
        def post_verify_with_external_write(path: Path):
            _external_overwrite(path, old="R1C1", new="EXTERNAL")

        result = mutations.atomic_replace_docx_parts(
            self.target, self._overrides(), post_verify=post_verify_with_external_write
        )
        self.assertTrue(result["external_change_during_write"])
        self.assertFalse(result["ledger_logged"])
        # revision_after is the token of the bytes we STAGED, not a re-read of
        # what an external writer left behind (the bug that hid #27's step 4).
        current = projection.compute_revision(self.target)["token"]
        self.assertNotEqual(result["revision_after"], current)
        # ...so the next call's revision_before, taken from it, conflicts.
        self.assertRefused(
            ErrorCode.REVISION_CONFLICT, self.write_cell, "next", revision_before=result["revision_after"]
        )
        # The ledger was NOT updated to claim bytes we did not produce.
        self.assertEqual(self.activity()["ledger"], "none")

    def test_verification_failure_after_external_change_does_not_roll_back(self):
        def failing_post_verify(path: Path):
            _external_overwrite(path, old="R1C1", new="EXTERNAL")
            raise ValueError("boom")

        envelope = self.assertRefused(
            ErrorCode.VERIFICATION_FAILED,
            mutations.atomic_replace_docx_parts,
            self.target,
            self._overrides(),
            post_verify=failing_post_verify,
        )
        self.assertTrue(envelope.diagnostics["rollback_skipped"])
        with zipfile.ZipFile(self.target) as zf:
            body = zf.read(projection.DEFAULT_PART)
        self.assertIn(b"EXTERNAL", body)  # the external save survived
        self.assertTrue(Path(envelope.diagnostics["jsbak_path"]).exists())

    def test_verification_failure_without_external_change_still_rolls_back(self):
        original = self.target.read_bytes()

        def failing_post_verify(path: Path):
            raise ValueError("boom")

        envelope = self.assertRefused(
            ErrorCode.VERIFICATION_FAILED,
            mutations.atomic_replace_docx_parts,
            self.target,
            self._overrides(),
            post_verify=failing_post_verify,
        )
        self.assertNotIn("rollback_skipped", envelope.diagnostics)
        self.assertEqual(self.target.read_bytes(), original)
        self.assertEqual(self._leftovers(), [])

    def test_revision_after_matches_the_installed_file_on_a_clean_write(self):
        evidence = self.write_cell("clean")
        self.assertEqual(evidence["revision_after"], projection.compute_revision(self.target)["token"])
        self.assertNotIn("external_change_during_write", evidence)


class FingerprintTests(_Base):
    def test_fingerprint_covers_parts_the_revision_token_ignores(self):
        before_fp = projection.compute_package_fingerprint(self.target)
        before_rev = projection.compute_revision(self.target)["token"]
        tmp = self.target.with_name("x.docx")
        with zipfile.ZipFile(self.target) as src, zipfile.ZipFile(tmp, "w") as dst:
            for item in src.infolist():
                data = src.read(item)
                if item.filename == "word/styles.xml":
                    data = data + b"<!-- changed -->"
                dst.writestr(item, data)
        self.assertEqual(projection.compute_revision(tmp)["token"], before_rev)
        self.assertNotEqual(projection.compute_package_fingerprint(tmp), before_fp)

    def test_fingerprint_ignores_compression_and_entry_order(self):
        base = projection.compute_package_fingerprint(self.target)
        tmp = self.target.with_name("y.docx")
        with zipfile.ZipFile(self.target) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_STORED) as dst:
            for item in reversed(src.infolist()):
                dst.writestr(item.filename, src.read(item))
        self.assertEqual(projection.compute_package_fingerprint(tmp), base)

    def test_fingerprint_accepts_bytes(self):
        self.assertEqual(
            projection.compute_package_fingerprint(self.target.read_bytes()),
            projection.compute_package_fingerprint(self.target),
        )


if __name__ == "__main__":
    unittest.main()
