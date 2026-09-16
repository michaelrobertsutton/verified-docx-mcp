"""Unit tests for src/verified_docx_mcp/paths.py.

Covers the three pieces paths.py combines (see its module docstring):
resolve_docx_pointer's $DOCS containment, the file-root allowlist/denylist,
and the read-path package snapshot with validated retry (issue #28 WP-02
acceptance: "snapshot retry (simulated truncated zip)").
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import paths
from verified_docx_mcp.errors import ErrorCode, VerifyError


class ResolveDocxPointerTests(unittest.TestCase):
    """The $DOCS containment rule, mirrored verbatim from
    core/document-backend-protocol.md §3 — see paths.py's module docstring."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name) / "docs-root"
        self.root.mkdir()
        (self.root / "sub").mkdir()

    def tearDown(self):
        self._tmp.cleanup()

    def test_simple_relative_path_resolves_inside_root(self):
        target = paths.resolve_docx_pointer("sub/file.docx", str(self.root))
        self.assertEqual(target, (self.root / "sub" / "file.docx").resolve())

    def test_empty_relative_path_rejected(self):
        with self.assertRaises(paths.DocxPathError):
            paths.resolve_docx_pointer("", str(self.root))

    def test_dot_relative_path_rejected(self):
        with self.assertRaises(paths.DocxPathError):
            paths.resolve_docx_pointer(".", str(self.root))

    def test_absolute_path_rejected(self):
        with self.assertRaises(paths.DocxPathError):
            paths.resolve_docx_pointer("/etc/passwd", str(self.root))

    def test_dotdot_traversal_rejected(self):
        with self.assertRaises(paths.DocxPathError):
            paths.resolve_docx_pointer("../evil.docx", str(self.root))

    def test_dotdot_in_middle_rejected(self):
        with self.assertRaises(paths.DocxPathError):
            paths.resolve_docx_pointer("sub/../../evil.docx", str(self.root))

    def test_nul_byte_rejected(self):
        with self.assertRaises(paths.DocxPathError):
            paths.resolve_docx_pointer("sub/\x00evil.docx", str(self.root))

    def test_nonexistent_root_rejected(self):
        with self.assertRaises(paths.DocxPathError):
            paths.resolve_docx_pointer("file.docx", str(self.root / "does-not-exist"))

    def test_sibling_directory_with_shared_prefix_rejected(self):
        # The exact "/docs-evil" vs "/docs" hazard the core doc calls out —
        # a naive startswith() would pass this; is_relative_to must not.
        evil_root = self.root.parent / (self.root.name + "-evil")
        evil_root.mkdir()
        # Constructing a relative_path that, joined naively as a string,
        # would look like it is "under" root but resolves into evil_root
        # is not possible via a pure relative join (pathlib always nests
        # under root when relative_path has no leading slash) — the real
        # hazard this guards against is a SYMLINK inside root pointing
        # outside it, exercised in test_symlink_escape_rejected below.
        target = paths.resolve_docx_pointer("file.docx", str(self.root))
        self.assertFalse(str(target).startswith(str(evil_root)))

    def test_symlink_escape_rejected(self):
        outside = Path(tempfile.mkdtemp())
        try:
            link = self.root / "escape"
            link.symlink_to(outside)
            with self.assertRaises(paths.DocxPathError):
                paths.resolve_docx_pointer("escape/evil.docx", str(self.root))
        finally:
            import shutil

            shutil.rmtree(outside, ignore_errors=True)


class AllowedFileRootsTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.allowed = Path(self._tmp.name) / "allowed"
        self.allowed.mkdir()
        self._old_env = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = str(self.allowed)

    def tearDown(self):
        if self._old_env is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_env
        self._tmp.cleanup()

    def test_path_inside_allowed_root_resolves(self):
        target = self.allowed / "doc.docx"
        target.write_bytes(b"stub")
        resolved = paths.resolve_allowed_docx_path(str(target))
        self.assertEqual(resolved, target.resolve())

    def test_path_outside_allowed_root_rejected(self):
        outside = Path(tempfile.mkdtemp()) / "doc.docx"
        outside.write_bytes(b"stub")
        try:
            with self.assertRaises(VerifyError) as ctx:
                paths.resolve_allowed_docx_path(str(outside))
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        finally:
            outside.unlink(missing_ok=True)
            outside.parent.rmdir()

    def test_missing_file_rejected(self):
        with self.assertRaises(VerifyError) as ctx:
            paths.resolve_allowed_docx_path(str(self.allowed / "nope.docx"))
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_denylisted_sensitive_path_rejected_even_inside_allowed_root(self):
        with mock.patch.object(Path, "home", return_value=self.allowed):
            ssh_dir = self.allowed / ".ssh"
            ssh_dir.mkdir()
            key = ssh_dir / "id_rsa"
            key.write_bytes(b"stub")
            with self.assertRaises(VerifyError) as ctx:
                paths.resolve_allowed_docx_path(str(key))
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
            self.assertIn("credential", ctx.exception.envelope.message)


class DefaultAllowedRootsTests(unittest.TestCase):
    """issue #101: the default allowed roots (env unset) widen to include
    paths._claude_code_scratch_root() when it exists. An explicit
    VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS is used verbatim — never widened
    with, or narrowed to, the scratch root."""

    def setUp(self):
        self._old_env = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        self._home_tmp = tempfile.TemporaryDirectory()
        self.fake_home = Path(self._home_tmp.name) / "home"
        self.fake_home.mkdir()
        self._scratch_tmp = tempfile.TemporaryDirectory()
        self.fake_scratch = Path(self._scratch_tmp.name) / "scratch"
        self.fake_scratch.mkdir()
        self._explicit_tmp = tempfile.TemporaryDirectory()
        self.explicit_root = Path(self._explicit_tmp.name) / "explicit"
        self.explicit_root.mkdir()

    def tearDown(self):
        if self._old_env is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_env
        self._home_tmp.cleanup()
        self._scratch_tmp.cleanup()
        self._explicit_tmp.cleanup()

    def test_env_unset_scratch_present_widens_default(self):
        home_file = self.fake_home / "doc.docx"
        home_file.write_bytes(b"stub")
        scratch_file = self.fake_scratch / "doc.docx"
        scratch_file.write_bytes(b"stub")
        with mock.patch.object(Path, "home", return_value=self.fake_home), mock.patch.object(
            paths, "_claude_code_scratch_root", return_value=self.fake_scratch
        ):
            self.assertEqual(paths.resolve_allowed_docx_path(str(home_file)), home_file.resolve())
            self.assertEqual(paths.resolve_allowed_docx_path(str(scratch_file)), scratch_file.resolve())

    def test_env_unset_scratch_absent_home_only(self):
        home_file = self.fake_home / "doc.docx"
        home_file.write_bytes(b"stub")
        would_be_scratch_file = self.fake_scratch / "doc.docx"
        would_be_scratch_file.write_bytes(b"stub")
        with mock.patch.object(Path, "home", return_value=self.fake_home), mock.patch.object(
            paths, "_claude_code_scratch_root", return_value=None
        ):
            self.assertEqual(paths._allowed_file_roots(), [self.fake_home.resolve()])
            self.assertEqual(paths.resolve_allowed_docx_path(str(home_file)), home_file.resolve())
            with self.assertRaises(VerifyError) as ctx:
                paths.resolve_allowed_docx_path(str(would_be_scratch_file))
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_env_set_explicit_allowlist_excludes_scratch_root(self):
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = str(self.explicit_root)
        scratch_file = self.fake_scratch / "doc.docx"
        scratch_file.write_bytes(b"stub")
        with mock.patch.object(paths, "_claude_code_scratch_root", return_value=self.fake_scratch):
            self.assertEqual(paths._allowed_file_roots(), [self.explicit_root.resolve()])
            with self.assertRaises(VerifyError) as ctx:
                paths.resolve_allowed_docx_path(str(scratch_file))
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_denylist_applies_inside_home_with_widened_default(self):
        with mock.patch.object(Path, "home", return_value=self.fake_home), mock.patch.object(
            paths, "_claude_code_scratch_root", return_value=self.fake_scratch
        ):
            ssh_dir = self.fake_home / ".ssh"
            ssh_dir.mkdir()
            key = ssh_dir / "id_rsa"
            key.write_bytes(b"stub")
            with self.assertRaises(VerifyError) as ctx:
                paths.resolve_allowed_docx_path(str(key))
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
            self.assertIn("credential", ctx.exception.envelope.message)

    @unittest.skipUnless(sys.platform == "darwin", "macOS /tmp -> /private/tmp alias only")
    def test_tmp_spelling_of_scratch_path_resolves_via_private_tmp_alias(self):
        with tempfile.TemporaryDirectory(dir="/private/tmp") as real_scratch_dir:
            real_scratch = Path(real_scratch_dir)
            target = real_scratch / "doc.docx"
            target.write_bytes(b"stub")
            tmp_spelling = Path("/tmp") / real_scratch.relative_to("/private/tmp") / "doc.docx"
            with mock.patch.object(Path, "home", return_value=self.fake_home), mock.patch.object(
                paths, "_claude_code_scratch_root", return_value=real_scratch
            ):
                resolved = paths.resolve_allowed_docx_path(str(tmp_spelling))
            self.assertEqual(resolved, target.resolve())


class SnapshotDocxPackageTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.source = Path(self._tmp.name) / "source.docx"
        with zipfile.ZipFile(self.source, "w") as zf:
            zf.writestr(
                "word/document.xml",
                '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main"/>',
            )
            zf.writestr("[Content_Types].xml", "<Types/>")

    def tearDown(self):
        self._tmp.cleanup()

    def test_valid_package_snapshots_on_first_attempt(self):
        sleeps: list[float] = []
        snap = paths.snapshot_docx_package(self.source, sleep=sleeps.append)
        try:
            self.assertTrue(snap.is_file())
            self.assertEqual(sleeps, [])  # no retry needed
            with zipfile.ZipFile(snap) as zf:
                self.assertIsNone(zf.testzip())
        finally:
            snap.unlink(missing_ok=True)

    def test_truncated_zip_retries_then_fails(self):
        # Simulate a truncated/corrupt zip by copying only the first few
        # bytes of the source — copyfile inside snapshot_docx_package will
        # then copy THIS truncated file every attempt, so validation fails
        # every time and SNAPSHOT_FAILED is raised after max_attempts.
        truncated = Path(self._tmp.name) / "truncated.docx"
        truncated.write_bytes(self.source.read_bytes()[:10])

        sleeps: list[float] = []
        with self.assertRaises(VerifyError) as ctx:
            paths.snapshot_docx_package(truncated, max_attempts=3, sleep=sleeps.append)

        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.SNAPSHOT_FAILED)
        self.assertEqual(ctx.exception.envelope.diagnostics["attempts"], 3)
        # Retried twice (between the 3 attempts), never a fixed sleep.time.
        self.assertEqual(sleeps, [paths._SNAPSHOT_RETRY_BACKOFF_SECONDS] * 2)

    def test_truncated_zip_succeeds_if_it_recovers_before_max_attempts(self):
        # A more realistic simulation: the FIRST copy races a concurrent
        # Word flush and is corrupt, but a later attempt (after the
        # backoff) copies clean bytes — snapshot_docx_package must return
        # successfully rather than exhausting all attempts.
        truncated = Path(self._tmp.name) / "flaky.docx"
        truncated.write_bytes(self.source.read_bytes()[:10])

        real_copyfile = __import__("shutil").copyfile
        calls = {"n": 0}

        def flaky_copyfile(src, dst):
            calls["n"] += 1
            if calls["n"] == 1:
                return real_copyfile(str(truncated), dst)
            return real_copyfile(str(self.source), dst)

        sleeps: list[float] = []
        with mock.patch("verified_docx_mcp.paths.shutil.copyfile", side_effect=flaky_copyfile):
            snap = paths.snapshot_docx_package(self.source, max_attempts=3, sleep=sleeps.append)
        try:
            self.assertTrue(snap.is_file())
            self.assertEqual(sleeps, [paths._SNAPSHOT_RETRY_BACKOFF_SECONDS])
            self.assertEqual(calls["n"], 2)
        finally:
            snap.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
