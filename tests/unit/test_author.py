"""Unit tests for src/verified_docx_mcp/author.py (issue #28 WP-07b-a):
author_name resolution from ~/.jennystack/config.json, falling back to
the macOS full name, falling back to a final literal default.

"author_name is currently absent from ~/.jennystack/config.json (it holds
only docs_root)" per the WP's own notes -- the fallback path is the one
this machine actually exercises; both paths are tested here directly via
VERIFIED_DOCX_MCP_JENNYSTACK_CONFIG (an isolated temp file, never the
real ~/.jennystack/config.json) and a patched _macos_full_name, so
neither test depends on this machine's actual config or account name.
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import author


class _IsolatedConfigCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_env = os.environ.get(author._CONFIG_PATH_ENV)
        self.config_path = Path(self._tmp.name) / "config.json"
        os.environ[author._CONFIG_PATH_ENV] = str(self.config_path)

    def tearDown(self):
        if self._old_env is None:
            os.environ.pop(author._CONFIG_PATH_ENV, None)
        else:
            os.environ[author._CONFIG_PATH_ENV] = self._old_env
        self._tmp.cleanup()


class ConfigAuthorNameTests(_IsolatedConfigCase):
    def test_reads_author_name_from_config(self):
        self.config_path.write_text(json.dumps({"docs_root": "/x", "author_name": "Jane Reviewer"}))
        self.assertEqual(author._read_config_author_name(), "Jane Reviewer")

    def test_resolve_prefers_config_over_macos_fallback(self):
        self.config_path.write_text(json.dumps({"author_name": "Jane Reviewer"}))
        with mock.patch("verified_docx_mcp.author._macos_full_name", return_value="Should Not Be Used"):
            self.assertEqual(author.resolve_author_name(), "Jane Reviewer")

    def test_missing_author_name_key_returns_none(self):
        # The documented current-machine state: only docs_root is set.
        self.config_path.write_text(json.dumps({"docs_root": "/x"}))
        self.assertIsNone(author._read_config_author_name())

    def test_empty_or_blank_author_name_returns_none(self):
        self.config_path.write_text(json.dumps({"author_name": "   "}))
        self.assertIsNone(author._read_config_author_name())

    def test_missing_config_file_returns_none(self):
        self.assertFalse(self.config_path.exists())
        self.assertIsNone(author._read_config_author_name())

    def test_malformed_json_returns_none_not_raise(self):
        self.config_path.write_text("{not valid json")
        self.assertIsNone(author._read_config_author_name())

    def test_non_dict_json_returns_none(self):
        self.config_path.write_text(json.dumps(["a", "list", "not", "a", "dict"]))
        self.assertIsNone(author._read_config_author_name())


class MacosFallbackTests(_IsolatedConfigCase):
    """The fallback path this WP's own notes call out as the live one on a
    machine with no author_name configured yet."""

    def test_resolve_falls_back_to_macos_full_name_when_config_absent(self):
        self.assertFalse(self.config_path.exists())
        with mock.patch("verified_docx_mcp.author._macos_full_name", return_value="Michael Sutton"):
            self.assertEqual(author.resolve_author_name(), "Michael Sutton")

    def test_falls_back_to_unknown_author_when_both_sources_unavailable(self):
        self.config_path.write_text(json.dumps({"docs_root": "/x"}))
        with mock.patch("verified_docx_mcp.author._macos_full_name", return_value=None):
            self.assertEqual(author.resolve_author_name(), author._FALLBACK_AUTHOR)

    def test_resolve_never_raises_and_never_returns_empty(self):
        self.config_path.write_text(json.dumps({"docs_root": "/x"}))
        with mock.patch("verified_docx_mcp.author._macos_full_name", return_value=None):
            name = author.resolve_author_name()
        self.assertIsInstance(name, str)
        self.assertTrue(name.strip())


class MacosFullNameGecosTests(unittest.TestCase):
    """_macos_full_name's own GECOS-parsing logic, independent of the
    config file entirely."""

    def test_takes_first_comma_separated_gecos_field(self):
        fake_pw = mock.Mock(pw_gecos="Michael Sutton,Room 4,555-1234,555-5678")
        with mock.patch("verified_docx_mcp.author.pwd.getpwuid", return_value=fake_pw):
            self.assertEqual(author._macos_full_name(), "Michael Sutton")

    def test_falls_back_to_id_dash_f_when_gecos_blank(self):
        fake_pw = mock.Mock(pw_gecos="")
        fake_proc = mock.Mock(stdout="Michael Sutton\n")
        with (
            mock.patch("verified_docx_mcp.author.pwd.getpwuid", return_value=fake_pw),
            mock.patch("verified_docx_mcp.author.subprocess.run", return_value=fake_proc),
        ):
            self.assertEqual(author._macos_full_name(), "Michael Sutton")

    def test_returns_none_when_nothing_available(self):
        fake_pw = mock.Mock(pw_gecos="")
        with (
            mock.patch("verified_docx_mcp.author.pwd.getpwuid", return_value=fake_pw),
            mock.patch("verified_docx_mcp.author.subprocess.run", side_effect=OSError("no id command")),
        ):
            self.assertIsNone(author._macos_full_name())


class LiveMachineSmokeTest(unittest.TestCase):
    """One unmocked, real-machine check -- not the contract (the mocked
    tests above are), just confirms the actual fallback this WP's notes
    describe ("which on this machine resolves to Michael Sutton") still
    holds, without hardcoding it as a requirement other machines must
    also satisfy."""

    def test_resolve_author_name_returns_a_non_empty_string_on_this_machine(self):
        name = author.resolve_author_name()
        self.assertIsInstance(name, str)
        self.assertTrue(name.strip())


if __name__ == "__main__":
    unittest.main()
